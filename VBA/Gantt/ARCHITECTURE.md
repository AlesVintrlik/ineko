# Architektura - Gantt

## Přehled architektury

Gantt aplikace je postavena na třívrstvé architektuře:

1. **Prezentační vrstva** - Excel listy a UserForms (UI)
2. **Business logika** - VBA moduly s výpočty a datovou logikou
3. **Datová vrstva** - SQL Server databáze přes ADODB

```
┌─────────────────────────────────────────────────────────┐
│                  Prezentační vrstva                     │
│  ┌──────────┐  ┌──────────┐  ┌─────────────────────┐   │
│  │frmLogin  │  │frmProgress│ │ Excel Listy         │   │
│  │          │  │           │  │ - Gantt             │   │
│  │          │  │           │  │ - Zakazky           │   │
│  └──────────┘  └──────────┘  │ - Konfigurace       │   │
│                               │ - Svátky            │   │
│                               └─────────────────────┘   │
└─────────────────────────────────────────────────────────┘
                        ↕
┌─────────────────────────────────────────────────────────┐
│                   Business logika                       │
│  ┌──────────────┐  ┌──────────┐  ┌────────────────┐    │
│  │Connection.bas│  │Data.bas  │  │Advanced.bas    │    │
│  │- Autentizace │  │- Načítání│  │- Výpočty       │    │
│  │- Šifrování   │  │- Aktuali-│  │- Kapacity      │    │
│  │- Připojení   │  │  zace    │  │- Formátování   │    │
│  └──────────────┘  └──────────┘  └────────────────┘    │
└─────────────────────────────────────────────────────────┘
                        ↕ ADODB
┌─────────────────────────────────────────────────────────┐
│                    Datová vrstva                        │
│              SQL Server - IN-EKO ERP DB                 │
│  ┌──────────────────────────────────────────────────┐   │
│  │ hvw_TerminyZakazekProPlanovani (VIEW)            │   │
│  │ - ID zakázky                                     │   │
│  │ - Termíny fází (Příprava, Svařování, Montáž...)  │   │
│  │ - Datum zahájení a ukončení                      │   │
│  └──────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

## Komponenty systému

### 1. Connection.bas - Správa připojení

**Zodpovědnost:**
- Vytváření a správa databázového připojení
- Autentizace uživatelů
- Šifrování a dešifrování přihlašovacích údajů

**Klíčové funkce:**
- `CreateConnection()` - Vytvoří ADODB.Connection objekt
- `VerifyCredentials()` - Ověří platnost přihlašovacích údajů
- `XOREncryptDecrypt()` - Šifrování/dešifrování pomocí XOR
- `GetDecryptedCredentials()` - Vrací dešifrované údaje z paměti

**Design pattern:**
- **Singleton pattern** - Přihlašovací údaje jsou uloženy v globální kolekci `loginCredentials`
- **Facade pattern** - `CreateConnection()` zapouzdřuje složitost vytváření připojení

**Datový tok:**
```
frmLogin → VerifyCredentials() → XOR šifrování → loginCredentials (paměť)
                                                          ↓
                    CreateConnection() ← GetDecryptedCredentials()
                           ↓
                    ADODB.Connection
```

### 2. Data.bas - Datová vrstva

**Zodpovědnost:**
- Načítání dat ze SQL Server
- Aktualizace Gantt diagramu
- Správa životního cyklu ADO objektů

**Klíčové procedury:**
- `LoadOrUpdateData()` - Načte data do listu "Zakazky"
- `AktualizovatSeznamZakazek()` - Aktualizuje Gantt diagram
- `RefreshAllConnections()` - Obnoví všechna datová připojení

**Datový tok:**
```
SQL Server                    Excel
    │                           │
    │ SELECT * FROM view        │
    ├───────────────────────────>│ List "Zakazky"
    │                           │ (raw data)
    │                           │
    │                           ↓
    │                    Dictionary (unikátní zakázky)
    │                           │
    │                           ↓
    │                    List "Gantt"
    │                    (1 řádek = 1 zakázka)
```

**Optimalizace:**
- Použití `Dictionary` objektu pro získání unikátních zakázek (O(n) vs O(n²))
- `CopyFromRecordset` pro rychlé načtení dat (rychlejší než řádek po řádku)
- Screen updating vypnuto během operací

### 3. Advanced.bas - Pokročilá logika

**Zodpovědnost:**
- Výpočet obsazenosti výrobních středisek
- Podmíněné formátování
- Správa svátků a víkendů

**Klíčové procedury:**
- `SumarizaceBoduProVsechnySloupce()` - Hlavní výpočetní procedura
- `AktualizaceKontrolyKapacit()` - Wrapper pro přepočet bez progress baru
- `VymazaniSouctuPodGrafem()` - Čištění předchozích výsledků
- `ZrusitVsechnyFiltry()` - Reset filtrů

**Algoritmus výpočtu kapacit:**

```
Pro každý den v časové ose:
  SKIP pokud je víkend (So, Ne)
  SKIP pokud je svátek (dle listu Svátky)

  Pro každé výrobní středisko (Příprava, Svařování, Montáž, Elektro):
    body = 0

    Pro každou zakázku:
      IF den >= start_fáze AND den <= konec_fáze THEN
        body += počet_pracovníků
      END IF

    Ulož body do příslušné buňky
    Aplikuj barevné formátování:
      - body = 0           → Bílá
      - body < maximum     → Zelená
      - body = maximum     → Oranžová
      - body > maximum     → Červená
```

**Výkonnost:**
- Časová složitost: O(dny × střediska × zakázky)
- Typicky: 365 × 4 × 50 = ~73,000 iterací
- Optimalizace: Podmíněné formátování vypnuto během výpočtu
- Progress bar pro vizuální feedback

**Konstanty kapacit:**
```vba
LidiPriprava = 1
LidiSvarovani = 2
LidiMontaz = 2
LidiElektro = 1
```

### 4. UserForms - Uživatelské rozhraní

#### frmLogin - Přihlašovací formulář

**Lifecycle:**
```
UserForm_Initialize()
    ↓
Načtení hodnot z listu "Konfigurace"
    ↓
Zobrazení formuláře
    ↓
btnLogin_Click()
    ↓
VerifyCredentials()
    ↓
Uložení šifrovaných údajů
    ↓
Unload formuláře
    ↓
Zobrazení aplikace
```

**Zabezpečení:**
- Hesla nejsou nikdy ukládána v plaintext
- XOR šifrování s 16-znakový klíčem
- Údaje existují pouze v paměti

#### frmProgress - Progress bar

**Účel:**
- Vizuální feedback během dlouhých operací
- Zobrazení % dokončení

**Použití:**
```vba
Load frmProgress
frmProgress.Show vbModeless
frmProgress.UpdateProgressBar(0.5) ' 50%
Unload frmProgress
```

### 5. ThisWorkbook - Události sešitu

**Události:**

**Workbook_Open:**
1. Skryje Excel aplikaci (`Application.Visible = False`)
2. Vypne mřížku, záhlaví, panel vzorců (minimalistický režim)
3. Zkontroluje exkluzivní přístup k souboru
4. Pokusí se získat Read/Write přístup
5. Pokud je soubor otevřen jiným uživatelem → varování a ukončení
6. Zobrazí `frmLogin`
7. Nastaví výchozí hodnotu pro časovou osu (P1 = 28 dní)

**Workbook_BeforeClose:**
1. Obnoví standardní zobrazení (mřížka, záhlaví, vzorce)

**Ochrana proti více uživatelům:**
```vba
Function GetFileInUseUserName() → Zjistí, kdo má soubor otevřený
  ↓
Pokud jiný uživatel → MsgBox a ukončení
Pokud stejný uživatel → Pokračovat
```

## Datový model

### Excel listy

**List "Gantt":**
- Řádek 1-3: Záhlaví a metadata
- Řádek 4+: Zakázky (1 řádek = 1 zakázka)
- Sloupec A: Informace o zakázce
- Sloupec B: Číslo zakázky
- Sloupce C-N: Metadata (termíny fází)
- Sloupce O+: Časová osa (každý sloupec = 1 den)
- Poslední řádky: Sumarizace obsazenosti středisek

**List "Zakazky":**
- Řádek 1: Hlavička (názvy sloupců z SQL)
- Řádek 2+: Data z `hvw_TerminyZakazekProPlanovani`

**List "Konfigurace":**
- Pojmenované oblasti:
  - `serverName`: SQL server
  - `databaseName`: Databáze
  - `login`: Uživatelské jméno

**List "Svátky":**
- Sloupec B: Datum svátku (Date)

### Globální proměnné

```vba
Public loginCredentials As Collection  ' Šifrované přihlašovací údaje
Public isLoggedIn As Boolean           ' Stav přihlášení
Public Const ENCRYPTION_KEY As String  ' Klíč pro XOR šifrování
Private x As Integer                   ' Flag pro potlačení progress baru
```

## Bezpečnostní architektura

### Autentizace

**Podporované metody:**
1. **Windows NT Authentication (doporučeno)**
   - `userName = ""` AND `password = ""`
   - Connection string: `Integrated Security=SSPI`
2. **SQL Authentication**
   - `userName` a `password` vyplněny
   - Connection string: `User ID=...; Password=...`

### Šifrování

**XOR Cipher:**
```vba
XOREncryptDecrypt(text, key):
  Pro každý znak v textu:
    výsledek += znak XOR key[(i MOD délka_klíče)]
```

**Vlastnosti:**
- Symetrické šifrování (stejná funkce pro šifrování i dešifrování)
- Klíč délky 16 znaků
- ⚠️ XOR není kryptograficky bezpečné pro citlivá data
- Vhodné pro základní ochranu v interním prostředí

### Správa připojení

**Connection lifetime:**
```
CreateConnection()
    ↓
Open connection
    ↓
Provést operaci (SELECT, atd.)
    ↓
Close connection
    ↓
Set conn = Nothing
```

**Error handling:**
- `On Error GoTo ErrorHandler` ve všech kritických funkcích
- MsgBox pro informování uživatele o chybách
- Zajištění uzavření připojení i při chybě

## Výkonnostní optimalizace

### 1. Screen updating
```vba
Application.ScreenUpdating = False
' ... operace ...
Application.ScreenUpdating = True
```

### 2. Calculation mode
```vba
Application.Calculation = xlCalculationManual
' ... změny dat ...
Application.Calculation = xlCalculationAutomatic
```

### 3. Conditional formatting
```vba
ws.EnableFormatConditionsCalculation = False
' ... nastavení pravidel ...
ws.EnableFormatConditionsCalculation = True
```

### 4. Datové struktury
- Dictionary pro O(1) vyhledávání
- Pole místo Range operací v cyklech
- CopyFromRecordset místo Range.Value v cyklu

## Rozšiřitelnost

### Přidání nového výrobního střediska

1. **Advanced.bas** - Přidat proměnné:
```vba
Dim totalPointsNoveStredisko As Long
Dim LidiNoveStredisko As Integer
```

2. **Advanced.bas** - Přidat do výpočetní smyčky:
```vba
If compareValue >= ws.Cells(row, "M").Value And _
   compareValue <= ws.Cells(row, "N").Value Then
    totalPointsNoveStredisko = totalPointsNoveStredisko + LidiNoveStredisko
End If
```

3. **Advanced.bas** - Přidat zápis výsledků:
```vba
ws.Cells(nextRow + 4, col).Value = totalPointsNoveStredisko
```

4. **Advanced.bas** - Přidat podmíněné formátování

### Změna šifrovacího algoritmu

Nahradit `XOREncryptDecrypt()` funkcí s jiným algoritmem (např. AES):

```vba
Function AESEncrypt(text As String, key As String) As String
    ' Implementace AES šifrování
End Function

Function AESDecrypt(text As String, key As String) As String
    ' Implementace AES dešifrování
End Function
```

## Závislosti

### Knihovny
- **Microsoft ActiveX Data Objects 2.x** (ADODB)
- **Microsoft Scripting Runtime** (Dictionary)

### Databáze
- **SQL Server** - jakákoli podporovaná verze
- **View**: `hvw_TerminyZakazekProPlanovani`

### Excel
- **Minimální verze**: Excel 2010
- **Podporované formáty**: .xlsm (Excel Macro-Enabled Workbook)

## Diagram toku dat (DFD)

```
┌─────────┐         ┌──────────────┐         ┌──────────┐
│Uživatel │────────>│ frmLogin     │────────>│Connection│
└─────────┘ údaje   └──────────────┘  ověření└──────────┘
                                                    │
                                                    ↓
                    ┌──────────────────────────────────┐
                    │       SQL Server                 │
                    │  hvw_TerminyZakazekProPlanovani  │
                    └──────────────────────────────────┘
                                    │
                                    ↓ SELECT
                    ┌──────────────────────────────────┐
                    │       Data.bas                   │
                    │  LoadOrUpdateData()              │
                    └──────────────────────────────────┘
                                    │
                    ┌───────────────┴───────────────┐
                    ↓                               ↓
            ┌──────────────┐              ┌──────────────┐
            │List "Zakazky"│              │List "Gantt"  │
            │(raw data)    │              │(agregace)    │
            └──────────────┘              └──────────────┘
                                                    │
                                                    ↓
                    ┌──────────────────────────────────┐
                    │       Advanced.bas               │
                    │  SumarizaceBoduProVsechnySloupce │
                    └──────────────────────────────────┘
                                    │
                                    ↓
                    ┌──────────────────────────────────┐
                    │  Kapacity + barevné formátování  │
                    └──────────────────────────────────┘
```

## Best practices

### 1. Error handling
Vždy používejte error handling v public proceduře:
```vba
Sub MojeProcedura()
    On Error GoTo ErrorHandler
    ' ... kód ...
    Exit Sub
ErrorHandler:
    MsgBox "Chyba: " & Err.Description
End Sub
```

### 2. Resource management
Vždy uzavírejte databázová připojení:
```vba
Set rs = Nothing
adoConn.Close
Set adoConn = Nothing
```

### 3. User feedback
Pro dlouhé operace používejte progress bar:
```vba
Load frmProgress
frmProgress.Show vbModeless
' ... operace ...
Unload frmProgress
```

### 4. Screen performance
Vypínejte aktualizaci pro dávkové operace:
```vba
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
' ... změny ...
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
```
