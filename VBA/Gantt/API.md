# API Reference - Gantt

Tato dokumentace popisuje všechny veřejné (Public) funkce a procedury dostupné v Gantt aplikaci.

## Obsah

- [Connection.bas](#connectionbas)
- [Data.bas](#databas)
- [Advanced.bas](#advancedbas)
- [frmLogin](#frmlogin)
- [frmProgress](#frmprogress)

---

## Connection.bas

Modul pro správu databázového připojení a autentizace.

### Globální proměnné

#### `loginCredentials`
```vba
Public loginCredentials As Collection
```
**Popis:** Kolekce obsahující šifrované přihlašovací údaje uživatele.

**Struktura:**
- `loginCredentials("login")` - Šifrované uživatelské jméno
- `loginCredentials("passw")` - Šifrované heslo

**Životnost:** Po dobu běhu aplikace (až do zavření sešitu)

---

#### `ENCRYPTION_KEY`
```vba
Public Const ENCRYPTION_KEY As String = "a7$D9!pR#3fG5&Z1"
```
**Popis:** Konstantní klíč pro XOR šifrování a dešifrování hesel.

**⚠️ Varování:** Změna této hodnoty zneplatní všechny uložené šifrované údaje.

---

#### `isLoggedIn`
```vba
Public isLoggedIn As Boolean
```
**Popis:** Indikátor, zda je uživatel úspěšně přihlášen.

**Hodnoty:**
- `True` - Uživatel je přihlášen
- `False` - Uživatel není přihlášen

---

### Funkce

#### `XOREncryptDecrypt`
```vba
Function XOREncryptDecrypt(text As String, key As String) As String
```
**Popis:** Šifruje nebo dešifruje text pomocí XOR operace.

**Parametry:**
- `text` (String) - Text k zašifrování nebo dešifrování
- `key` (String) - Šifrovací klíč

**Návratová hodnota:**
- (String) Zašifrovaný nebo dešifrovaný text

**Poznámka:** Funkce je symetrická - použije se stejně pro šifrování i dešifrování.

**Příklad:**
```vba
Dim encrypted As String
Dim decrypted As String
Dim key As String

key = ENCRYPTION_KEY
encrypted = XOREncryptDecrypt("heslo123", key)  ' Šifrování
decrypted = XOREncryptDecrypt(encrypted, key)   ' Dešifrování
' decrypted = "heslo123"
```

**Algoritmus:**
```
Pro každý znak v textu:
  výsledek[i] = text[i] XOR key[i MOD length(key)]
```

---

#### `VerifyCredentials`
```vba
Function VerifyCredentials(userName As String, password As String, _
                          ServerName As String, databaseName As String) As Boolean
```
**Popis:** Ověří platnost přihlašovacích údajů pokusem o připojení k databázi.

**Parametry:**
- `userName` (String) - Uživatelské jméno (prázdné pro NT autentizaci)
- `password` (String) - Heslo (prázdné pro NT autentizaci)
- `ServerName` (String) - Název SQL serveru
- `databaseName` (String) - Název databáze

**Návratová hodnota:**
- `True` - Připojení úspěšné, údaje jsou platné
- `False` - Připojení selhalo, neplatné údaje

**Chybové stavy:**
- Zobrazí MsgBox s chybovou zprávou při selhání připojení

**Příklad:**
```vba
Dim isValid As Boolean
isValid = VerifyCredentials("admin", "heslo", "SERVER01", "ERP_DB")

If isValid Then
    MsgBox "Připojení úspěšné"
Else
    MsgBox "Neplatné údaje"
End If
```

**NT Autentizace:**
```vba
' Pro Windows autentizaci použijte prázdné řetězce
isValid = VerifyCredentials("", "", "SERVER01", "ERP_DB")
```

---

#### `GetDecryptedCredentials`
```vba
Function GetDecryptedCredentials() As Collection
```
**Popis:** Vrací kolekci s dešifrovanými přihlašovacími údaji.

**Parametry:** Žádné

**Návratová hodnota:**
- (Collection) Kolekce obsahující:
  - `"login"` - Zašifrované uživatelské jméno
  - `"passw"` - Zašifrované heslo
- `Nothing` - Pokud údaje nejsou dostupné

**Chybové stavy:**
- Zobrazí MsgBox a otevře `frmLogin` pokud údaje chybí

**Příklad:**
```vba
Dim creds As Collection
Set creds = GetDecryptedCredentials()

If Not creds Is Nothing Then
    Dim login As String
    login = XOREncryptDecrypt(creds("login"), ENCRYPTION_KEY)
End If
```

**⚠️ Poznámka:** Funkce vrací **šifrované** údaje. Pro získání čitelného textu je nutné je dešifrovat pomocí `XOREncryptDecrypt()`.

---

#### `CreateConnection`
```vba
Function CreateConnection() As Object
```
**Popis:** Vytvoří a vrátí aktivní připojení k SQL Server databázi.

**Parametry:** Žádné (používá globální `loginCredentials`)

**Návratová hodnota:**
- (ADODB.Connection) Otevřené databázové připojení
- `Nothing` - Pokud připojení selhalo

**Prerekvizity:**
- Uživatel musí být přihlášen (`isLoggedIn = True`)
- `loginCredentials` musí být naplněny
- List "Konfigurace" musí obsahovat pojmenované oblasti:
  - `serverName`
  - `databaseName`

**Chybové stavy:**
- Zobrazí MsgBox pokud se nepodařilo načíst credentials
- Zobrazí MsgBox pokud se nepodařilo otevřít připojení

**Příklad:**
```vba
Dim conn As Object
Set conn = CreateConnection()

If Not conn Is Nothing Then
    ' Připojení úspěšné
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM MojeTabuka", conn
    ' ... práce s daty ...
    rs.Close
    conn.Close
Else
    MsgBox "Nelze se připojit k databázi"
End If

Set rs = Nothing
Set conn = Nothing
```

**Connection String formáty:**

NT Autentizace:
```
Provider=SQLOLEDB;Data Source=SERVER;Initial Catalog=DB;Integrated Security=SSPI;
```

SQL Autentizace:
```
Provider=SQLOLEDB;Data Source=SERVER;Initial Catalog=DB;User ID=user;Password=pass;
```

---

## Data.bas

Modul pro načítání a aktualizaci dat z databáze.

### `LoadOrUpdateData`
```vba
Sub LoadOrUpdateData()
```
**Popis:** Načte data z databázového view `hvw_TerminyZakazekProPlanovani` do listu "Zakazky".

**Parametry:** Žádné

**Funkce:**
1. Vymaže existující data na listu "Zakazky"
2. Připojí se k databázi
3. Spustí SQL dotaz
4. Zapíše hlavičku (názvy sloupců)
5. Načte data pomocí `CopyFromRecordset`
6. Zobrazí progress bar během operace

**Prerekvizity:**
- Aktivní přihlášení
- Existence listu "Zakazky"
- Přístup k view `hvw_TerminyZakazekProPlanovani`

**Chybové stavy:**
- MsgBox pokud připojení selže
- Automatické ukončení při chybě

**Příklad použití:**
```vba
' Ruční volání z VBA
Call LoadOrUpdateData

' Nebo z tlačítka/události
Private Sub btnNacist_Click()
    LoadOrUpdateData
End Sub
```

**SQL dotaz:**
```sql
SELECT * FROM hvw_TerminyZakazekProPlanovani ORDER BY DatumUkonceni
```

**Výstup:**
- Data na listu "Zakazky", počínaje řádkem 1 (hlavička) a 2 (data)

---

### `AktualizovatSeznamZakazek`
```vba
Sub AktualizovatSeznamZakazek()
```
**Popis:** Aktualizuje Gantt diagram na základě dat z listu "Zakazky". Vytvoří seznam unikátních zakázek a vygeneruje řádky s vzorci.

**Parametry:** Žádné

**Funkce:**
1. Vypne screen updating a automatické výpočty (optimalizace)
2. Zruší všechny filtry
3. Vymaže předchozí sumarizaci pod grafem
4. Vytvoří Dictionary s unikátními čísly zakázek
5. Zapíše zakázky do sloupce B na listu "Gantt"
6. Zkopíruje vzorce z řádku 4 do nových řádků
7. Zkopíruje formátování z řádku 4
8. Odstraní přebytečné řádky
9. Spustí sumarizaci kapacit
10. Zapne zpět screen updating a výpočty

**Prerekvizity:**
- List "Zakazky" obsahuje data (běžte nejprve `LoadOrUpdateData`)
- List "Gantt" existuje a má správnou strukturu
- Řádek 4 na listu "Gantt" obsahuje vzorový řádek se vzorci

**Chybové stavy:**
- Error handler zajistí zapnutí screen updating i při chybě
- MsgBox s popisem chyby a číslem

**Příklad:**
```vba
' Kompletní aktualizace dat
Call LoadOrUpdateData           ' 1. Načíst data ze serveru
Call AktualizovatSeznamZakazek  ' 2. Aktualizovat Gantt

' Nebo pouze aktualizace Ganttu (pokud už jsou data načtená)
Call AktualizovatSeznamZakazek
```

**Optimalizace:**
- Používá Dictionary pro O(1) vyhledávání unikátních hodnot
- Screen updating vypnut během operací
- Calculation mode na Manual

**Výstup:**
- Aktualizovaný Gantt diagram s řádky pro každou zakázku
- Automaticky spočítané kapacity

---

### `RefreshAllConnections`
```vba
Sub RefreshAllConnections()
```
**Popis:** Obnoví všechna datová připojení v sešitu.

**Parametry:** Žádné

**Funkce:**
- Volá `ThisWorkbook.RefreshAll`
- Zobrazí potvrzovací MsgBox

**Příklad:**
```vba
Call RefreshAllConnections
```

**Poznámka:** Tato funkce obnovuje všechna Excel datová připojení (Power Query, externí data, atd.), ne pouze ADODB připojení z VBA.

---

### `IsInArray` (Helper)
```vba
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
```
**Popis:** Pomocná funkce pro kontrolu, zda je hodnota přítomna v poli.

**Parametry:**
- `valToBeFound` (Variant) - Hledaná hodnota
- `arr` (Variant) - Pole k prohledání

**Návratová hodnota:**
- `True` - Hodnota nalezena
- `False` - Hodnota nenalezena

**Příklad:**
```vba
Dim cisla() As Variant
cisla = Array(1, 2, 3, 4, 5)

If IsInArray(3, cisla) Then
    MsgBox "Hodnota 3 je v poli"
End If
```

---

### `BubbleSort` (Helper)
```vba
Sub BubbleSort(arr As Variant)
```
**Popis:** Seřadí pole vzestupně pomocí Bubble Sort algoritmu.

**Parametry:**
- `arr` (Variant) - Pole k seřazení (ByRef - mění originál)

**Návratová hodnota:** Žádná (pole je změněno přímo)

**Příklad:**
```vba
Dim cisla() As Variant
cisla = Array(5, 2, 8, 1, 9)

BubbleSort cisla
' cisla = [1, 2, 5, 8, 9]
```

**Složitost:** O(n²) - vhodné pro malá pole (< 1000 prvků)

**⚠️ Poznámka:** V současné verzi není funkce používána (zakomentováno v `AktualizovatSeznamZakazek`).

---

## Advanced.bas

Modul pro pokročilé funkce - výpočet kapacit, formátování, filtry.

### Globální proměnné

#### `x`
```vba
Dim x As Integer
```
**Popis:** Flag pro potlačení progress baru.

**Hodnoty:**
- `0` nebo `Nothing` - Zobrazit progress bar
- `1` - Potlačit progress bar

**Použití:** Interní řízení toku

---

### `AktualizaceKontrolyKapacit`
```vba
Sub AktualizaceKontrolyKapacit()
```
**Popis:** Aktualizuje kontrolu kapacit bez zobrazení progress baru.

**Parametry:** Žádné

**Funkce:**
1. Nastaví `x = 1` (potlačení progress baru)
2. Vymaže předchozí součty
3. Spustí sumarizaci kapacit

**Příklad:**
```vba
' Rychlá aktualizace bez progress baru
Call AktualizaceKontrolyKapacit
```

**Použití:** Vhodné pro opakované nebo rychlé přepočty, kde progress bar není nutný.

---

### `VymazaniSouctuPodGrafem`
```vba
Sub VymazaniSouctuPodGrafem()
```
**Popis:** Vymaže řádky se sumarizací kapacit pod Gantt diagramem.

**Parametry:** Žádné

**Funkce:**
1. Najde poslední řádek se zakázkou (sloupec B)
2. Smaže následující 4 řádky (řádky s kapacitami středisek)

**Bezpečnost:**
- Kontroluje, zda jsou řádky k smazání v rozsahu listu
- MsgBox s chybou pokud rozsah není platný

**Příklad:**
```vba
Call VymazaniSouctuPodGrafem
```

**⚠️ Varování:** Smaže vždy 4 řádky. Pokud změníte počet středisek, upravte konstantu.

---

### `SumarizaceBoduProVsechnySloupce`
```vba
Sub SumarizaceBoduProVsechnySloupce()
```
**Popis:** Hlavní procedura pro výpočet obsazenosti všech výrobních středisek pro všechny dny v časové ose.

**Parametry:** Žádné

**Funkce:**
1. Načte seznam svátků z listu "Svátky"
2. Definuje počty pracovníků pro každé středisko
3. Zobrazí progress bar (pokud `x <> 1`)
4. Projde všechny dny v časové ose (sloupce)
5. Pro každý pracovní den (ne víkend, ne svátek):
   - Projde všechny zakázky
   - Spočítá, kolik zakázek má danou fázi v daný den
   - Sečte body (počet pracovníků) pro každé středisko
6. Zapíše výsledky do řádků pod Ganttem
7. Aplikuje podmíněné formátování (zelená/oranžová/červená)
8. Zavře progress bar

**Konstanty kapacit:**
```vba
LidiPriprava = 1      ' Příprava: 1 pracovník
LidiSvarovani = 2     ' Svařování: 2 pracovníci
LidiMontaz = 2        ' Montáž: 2 pracovníci
LidiElektro = 1       ' Elektro: 1 pracovník
```

**Prerekvizity:**
- List "Gantt" s daty zakázek
- List "Svátky" se sloupcem B obsahujícím data svátků
- Pojmenované oblasti nebo globální konstanty `LidiPriprava`, `LidiSvarovani`, atd.

**Výstup:**
- 4 řádky pod Ganttem s hodnotami obsazenosti
- Barevné formátování:
  - **Zelená**: Kapacita pod 100%
  - **Oranžová**: Kapacita na 100%
  - **Červená**: Překročení kapacity
  - **Bílá**: Žádná zakázka nebo víkend/svátek

**Příklad:**
```vba
' Ruční spuštění
Call SumarizaceBoduProVsechnySloupce

' Volání bez progress baru (přes wrapper)
Call AktualizaceKontrolyKapacit
```

**Výkonnost:**
- Pro 365 dní a 50 zakázek: ~73,000 iterací
- Trvání: typicky 2-10 sekund (závisí na množství dat)
- Screen updating vypnut pro optimalizaci

**Logika výpočtu:**
```vba
Pro každý sloupec (den) od O do posledního:
  compareValue = datum z řádku 2

  SKIP pokud je víkend (Sobota nebo Neděle)
  SKIP pokud je svátek

  Pro každé středisko (Příprava, Svařování, Montáž, Elektro):
    body = 0

    Pro každou zakázku (řádek 4 až poslední):
      IF compareValue >= start_fáze AND compareValue <= konec_fáze THEN
        body = body + počet_pracovníků_střediska
      END IF

    Zapiš body do příslušné buňky

  Aplikuj podmíněné formátování
```

---

### `ZrusitVsechnyFiltry`
```vba
Sub ZrusitVsechnyFiltry()
```
**Popis:** Zruší všechny aktivní filtry na listu "Gantt".

**Parametry:** Žádné

**Funkce:**
- Zkontroluje, zda je AutoFilter aktivní
- Pokud ano, zobrazí všechna data (`ShowAllData`)

**Příklad:**
```vba
Call ZrusitVsechnyFiltry
```

**Bezpečnost:**
- Nekončí chybou, pokud filtry nejsou aktivní

---

### `ZobrazitMinimalistickyGantt`
```vba
Sub ZobrazitMinimalistickyGantt()
```
**Popis:** Aktivuje minimalistický režim zobrazení Gantt diagramu.

**Parametry:** Žádné

**Funkce:**
1. Aktivuje list "Gantt"
2. Skryje mřížku (`DisplayGridlines = False`)
3. Skryje záhlaví řádků a sloupců (`DisplayHeadings = False`)
4. Skryje panel vzorců (`DisplayFormulaBar = False`)

**Příklad:**
```vba
Call ZobrazitMinimalistickyGantt
```

**Poznámka:** Minimalistický režim se automaticky aktivuje v `Workbook_Open` a ruší v `Workbook_BeforeClose`.

---

## frmLogin

Přihlašovací formulář pro autentizaci uživatele.

### Události

#### `UserForm_Initialize`
```vba
Private Sub UserForm_Initialize()
```
**Popis:** Inicializuje formulář při načtení.

**Funkce:**
- Načte hodnoty z listu "Konfigurace" do polí formuláře:
  - `txtServerName` ← `serverName`
  - `txtDBName` ← `databaseName`
  - `txtUsername` ← `login`

**Automaticky voláno:** Při `Load frmLogin` nebo `frmLogin.Show`

---

#### `btnLogin_Click`
```vba
Private Sub btnLogin_Click()
```
**Popis:** Zpracuje kliknutí na tlačítko "Přihlásit".

**Funkce:**
1. Načte hodnoty z formulářových polí
2. Uloží hodnoty do pojmenovaných oblastí na listu "Konfigurace"
3. Ověří credentials pomocí `VerifyCredentials()`
4. Pokud platné:
   - Zašifruje a uloží do `loginCredentials`
   - Nastaví `isLoggedIn = True`
   - Zavře formulář
   - Maximalizuje a zobrazí Excel
   - Aktivuje list "Gantt"
5. Pokud neplatné:
   - Zobrazí chybovou zprávu
   - Vrátí focus na pole username

**Formulářová pole:**
- `txtUsername` - Uživatelské jméno
- `txtPassword` - Heslo
- `txtServerName` - SQL server
- `txtDBName` - Databáze

---

#### `UserForm_QueryClose`
```vba
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
```
**Popis:** Zpracuje pokus o zavření formuláře.

**Funkce:**
- Pokud uživatel není přihlášen (`isLoggedIn = False`):
  - Zavře celý sešit bez uložení

**Účel:** Zajistit, že uživatel nemůže používat aplikaci bez přihlášení.

---

## frmProgress

Formulář pro zobrazení průběhu dlouhých operací.

### Vlastnosti

#### `lblProgress`
**Popis:** Label s textem "Načítání dat, prosím čekejte..."

#### `lblProgressText`
**Popis:** Label pro zobrazení % dokončení

#### `lblProgressBar`
**Popis:** Label sloužící jako progress bar (mění se šířka)

#### `FrameProgress`
**Popis:** Frame obsahující progress bar

---

### Události

#### `UserForm_Initialize`
```vba
Private Sub UserForm_Initialize()
```
**Popis:** Inicializuje formulář.

**Funkce:**
- Nastaví text na "Načítání dat, prosím čekejte..."
- Nastaví šířku progress baru na 0

---

### Metody

#### `UpdateProgressBar`
```vba
Public Sub UpdateProgressBar(progress As Double)
```
**Popis:** Aktualizuje progress bar na zadané % dokončení.

**Parametry:**
- `progress` (Double) - Hodnota 0.0 až 1.0 (0% až 100%)

**Funkce:**
- Vypočítá novou šířku progress baru: `progress × (šířka_framu - 10)`
- Aktualizuje `lblProgressBar.Width`
- Volá `DoEvents` pro překreslení

**Příklad:**
```vba
Load frmProgress
frmProgress.Show vbModeless

For i = 1 To 100
    ' ... nějaká operace ...
    frmProgress.UpdateProgressBar i / 100  ' 0.01, 0.02, ..., 1.0
    frmProgress.lblProgressText.Caption = i & " %"
Next i

Unload frmProgress
```

**Best practices:**
- Vždy zobrazujte modeless: `frmProgress.Show vbModeless`
- Aktualizujte progress bar v rozumných intervalech (ne po každé iteraci v tisících iterací)
- Nezapomeňte zavřít: `Unload frmProgress`

---

## Konstanty a konfigurace

### Pojmenované oblasti (List "Konfigurace")

| Název | Popis | Datový typ | Příklad |
|-------|-------|------------|---------|
| `serverName` | Název SQL serveru | String | "SERVER01" |
| `databaseName` | Název databáze | String | "ERP_IN_EKO" |
| `login` | Uživatelské jméno | String | "admin" nebo "" |
| `localServer` | *(nepoužito)* | String | - |

### Sloupce Gantt diagramu

| Sloupec | Popis | Datový typ |
|---------|-------|------------|
| B | Číslo zakázky | String |
| C | *(závisí na datech)* | Variant |
| E | Datum zahájení Přípravy | Date |
| F | Datum ukončení Přípravy | Date |
| G | Datum zahájení Svařování | Date |
| H | Datum ukončení Svařování | Date |
| I | Datum zahájení Montáže | Date |
| J | Datum ukončení Montáže | Date |
| K | Datum zahájení Elektro | Date |
| L | Datum ukončení Elektro | Date |
| O+ | Časová osa (každý sloupec = 1 den) | Calculated |

---

## Vývojářské poznámky

### Běžné úkoly

**Přidání nového SQL dotazu:**
```vba
Sub NactiMojeData()
    Dim conn As Object
    Dim rs As Object

    Set conn = CreateConnection()
    If conn Is Nothing Then Exit Sub

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT * FROM MujView", conn

    ' Zpracování dat...

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub
```

**Úprava počtu pracovníků střediska:**
```vba
' V Advanced.bas, procedura SumarizaceBoduProVsechnySloupce
LidiPriprava = 2      ' Změna z 1 na 2 pracovníky
```

**Změna SQL view:**
```vba
' V Data.bas, procedura LoadOrUpdateData
sql = "SELECT * FROM MujNovyView ORDER BY Datum"
```

### Debugování

**Zapnutí debug výpisů:**
```vba
' V Connection.bas, řádek 115 (odkomentovat)
Debug.Print "Connection String: " & connectionString

' V Advanced.bas, různé řádky
Debug.Print "Posledn sloupec: " & lastCol
```

**Sledování hodnot:**
- Immediate Window (Ctrl+G): `? loginCredentials("login")`
- Locals Window: Ctrl+L
- Watch Window: Přidat proměnnou do Watch

---

## Verzování

**Aktuální verze:** Export 2026-01-16

**Changelog API:**
- 2026-01-16: Iniciální dokumentace
