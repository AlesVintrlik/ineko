# Architektura - Plánování

## Přehled architektury

Aplikace Plánování je postavena na vícevrstvé architektuře s integrací do Helios ERP systému.

```
┌───────────────────────────────────────────────────────────────┐
│                   Prezentační vrstva                          │
│  ┌────────────┐  ┌──────────┐  ┌──────────────────────────┐  │
│  │ UserForms  │  │ Ribbons  │  │ Excel Listy              │  │
│  │ - frmLogin │  │ - Zkratky│  │ - Plan                   │  │
│  │ - frmZakazka│ │ - Toolbar│  │ - Zakazky                │  │
│  │ - frmHodiny│  │          │  │ - Operace                │  │
│  │ - frmProgress│ │          │  │ - Obsazenost úseků       │  │
│  │ - frmChangelog│ │         │  │ - Obsazenost pracovišť   │  │
│  │ - frmReady │  │          │  │ - Kontrola               │  │
│  └────────────┘  └──────────┘  └──────────────────────────┘  │
└───────────────────────────────────────────────────────────────┘
                          ↕
┌───────────────────────────────────────────────────────────────┐
│                   Business logika                             │
│  ┌──────────────┐  ┌───────────────┐  ┌──────────────────┐   │
│  │Main.bas      │  │VypocetPlanu   │  │PracnostZakazek   │   │
│  │- Workflow    │  │.bas           │  │.bas              │   │
│  │- UI helpers  │  │- Výpočet      │  │- Evidence hodin  │   │
│  │- Font mgmt   │  │  termínů      │  │- Uložení do DB   │   │
│  └──────────────┘  └───────────────┘  └──────────────────┘   │
│  ┌──────────────┐  ┌───────────────┐  ┌──────────────────┐   │
│  │Connection.bas│  │Data.bas       │  │Helios.bas        │   │
│  │- Autentizace │  │- Načítání dat │  │- Synchronizace   │   │
│  │- Šifrování   │  │- Aktualizace  │  │- SP volání       │   │
│  └──────────────┘  └───────────────┘  └──────────────────┘   │
└───────────────────────────────────────────────────────────────┘
                          ↕ ADODB
┌───────────────────────────────────────────────────────────────┐
│                    Datová vrstva                              │
│                 SQL Server - IN-EKO ERP DB                    │
│  ┌──────────────────────────────────────────────────────┐    │
│  │ Views                                                 │    │
│  │ - hvw_TerminyZakazekProPlanovani                     │    │
│  └──────────────────────────────────────────────────────┘    │
│  ┌──────────────────────────────────────────────────────┐    │
│  │ Tables                                                │    │
│  │ - TabZakazka (READ)                                   │    │
│  │ - TabZakazka_EXT (READ/WRITE)                        │    │
│  └──────────────────────────────────────────────────────┘    │
│  ┌──────────────────────────────────────────────────────┐    │
│  │ Stored Procedures                                     │    │
│  │ - EP_PlanTerminyVyrobyDoZakazek                      │    │
│  │ - ep_DoplnTerminOperacePodleUseku                    │    │
│  └──────────────────────────────────────────────────────┘    │
│                                                               │
│                   Helios ERP System                           │
└───────────────────────────────────────────────────────────────┘
```

## Komponenty systému

### 1. Connection.bas - Správa připojení

**Zodpovědnost:**
- Autentizace uživatelů
- Vytváření databázových připojení
- Šifrování credentials (XOR)

**Stejný jako v Gantt**, viz Gantt/ARCHITECTURE.md pro detaily.

---

### 2. Data.bas - Datová vrstva

**Zodpovědnost:**
- Načítání dat ze SQL Server
- Aktualizace Excel listů
- Správa životního cyklu ADO objektů

**Klíčové funkce:**
- `LoadOrUpdateData()` - Načte zakázky z view do listu "Zakazky"
- Podobné jako v Gantt, ale s rozšířenou funkcionalitou

---

### 3. Main.bas - Hlavní modul aplikace

**Zodpovědnost:**
- Hlavní workflow aplikace
- Správa UI (fonty, formátování)
- Utility funkce
- Win32 API volání pro font detection

#### Klíčové procedury

**`AktualizovatSeznamZakazek()`**
- Načte zakázky z listu "Zakazky"
- Vytvoří Dictionary s unikátními zakázkami
- Seřadí podle termínu expedice (ElektroKonec)
- Aktualizuje list "Plan"
- Zkopíruje vzorce a formátování

**Algoritmus:**
```
1. Odemkni list "Plan"
2. Najdi sloupce "ElektroKonec" a "Firma" v listu "Zakazky"
3. Pro každou zakázku:
     Vytvoř klíč = číslo zakázky
     Ulož do Dictionary: [číslo, firma, datum expedice]
4. Převeď Dictionary na pole
5. Seřaď pole podle data expedice (BubbleSort)
6. Zap IS do listu "Plan" (sloupec B = zakázka, C = firma)
7. Zkopíruj vzorce z řádku 15 do nových řádků
8. Zkopíruj formátování
9. Smaž přebytečné řádky
10. Zapni AutoFilter
11. Zamkni list "Plan"
12. Nastav font a formát datumu
```

**`BubbleSortZakazky(arr)`**
- Seřadí zakázky podle data expedice
- Prázdné datumy (NULL) jsou vždy na konci
- Používá Bubble Sort algoritmus

**Logika řazení:**
```vba
Pro i od 0 do N-1:
  Pro j od i+1 do N:
    IF arr[i] je prázdné THEN
      NIC (prázdné zůstává na konci)
    ELSE IF arr[j] je prázdné THEN
      SWAP arr[i], arr[j]  (posuň prázdné dozadu)
    ELSE IF arr[i] > arr[j] THEN
      SWAP arr[i], arr[j]  (standardní řazení)
```

**`NaplnitKontrolu()`**
- Vytvoří kontrolní list pro synchronizaci do Helios
- Projde všechny zakázky a úseky
- Zapíše změny termínů do listu "Kontrola"
- Vypočítá týdny (ISO week number)

**Struktura listu "Kontrola":**
| Sloupec | Popis |
|---------|-------|
| A | Úsek (ID) |
| B | Zakázka (číslo) |
| C | Původní termín |
| D | Termín výroby do |
| E | Potřebný čas celkem (hod) |
| F | Týden původní |
| G | Týden nový |

**`PlusTyden()` a `MinusTyden()`**
- Klávesové zkratky Ctrl+N a Ctrl+D
- Přidají/odeberou 7 dní k/od termínu
- Fungují pouze ve sloupcích M, P, S, V, Y (termíny výroby)
- Kontrolují, zda levá buňka obsahuje datum (původní termín)

**Logika:**
```vba
IF aktivní sloupec IN [13, 16, 19, 22, 25] AND řádek >= 13 THEN
  levá_buňka = buňka vlevo od aktivní

  IF levá_buňka je datum THEN
    IF aktivní_buňka je prázdná THEN
      aktivní_buňka = levá_buňka + 7 dní
    ELSE IF (aktivní_buňka - levá_buňka) MOD 7 = 0 THEN
      aktivní_buňka = aktivní_buňka + 7 dní
    ELSE
      aktivní_buňka = levá_buňka + 7 dní
    END IF
  END IF
END IF
```

**`NastavFontPlusFormatDatumu()`**
- Detekuje dostupné fonty v systému
- Vybere nejlepší font (Segoe UI Semilight → Arial)
- Aplikuje font na celý list
- Nastaví formát datumu `dd.mm.yy`
- Automaticky přizpůsobí šířku sloupců
- Zobrazuje progress během operace

**Font Selection Algorithm:**
```
favFonts = ["Segoe UI Semilight", "Segoe UI Light", "Calibri Light", "Arial Narrow", "Arial"]

Pro každý font v favFonts:
  IF IsFontInstalled(font) AND SupportsCzech(font) THEN
    chosenFont = font
    BREAK
  END IF

IF chosenFont = "" THEN
  chosenFont = "Arial"  (fallback)
END IF

Aplikuj font na ws.Cells
```

**Win32 API - Font Detection:**
```vba
' Deklarace API funkcí
EnumFontFamiliesEx()  - Enumeruje všechny fonty
GetDC()                - Získá device context
ReleaseDC()            - Uvolní device context
CopyMemory()           - Kopíruje paměť (LOGFONT struktura)

' Callback funkce
EnumFontsProc()       - Volána pro každý font
  - Porovná název fontu s hledaným
  - Nastaví m_fontFound = True při shodě
```

**Utility funkce:**
- `ObsazenostUseku()` - Zobrazí list "Obsazenost úseků"
- `ObsazenostPracovist()` - Zobrazí list "Obsazenost pracovišť"
- `Konfigurace()` - Zobrazí list "Konfigurace"
- `Ribbon()` - Zobrazí Ribbon (Ctrl+L)
- `Ladeni()` - Toggle Ribbon + Formula Bar + Status Bar
- `Delay(seconds)` - Čekání s DoEvents
- `ToggleLabel()` - Blikání labelu ve frmReady
- `ZobrazFormularZpozdene()` - Zobrazí formulář po zpoždění

---

### 4. VypocetPlanu.bas - Výpočet termínů výroby

**Zodpovědnost:**
- Volání stored procedure pro výpočet termínů
- Zpětné plánování od data expedice
- Získání ID zakázky

#### Klíčové funkce

**`OtevritFormularZakazky()`**
- Otevře formulář pro detail zakázky
- Předá číslo zakázky z aktivního řádku

**`CallPlanTerminyVyrobyDoZakazek(zakazkaID, datumUkonceni)`**
- Volá SP `EP_PlanTerminyVyrobyDoZakazek`
- Parametry:
  - `@ID` (Long) - ID zakázky z TabZakazka
  - `@DatumUkonceni` (Date/NULL) - Datum expedice nebo NULL

**Logika:**
```vba
IF datumUkonceni je prázdné OR NULL THEN
  param_@DatumUkonceni = NULL
ELSE
  Validuj datum
  param_@DatumUkonceni = Format(datum, "YYYY-MM-DD")
END IF

Vytvoř ADODB.Command
Nastav CommandType = adCmdStoredProc
Přidej parametry
Spusť cmd.Execute
```

**`GetZakazkaID(cisloZakazky)`**
- Získá ID zakázky podle čísla zakázky
- SQL: `SELECT ID FROM TabZakazka WHERE CisloZakazky = '...'`
- Vrací: Long (ID) nebo 0 (nenalezeno)

---

### 5. PracnostZakazek.bas - Evidence pracnosti

**Zodpovědnost:**
- Evidence plánovaných hodin
- Uložení do `TabZakazka_EXT`

#### Klíčové funkce

**`OtevritFormularHodiny()`**
- Otevře formulář pro evidenci hodin
- Předá číslo zakázky z aktivního řádku

**`UlozitHodiny(...)`**
- Parametry:
  - `zakazkaID` (Long)
  - `HodCelkem` (Long)
  - `HodSkPrac1-5` (Long) - Hodiny podle skupin pracovníků
  - `HodKoop` (Long) - Hodiny kooperací

**Logika:**
```sql
-- 1. Zajistí existenci záznamu
IF NOT EXISTS (SELECT 1 FROM TabZakazka_EXT WHERE ID = ...)
  INSERT INTO TabZakazka_EXT (ID) VALUES (...)

-- 2. Aktualizuje hodnoty
UPDATE TabZakazka_EXT
SET _HodCelkem = ...,
    _HodSkPrac1 = ...,
    _HodSkPrac2 = ...,
    _HodSkPrac3 = ...,
    _HodSkPrac4 = ...,
    _HodSkPrac5 = ...,
    _HodKoop = ...
WHERE ID = ...
```

---

### 6. Helios.bas - Integrace s Helios ERP

**Zodpovědnost:**
- Synchronizace termínů zpět do Helios
- Volání SP pro přepočet operací
- Batch aktualizace všech změn

#### Klíčové funkce

**`AktualizovatData()`**
- Hlavní procedura pro synchronizaci do Helios
- Projde všechny řádky v listu "Kontrola"
- Pro každý řádek:
  1. Aktualizuje termín konce úseku (`_U{usek}Konec`)
  2. Vypočítá rozdíl dnů (dateDiff)
  3. Aktualizuje termín začátku úseku (`_U{usek}Start = konec - dateDiff`)
  4. Zavolá `ep_DoplnTerminOperacePodleUseku` pro přepočet operací
- Obnoví všechna datová připojení v Excelu

**Workflow:**
```
Pro každý řádek v "Kontrola":
  usek = A[i]
  zakazka = B[i]
  terminDo = D[i]

  1. UPDATE TabZakazka_EXT
     SET _U{usek}Konec = terminDo
     WHERE CisloZakazky = zakazka

  2. dateDiff = GetDateDiffAndID(usek, zakazka)

  3. UPDATE TabZakazka_EXT
     SET _U{usek}Start = terminDo - dateDiff dni
     WHERE CisloZakazky = zakazka

  4. CallStoredProcedure(ID)  -- ep_DoplnTerminOperacePodleUseku

  Wait 1 sekunda mezi iteracemi
```

**`GetDateDiffAndID(usek, zakazka, ByRef ID)`**
- Vypočítá rozdíl ve dnech mezi začátkem a koncem úseku
- Vrací také ID zakázky (ByRef parametr)

**SQL:**
```sql
SELECT TZ.ID, DATEDIFF(day, _U{usek}Start, _U{usek}Konec) as DateDiff
FROM TabZakazka_EXT AS TZE
JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID
WHERE TZ.CisloZakazky = '...'
```

**`CallStoredProcedure(ID)`**
- Volá `ep_DoplnTerminOperacePodleUseku`
- Přepočítá termíny jednotlivých operací podle úseků
- Skryje list "Kontrola"
- Aktivuje list "Plan"

---

## Datový model

### Excel listy

**List "Plan"** (hlavní list)
- Řádek 1-14: Záhlaví, metadata, filtry
- Řádek 15+: Zakázky (1 řádek = 1 zakázka)
- Sloupce:
  - `B` - Číslo zakázky
  - `C` - Firma
  - `D-G` - Metadata
  - `H-I` - Dodací podmínky
  - `J` - Pomocný sloupec (šířka 1)
  - `K-Y` - Normované časy a termíny
  - `L, O, R, U, X, AA` - Původní termíny
  - `M, P, S, V, Y, AB` - Plánované termíny (editovatelné)
  - `N, Q, T, W, Z, AC` - Normované časy

**Mapování úseků:**
| Úsek | ID | Sloupec původní | Sloupec plánovaný | Normované časy |
|------|----|-----------------|-------------------|----------------|
| Příprava | 1 | L | M | N |
| Pila | 4 | O | P | Q |
| Svařování | 2 | R | S | T |
| Montáž | 3 | U | V | W |
| Elektro | 5 | X | Y | Z |
| Balení | 8 | AA | AB | AC |

**List "Zakazky"** (cache dat z DB)
- Raw data z view `hvw_TerminyZakazekProPlanovani`
- Slouží jako datový zdroj pro list "Plan"

**List "Operace"** (data z Helios)
- Data operací a normovaných časů
- Použito v `NaplnitKontrolu()` pro SUMIFS

**List "Kontrola"** (dočasný)
- Vytváří se pomocí `NaplnitKontrolu()`
- Obsahuje změny termínů před synchronizací do Helios
- Smaže se po úspěšné synchronizaci

**List "Obsazenost úseků"**
- Přehled obsazenosti výrobních úseků
- Pivotové tabulky a grafy

**List "Obsazenost pracovišť"**
- Detailní přehled jednotlivých pracovišť
- Kapacity a aktuální vytížení

**List "Konfigurace"**
- Pojmenované oblasti pro databázové připojení
- Stejné jako v Gantt

**List "Svátky"** (pokud existuje)
- Seznam svátků pro výpočty

---

### Databázový model

**TabZakazka** (READ)
```sql
CREATE TABLE TabZakazka (
    ID INT PRIMARY KEY,
    CisloZakazky NVARCHAR(50),
    Firma NVARCHAR(200),
    ...
)
```

**TabZakazka_EXT** (READ/WRITE)
```sql
CREATE TABLE TabZakazka_EXT (
    ID INT PRIMARY KEY,  -- FK na TabZakazka.ID
    -- Termíny úseků
    _U1Start DATETIME,
    _U1Konec DATETIME,
    _U2Start DATETIME,
    _U2Konec DATETIME,
    _U3Start DATETIME,
    _U3Konec DATETIME,
    _U4Start DATETIME,
    _U4Konec DATETIME,
    _U5Start DATETIME,
    _U5Konec DATETIME,
    _U8Start DATETIME,
    _U8Konec DATETIME,
    -- Hodiny
    _HodCelkem INT,
    _HodSkPrac1 INT,
    _HodSkPrac2 INT,
    _HodSkPrac3 INT,
    _HodSkPrac4 INT,
    _HodSkPrac5 INT,
    _HodKoop INT,
    ...
)
```

---

## Diagram toku dat (DFD)

### Načítání dat

```
┌──────────┐
│ Uživatel │
└─────┬────┘
      │ Otevře aplikaci
      ↓
┌──────────────┐
│ frmLogin     │ ─→ Ověření credentials
└──────┬───────┘
       │ Úspěch
       ↓
┌─────────────────────────────┐
│ Data.LoadOrUpdateData()     │
└──────┬──────────────────────┘
       │ SQL: SELECT * FROM hvw_TerminyZakazekProPlanovani
       ↓
┌────────────────┐
│ List "Zakazky" │ (raw data)
└──────┬─────────┘
       │
       ↓
┌──────────────────────────────────┐
│ Main.AktualizovatSeznamZakazek() │
└──────┬───────────────────────────┘
       │ Dictionary + BubbleSort
       ↓
┌────────────────┐
│ List "Plan"    │ (zakázky + vzorce)
└──────┬─────────┘
       │
       ↓
┌──────────────────────────────────┐
│ Main.NastavFontPlusFormatDatumu() │
└──────┬───────────────────────────┘
       │ Font detection + formátování
       ↓
┌──────────────┐
│ frmReady     │ (blikající "Hotovo!")
└──────────────┘
```

### Výpočet termínů výroby

```
┌──────────┐
│ Uživatel │ klikne na zakázku
└─────┬────┘
      ↓
┌──────────────────────┐
│ frmZakazka           │
│ - Číslo zakázky      │
│ - Datum expedice     │
└─────┬────────────────┘
      │ Klikne "Vypočítat"
      ↓
┌─────────────────────────────────────┐
│ VypocetPlanu.GetZakazkaID()         │
│ SQL: SELECT ID FROM TabZakazka...   │
└─────┬───────────────────────────────┘
      │ ID = 12345
      ↓
┌───────────────────────────────────────────────┐
│ VypocetPlanu.CallPlanTerminyVyrobyDoZakazek() │
│ EXEC EP_PlanTerminyVyrobyDoZakazek            │
│   @ID = 12345,                                │
│   @DatumUkonceni = '2024-12-31'               │
└───────┬───────────────────────────────────────┘
        │ SP vypočítá termíny zpětně
        ↓
┌──────────────────────┐
│ TabZakazka_EXT       │
│ - _U1Start, _U1Konec │
│ - _U2Start, _U2Konec │
│ - ...                │
└───────┬──────────────┘
        │
        ↓
┌──────────────────────┐
│ Data.LoadOrUpdateData│ (refresh)
└───────┬──────────────┘
        │
        ↓
┌──────────────────────┐
│ List "Plan"          │ (aktualizované termíny)
└──────────────────────┘
```

### Synchronizace do Helios

```
┌──────────┐
│ Uživatel │ upraví termíny v listu "Plan"
└─────┬────┘
      ↓
┌──────────────────────┐
│ Main.NaplnitKontrolu()│
└─────┬────────────────┘
      │ Projde všechny úseky
      ↓
┌──────────────────┐
│ List "Kontrola"  │ (změny termínů)
│ Úsek │ Zakázka │ Původní │ Nový
│  1   │ 12345   │ 01.12.  │ 08.12.
│  2   │ 12345   │ 05.12.  │ 12.12.
└─────┬────────────┘
      │ Uživatel zkontroluje a potvrdí
      ↓
┌─────────────────────────┐
│ Helios.AktualizovatData()│
└─────┬───────────────────┘
      │ Pro každý řádek:
      ↓
┌──────────────────────────────────────────┐
│ 1. UPDATE TabZakazka_EXT                 │
│    SET _U{usek}Konec = novy_termin       │
└─────┬────────────────────────────────────┘
      ↓
┌──────────────────────────────────────────┐
│ 2. GetDateDiffAndID() → dateDiff         │
└─────┬────────────────────────────────────┘
      ↓
┌──────────────────────────────────────────┐
│ 3. UPDATE TabZakazka_EXT                 │
│    SET _U{usek}Start = konec - dateDiff  │
└─────┬────────────────────────────────────┘
      ↓
┌──────────────────────────────────────────┐
│ 4. EXEC ep_DoplnTerminOperacePodleUseku  │
│    @ID = ...                             │
└─────┬────────────────────────────────────┘
      │ Přepočítá operace
      ↓
┌──────────────────┐
│ Helios ERP       │ (synchronizováno)
└──────────────────┘
```

---

## Bezpečnostní architektura

### Ochrana listů

**Password:** `MrkevNeniOvoce123`

**⚠️ Bezpečnostní riziko:** Heslo je uloženo v plaintext v kódu!

**Ochrana listu "Plan":**
```vba
ws.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True

' Odemknutí před úpravami
ws.Unprotect password:="MrkevNeniOvoce123"

' Zamknutí po úpravách
ws.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True
```

**Povolené akce při zamčení:**
- Filtrování (`AllowFiltering:=True`)
- Úprava odemčených buněk
- Spouštění maker

**Zakázané akce:**
- Smazání/úprava vzorců
- Přesunutí řádků/sloupců
- Změna struktury

---

## Výkonnostní optimalizace

### 1. Screen updating a Calculation

Stejné jako v Gantt - vypínání během batch operací.

### 2. Dictionary vs Array

Použití `Scripting.Dictionary` pro:
- O(1) vyhledávání unikátních zakázek
- Eliminace duplicit

### 3. Font Detection

Ukládání do globální proměnné `m_fontFound` místo opakovaného volání API.

### 4. Progress Feedback

Použití `frmProgress` pro vizuální feedback:
- Načítání dat
- Font detection
- Formátování

### 5. Delayed Execution

Použití `Application.OnTime` pro:
- Zpožděné zobrazení formulářů
- Blikání labelů bez blokování UI

---

## Design Patterns

### 1. Facade Pattern
`CreateConnection()` zapouzdřuje složitost vytváření ADO připojení.

### 2. Template Method
`AktualizovatSeznamZakazek()` definuje workflow, ale deleguje specifické úkoly:
- `BubbleSortZakazky()` - řazení
- `NastavFontPlusFormatDatumu()` - formátování

### 3. Strategy Pattern
Font selection - vybírá strategii podle dostupnosti fontů.

### 4. Observer Pattern (lehká forma)
`ToggleLabel()` pomocí `Application.OnTime` simuluje observer pro blikání.

---

## Závislosti

### Knihovny
- **Microsoft ActiveX Data Objects 2.x** (ADODB)
- **Microsoft Scripting Runtime** (Dictionary)
- **VBA7** pro 64-bit Excel (Win32 API deklarace)

### Win32 API
- `gdi32.dll` - EnumFontFamiliesEx (font detection)
- `user32.dll` - GetDC, ReleaseDC
- `kernel32.dll` - CopyMemory (RtlMoveMemory)

### Databáze
- **SQL Server** - jakákoli podporovaná verze
- **View**: `hvw_TerminyZakazekProPlanovani`
- **Tables**: `TabZakazka`, `TabZakazka_EXT`
- **SP**: `EP_PlanTerminyVyrobyDoZakazek`, `ep_DoplnTerminOperacePodleUseku`

### Excel
- **Minimální verze**: Excel 2013
- **Doporučeno**: Excel 2016+ (64-bit)

---

## Rozšiřitelnost

### Přidání nového výrobního úseku

1. **Databáze**: Přidat sloupce `_U{ID}Start`, `_U{ID}Konec` do `TabZakazka_EXT`
2. **Excel**: Přidat sloupce do listu "Plan"
3. **Main.bas**: Aktualizovat `columnCodes` a `usekValues` v `NaplnitKontrolu()`
4. **Main.bas**: Přidat sloupec do array v `PlusTyden()` a `MinusTyden()`

### Změna šifrovacího algoritmu

Stejné jako v Gantt - nahradit `XOREncryptDecrypt()`.

### Přidání nové stored procedure

```vba
Sub CallMojaStoredProcedure(param1 As Long)
    Dim conn As Object
    Dim cmd As Object

    Set conn = CreateConnection()
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = 4  ' adCmdStoredProc
    cmd.CommandText = "dbo.MojaSP"

    cmd.Parameters.Append cmd.CreateParameter("@Param1", 3, 1, , param1)
    cmd.Execute

    conn.Close
End Sub
```

---

## Logování

### WriteLog funkce

Aplikace obsahuje logování (implementace není vidět, ale je volána):

```vba
WriteLog "zpráva"
```

Volána v:
- `Main.AktualizovatSeznamZakazek()` - "Inserting formulas..."
- `VypocetPlanu.CallPlanTerminyVyrobyDoZakazek()` - "Starting procedures..."
- `PracnostZakazek.UlozitHodiny()` - "Starting procedures..."
- `Helios.AktualizovatData()` - "Deadlines set in Helios..."

---

## Best Practices

### 1. Vždy odemkni před úpravou
```vba
ws.Unprotect password:="MrkevNeniOvoce123"
' ... úpravy ...
ws.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True
```

### 2. Používej progress bar pro dlouhé operace
```vba
frmProgress.lblProgress.Caption = "Zpracovávám..."
frmProgress.UpdateProgressBar 0.5
DoEvents
```

### 3. Kontroluj návratové hodnoty
```vba
Dim ID As Long
ID = GetZakazkaID(cislo)
If ID = 0 Then
    MsgBox "Zakázka nenalezena"
    Exit Sub
End If
```

### 4. Loguj klíčové operace
```vba
WriteLog "Starting operation..."
' ... operace ...
WriteLog "Operation completed."
```

### 5. Uzavírej připojení
```vba
If Not conn Is Nothing Then
    If conn.State = 1 Then conn.Close
    Set conn = Nothing
End If
```

---

## Diagram komponent

```
Main.bas ────┬──→ Connection.bas (CreateConnection)
             ├──→ Data.bas (LoadOrUpdateData)
             └──→ VypocetPlanu.bas (CallPlan...)

VypocetPlanu.bas ──→ Connection.bas (CreateConnection)

PracnostZakazek.bas ──→ Connection.bas (CreateConnection)

Helios.bas ────┬──→ Connection.bas (CreateConnection)
               └──→ VypocetPlanu.bas (GetDateDiffAndID)

frmZakazka ──→ VypocetPlanu.bas (OtevritFormularZakazky)

frmHodiny ──→ PracnostZakazek.bas (OtevritFormularHodiny)
```

---

## Verze

**Export:** 2026-01-16
**Řádků kódu:** ~2,193
**Modulů:** 7
**Formulářů:** 6
**ExcelObjects:** 11
