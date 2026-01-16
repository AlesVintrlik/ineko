# API Reference - Plánování

Tato dokumentace popisuje všechny veřejné (Public) funkce a procedury dostupné v aplikaci Plánování.

## Obsah

- [Connection.bas](#connectionbas)
- [Data.bas](#databas)
- [Main.bas](#mainbas)
- [VypocetPlanu.bas](#vypocetplanubas)
- [PracnostZakazek.bas](#pracnostzakazekbas)
- [Helios.bas](#heliosbas)
- [UserForms](#userforms)

---

## Connection.bas

**Poznámka:** Connection.bas je **identický** s verzí v Gantt projektu.

Pro kompletní dokumentaci viz: [Gantt/API.md - Connection.bas](../Gantt/API.md#connectionbas)

**Stručný přehled:**
- `XOREncryptDecrypt(text, key)` - Šifrování/dešifrování
- `VerifyCredentials(userName, password, ServerName, databaseName)` - Ověření údajů
- `GetDecryptedCredentials()` - Vrací šifrované credentials
- `CreateConnection()` - Vytvoří ADODB.Connection

---

## Data.bas

**Poznámka:** Data.bas je **podobný** jako v Gantt, ale s rozšířeními.

### `LoadOrUpdateData`
```vba
Sub LoadOrUpdateData()
```
**Popis:** Načte data z databázového view do listu "Zakazky".

**View:** `hvw_TerminyZakazekProPlanovani`

**Funkce:** Stejná jako v Gantt - viz [Gantt/API.md](../Gantt/API.md#loadorupdatedata)

---

## Main.bas

Hlavní modul aplikace s workflow funkcemi, UI helpery a Win32 API voláními.

### Globální proměnné

#### Font Detection
```vba
Dim m_targetFont As String
Dim m_fontFound As Boolean
```
**Popis:** Globální proměnné pro detekci fontů přes Win32 API.
- `m_targetFont` - Hledaný font
- `m_fontFound` - True pokud font nalezen

---

### Workflow procedury

#### `AktualizovatSeznamZakazek`
```vba
Sub AktualizovatSeznamZakazek()
```
**Popis:** Aktualizuje seznam zakázek v listu "Plan" na základě dat z listu "Zakazky".

**Funkce:**
1. Načte zakázky z listu "Zakazky"
2. Vytvoří Dictionary s unikátními zakázkami
3. Seřadí podle termínu expedice (BubbleSortZakazky)
4. Zapíše do listu "Plan"
5. Zkopíruje vzorce z řádku 15
6. Zkopíruje formátování
7. Smaže přebytečné řádky
8. Nastaví AutoFilter
9. Zavolá NastavFontPlusFormatDatumu

**Prerekvizity:**
- List "Zakazky" obsahuje data
- List "Plan" existuje
- Sloupce "ElektroKonec" a "Firma" existují v "Zakazky"

**Chybové stavy:**
- MsgBox pokud sloupec "ElektroKonec" nebo "Firma" neexistuje
- Error handler pro ostatní chyby

**Ochrana:**
- Odemkne list "Plan" na začátku (heslo: `MrkevNeniOvoce123`)
- Zamkne list "Plan" na konci s `AllowFiltering:=True`

**Příklad:**
```vba
' Ruční aktualizace
Call AktualizovatSeznamZakazek
```

**Výstup:**
- List "Plan" s aktualizovaným seznamem zakázek

**Performance:**
- Screen updating OFF
- Calculation MANUAL
- Typický čas: 5-15 sekund

---

#### `BubbleSortZakazky`
```vba
Sub BubbleSortZakazky(arr As Variant)
```
**Popis:** Seřadí pole zakázek podle termínu expedice (vzestupně).

**Parametry:**
- `arr` (Variant) - Pole zakázek k seřazení
  - Každý prvek je pole: `[číslo zakázky, firma, datum expedice]`

**Logika řazení:**
- Prázdné nebo NULL datumy jsou posunuty na konec
- Ostatní jsou seřazeny vzestupně podle data

**Algoritmus:**
```
Pro i od 0 do N-1:
  Pro j od i+1 do N:
    val1 = arr[i][2]  (datum expedice)
    val2 = arr[j][2]

    IF val1 je prázdné THEN
      SKIP (ponechej na konci)
    ELSE IF val2 je prázdné THEN
      SWAP arr[i], arr[j]
    ELSE IF val1 > val2 THEN
      SWAP arr[i], arr[j]
```

**Složitost:** O(n²)

**Příklad:**
```vba
Dim zakazky() As Variant
zakazky = Array( _
    Array("12345", "Firma A", #12/31/2024#), _
    Array("12346", "Firma B", #12/15/2024#), _
    Array("12347", "Firma C", Null) _
)

BubbleSortZakazky zakazky
' Výsledek: 12346, 12345, 12347
```

---

#### `NaplnitKontrolu`
```vba
Sub NaplnitKontrolu()
```
**Popis:** Vytvoří kontrolní list pro přehled změn termínů před synchronizací do Helios.

**Funkce:**
1. Smaže obsah listu "Kontrola" (ponechá hlavičku)
2. Projde všechny zakázky v listu "Plan"
3. Pro každý úsek (Příprava, Pila, Svařování, Montáž, Elektro, Balení):
   - Pokud existuje termín výroby do (plánovaný)
   - Zapíše řádek do "Kontrola"
4. Vypočítá ISO týdny pro původní a nový termín
5. Načte potřebný čas z listu "Operace" (SUMIFS)

**Struktura výstupu (list "Kontrola"):**
| Sloupec | Název | Popis |
|---------|-------|-------|
| A | Usek | ID úseku (1, 2, 3, 4, 5, 8) |
| B | Zakazka | Číslo zakázky |
| C | PuvodniTermin | Původní termín z Helios |
| D | TerminVyrobyDo | Nový plánovaný termín |
| E | PotrebnyCasCelkemHod | Součet hodin z "Operace" |
| F | TydenPuvodni | ISO week number původního termínu |
| G | TydenNovy | ISO week number nového termínu |

**Mapování sloupců a úseků:**
```vba
columnCodes = Array("M", "P", "S", "V", "Y", "AB")
usekValues = Array(1, 4, 2, 3, 5, 8)
```

**Prerekvizity:**
- List "Plan" obsahuje zakázky s termíny
- List "Operace" obsahuje data (pro SUMIFS)
- List "Kontrola" existuje s hlavičkou

**Příklad:**
```vba
Call NaplnitKontrolu

' Otevře list "Kontrola" s přehledem změn
' Uživatel zkontroluje a pak může zavolat Helios.AktualizovatData()
```

**Výstup:**
- List "Kontrola" je viditelný a aktivní
- Obsahuje řádky se změnami termínů

---

### Klávesové zkratky

#### `PlusTyden`
```vba
Sub PlusTyden()
Attribute PlusTyden.VB_ProcData.VB_Invoke_Func = "n\n14"
```
**Popis:** Přidá 7 dní k termínu v aktivní buňce.

**Zkratka:** `Ctrl+N`

**Parametry:** Žádné (použije ActiveCell)

**Funkce:**
- Kontrola, zda je aktivní buňka ve sloupcích M, P, S, V, Y (13, 16, 19, 22, 25)
- Kontrola, zda je řádek >= 13
- Kontrola, zda levá buňka obsahuje datum (původní termín)
- Přidá 7 dní k:
  - Aktivní buňce, pokud už obsahuje datum a rozdíl je násobek 7
  - Nebo k datu v levé buňce

**Logika:**
```vba
IF sloupec IN [13, 16, 19, 22, 25] AND řádek >= 13 THEN
  levá_buňka = ActiveCell.Offset(0, -1)

  IF IsDate(levá_buňka) THEN
    IF NOT IsEmpty(ActiveCell) AND IsDate(ActiveCell) AND
       ((ActiveCell - levá_buňka) MOD 7 = 0) THEN
      ActiveCell = ActiveCell + 7 dní
    ELSE
      ActiveCell = levá_buňka + 7 dní
    END IF
  END IF
END IF
```

**Použití:**
1. Klikněte na buňku s plánovaným termínem (M, P, S, V, Y)
2. Stiskněte `Ctrl+N`
3. Termín se posune o týden dopředu

**Příklad:**
```
Původní (L): 01.12.24
Plánovaný (M): 08.12.24   <- Aktivní buňka
Stisknout Ctrl+N
Plánovaný (M): 15.12.24   <- Nová hodnota
```

---

#### `MinusTyden`
```vba
Sub MinusTyden()
Attribute MinusTyden.VB_ProcData.VB_Invoke_Func = "d\n14"
```
**Popis:** Odebere 7 dní od termínu v aktivní buňce.

**Zkratka:** `Ctrl+D`

**Funkce:** Stejná jako `PlusTyden`, ale odebírá místo přidávání.

**Použití:**
1. Klikněte na buňku s plánovaným termínem
2. Stiskněte `Ctrl+D`
3. Termín se posune o týden zpět

---

### UI Helper procedury

#### `ObsazenostUseku`
```vba
Sub ObsazenostUseku()
```
**Popis:** Zobrazí list "Obsazenost úseků".

**Funkce:**
- Nastaví `Visible = True`
- Aktivuje list

**Příklad:**
```vba
Call ObsazenostUseku
```

---

#### `ObsazenostPracovist`
```vba
Sub ObsazenostPracovist()
```
**Popis:** Zobrazí list "Obsazenost pracovišť".

**Funkce:** Stejná jako `ObsazenostUseku`

---

#### `Konfigurace`
```vba
Sub Konfigurace()
```
**Popis:** Zobrazí list "Konfigurace" s nastavením připojení.

---

#### `Ribbon`
```vba
Sub Ribbon()
Attribute Ribbon.VB_ProcData.VB_Invoke_Func = "l\n14"
```
**Popis:** Zobrazí Excel Ribbon.

**Zkratka:** `Ctrl+L`

**Funkce:**
- Volá Excel 4 Macro: `SHOW.TOOLBAR("Ribbon", True)`

---

#### `Ladeni`
```vba
Sub Ladeni()
Attribute Ladeni.VB_ProcData.VB_Invoke_Func = "l\n14"
```
**Popis:** Toggle Ribbon, Formula Bar a Status Bar (debug režim).

**Zkratka:** `Ctrl+L`

**Funkce:**
- Přepíná viditelnost Ribbon
- Přepíná `DisplayFormulaBar`
- Přepíná `DisplayStatusBar`

**Static proměnná:** `ribbonHidden` uchov Keterá stav

---

### Font Management

#### `IsFontInstalled`
```vba
Function IsFontInstalled(ByVal FontName As String) As Boolean
```
**Popis:** Ověří, zda je font nainstalován v systému pomocí Win32 API.

**Parametry:**
- `FontName` (String) - Název fontu (např. "Segoe UI Light")

**Návratová hodnota:**
- `True` - Font je nainstalován
- `False` - Font není nainstalován

**Implementace:**
- Používá `EnumFontFamiliesEx` z `gdi32.dll`
- Callback: `EnumFontsProc`
- Globální proměnné: `m_targetFont`, `m_fontFound`

**Příklad:**
```vba
If IsFontInstalled("Segoe UI Light") Then
    MsgBox "Font je k dispozici"
Else
    MsgBox "Font není nainstalován"
End If
```

**⚠️ Poznámka:** Vyžaduje Win32 API - funguje pouze na Windows.

---

#### `SupportsCzech`
```vba
Function SupportsCzech(ByVal FontName As String) As Boolean
```
**Popis:** Kontroluje, zda font podporuje českou diakritiku.

**Parametry:**
- `FontName` (String) - Název fontu

**Návratová hodnota:**
- `True` - Font podporuje češtinu
- `False` - Font nepodporuje češtinu

**Podporované fonty:**
- Segoe UI Light
- Segoe UI Semilight
- Calibri Light
- Arial Narrow
- Arial

**Příklad:**
```vba
If SupportsCzech("Segoe UI Light") Then
    Debug.Print "Podporuje češtinu"
End If
```

**⚠️ Poznámka:** Hardcoded seznam - pro jiné fonty vrací `False`.

---

#### `NastavFontPlusFormatDatumu`
```vba
Sub NastavFontPlusFormatDatumu()
```
**Popis:** Detekuje a aplikuje nejlepší dostupný font, nastaví formát datumu a přizpůsobí šířku sloupců.

**Funkce:**
1. Odemkne list "Plan"
2. Projde seznam oblíbených fontů:
   - Segoe UI Semilight
   - Segoe UI Light
   - Calibri Light
   - Arial Narrow
   - Arial
3. Pro každý font:
   - Zkontroluje instalaci (`IsFontInstalled`)
   - Zkontroluje podporu češtiny (`SupportsCzech`)
   - První vyhovující font je vybrán
4. Aplikuje font na celý list
5. Nastaví velikost: 10
6. Nastaví formát datumu: `dd.mm.yy` pro sloupce H:AB
7. AutoFit sloupce H:AA
8. Nastaví jednotnou šířku H:AB podle nejširšího sloupce + 1
9. Sloupec J = šířka 1
10. Formát čísel K5:Y11 = `0`
11. Zobrazí frmReady
12. Zamkne list "Plan"

**Progress feedback:**
- Zobrazuje `frmProgress` s textem a progress barem
- Aktualizace při každém kroku

**Prerekvizity:**
- frmProgress je načtený (Load frmProgress)

**Výstup:**
- List "Plan" s optimálním fontem a formátováním

**Příklad:**
```vba
' Automaticky volána po AktualizovatSeznamZakazek
Call NastavFontPlusFormatDatumu
```

**Čas trvání:** ~3-5 sekund (závisí na počtu fontů)

---

#### `Delay`
```vba
Sub Delay(Seconds As Single)
```
**Popis:** Čeká zadaný počet sekund s `DoEvents`.

**Parametry:**
- `Seconds` (Single) - Počet sekund (může být desetinné číslo)

**Funkce:**
- Použije `Timer` pro měření času
- Volá `DoEvents` v loop

**Příklad:**
```vba
Delay 0.5  ' Čeká 0.5 sekundy
Delay 2    ' Čeká 2 sekundy
```

**⚠️ Poznámka:** Blokuje thread - pro produkci zvážit `Application.OnTime`.

---

#### `ToggleLabel`
```vba
Public Sub ToggleLabel()
```
**Popis:** Přepíná barvu labelu ve frmReady (blikání).

**Funkce:**
- Přepíná `Label2.ForeColor` mezi červenou a černou
- Naplánuje další volání pomocí `Application.OnTime`

**Rekurze:** Ano - volá sebe sama po 1 sekundě

**Stop podmínka:** frmReady není viditelný

**Použití:** Automaticky spuštěno při zobrazení frmReady

---

#### `ZobrazFormularZpozdene`
```vba
Public Sub ZobrazFormularZpozdene(nazevFormulare As String, Optional sekundyZpozdeni As Double = 1)
```
**Popis:** Naplánuje zobrazení formuláře po zadaném zpoždění.

**Parametry:**
- `nazevFormulare` (String) - Název formuláře (např. "frmReady")
- `sekundyZpozdeni` (Double) - Zpoždění v sekundách (default: 1)

**Funkce:**
- Používá `Application.OnTime`
- Volá `ZobrazFormular` po uplynutí času

**Příklad:**
```vba
ZobrazFormularZpozdene "frmReady"        ' 1 sekunda
ZobrazFormularZpozdene "frmZakazka", 0.5 ' 0.5 sekundy
```

---

#### `ZobrazFormular`
```vba
Public Sub ZobrazFormular(formName As String)
```
**Popis:** Zobrazí zadaný formulář (voláno z `ZobrazFormularZpozdene`).

**Parametry:**
- `formName` (String) - Název formuláře

**Podporované formuláře:**
- "frmReady"

**Funkce:**
- Kontroluje, zda formulář už není viditelný
- Zobrazí modálně: `frmReady.Show vbModal`

---

### Utility funkce

#### `IsInArray`
```vba
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
```
**Popis:** Kontroluje, zda je hodnota v poli.

**Stejná jako v Gantt** - viz [Gantt/API.md](../Gantt/API.md#isinarray)

---

## VypocetPlanu.bas

Modul pro výpočet termínů výroby pomocí stored procedure.

### `OtevritFormularZakazky`
```vba
Sub OtevritFormularZakazky()
```
**Popis:** Otevře formulář pro detail zakázky.

**Funkce:**
1. Kontroluje, zda je vybrána pouze jedna buňka
2. Získá číslo zakázky ze sloupce B
3. Předá hodnotu do `frmZakazka.txtZakazka`
4. Zobrazí formulář se zpožděním

**Chybové stavy:**
- MsgBox pokud je vybráno více řádků
- MsgBox pokud číslo zakázky chybí

**Příklad:**
```vba
' Automaticky volána z UI (tlačítko, pravé click menu)
Call OtevritFormularZakazky
```

---

### `CallPlanTerminyVyrobyDoZakazek`
```vba
Sub CallPlanTerminyVyrobyDoZakazek(zakazkaID As Long, Optional datumUkonceni As Variant)
```
**Popis:** Volá stored proceduru pro výpočet termínů výroby.

**Parametry:**
- `zakazkaID` (Long) - ID zakázky z TabZakazka
- `datumUkonceni` (Variant, Optional) - Datum expedice nebo Empty/NULL

**Stored Procedure:** `dbo.EP_PlanTerminyVyrobyDoZakazek`

**SP Parametry:**
- `@ID` (INT) - ID zakázky
- `@DatumUkonceni` (DATE/NULL) - Datum expedice

**Funkce:**
1. Vytvoří ADODB.Connection
2. Vytvoří ADODB.Command
3. Nastaví CommandType = adCmdStoredProc
4. Přidá parametr @ID
5. Pokud je datum zadáno:
   - Validuje datum
   - Formátuje na "YYYY-MM-DD"
   - Přidá parametr @DatumUkonceni
6. Jinak: Přidá @DatumUkonceni = NULL
7. Spustí `cmd.Execute`
8. Loguje do Debug.Print

**Chybové stavy:**
- MsgBox pokud datum není platné
- MsgBox pokud execute selže

**Příklad:**
```vba
Dim ID As Long
ID = GetZakazkaID("12345")

' S datem expedice
CallPlanTerminyVyrobyDoZakazek ID, #12/31/2024#

' Bez data (SP použije výchozí)
CallPlanTerminyVyrobyDoZakazek ID
```

**Debug výstup:**
```
EXEC dbo.EP_PlanTerminyVyrobyDoZakazek @ID = 12345, @DatumUkonceni = '2024-12-31'
```

**⚠️ Poznámka:** SP musí existovat v databázi a uživatel musí mít EXECUTE oprávnění.

---

### `GetZakazkaID`
```vba
Function GetZakazkaID(cisloZakazky As String) As Long
```
**Popis:** Získá ID zakázky podle čísla zakázky.

**Parametry:**
- `cisloZakazky` (String) - Číslo zakázky (např. "12345")

**Návratová hodnota:**
- (Long) ID zakázky nebo 0 pokud nenalezena

**SQL:**
```sql
SELECT ID FROM TabZakazka WHERE CisloZakazky = '...'
```

**Chybové stavy:**
- MsgBox pokud SQL selže
- Vrací 0 pokud zakázka neexistuje

**Příklad:**
```vba
Dim ID As Long
ID = GetZakazkaID("12345")

If ID = 0 Then
    MsgBox "Zakázka nenalezena"
Else
    Debug.Print "ID: " & ID
End If
```

---

## PracnostZakazek.bas

Modul pro evidenci pracnosti (hodin) zakázek.

### `OtevritFormularHodiny`
```vba
Sub OtevritFormularHodiny()
```
**Popis:** Otevře formulář pro evidenci hodin.

**Funkce:**
1. Kontroluje výběr jedné buňky
2. Získá číslo zakázky ze sloupce B
3. Zavolá `GetZakazkaID` (z VypocetPlanu.bas)
4. Předá číslo do `frmHodiny.txtCisloZakazky`
5. Zobrazí formulář se zpožděním

**Chybové stavy:** Stejné jako `OtevritFormularZakazky`

---

### `UlozitHodiny`
```vba
Sub UlozitHodiny(zakazkaID As Long, HodCelkem As Long, HodSkPrac1 As Long, _
                 HodSkPrac2 As Long, HodSkPrac3 As Long, HodSkPrac4 As Long, _
                 HodSkPrac5 As Long, HodKoop As Long)
```
**Popis:** Uloží plánované hodiny do databáze.

**Parametry:**
- `zakazkaID` (Long) - ID zakázky
- `HodCelkem` (Long) - Celkový počet hodin
- `HodSkPrac1-5` (Long) - Hodiny pro skupiny pracovníků 1-5
- `HodKoop` (Long) - Hodiny kooperací

**Tabulka:** `TabZakazka_EXT`

**Funkce:**
1. Vytvoří ADODB.Connection
2. SQL INSERT: Zajistí existenci záznamu
   ```sql
   IF NOT EXISTS (SELECT 1 FROM TabZakazka_EXT WHERE ID = ...)
     INSERT INTO TabZakazka_EXT (ID) VALUES (...)
   ```
3. SQL UPDATE: Aktualizuje hodnoty
   ```sql
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
4. Zobrazí MsgBox "Hodiny byly úspěšně uloženy."

**Chybové stavy:**
- MsgBox pokud SQL selže

**Příklad:**
```vba
Dim ID As Long
ID = GetZakazkaID("12345")

UlozitHodiny ID, 100, 20, 30, 20, 15, 15, 0
' Celkem: 100, Skup1: 20, Skup2: 30, ..., Koop: 0
```

---

## Helios.bas

Modul pro synchronizaci dat s Helios ERP systémem.

### `AktualizovatData`
```vba
Sub AktualizovatData()
```
**Popis:** Aktualizuje termíny v Helios na základě listu "Kontrola".

**Funkce:**
1. Vytvoří ADODB.Connection
2. Projde všechny řádky v listu "Kontrola" (od řádku 2)
3. Pro každý řádek:
   a. Načte: usek, zakazka, terminDo, dateDiff
   b. UPDATE `_U{usek}Konec` = terminDo
   c. `dateDiff = GetDateDiffAndID(...)`
   d. UPDATE `_U{usek}Start` = terminDo - dateDiff
   e. `CallStoredProcedure(ID)` - přepočítá operace
   f. Čeká 1 sekundu (`Application.Wait`)
4. Obnoví všechna datová připojení (`RefreshAll`)
5. MsgBox "Nastavil jsem termíny zakázek v Heliosu."

**Prerekvizity:**
- List "Kontrola" obsahuje změny (vytvořený přes `NaplnitKontrolu`)

**SQL UPDATE pro konec:**
```sql
UPDATE TabZakazka_EXT
SET _U{usek}Konec = '20241231'
FROM TabZakazka_EXT AS TZE
JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID
WHERE TZ.CisloZakazky = '...'
```

**SQL UPDATE pro začátek:**
```sql
UPDATE TabZakazka_EXT
SET _U{usek}Start = DATEADD(day, -{dateDiff}, CONVERT(datetime, '20241231', 102))
FROM TabZakazka_EXT AS TZE
JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID
WHERE TZ.CisloZakazky = '...'
```

**Příklad:**
```vba
' 1. Uživatel upraví termíny v listu "Plan"
' 2. Spustí NaplnitKontrolu
Call NaplnitKontrolu

' 3. Zkontroluje změny v listu "Kontrola"
' 4. Potvrdí synchronizaci
Call AktualizovatData
```

**⚠️ Poznámka:** Obsahuje `Application.Wait` 1 sekunda mezi iteracemi - pro velké množství zakázek může trvat dlouho!

---

### `GetDateDiffAndID`
```vba
Function GetDateDiffAndID(ws As Worksheet, usek As String, zakazka As String, ByRef ID As Long) As Long
```
**Popis:** Vypočítá rozdíl ve dnech mezi začátkem a koncem úseku.

**Parametry:**
- `ws` (Worksheet) - List "Kontrola" (nepoužito, legacy parametr)
- `usek` (String) - ID úseku (např. "1", "2", "3")
- `zakazka` (String) - Číslo zakázky
- `ID` (Long, ByRef) - Výstupní parametr pro ID zakázky

**Návratová hodnota:**
- (Long) Počet dnů mezi start a konec nebo 0 pokud nenalezeno

**SQL:**
```sql
SELECT TZ.ID, DATEDIFF(day, _U{usek}Start, _U{usek}Konec) as DateDiff
FROM TabZakazka_EXT AS TZE
JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID
WHERE TZ.CisloZakazky = '...'
```

**Příklad:**
```vba
Dim ID As Long
Dim dateDiff As Long

dateDiff = GetDateDiffAndID(Nothing, "1", "12345", ID)
' dateDiff = 7 (např. 7 dní mezi začátkem a koncem Přípravy)
' ID = 12345 (ID zakázky)
```

---

### `CallStoredProcedure`
```vba
Sub CallStoredProcedure(ID As Long)
```
**Popis:** Volá stored proceduru pro doplnění termínů operací.

**Parametry:**
- `ID` (Long) - ID zakázky

**Stored Procedure:** `dbo.ep_DoplnTerminOperacePodleUseku`

**SP Parametry:**
- `@ID` (INT) - ID zakázky

**Funkce:**
1. Vytvoří ADODB.Connection
2. Vytvoří ADODB.Command
3. Nastaví CommandType = adCmdStoredProc
4. Přidá parametr @ID
5. Spustí `cmd.Execute`
6. Skryje list "Kontrola"
7. Aktivuje list "Plan"

**Příklad:**
```vba
CallStoredProcedure 12345
```

**⚠️ Poznámka:** Automaticky voláno z `AktualizovatData`, není nutné volat ručně.

---

## UserForms

### frmLogin
**Stejný jako v Gantt** - viz [Gantt/API.md - frmLogin](../Gantt/API.md#frmlogin)

---

### frmProgress
**Stejný jako v Gantt** - viz [Gantt/API.md - frmProgress](../Gantt/API.md#frmprogress)

**Rozšíření:**
- Používán více místech (font detection, formátování, atd.)
- Property: `lblProgress.Caption` - text zprávy

---

### frmZakazka
Detail zakázky pro výpočet termínů.

**Properties:**
- `txtZakazka` (TextBox) - Číslo zakázky
- `txtDatumUkonceni` (TextBox) - Datum expedice

**Metody:**
- `btnVypocitat_Click()` - Zavolá `CallPlanTerminyVyrobyDoZakazek`

**Workflow:**
1. Otevře se přes `OtevritFormularZakazky`
2. Uživatel upraví datum expedice
3. Klikne "Vypočítat"
4. SP přepočítá termíny
5. Formulář se zavře
6. Data se obnoví v listu "Plan"

---

### frmHodiny
Evidence pracnosti zakázky.

**Properties:**
- `txtCisloZakazky` (TextBox) - Číslo zakázky
- `txtHodCelkem` (TextBox) - Celkové hodiny
- `txtHodSkPrac1-5` (TextBox) - Hodiny skupin pracovníků
- `txtHodKoop` (TextBox) - Hodiny kooperací

**Metody:**
- `btnUlozit_Click()` - Zavolá `UlozitHodiny`

**Workflow:**
1. Otevře se přes `OtevritFormularHodiny`
2. Uživatel vyplní hodiny
3. Klikne "Uložit"
4. Data se uloží do `TabZakazka_EXT`
5. Formulář se zavře

---

### frmChangelog
Zobrazení changelogu verzí.

**Funkce:**
- Zobrazí seznam změn mezi verzemi
- Modální formulář

---

### frmReady
Oznámení o dokončení načítání dat.

**Properties:**
- `Label2` - Blikající label "Hotovo!"

**Funkce:**
- Zobrazí se po dokončení `NastavFontPlusFormatDatumu`
- `ToggleLabel` způsobuje blikání
- Uživatel klikne "OK" pro zavření

---

## Konstanty a mapování

### Výrobní úseky

```vba
' IDs
Dim usekValues As Variant
usekValues = Array(1, 4, 2, 3, 5, 8)
' Příprava, Pila, Svařování, Montáž, Elektro, Balení

' Sloupce plánovaných termínů
Dim columnCodes As Variant
columnCodes = Array("M", "P", "S", "V", "Y", "AB")

' Sloupce pro klávesové zkratky
Dim columnCodes As Variant
columnCodes = Array(13, 16, 19, 22, 25)
' M, P, S, V, Y
```

### Ochrana listu

```vba
Const PASSWORD As String = "MrkevNeniOvoce123"

ws.Protect password:=PASSWORD, AllowFiltering:=True
ws.Unprotect password:=PASSWORD
```

### Oblíbené fonty

```vba
favFonts = Array("Segoe UI Semilight", "Segoe UI Light", _
                 "Calibri Light", "Arial Narrow", "Arial")
```

---

## Win32 API Deklarace

### VBA7 (64-bit Excel)

```vba
Private Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32" _
    Alias "EnumFontFamiliesExA" ( _
    ByVal hdc As LongPtr, _
    lpLogFont As LOGFONT, _
    ByVal lpEnumFontFamExProc As LongPtr, _
    ByVal lParam As LongPtr, _
    ByVal dwFlags As Long _
) As Long

Private Declare PtrSafe Function GetDC Lib "user32" ( _
    ByVal hwnd As LongPtr _
) As LongPtr

Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal hdc As LongPtr _
) As Long

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    Destination As Any, _
    Source As Any, _
    ByVal Length As LongPtr _
)
```

### LOGFONT Structure

```vba
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type
```

### Callback funkce

```vba
Function EnumFontsProc(ByVal lpelfe As LongPtr, ByVal lpntme As LongPtr, _
                       ByVal FontType As Long, ByVal lParam As LongPtr) As Long
    ' Vrací 0 pro zastavení enumerace, 1 pro pokračování
End Function
```

---

## Changelog API

**Verze:** 1.0 (2026-01-16)

**Změny:**
- Iniciální dokumentace
- Win32 API pro font detection
- Integrace s Helios ERP
- Klávesové zkratky Ctrl+N, Ctrl+D

---

## Poznámky pro vývojáře

### Přidání nového úseku

1. Databáze: Přidat `_U{ID}Start`, `_U{ID}Konec` do `TabZakazka_EXT`
2. `Main.bas NaplnitKontrolu`: Přidat do `usekValues` a `columnCodes`
3. `Main.bas PlusTyden/MinusTyden`: Přidat sloupec do array
4. Excel: Přidat sloupce do listu "Plan"

### Změna hesla ochrany

```vba
' Main.bas, hledat všechny výskyty:
"MrkevNeniOvoce123"

' Nahradit novým heslem
```

**⚠️ Bezpečnost:** Heslo v kódu je bezpečnostní riziko!

### Logování

```vba
' Přidat logování
WriteLog "Moje zpráva"

' Implementace WriteLog není v poskytnutém kódu
' Pravděpodobně v skrytém modulu nebo external library
```

---

## Revize

**Autor:** IN-EKO VBA Development Team
**Datum:** 2026-01-16
**Verze:** 1.0
