# Architektura - Rozpočet

## Přehled architektury

Rozpočet je založen na podobné architektuře jako Gantt a Plánování, ale s důrazem na vizualizaci dat pomocí pivotových tabulek a grafů.

```
┌─────────────────────────────────────────────────┐
│          Prezentační vrstva                     │
│  ┌──────────┐  ┌───────────┐  ┌─────────────┐  │
│  │ Ikony    │  │ Forms     │  │ Listy       │  │
│  │ Navigace │  │ - Login   │  │ - Aplikace  │  │
│  │          │  │ - Progress│  │ - Kumulace  │  │
│  │          │  │ - Grafy   │  │ - Pivot     │  │
│  └──────────┘  └───────────┘  └─────────────┘  │
└─────────────────────────────────────────────────┘
                    ↕
┌─────────────────────────────────────────────────┐
│         Business logika                         │
│  ┌──────────┐  ┌───────┐  ┌──────┐  ┌───────┐ │
│  │Navigace  │  │Copy   │  │Grafy │  │Rutiny │ │
│  │Data      │  │       │  │      │  │       │ │
│  └──────────┘  └───────┘  └──────┘  └───────┘ │
└─────────────────────────────────────────────────┘
                    ↕ ADODB + Queries
┌─────────────────────────────────────────────────┐
│         Datová vrstva                           │
│         SQL Server - IN-EKO ERP DB              │
└─────────────────────────────────────────────────┘
```

## Komponenty systému

### 1. Connection.bas
**Stejný jako Gantt** - viz [Gantt/ARCHITECTURE.md](../Gantt/ARCHITECTURE.md#1-connectionbas---správa-připojení)

---

### 2. Data.bas
**Podobný jako Gantt** - načítání dat z databáze pomocí queries nebo view.

---

### 3. Navigace.bas - Správa navigace

**Zodpovědnost:**
- Přiřazování maker ikonkám
- Navigace mezi listy
- Vizuální feedback (barvy ikonek)
- Ovládání pivotových tabulek

**Klíčové funkce:**

**`AssignMacroToShapes(targetSheet)`**
- Dynamicky přiřadí makra ikonkám na základě názvu ikony a listu
- Nastaví barvu ikony (aktivní/neaktivní)

**Algoritmus:**
```
Najdi skupinu "navigace" na listu

Pro každou ikonu ve skupině:
  IF název začíná "ico_" THEN
    zbytek_názvu = název bez "ico_"

    SELECT CASE zbytek_názvu:
      CASE "aplikace":
        IF list = "Aplikace" OR "Kumulace" OR "Kontingentní tabulka"
          OnAction = "NavigateSheet"
          Barva = ICO_ENABLE_COLOR
        ELSE
          Barva = ICO_DISABLE_COLOR
        END IF

      CASE "load":
        IF list = "Aplikace"
          OnAction = "LoadDataFromQueries"
          Barva = ICO_ENABLE_COLOR
        ELSE
          OnAction = "TotoMakroNicNedela"
          Barva = ICO_DISABLE_COLOR
        END IF

      ' ... další ikony
    END SELECT
  END IF
```

**Ovládání pivotových tabulek:**
```vba
Sub Plus()
    ActiveSheet.PivotTables("Rozpočet").PivotFields( _
        "[Rozpočet].[Skupina].[Skupina]").DrilledDown = True
End Sub

Sub Minus()
    ActiveSheet.PivotTables("Rozpočet").PivotFields( _
        "[Rozpočet].[Skupina].[Skupina]").DrilledDown = False
End Sub
```

---

### 4. Copy.bas - Export dat

**Zodpovědnost:**
- Kopírování dat do nového sešitu
- Převod vzorců na hodnoty
- Zachování formátování

**`KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty()`**

**Algoritmus:**
```
1. Najdi poslední řádek v "Aplikace" (sloupce B:AQ)
2. Najdi poslední řádek v "Kumulace" (sloupce B:AQ)

3. Vytvoř nový sešit
4. Přejmenuj první list na "Aplikace"

5. Zkopíruj šířky sloupců B:AQ z "Aplikace"
6. Zkopíruj data B4:AQ{N} z "Aplikace"
7. Převeď vzorce na hodnoty

8. Přidej nový list "Kumulace"
9. Zkopíruj šířky sloupců B:AQ z "Kumulace"
10. Zkopíruj data B4:AQ{N} z "Kumulace"
11. Převeď vzorce na hodnoty

12. Smaž přebytečné listy
13. Zobraz dialog "Uložit jako"
14. Výchozí název: "Kopie rozpočtu ze dne YYYYMMDD.xlsx"
```

**Optimalizace:**
- Používá `.Value = .Value` pro převod vzorců
- Vypne DisplayAlerts při mazání listů
- Kopíruje šířky sloupců najednou

---

### 5. Grafy.bas - Správa grafů

**Zodpovědnost:**
- Inicializace barev grafů
- Výběr vlastních barev
- Aplikace barev na grafy
- Export grafů jako obrázky

**Barvy:**
```vba
Public BarvyGrafu(1 To 3) As Long

BarvyGrafu(1) = Konfigurace!C7  ' Hlavní (default: RGB 35, 176, 160)
BarvyGrafu(2) = Konfigurace!C8  ' Doplňková (default: RGB 209, 209, 209)
```

**`VyberBarvyGrafu()`**
- Otevře dialog pro výběr barev
- Uloží do listu "Konfigurace"
- Aplikuje na všechny grafy

**`NastavBarvyGrafu()`**
- Načte barvy z konfigurace
- Aplikuje na:
  - GrafKategorie (3 datové řady)
  - GrafKategorieKumulativni (3 datové řady)
  - Graphic 8 na listu "Kumulace"

**Struktura grafů:**
```
GrafKategorie:
  Series 1: Doplňková barva (fill)
  Series 2: Hlavní barva (fill)
  Series 3: Doplňková barva (fill) + Hlavní barva (line)

GrafKategorieKumulativni:
  (stejné jako GrafKategorie)
```

---

### 6. Rutiny.bas - Utility funkce

**Zodpovědnost:**
- Ladění (Ctrl+L)
- Zamykání/odemykání listů
- Toggle screen updating
- Dočasná změna buňky (pro refresh)
- Changelog

**`LockSpecificSheets(sh)`**
```vba
sh.Protect UserInterfaceOnly:=True, _
           DrawingObjects:=True, _
           Contents:=True, _
           AllowUsingPivotTables:=True, _
           AllowFormattingColumns:=True

sh.EnableSelection = xlNoRestrictions
```

**`DocasnaZmenaBunky()`**
- Změní B2 na listu "Rozpočet" na 13 (neexistující měsíc)
- Vrátí původní hodnotu
- **Účel:** Vynucení refreshu po importu (workaround pro kompilaci bug)

**`ToggleScreenUpdating(turnOff)`**
```vba
IF turnOff THEN
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual
ELSE
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic
END IF
```

---

## Datový model

### Excel listy

**List "Aplikace"**
- Rozsah dat: B4:AQ{N}
- Obsahuje vzorce
- Hlavní pracovní list

**List "Kumulace"**
- Rozsah dat: B4:AQ{N}
- Kumulativní data
- Shape "Graphic 8" - grafický prvek s barvou

**List "Kontingentní tabulka"**
- Pivotová tabulka "Rozpočet"
- Pole: `[Rozpočet].[Skupina].[Skupina]`

**List "Grafy"**
- ChartObject "GrafKategorie"
- ChartObject "GrafKategorieKumulativni"

**List "Rozpočet"**
- B2: Měsíc (1-12)
- B5: Hyperlink na changelog

**List "Konfigurace"**
- serverName, databaseName (připojení)
- C7: Hlavní barva grafu (Long)
- C8: Doplňková barva grafu (Long)

---

### Navigační struktura

```
Skupina "navigace" (Shape Group)
  ├── ico_aplikace (přepnout na Aplikace)
  ├── ico_load (načíst data)
  ├── ico_kontingencni (otevřít pivot)
  ├── ico_grafy (otevřít grafy)
  ├── ico_copy (exportovat)
  └── ... (další ikony)
```

**Barvy ikonek:**
```vba
ICO_DISABLE_COLOR = RGB(250, 250, 250)  ' Světle šedá
ICO_ENABLE_COLOR = RGB(134, 134, 134)   ' Středně šedá
```

---

## Diagram toku dat

### Export dat

```
Uživatel → Ctrl+K / Ikona "Copy"
    ↓
KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty()
    ↓
1. Najdi poslední řádky v "Aplikace" a "Kumulace"
    ↓
2. Vytvoř nový sešit
    ↓
3. Zkopíruj "Aplikace" (šířky + data)
    ↓
4. Převeď vzorce → hodnoty
    ↓
5. Zkopíruj "Kumulace" (šířky + data)
    ↓
6. Převeď vzorce → hodnoty
    ↓
7. Smaž přebytečné listy
    ↓
8. Dialog "Uložit jako"
    ↓
9. Uložit jako XLSX (bez maker)
```

### Aplikace barev na grafy

```
VyberBarvyGrafu()
    ↓
Dialog pro výběr barvy 1 (hlavní)
    ↓
Dialog pro výběr barvy 2 (doplňková)
    ↓
UlozBarvy() → Konfigurace!C7, C8
    ↓
NastavBarvyGrafu()
    ↓
InicializujBarvyGrafu() - načti z Konfigurace
    ↓
Aplikuj na GrafKategorie (Series 1, 2, 3)
    ↓
Aplikuj na GrafKategorieKumulativni (Series 1, 2, 3)
    ↓
Aplikuj na Graphic 8 (Kumulace)
```

---

## Bezpečnostní architektura

### Ochrana listů

**Rozdíl od Plánování:**
- **Žádné heslo** - listy jsou chráněny, ale bez hesla
- Lze snadno odemknout voláním `UnlockAllSheets`

**Parametry ochrany:**
```vba
UserInterfaceOnly:=True    ' Makra mohou upravovat
DrawingObjects:=True        ' Chránit grafické objekty
Contents:=True              ' Chránit obsah
AllowUsingPivotTables:=True ' Povolit pivoty
AllowFormattingColumns:=True ' Povolit formátování sloupců
EnableSelection = xlNoRestrictions ' Povolit výběr všeho
```

---

## Výkonnostní optimalizace

### 1. Batch operations

```vba
Call ToggleScreenUpdating(True)  ' Vypnout
' ... operace ...
Call ToggleScreenUpdating(False) ' Zapnout
```

### 2. Přímý převod vzorců

```vba
' Místo iterace řádek po řádku
Range.Value = Range.Value  ' Instant převod
```

### 3. Single copy pro šířky sloupců

```vba
ws.Range("B:AQ").Copy
wsNovy.Range("B1").PasteSpecial Paste:=xlPasteColumnWidths
```

---

## Design Patterns

### 1. Factory Pattern (lehká forma)
`AssignMacroToShapes` - vytváří různá makra podle kontextu

### 2. Strategy Pattern
Navigace - různé akce pro stejnou ikonu na různých listech

### 3. Template Method
Export - definuje workflow, ale deleguje specifické kroky

---

## Rozšiřitelnost

### Přidání nové ikony navigace

1. **Excel**: Přidat shape do skupiny "navigace" s názvem "ico_{název}"
2. **Navigace.bas**: Přidat CASE pro nový název:
```vba
Case "{název}"
    Select Case targetSheet.Name
        Case "MujList"
            shp.OnAction = "MojeFunkce"
            shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
    End Select
```
3. **Vytvořit funkci**: `Sub MojeFunkce()` v příslušném modulu

### Přidání nového grafu

1. **Excel**: Vytvořit graf na listu "Grafy"
2. **Grafy.bas**: Přidat do `NastavBarvyGrafu`:
```vba
Set graf_novy = ws.ChartObjects("MujGraf")
graf_novy.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = BarvyGrafu(1)
```

---

## Závislosti

### Knihovny
- **Microsoft ActiveX Data Objects 2.x** (ADODB)
- **Microsoft Excel Object Library**

### Databáze
- **SQL Server**
- Specifické view/queries pro rozpočtová data

### Excel
- **Minimální verze**: Excel 2010
- **Pivotové tabulky** - vyžaduje podporu OLAP nebo SQL queries

---

## Best Practices

### 1. Vždy refreshuj data před exportem
```vba
Call LoadDataFromQueries  ' nebo RefreshAll
Call KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty
```

### 2. Používej utility funkce
```vba
Call ToggleScreenUpdating(True)
' ... operace ...
Call ToggleScreenUpdating(False)
```

### 3. Kontroluj existenci listů
```vba
If SheetExists("MujList") Then
    ' ...
End If
```

---

## Verze

**Export:** 2026-01-16
**Řádků kódu:** ~2,430
**Modulů:** 11
**Formulářů:** 4
