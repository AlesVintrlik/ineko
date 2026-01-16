# API Reference - Rozpočet

Tato dokumentace popisuje veřejné funkce a procedury v aplikaci Rozpočet.

## Obsah

- [Connection.bas](#connectionbas)
- [Navigace.bas](#navigacebas)
- [Copy.bas](#copybas)
- [Grafy.bas](#grafybas)
- [Rutiny.bas](#rutinybas)

---

## Connection.bas

**Stejný jako Gantt** - viz [Gantt/API.md - Connection.bas](../Gantt/API.md#connectionbas)

---

## Navigace.bas

Modul pro navigaci mezi listy a správu ikonek.

### Globální proměnné

```vba
Public ICO_DISABLE_COLOR As Long  ' RGB(250, 250, 250) - Světle šedá
Public ICO_ENABLE_COLOR As Long   ' RGB(134, 134, 134) - Středně šedá
```

---

### `InitializeColors`
```vba
Sub InitializeColors()
```
**Popis:** Inicializuje barvy pro ikony navigace.

**Volá se:** Automaticky z `AssignMacroToShapes`

---

### `SheetExists`
```vba
Function SheetExists(sheetName As String) As Boolean
```
**Popis:** Kontroluje, zda list existuje v sešitu.

**Parametry:**
- `sheetName` (String) - Název listu

**Návratová hodnota:**
- `True` pokud list existuje
- `False` pokud neexistuje

**Příklad:**
```vba
If SheetExists("Grafy") Then
    Sheets("Grafy").Activate
End If
```

---

### `Plus`
```vba
Sub Plus()
```
**Popis:** Rozbalí skupiny v pivotové tabulce "Rozpočet".

**Funkce:**
- Nastaví `DrilledDown = True` pro pole `[Rozpočet].[Skupina].[Skupina]`

**Prerekvizity:**
- Aktivní list obsahuje pivotovou tabulku "Rozpočet"

---

### `Minus`
```vba
Sub Minus()
```
**Popis:** Sbalí skupiny v pivotové tabulce "Rozpočet".

**Funkce:**
- Nastaví `DrilledDown = False`

---

### `TotoMakroNicNedela`
```vba
Sub TotoMakroNicNedela()
```
**Popis:** Placeholder makro pro neaktivní tlačítka.

**Zobrazí:** MsgBox "Toto tlačítko není na listu {název} aktivní"

---

### `KontingencniTabulka`
```vba
Sub KontingencniTabulka()
```
**Popis:** Zobrazí a obnoví list "Kontingentní tabulka".

**Funkce:**
1. Nastaví Visible = True
2. Aktivuje list
3. Refreshne pivotovou tabulku "Rozpočet"

**Příklad:**
```vba
Call KontingencniTabulka
```

---

### `AssignMacroToShapes`
```vba
Sub AssignMacroToShapes(targetSheet As Worksheet)
```
**Popis:** Dynamicky přiřadí makra ikonkám navigace podle názvu ikony a aktivního listu.

**Parametry:**
- `targetSheet` (Worksheet) - List, na kterém přiřadit makra

**Funkce:**
1. Najde skupinu "navigace"
2. Projde všechny objekty ve skupině
3. Pro ikony začínající "ico_":
   - Určí makro podle názvu a listu
   - Nastaví OnAction
   - Nastaví barvu (aktivní/neaktivní)

**Podporované ikony:**
| Název | Listy | Makro | Popis |
|-------|-------|-------|-------|
| ico_aplikace | Aplikace, Kumulace, Kontingentní tabulka | NavigateSheet | Přepnout na hlavní pohled |
| ico_load | Aplikace | LoadDataFromQueries | Načíst data |
| ico_kontingencni | - | KontingencniTabulka | Otevřít pivot |
| ico_grafy | - | - | Otevřít grafy |
| ico_copy | - | KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty | Exportovat |

**Příklad:**
```vba
' Automaticky voláno při aktivaci listu
Private Sub Worksheet_Activate()
    Call AssignMacroToShapes(Me)
End Sub
```

---

## Copy.bas

Modul pro export dat do nového sešitu.

### `KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty`
```vba
Sub KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty()
Attribute .VB_ProcData.VB_Invoke_Func = "k\n14"
```
**Popis:** Vytvoří nový sešit s listy "Aplikace" a "Kumulace", zkopíruje data a převede vzorce na hodnoty.

**Zkratka:** `Ctrl+K`

**Funkce:**
1. Najde poslední řádek v "Aplikace" (sloupce B:AQ)
2. Najde poslední řádek v "Kumulace" (sloupce B:AQ)
3. Vytvoří nový sešit
4. Zkopíruje šířky sloupců B:AQ
5. Zkopíruje data B4:AQ{N} z obou listů
6. Převede všechny vzorce na hodnoty
7. Smaže přebytečné listy
8. Zobrazí dialog "Uložit jako"
9. Výchozí název: `Kopie rozpočtu ze dne YYYYMMDD.xlsx`

**Rozsahy:**
- **Aplikace**: B4:AQ{poslední řádek}
- **Kumulace**: B4:AQ{poslední řádek}

**Formát:** `.xlsx` (Excel Open XML bez maker)

**Chybové stavy:**
- Pokud uživatel zruší uložení, nový sešit se zavře bez uložení
- MsgBox s informací o výsledku

**Příklad:**
```vba
' Ruční volání
Call KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty

' Nebo Ctrl+K
```

**⚠️ Poznámka:** Funkce nekontroluje, zda listy existují nebo obsahují data!

---

## Grafy.bas

Modul pro správu grafů a barev.

### Globální proměnné

```vba
Public BarvyGrafu(1 To 3) As Long
```

---

### `InicializujBarvyGrafu`
```vba
Sub InicializujBarvyGrafu()
```
**Popis:** Načte barvy grafů z listu "Konfigurace".

**Funkce:**
```vba
BarvyGrafu(1) = Konfigurace!C7  ' Hlavní barva
BarvyGrafu(2) = Konfigurace!C8  ' Doplňková barva
```

**Default hodnoty:**
- C7: RGB(35, 176, 160) - Tyrkysová
- C8: RGB(209, 209, 209) - Světle šedá

---

### `VyberBarvyGrafu`
```vba
Sub VyberBarvyGrafu()
```
**Popis:** Umožní uživateli vybrat vlastní barvy pro grafy.

**Funkce:**
1. Zobrazí MsgBox s instrukcemi
2. Pro každou barvu (1-2):
   - Otevře dialog výběru barvy
   - Uloží vybranou barvu
3. Uloží barvy do Konfigurace
4. Aplikuje na všechny grafy

**Příklad:**
```vba
Call VyberBarvyGrafu
' → Dialog pro výběr hlavní barvy
' → Dialog pro výběr doplňkové barvy
' → Barvy se aplikují na grafy
```

---

### `UlozBarvy`
```vba
Sub UlozBarvy(Barva() As Long)
```
**Popis:** Uloží barvy do listu "Konfigurace".

**Parametry:**
- `Barva()` (Long Array) - Pole s barvami

**Funkce:**
```vba
Konfigurace!C7 = Barva(1)
Konfigurace!C8 = Barva(2)
```

---

### `NastavBarvyGrafu`
```vba
Sub NastavBarvyGrafu()
```
**Popis:** Aplikuje barvy na všechny grafy v sešitu.

**Funkce:**
1. Načte barvy pomocí `InicializujBarvyGrafu()`
2. Aplikuje na:
   - **GrafKategorie** (list "Grafy")
     - Series 1: Doplňková (fill)
     - Series 2: Hlavní (fill)
     - Series 3: Doplňková (fill) + Hlavní (line)
   - **GrafKategorieKumulativni** (list "Grafy")
     - Series 1-3: Stejné jako GrafKategorie
   - **Graphic 8** (list "Kumulace")
     - Fill: Hlavní barva

**Příklad:**
```vba
' Po změně barev v konfiguraci
Call NastavBarvyGrafu
```

---

### `SaveChartAsImage`
```vba
Sub SaveChartAsImage()
```
**Popis:** Uloží grafy jako PNG obrázky.

**Funkce:**
1. Nastaví popisky datových řad
2. Uloží GrafKategorie jako PNG
3. Uloží GrafKategorieKumulativni jako PNG
4. Dialog pro výběr umístění

**Prerekvizity:**
- Grafy "GrafKategorie" a "GrafKategorieKumulativni" existují na listu "Grafy"

**⚠️ Poznámka:** Kód není kompletní v poskytnutém výpisu (řádek 100 limit).

---

## Rutiny.bas

Modul s utility funkcemi.

### `Ladeni`
```vba
Sub Ladeni()
Attribute Ladeni.VB_ProcData.VB_Invoke_Func = "l\n14"
```
**Popis:** Přepíná viditelnost Ribbon, Formula Bar a Status Bar.

**Zkratka:** `Ctrl+L`

**Stejné jako v Plánování** - viz [Planovani/API.md](../Planovani/API.md#ladeni)

---

### `LockSpecificSheets`
```vba
Sub LockSpecificSheets(ByVal Sh As Object)
```
**Popis:** Zamkne zadaný list s povoleními pro pivoty a formátování.

**Parametry:**
- `Sh` (Object) - Worksheet k zamknutí

**Funkce:**
```vba
Sh.Protect UserInterfaceOnly:=True, _
           DrawingObjects:=True, _
           Contents:=True, _
           AllowUsingPivotTables:=True, _
           AllowFormattingColumns:=True

Sh.EnableSelection = xlNoRestrictions
```

**Rozdíl od Plánování:**
- **Žádné heslo** - lze snadno odemknout

**Příklad:**
```vba
Call LockSpecificSheets(Sheets("Aplikace"))
```

---

### `UnlockAllSheets`
```vba
Sub UnlockAllSheets()
```
**Popis:** Odemkne všechny zamčené listy v sešitu.

**Funkce:**
- Projde všechny listy
- Pro zamčené listy zavolá `ws.Unprotect`

**Příklad:**
```vba
Call UnlockAllSheets
' Všechny listy jsou nyní odemčené
```

---

### `ToggleScreenUpdating`
```vba
Sub ToggleScreenUpdating(turnOff As Boolean)
```
**Popis:** Zapne/vypne screen updating, events a calculation.

**Parametry:**
- `turnOff` (Boolean)
  - `True` = Vypnout (optimalizace)
  - `False` = Zapnout (normální režim)

**Funkce:**
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

**Příklad:**
```vba
Call ToggleScreenUpdating(True)  ' Vypnout pro rychlost
' ... dávkové operace ...
Call ToggleScreenUpdating(False) ' Zapnout zpět
```

---

### `DocasnaZmenaBunky`
```vba
Sub DocasnaZmenaBunky()
```
**Popis:** Dočasně změní buňku B2 na listu "Rozpočet" pro vynucení refreshu.

**Funkce:**
1. Uloží původní hodnotu B2
2. Změní B2 na 13 (neexistující měsíc)
3. Vrátí původní hodnotu

**Účel:** Workaround pro bug při kompilaci - vynutí aktualizaci dat po importu.

**Příklad:**
```vba
' Po importu dat
Call LoadData
Call DocasnaZmenaBunky  ' Vynutí refresh
```

---

### `FollowHyperlink`
```vba
Sub FollowHyperlink(ByVal Target As Range)
```
**Popis:** Zpracuje kliknutí na specifické buňky (např. changelog).

**Parametry:**
- `Target` (Range) - Kliknutá buňka

**Funkce:**
```vba
IF Target.Address = "$B$5" THEN
  frmChangelog.txtChangelog.Value = "Historie verzí..."
  frmChangelog.Show
END IF
```

**Použití:** Volat z události `Worksheet_FollowHyperlink`

---

## Konstanty a konfigurace

### Rozsahy

```vba
' Hlavní datový rozsah
Const DATA_RANGE As String = "B4:AQ"  ' Sloupce B až AQ, od řádku 4
```

### Barvy ikonek

```vba
ICO_DISABLE_COLOR = RGB(250, 250, 250)
ICO_ENABLE_COLOR = RGB(134, 134, 134)
```

### Barvy grafů (default)

```vba
' List Konfigurace
C7: RGB(35, 176, 160)   ' Hlavní - Tyrkysová
C8: RGB(209, 209, 209)  ' Doplňková - Světle šedá
```

---

## Klávesové zkratky

| Zkratka | Funkce | Popis |
|---------|--------|-------|
| `Ctrl+K` | KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty | Export dat |
| `Ctrl+L` | Ladeni | Toggle Ribbon/Formula Bar/Status Bar |

---

## UserForms

### frmLogin
**Stejný jako Gantt** - viz [Gantt/API.md - frmLogin](../Gantt/API.md#frmlogin)

### frmProgress
**Stejný jako Gantt** - viz [Gantt/API.md - frmProgress](../Gantt/API.md#frmprogress)

### frmChangelog
Zobrazení historie verzí aplikace.

**Properties:**
- `txtChangelog` (TextBox) - Text s changelogem

**Zobrazení:** Kliknutím na buňku B5 nebo ruční volání

### frmGraf
Formulář pro práci s grafy (detaily nejsou v poskytnutém kódu).

---

## Poznámky pro vývojáře

### Přidání nové ikony

1. **Excel**: Vytvořit shape s názvem `ico_{název}`
2. **Přidat do skupiny** "navigace"
3. **Navigace.bas**: Přidat CASE:
```vba
Case "{název}"
    Select Case targetSheet.Name
        Case "MujList"
            shp.OnAction = "MojeFunkce"
            shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
    End Select
```

### Změna rozsahu exportu

V `Copy.bas`:
```vba
' Změnit AQ na jiný sloupec
Set rozsahAplikace = wsAplikace.Range( _
    wsAplikace.Cells(4, "B"), _
    wsAplikace.Cells(posledniRadekAplikace, "AZ")  ' ← Změnit zde
)
```

---

## Changelog API

**Verze:** 1.0 (2026-01-16)

**Změny:**
- Iniciální dokumentace
- Export bez maker
- Vlastní barvy grafů

---

## Revize

**Autor:** IN-EKO VBA Development Team
**Datum:** 2026-01-16
