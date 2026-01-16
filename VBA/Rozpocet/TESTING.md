# Testing Guide - Rozpočet

Tento dokument popisuje postupy pro testování VBA aplikace Rozpočet.

## Testovací scénáře

### TS-R001: Přihlášení
**Stejný jako Gantt TS-001** - viz [Gantt/TESTING.md](../Gantt/TESTING.md#ts-001-první-přihlášení)

---

### TS-R002: Navigace mezi listy

**Popis:** Testování navigace pomocí ikonek.

**Kroky:**
1. Otevřít Rozpočet.xlsm
2. Přihlásit se
3. Ověřit, že jste na listu "Aplikace"
4. Kliknout na ikonu "Kontingentní tabulka"
5. Ověřit přechod na list "Kontingentní tabulka"

**Očekávaný výsledek:**
- ✅ List se přepne
- ✅ Ikony se přebarví (aktivní/neaktivní)
- ✅ Pivotová tabulka se refreshne

**Validace:**
```vba
Sub ValidateTS_R002()
    ' Test: Aktivní list je správný
    If ActiveSheet.Name = "Kontingentní tabulka" Then
        Debug.Print "✅ Navigace OK"
    Else
        Debug.Print "❌ Špatný list: " & ActiveSheet.Name
    End If

    ' Test: Ikony mají správné barvy
    Dim nav As Shape
    Set nav = Sheets("Aplikace").Shapes("navigace")

    Dim ico As Shape
    For Each ico In nav.GroupItems
        If ico.Name = "ico_kontingencni" Then
            If ico.Fill.ForeColor.RGB = RGB(134, 134, 134) Then
                Debug.Print "✅ Barva ikony OK"
            Else
                Debug.Print "❌ Špatná barva ikony"
            End If
            Exit For
        End If
    Next ico
End Sub
```

---

### TS-R003: Export dat (Ctrl+K)

**Popis:** Testování exportu do nového sešitu.

**Prerekvizity:**
- Listy "Aplikace" a "Kumulace" obsahují data
- Sloupce B:AQ mají data

**Kroky:**
1. Přejít na list "Aplikace"
2. Stisknout `Ctrl+K`
3. Počkat na vytvoření nového sešitu
4. V dialogu vybrat umístění a kliknout "Uložit"

**Očekávaný výsledek:**
- ✅ Vytvoří se nový sešit
- ✅ Obsahuje listy "Aplikace" a "Kumulace"
- ✅ Data jsou zkopírována (rozsah B4:AQ)
- ✅ Všechny vzorce jsou převedeny na hodnoty
- ✅ Šířky sloupců jsou zachovány
- ✅ Soubor je uložen jako `.xlsx` (bez maker)
- ✅ MsgBox "Nový sešit byl úspěšně vytvořen a uložen."

**Validace:**
```vba
Sub ValidateTS_R003()
    ' Musí se provést manuálně po exportu

    ' Test 1: Nový sešit má 2 listy
    If ActiveWorkbook.Worksheets.Count = 2 Then
        Debug.Print "✅ Počet listů OK"
    Else
        Debug.Print "❌ Počet listů: " & ActiveWorkbook.Worksheets.Count
    End If

    ' Test 2: První list je "Aplikace"
    If ActiveWorkbook.Sheets(1).Name = "Aplikace" Then
        Debug.Print "✅ Název listu OK"
    Else
        Debug.Print "❌ Název listu: " & ActiveWorkbook.Sheets(1).Name
    End If

    ' Test 3: Buňky obsahují hodnoty, ne vzorce
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("Aplikace")

    If ws.Range("C4").HasFormula Then
        Debug.Print "❌ Vzorce nebyly převedeny"
    Else
        Debug.Print "✅ Vzorce převedeny na hodnoty"
    End If
End Sub
```

**Test edge case - Zrušení uložení:**
1. Stisknout Ctrl+K
2. V dialogu kliknout "Zrušit"
3. Očekávaný výsledek:
   - ✅ MsgBox "Uživatel zrušil uložení nového sešitu."
   - ✅ Nový sešit se zavře bez uložení
   - ✅ Původní sešit zůstane aktivní

---

### TS-R004: Pivotová tabulka Plus/Minus

**Popis:** Testování rozbalování/sbalování skupin.

**Prerekvizity:**
- List "Kontingentní tabulka" je aktivní
- Pivotová tabulka "Rozpočet" existuje

**Kroky:**
1. Přejít na list "Kontingentní tabulka"
2. Spustit makro `Plus`
3. Ověřit, že skupiny jsou rozbalené
4. Spustit makro `Minus`
5. Ověřit, že skupiny jsou sbalené

**Očekávaný výsledek:**
- ✅ Plus rozbalí pole `[Rozpočet].[Skupina].[Skupina]`
- ✅ Minus sbalí pole

**Příklad:**
```vba
Sub TestTS_R004()
    Sheets("Kontingentní tabulka").Activate

    ' Test Plus
    Call Plus
    Dim isExpanded As Boolean
    isExpanded = ActiveSheet.PivotTables("Rozpočet").PivotFields( _
        "[Rozpočet].[Skupina].[Skupina]").DrilledDown

    If isExpanded Then
        Debug.Print "✅ Plus OK"
    Else
        Debug.Print "❌ Plus FAILED"
    End If

    ' Test Minus
    Call Minus
    isExpanded = ActiveSheet.PivotTables("Rozpočet").PivotFields( _
        "[Rozpočet].[Skupina].[Skupina]").DrilledDown

    If Not isExpanded Then
        Debug.Print "✅ Minus OK"
    Else
        Debug.Print "❌ Minus FAILED"
    End If
End Sub
```

---

### TS-R005: Změna barev grafů

**Popis:** Testování vlastního výběru barev.

**Kroky:**
1. Otevřít VBA Editor (Alt+F11)
2. Spustit makro `VyberBarvyGrafu`
3. Vybrat hlavní barvu (např. modrá)
4. Vybrat doplňkovou barvu (např. šedá)
5. Zkontrolovat grafy na listu "Grafy"

**Očekávaný výsledek:**
- ✅ Zobrazí se MsgBox "Vyber postupně hlavní a doplňkovou barvu pro graf"
- ✅ Otevře se dialog výběru barvy (2×)
- ✅ Barvy se uloží do Konfigurace!C7 a C8
- ✅ Grafy se přebarví

**Validace:**
```vba
Sub ValidateTS_R005()
    Dim ws As Worksheet
    Set ws = Sheets("Grafy")

    Dim graf As ChartObject
    Set graf = ws.ChartObjects("GrafKategorie")

    ' Test: Series 2 má hlavní barvu
    Dim expectedColor As Long
    expectedColor = Sheets("Konfigurace").Range("C7").Value

    Dim actualColor As Long
    actualColor = graf.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB

    If actualColor = expectedColor Then
        Debug.Print "✅ Barva grafu OK"
    Else
        Debug.Print "❌ Barva grafu nesedí"
        Debug.Print "Očekáváno: " & expectedColor
        Debug.Print "Aktuálně: " & actualColor
    End If
End Sub
```

---

### TS-R006: Ochrana listů

**Popis:** Testování zamykání/odemykání listů.

**Kroky:**
1. Spustit `Call LockSpecificSheets(Sheets("Aplikace"))`
2. Pokusit se smazat sloupec
3. Pokusit se použít pivotovou tabulku
4. Spustit `Call UnlockAllSheets`
5. Pokusit se smazat sloupec

**Očekávaný výsledek:**
- ✅ Po zamknutí: Sloupec nelze smazat
- ✅ Po zamknutí: Pivotová tabulka funguje
- ✅ Po zamknutí: Formátování sloupců funguje
- ✅ Po odemknutí: Sloupec lze smazat

**Test:**
```vba
Sub TestTS_R006()
    Dim ws As Worksheet
    Set ws = Sheets("Aplikace")

    ' Test 1: Zamkni list
    Call LockSpecificSheets(ws)

    If ws.ProtectContents Then
        Debug.Print "✅ List je zamčený"
    Else
        Debug.Print "❌ List není zamčený"
    End If

    ' Test 2: Odemkni listy
    Call UnlockAllSheets

    If Not ws.ProtectContents Then
        Debug.Print "✅ List je odemčený"
    Else
        Debug.Print "❌ List je stále zamčený"
    End If
End Sub
```

---

### TS-R007: Dočasná změna buňky

**Popis:** Testování workaround pro refresh.

**Kroky:**
1. Zaznamenat hodnotu Rozpočet!B2
2. Spustit `Call DocasnaZmenaBunky`
3. Zkontrolovat hodnotu Rozpočet!B2

**Očekávaný výsledek:**
- ✅ Hodnota B2 je stejná jako na začátku
- ✅ Během běhu byla dočasně změněna na 13

**Test:**
```vba
Sub TestTS_R007()
    Dim ws As Worksheet
    Set ws = Sheets("Rozpočet")

    Dim originalValue As Variant
    originalValue = ws.Range("B2").Value

    Call DocasnaZmenaBunky

    If ws.Range("B2").Value = originalValue Then
        Debug.Print "✅ Hodnota B2 obnovena OK"
    Else
        Debug.Print "❌ Hodnota B2 změněna!"
        Debug.Print "Původní: " & originalValue
        Debug.Print "Aktuální: " & ws.Range("B2").Value
    End If
End Sub
```

---

## Unit Testy

### Test: SheetExists

```vba
Sub Test_SheetExists()
    ' Test 1: Existující list
    If SheetExists("Aplikace") Then
        Debug.Print "✅ SheetExists (existující) PASSED"
    Else
        Debug.Print "❌ SheetExists (existující) FAILED"
    End If

    ' Test 2: Neexistující list
    If Not SheetExists("NeexistujiciList") Then
        Debug.Print "✅ SheetExists (neexistující) PASSED"
    Else
        Debug.Print "❌ SheetExists (neexistující) FAILED"
    End If
End Sub
```

---

### Test: ToggleScreenUpdating

```vba
Sub Test_ToggleScreenUpdating()
    ' Test 1: Vypnout
    Call ToggleScreenUpdating(True)

    If Not Application.ScreenUpdating And _
       Not Application.EnableEvents And _
       Application.Calculation = xlCalculationManual Then
        Debug.Print "✅ ToggleScreenUpdating (OFF) PASSED"
    Else
        Debug.Print "❌ ToggleScreenUpdating (OFF) FAILED"
    End If

    ' Test 2: Zapnout
    Call ToggleScreenUpdating(False)

    If Application.ScreenUpdating And _
       Application.EnableEvents And _
       Application.Calculation = xlCalculationAutomatic Then
        Debug.Print "✅ ToggleScreenUpdating (ON) PASSED"
    Else
        Debug.Print "❌ ToggleScreenUpdating (ON) FAILED"
    End If
End Sub
```

---

## Integrační testy

### INT-R001: End-to-End workflow

**Popis:** Kompletní průchod aplikací.

**Kroky:**
1. Zavřít Excel
2. Otevřít Rozpočet.xlsm
3. Přihlásit se
4. Navigovat na "Grafy"
5. Změnit barvy grafů
6. Navigovat na "Aplikace"
7. Exportovat data (Ctrl+K)
8. Zkontrolovat exportovaný soubor

**Očekávaný čas:** ~3 minuty

**Očekávaný výsledek:**
- ✅ Všechny kroky proběhnou bez chyby

---

## Známé problémy

### Issue #1: Pevně kódované rozsahy
**Popis:** Rozsah B4:AQ je hardcoded v Copy.bas

**Riziko:** Pokud se přidají sloupce za AQ, nebudou exportovány

**Řešení:** Dynamicky najít poslední sloupec

---

### Issue #2: Žádná validace exportu
**Popis:** Export nekontroluje, zda data existují

**Riziko:** Prázdný export

**Řešení:** Přidat kontrolu před exportem

---

### Issue #3: Absence error handlingu
**Popis:** Některé funkce nemají `On Error`

**Riziko:** Nezachycené chyby

**Řešení:** Přidat error handling

---

## Test Checklist

Před nasazením do produkce:

- [ ] TS-R001: Přihlášení
- [ ] TS-R002: Navigace
- [ ] TS-R003: Export (Ctrl+K)
- [ ] TS-R004: Pivot Plus/Minus
- [ ] TS-R005: Barvy grafů
- [ ] TS-R006: Ochrana listů
- [ ] TS-R007: Dočasná změna buňky
- [ ] Test_SheetExists
- [ ] Test_ToggleScreenUpdating
- [ ] INT-R001: End-to-End

---

## Reporting chyb

Při nalezení chyby uveďte:

1. **Číslo testu:** např. TS-R003
2. **Popis chyby**
3. **Očekávané chování**
4. **Kroky k reprodukci**
5. **Prostředí:** Excel verze, Windows verze
6. **Debug výstupy**

---

## Revize

**Verze:** 1.0
**Datum:** 2026-01-16
**Autor:** IN-EKO VBA Development Team
