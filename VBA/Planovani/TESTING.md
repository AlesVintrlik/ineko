# Testing Guide - Plánování

Tento dokument popisuje postupy pro testování VBA aplikace Plánování.

## Obsah

- [Příprava testovacího prostředí](#příprava-testovacího-prostředí)
- [Test Scenarios](#testovací-scénáře)
- [Unit testy](#unit-testy)
- [Integrační testy](#integrační-testy)
- [Helios integrace testy](#helios-integrace-testy)
- [UI/UX testy](#uiux-testy)
- [Známé problémy](#známé-problémy)

---

## Příprava testovacího prostředí

### Požadavky

1. **Testovací databáze**
   - Kopie produkční DB nebo testovací data
   - View: `hvw_TerminyZakazekProPlanovani`
   - Tables: `TabZakazka`, `TabZakazka_EXT`
   - SP: `EP_PlanTerminyVyrobyDoZakazek`, `ep_DoplnTerminOperacePodleUseku`

2. **Testovací účet**
   - SQL login s potřebnými oprávněními
   - Nebo NT autentizace

3. **Excel soubor**
   - Backup souboru Plánování.xlsm
   - Testovací kopie

### Setup

```vba
Sub SetupTestEnvironment()
    ' Přidat testovací zakázky
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Zakazky")

    ' Testovací data
    ws.Range("A2").Value = "TEST-001"
    ws.Range("B2").Value = "Testovací firma"
    ' ... další sloupce
End Sub
```

---

## Testovací scénáře

### TS-P001: První přihlášení

**Stejný jako Gantt TS-001** - viz [Gantt/TESTING.md](../Gantt/TESTING.md#ts-001-první-přihlášení)

**Dodatek pro Plánování:**
- ✅ Po přihlášení se automaticky načtou zakázky
- ✅ Zobrazí se progress bar během načítání
- ✅ Aplikuje se optimální font
- ✅ Zobrazí se frmReady s blikajícím "Hotovo!"

---

### TS-P002: Aktualizace seznamu zakázek

**Popis:** Testování aktualizace listu "Plan" z dat v "Zakazky".

**Prerekvizity:**
- List "Zakazky" obsahuje alespoň 20 zakázek
- Různá data expedice (včetně NULL)

**Kroky:**
1. Spustit `AktualizovatSeznamZakazek`
2. Sledovat progress bar
3. Počkat na dokončení

**Očekávaný výsledek:**
- ✅ Zakázky jsou seřazeny podle data expedice (vzestupně)
- ✅ Zakázky s NULL datem jsou na konci
- ✅ Sloupec B obsahuje čísla zakázek
- ✅ Sloupec C obsahuje názvy firem
- ✅ Vzorce jsou zkopírovány do všech řádků
- ✅ Formátování je konzistentní
- ✅ AutoFilter je aktivní
- ✅ List je zamčený (ochrana zapnuta)

**Validace:**
```vba
Sub ValidateTS_P002()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plan")

    ' Test 1: Data jsou seřazena
    Dim i As Long
    Dim prevDate As Variant
    prevDate = 0

    For i = 15 To ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        Dim currentDate As Variant
        currentDate = ws.Range("F" & i).Value ' ElektroKonec (předpokládaný sloupec)

        If Not IsEmpty(currentDate) And Not IsEmpty(prevDate) Then
            If currentDate < prevDate Then
                Debug.Print "❌ Řazení FAILED na řádku " & i
                Exit Sub
            End If
        End If
        If Not IsEmpty(currentDate) Then prevDate = currentDate
    Next i

    Debug.Print "✅ Řazení OK"

    ' Test 2: Vzorce jsou zkopírovány
    If ws.Range("D16").HasFormula Then
        Debug.Print "✅ Vzorce OK"
    Else
        Debug.Print "❌ Vzorce chybí"
    End If

    ' Test 3: Ochrana je zapnuta
    If ws.ProtectContents Then
        Debug.Print "✅ Ochrana OK"
    Else
        Debug.Print "❌ Ochrana chybí"
    End If
End Sub
```

---

### TS-P003: Klávesové zkratky +/- týden

**Popis:** Testování Ctrl+N a Ctrl+D pro úpravu termínů.

**Kroky:**
1. Kliknout na buňku M15 (plánovaný termín Přípravy)
2. Ověřit, že L15 obsahuje datum (původní termín)
3. Stisknout `Ctrl+N`
4. Ověřit, že M15 = L15 + 7 dní
5. Stisknout `Ctrl+N` znovu
6. Ověřit, že M15 = původní hodnota + 14 dní
7. Stisknout `Ctrl+D` dvakrát
8. Ověřit, že M15 = L15 (zpět na začátek)

**Očekávaný výsledek:**
- ✅ Ctrl+N přidává týden
- ✅ Ctrl+D odebírá týden
- ✅ Zkratky fungují pouze ve sloupcích M, P, S, V, Y
- ✅ Zkratky fungují pouze pokud levá buňka obsahuje datum

**Test edge cases:**
```vba
Sub TestTS_P003_EdgeCases()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plan")

    ' Edge case 1: Prázdná aktivní buňka
    ws.Range("M15").Clear
    ws.Range("L15").Value = #12/1/2024#
    ws.Range("M15").Select

    Application.SendKeys "^n"  ' Ctrl+N
    Application.Wait Now + TimeValue("00:00:01")

    If ws.Range("M15").Value = #12/8/2024# Then
        Debug.Print "✅ Prázdná buňka OK"
    Else
        Debug.Print "❌ Prázdná buňka FAILED"
    End If

    ' Edge case 2: Jiný sloupec (nefunguje)
    ws.Range("C15").Select
    Dim oldValue As Variant
    oldValue = ws.Range("C15").Value

    Application.SendKeys "^n"
    Application.Wait Now + TimeValue("00:00:01")

    If ws.Range("C15").Value = oldValue Then
        Debug.Print "✅ Jiný sloupec ignorován OK"
    Else
        Debug.Print "❌ Jiný sloupec FAILED (neměl se změnit)"
    End If
End Sub
```

---

### TS-P004: Výpočet termínů výroby

**Popis:** Testování volání SP pro výpočet termínů.

**Prerekvizity:**
- Zakázka existuje v DB (např. "TEST-001")
- SP `EP_PlanTerminyVyrobyDoZakazek` existuje

**Kroky:**
1. Kliknout na řádek se zakázkou "TEST-001"
2. Otevřít formulář zakázky
3. Ověřit, že txtZakazka = "TEST-001"
4. Zadat datum expedice: `31.12.2024`
5. Kliknout "Vypočítat termíny"

**Očekávaný výsledek:**
- ✅ SP byla zavolána úspěšně
- ✅ Termíny v `TabZakazka_EXT` byly aktualizovány
- ✅ Po refresh dat se zobrazí nové termíny

**Test:**
```vba
Sub TestTS_P004()
    Dim ID As Long
    ID = GetZakazkaID("TEST-001")

    If ID = 0 Then
        Debug.Print "❌ Zakázka nenalezena"
        Exit Sub
    End If

    ' Zavolej SP
    CallPlanTerminyVyrobyDoZakazek ID, #12/31/2024#

    ' Počkej na dokončení
    Application.Wait Now + TimeValue("00:00:03")

    ' Ověř, že termíny byly aktualizovány
    Dim conn As Object
    Dim rs As Object
    Set conn = CreateConnection()

    Dim sql As String
    sql = "SELECT _U1Konec FROM TabZakazka_EXT WHERE ID = " & ID

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0).Value) Then
            Debug.Print "✅ Termíny aktualizovány OK"
            Debug.Print "U1Konec: " & rs.Fields(0).Value
        Else
            Debug.Print "❌ Termíny NULL"
        End If
    Else
        Debug.Print "❌ Záznam nenalezen"
    End If

    rs.Close
    conn.Close
End Sub
```

---

### TS-P005: Evidence pracnosti

**Popis:** Testování uložení hodin do DB.

**Kroky:**
1. Kliknout na řádek se zakázkou
2. Otevřít formulář "Hodiny"
3. Vyplnit:
   - Hodiny celkem: 100
   - Skup. prac. 1: 20
   - Skup. prac. 2: 30
   - Skup. prac. 3: 20
   - Skup. prac. 4: 15
   - Skup. prac. 5: 15
   - Kooperace: 0
4. Kliknout "Uložit"

**Očekávaný výsledek:**
- ✅ MsgBox "Hodiny byly úspěšně uloženy."
- ✅ Data jsou v `TabZakazka_EXT`

**Validace:**
```vba
Sub ValidateTS_P005()
    Dim conn As Object
    Dim rs As Object
    Set conn = CreateConnection()

    Dim ID As Long
    ID = GetZakazkaID("TEST-001")

    Dim sql As String
    sql = "SELECT _HodCelkem, _HodSkPrac1, _HodSkPrac2, _HodSkPrac3, " & _
          "_HodSkPrac4, _HodSkPrac5, _HodKoop " & _
          "FROM TabZakazka_EXT WHERE ID = " & ID

    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        If rs.Fields("_HodCelkem").Value = 100 And _
           rs.Fields("_HodSkPrac1").Value = 20 And _
           rs.Fields("_HodSkPrac2").Value = 30 Then
            Debug.Print "✅ Hodiny uloženy správně"
        Else
            Debug.Print "❌ Hodnoty nesedí"
        End If
    Else
        Debug.Print "❌ Záznam nenalezen"
    End If

    rs.Close
    conn.Close
End Sub
```

---

### TS-P006: Naplnění kontroly

**Popis:** Testování vytvoření listu "Kontrola" před synchronizací.

**Prerekvizity:**
- List "Plan" obsahuje zakázky s upravenými termíny

**Kroky:**
1. Upravit nějaké termíny v listu "Plan"
2. Spustit `NaplnitKontrolu`
3. Zkontrolovat list "Kontrola"

**Očekávaný výsledek:**
- ✅ List "Kontrola" je viditelný
- ✅ Obsahuje řádky pro každý úsek každé zakázky
- ✅ Sloupce: Usek, Zakazka, PuvodniTermin, TerminVyrobyDo, PotrebnyCas, TydenPuvodni, TydenNovy
- ✅ ISO týdny jsou správně vypočítány

**Validace:**
```vba
Sub ValidateTS_P006()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Kontrola")

    ' Test 1: List je viditelný
    If ws.Visible = xlSheetVisible Then
        Debug.Print "✅ List viditelný OK"
    Else
        Debug.Print "❌ List není viditelný"
    End If

    ' Test 2: Hlavička existuje
    If ws.Range("A1").Value = "Usek" Then
        Debug.Print "✅ Hlavička OK"
    Else
        Debug.Print "❌ Hlavička chybí"
    End If

    ' Test 3: Data existují
    If ws.Range("A2").Value <> "" Then
        Debug.Print "✅ Data OK"
        Debug.Print "Počet změn: " & (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1)
    Else
        Debug.Print "❌ Žádná data"
    End If

    ' Test 4: ISO týdny
    If IsNumeric(ws.Range("F2").Value) Then
        Debug.Print "✅ ISO týdny OK"
    Else
        Debug.Print "❌ ISO týdny FAILED"
    End If
End Sub
```

---

### TS-P007: Synchronizace do Helios

**Popis:** Testování kompletní synchronizace změn zpět do Helios.

**⚠️ POZOR:** Tento test mění data v databázi! Používat pouze na testovacím prostředí!

**Prerekvizity:**
- List "Kontrola" je naplněn (TS-P006)
- Testovací zakázky

**Kroky:**
1. Zkontrolovat data v listu "Kontrola"
2. Zaznamenat původní hodnoty `_U1Konec` v DB
3. Spustit `AktualizovatData`
4. Počkat na dokončení (může trvat 1 sekunda × počet řádků)
5. Ověřit změny v DB

**Očekávaný výsledek:**
- ✅ MsgBox "Nastavil jsem termíny zakázek v Heliosu."
- ✅ `_U{usek}Konec` je aktualizován v DB
- ✅ `_U{usek}Start` je přepočítán (konec - dateDiff)
- ✅ SP `ep_DoplnTerminOperacePodleUseku` byla zavolána
- ✅ List "Kontrola" je skrytý
- ✅ List "Plan" je aktivní

**Test:**
```vba
Sub TestTS_P007()
    ' ⚠️ POUZE NA TESTOVACÍM PROSTŘEDÍ!

    Dim ID As Long
    ID = GetZakazkaID("TEST-001")

    ' Zaznamenej původní hodnotu
    Dim conn As Object
    Dim rs As Object
    Set conn = CreateConnection()

    Dim sql As String
    sql = "SELECT _U1Konec FROM TabZakazka_EXT WHERE ID = " & ID
    Set rs = conn.Execute(sql)

    Dim originalDate As Variant
    If Not rs.EOF Then
        originalDate = rs.Fields(0).Value
    End If
    rs.Close
    conn.Close

    ' Spusť synchronizaci
    Call AktualizovatData

    ' Počkej
    Application.Wait Now + TimeValue("00:00:05")

    ' Ověř změnu
    Set conn = CreateConnection()
    Set rs = conn.Execute(sql)

    If Not rs.EOF Then
        If rs.Fields(0).Value <> originalDate Then
            Debug.Print "✅ Synchronizace OK"
            Debug.Print "Původní: " & originalDate
            Debug.Print "Nový: " & rs.Fields(0).Value
        Else
            Debug.Print "❌ Žádná změna"
        End If
    End If

    rs.Close
    conn.Close
End Sub
```

---

### TS-P008: Font detection

**Popis:** Testování detekce a aplikace optimálního fontu.

**Kroky:**
1. Spustit `NastavFontPlusFormatDatumu`
2. Sledovat progress bar
3. Zkontrolovat aplikovaný font

**Očekávaný výsledek:**
- ✅ Progress bar zobrazuje aktuální font
- ✅ Aplikován první dostupný font z: Segoe UI Semilight, Segoe UI Light, Calibri Light, Arial Narrow, Arial
- ✅ Font podporuje češtinu
- ✅ Velikost fontu = 10
- ✅ Formát datumu = `dd.mm.yy`
- ✅ Sloupce H:AB mají jednotnou šířku
- ✅ Sloupec J má šířku 1

**Test:**
```vba
Sub TestTS_P008()
    Call NastavFontPlusFormatDatumu

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Plan")

    ' Test 1: Font
    Dim fontName As String
    fontName = ws.Range("A1").Font.Name

    Debug.Print "Aplikovaný font: " & fontName

    If fontName = "Segoe UI Semilight" Or _
       fontName = "Segoe UI Light" Or _
       fontName = "Calibri Light" Or _
       fontName = "Arial Narrow" Or _
       fontName = "Arial" Then
        Debug.Print "✅ Font OK"
    Else
        Debug.Print "❌ Neočekávaný font: " & fontName
    End If

    ' Test 2: Velikost
    If ws.Range("A1").Font.Size = 10 Then
        Debug.Print "✅ Velikost OK"
    Else
        Debug.Print "❌ Velikost: " & ws.Range("A1").Font.Size
    End If

    ' Test 3: Formát datumu
    If InStr(ws.Range("H15").NumberFormat, "dd.mm.yy") > 0 Then
        Debug.Print "✅ Formát datumu OK"
    Else
        Debug.Print "❌ Formát: " & ws.Range("H15").NumberFormat
    End If

    ' Test 4: Šířka sloupce J
    If ws.Columns("J").ColumnWidth = 1 Then
        Debug.Print "✅ Sloupec J OK"
    Else
        Debug.Print "❌ Sloupec J: " & ws.Columns("J").ColumnWidth
    End If
End Sub
```

---

## Unit testy

### Test: BubbleSortZakazky

```vba
Sub Test_BubbleSortZakazky()
    Dim arr() As Variant
    arr = Array( _
        Array("Z003", "Firma C", #12/31/2024#), _
        Array("Z001", "Firma A", #12/15/2024#), _
        Array("Z004", "Firma D", Null), _
        Array("Z002", "Firma B", #12/20/2024#) _
    )

    BubbleSortZakazky arr

    ' Očekávané pořadí: Z001, Z002, Z003, Z004 (NULL na konci)
    If arr(0)(0) = "Z001" And arr(3)(0) = "Z004" Then
        Debug.Print "✅ BubbleSortZakazky PASSED"
    Else
        Debug.Print "❌ BubbleSortZakazky FAILED"
        Debug.Print "Pořadí: " & arr(0)(0) & ", " & arr(1)(0) & ", " & arr(2)(0) & ", " & arr(3)(0)
    End If
End Sub
```

---

### Test: GetZakazkaID

```vba
Sub Test_GetZakazkaID()
    ' Test 1: Existující zakázka
    Dim ID As Long
    ID = GetZakazkaID("TEST-001")

    If ID > 0 Then
        Debug.Print "✅ GetZakazkaID (existující) PASSED, ID: " & ID
    Else
        Debug.Print "❌ GetZakazkaID (existující) FAILED"
    End If

    ' Test 2: Neexistující zakázka
    ID = GetZakazkaID("NEEXISTUJE-999")

    If ID = 0 Then
        Debug.Print "✅ GetZakazkaID (neexistující) PASSED"
    Else
        Debug.Print "❌ GetZakazkaID (neexistující) FAILED, vrátilo: " & ID
    End If
End Sub
```

---

### Test: IsFontInstalled

```vba
Sub Test_IsFontInstalled()
    ' Test 1: Arial (měl by vždy existovat)
    If IsFontInstalled("Arial") Then
        Debug.Print "✅ IsFontInstalled (Arial) PASSED"
    Else
        Debug.Print "❌ IsFontInstalled (Arial) FAILED"
    End If

    ' Test 2: Neexistující font
    If Not IsFontInstalled("NonExistentFont12345") Then
        Debug.Print "✅ IsFontInstalled (neexistující) PASSED"
    Else
        Debug.Print "❌ IsFontInstalled (neexistující) FAILED"
    End If
End Sub
```

---

## Integrační testy

### INT-P001: End-to-End workflow

**Popis:** Kompletní průchod aplikací od přihlášení po synchronizaci.

**Kroky:**
1. Zavřít Excel
2. Otevřít Plánování.xlsm
3. Přihlásit se
4. Počkat na načtení dat
5. Upravit termín u zakázky "TEST-001"
6. Použít Ctrl+N pro posun o týden
7. Naplnit kontrolu
8. Zkontrolovat změny
9. Synchronizovat do Helios
10. Obnovit data
11. Ověřit změny

**Očekávaný čas:** ~5 minut

**Očekávaný výsledek:**
- ✅ Všechny kroky proběhnou bez chyby
- ✅ Změna je viditelná v DB i v Excelu

---

## Helios Integrace testy

### HEL-001: Volání SP s NULL parametrem

```vba
Sub TestHEL_001()
    Dim ID As Long
    ID = GetZakazkaID("TEST-001")

    ' Zavolej SP s NULL datem
    CallPlanTerminyVyrobyDoZakazek ID  ' druhý parametr chybí = NULL

    Debug.Print "✅ SP zavolána s NULL - check DB pro výsledek"
End Sub
```

---

### HEL-002: Zpětné plánování

**Manuální test:**
1. Nastav datum expedice: 31.12.2024
2. Zavolej SP
3. Zkontroluj termíny v DB:
   - Balení (_U8) by měl být nejblíže expedici
   - Příprava (_U1) by měla být nejdříve

---

## UI/UX testy

### UI-001: Progress bar feedback

**Testovat:**
- Načítání dat zobrazuje progress
- Font detection zobrazuje aktuální font
- Formátování zobrazuje % dokončení

---

### UI-002: frmReady blikání

**Testovat:**
- Label2 bliká červeně/černě
- Blikání se zastaví po zavření formuláře

---

### UI-003: Ochrana listu

**Testovat:**
- List "Plan" je zamčený
- AutoFilter funguje
- Buňky s termíny jsou editovatelné
- Vzorce nelze smazat

---

## Známé problémy

### Issue #1: Delay v Helios.AktualizovatData

**Popis:** 1 sekunda pauza mezi každou aktualizací je pomalá.

**Impact:** Pro 100 změn = 100 sekund (~1.7 minuty)

**Workaround:** Batch UPDATE místo iterace

---

### Issue #2: Heslo v plaintext

**Popis:** `MrkevNeniOvoce123` je v kódu.

**Riziko:** Každý s přístupem k VBA může odemknout list.

**Řešení:** Použít šifrování nebo Windows permission.

---

### Issue #3: Font detection vyžaduje Windows

**Popis:** Win32 API nefunguje na Mac.

**Impact:** Na Mac se použije Arial (fallback).

---

### Issue #4: Concurrent edit

**Popis:** Aplikace nekontroluje, zda někdo jiný upravuje stejnou zakázku.

**Riziko:** Přepsání změn jiného uživatele.

**Řešení:** Implementovat locking nebo timestamp check.

---

## Test Checklist

Před nasazením do produkce:

- [ ] TS-P001: Přihlášení
- [ ] TS-P002: Aktualizace zakázek
- [ ] TS-P003: Klávesové zkratky
- [ ] TS-P004: Výpočet termínů
- [ ] TS-P005: Evidence pracnosti
- [ ] TS-P006: Naplnění kontroly
- [ ] TS-P007: Synchronizace do Helios (POUZE TESTOVACÍ DB!)
- [ ] TS-P008: Font detection
- [ ] INT-P001: End-to-End
- [ ] Test_BubbleSortZakazky
- [ ] Test_GetZakazkaID
- [ ] Test_IsFontInstalled
- [ ] UI-001, UI-002, UI-003

---

## Reporting chyb

Při nalezení chyby uveďte:

1. **Číslo testu:** např. TS-P004
2. **Popis chyby**
3. **Očekávané chování**
4. **Kroky k reprodukci**
5. **Prostředí:** Excel verze, SQL Server verze, Windows verze
6. **Debug.Print výstupy**

**Příklad:**
```
Bug Report: TS-P004 Failed

Popis: SP vrátila chybu "Invalid date format"
Očekávané: SP by měla přijmout formát YYYY-MM-DD
Aktuální: SP vrací chybu

Kroky:
1. Otevřít formulář zakázky
2. Zadat datum: 31.12.2024
3. Kliknout "Vypočítat"

Prostředí:
- Excel 2019
- SQL Server 2016
- Windows 10

Debug output:
EXEC dbo.EP_PlanTerminyVyrobyDoZakazek @ID = 12345, @DatumUkonceni = '2024-12-31'
Msg 241, Level 16, State 1, Procedure EP_PlanTerminyVyrobyDoZakazek, Line 10
Conversion failed when converting date and/or time from character string.
```

---

## Revize

**Verze:** 1.0
**Datum:** 2026-01-16
**Autor:** IN-EKO VBA Development Team
