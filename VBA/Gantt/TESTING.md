# Testing Guide - Gantt

Tento dokument popisuje postupy pro testování VBA aplikace Gantt, včetně manuálního testování, testovacích scénářů a běžných problémů.

## Obsah

- [Typy testování](#typy-testování)
- [Příprava testovacího prostředí](#příprava-testovacího-prostředí)
- [Testovací scénáře](#testovací-scénáře)
- [Unit testy](#unit-testy)
- [Integrační testy](#integrační-testy)
- [Testování výkonu](#testování-výkonu)
- [Bezpečnostní testování](#bezpečnostní-testování)
- [Známé problémy](#známé-problémy)
- [Debug nástroje](#debug-nástroje)

---

## Typy testování

### 1. Manuální testování
- Funkční testování uživatelského rozhraní
- Vizuální kontrola výstupů
- Testování workflow

### 2. Unit testování
- Testování jednotlivých funkcí a procedur
- Izolované testování bez závislostí

### 3. Integrační testování
- Testování komunikace s databází
- Testování interakce mezi moduly
- End-to-end testování

### 4. Performance testování
- Měření rychlosti operací
- Testování s velkým množstvím dat
- Identifikace bottlenecks

### 5. Bezpečnostní testování
- Testování autentizace
- Testování šifrování
- SQL injection testy

---

## Příprava testovacího prostředí

### Požadavky

1. **Testovací databáze**
   - Kopie produkční databáze nebo testovací data
   - Testovací SQL Server instance
   - View `hvw_TerminyZakazekProPlanovani` s testovacími daty

2. **Testovací účet**
   - SQL login s READ oprávněními
   - Nebo NT autentizace

3. **Excel soubor**
   - Kopie Gantt.xlsm pro testování
   - Backup originálních dat

### Setup testovacího prostředí

```vba
' 1. Vytvoření testovacího listu
Sub SetupTestEnvironment()
    On Error Resume Next

    ' Přidat testovací list pokud neexistuje
    Dim wsTest As Worksheet
    Set wsTest = ThisWorkbook.Sheets("TEST_DATA")

    If wsTest Is Nothing Then
        Set wsTest = ThisWorkbook.Sheets.Add
        wsTest.Name = "TEST_DATA"
    End If

    ' Inicializovat testovací data
    wsTest.Range("A1").Value = "Test Data Initialized: " & Now()
End Sub

' 2. Cleanup po testech
Sub CleanupTestEnvironment()
    On Error Resume Next
    ThisWorkbook.Sheets("TEST_DATA").Delete
    Application.DisplayAlerts = True
End Sub
```

---

## Testovací scénáře

### TS-001: První přihlášení

**Popis:** Testování prvního přihlášení uživatele.

**Prerekvizity:**
- Gantt.xlsm není otevřen
- Testovací SQL účet je aktivní

**Kroky:**
1. Otevřít Gantt.xlsm
2. Ověřit, že se zobrazí frmLogin
3. Vyplnit údaje:
   - Server: `TESTSERVER`
   - Databáze: `TEST_ERP`
   - Uživatel: `testuser`
   - Heslo: `testpass`
4. Kliknout "Přihlásit"

**Očekávaný výsledek:**
- ✅ Formulář se zavře
- ✅ Excel se maximalizuje a zobrazí
- ✅ Aktivuje se list "Gantt"
- ✅ `isLoggedIn = True`
- ✅ `loginCredentials` obsahuje šifrované údaje

**Test data:**
```
Server: TESTSERVER
Database: TEST_ERP
User: testuser
Password: testpass123
```

---

### TS-002: Neplatné přihlášení

**Popis:** Testování přihlášení s neplatnými údaji.

**Kroky:**
1. Otevřít Gantt.xlsm
2. Vyplnit neplatné údaje:
   - Server: `TESTSERVER`
   - Databáze: `TEST_ERP`
   - Uživatel: `invaliduser`
   - Heslo: `wrongpassword`
3. Kliknout "Přihlásit"

**Očekávaný výsledek:**
- ✅ Zobrazí se MsgBox: "Nesprávné přihlašovací údaje. Zkuste to prosím znovu."
- ✅ Formulář zůstane otevřený
- ✅ Focus se vrátí na pole username
- ✅ `isLoggedIn = False`

---

### TS-003: NT Autentizace

**Popis:** Testování Windows autentizace.

**Kroky:**
1. Otevřít Gantt.xlsm
2. Vyplnit údaje:
   - Server: `TESTSERVER`
   - Databáze: `TEST_ERP`
   - Uživatel: *(prázdné)*
   - Heslo: *(prázdné)*
3. Kliknout "Přihlásit"

**Očekávaný výsledek:**
- ✅ Přihlášení úspěšné pomocí Windows účtu
- ✅ Connection string obsahuje `Integrated Security=SSPI`

---

### TS-004: Načtení dat ze serveru

**Popis:** Testování načítání dat z databáze.

**Prerekvizity:**
- Uživatel je přihlášen
- View `hvw_TerminyZakazekProPlanovani` obsahuje alespoň 10 záznamů

**Kroky:**
1. Spustit makro `LoadOrUpdateData`
2. Počkat na dokončení

**Očekávaný výsledek:**
- ✅ Zobrazí se frmProgress
- ✅ Data se načtou do listu "Zakazky"
- ✅ Řádek 1 obsahuje hlavičku (názvy sloupců)
- ✅ Řádek 2+ obsahuje data
- ✅ MsgBox: "Data byla načtena do listu s hlavičkou."
- ✅ frmProgress se zavře

**Validace:**
```vba
Sub ValidateTS004()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Zakazky")

    ' Test 1: Hlavička existuje
    If ws.Range("A1").Value <> "" Then
        Debug.Print "✅ Hlavička OK"
    Else
        Debug.Print "❌ Hlavička chybí"
    End If

    ' Test 2: Data existují
    If ws.Range("A2").Value <> "" Then
        Debug.Print "✅ Data OK"
    Else
        Debug.Print "❌ Data chybí"
    End If

    ' Test 3: Počet řádků
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Debug.Print "Počet řádků s daty: " & (lastRow - 1)
End Sub
```

---

### TS-005: Aktualizace Gantt diagramu

**Popis:** Testování vytvoření Gantt diagramu ze zakázek.

**Prerekvizity:**
- List "Zakazky" obsahuje data (běžte TS-004)

**Kroky:**
1. Spustit makro `AktualizovatSeznamZakazek`
2. Počkat na dokončení

**Očekávaný výsledek:**
- ✅ Zobrazí se MsgBox: "Načetl jsem aktuální zakázky."
- ✅ List "Gantt" obsahuje řádky se zakázkami (od řádku 4)
- ✅ Sloupec B obsahuje čísla zakázek
- ✅ Sloupce C-O obsahují vzorce zkopírované z řádku 4
- ✅ Formátování je konzistentní
- ✅ Pod zakázkami jsou 4 řádky se sumarizací kapacit

**Validace:**
```vba
Sub ValidateTS005()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")

    ' Test 1: Zakázky existují
    If ws.Range("B4").Value <> "" Then
        Debug.Print "✅ Zakázky OK"
    Else
        Debug.Print "❌ Zakázky chybí"
    End If

    ' Test 2: Vzorce zkopírovány
    If ws.Range("C5").HasFormula Then
        Debug.Print "✅ Vzorce OK"
    Else
        Debug.Print "❌ Vzorce chybí"
    End If

    ' Test 3: Počet zakázek
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    Debug.Print "Počet zakázek: " & (lastRow - 3)
End Sub
```

---

### TS-006: Výpočet kapacit

**Popis:** Testování výpočtu obsazenosti výrobních středisek.

**Prerekvizity:**
- Gantt diagram je vytvořený (běžte TS-005)
- List "Svátky" obsahuje data

**Kroky:**
1. Spustit makro `SumarizaceBoduProVsechnySloupce`
2. Počkat na dokončení (sledovat progress bar)

**Očekávaný výsledek:**
- ✅ Zobrazí se frmProgress s % dokončení
- ✅ Pod zakázkami se vytvoří 4 řádky:
  - Příprava (kapacita 1)
  - Svařování (kapacita 2)
  - Montáž (kapacita 2)
  - Elektro (kapacita 1)
- ✅ Buňky obsahují číselné hodnoty (0, 1, 2, ...)
- ✅ Barevné formátování je aplikováno:
  - Zelená: hodnota < kapacita
  - Oranžová: hodnota = kapacita
  - Červená: hodnota > kapacita
  - Bílá: hodnota = 0
- ✅ Víkendy a svátky mají hodnotu 0 nebo jsou prázdné

**Validace:**
```vba
Sub ValidateTS006()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Test: Řádky se součty existují
    If IsNumeric(ws.Cells(lastRow + 1, 16).Value) Then
        Debug.Print "✅ Součty OK"
        Debug.Print "Příprava (sloupec O): " & ws.Cells(lastRow + 1, 16).Value
        Debug.Print "Svařování (sloupec O): " & ws.Cells(lastRow + 2, 16).Value
        Debug.Print "Montáž (sloupec O): " & ws.Cells(lastRow + 3, 16).Value
        Debug.Print "Elektro (sloupec O): " & ws.Cells(lastRow + 4, 16).Value
    Else
        Debug.Print "❌ Součty chybí"
    End If
End Sub
```

---

### TS-007: Barevné formátování kapacit

**Popis:** Testování správnosti barevného kódování.

**Prerekvizity:**
- Kapacity jsou spočítané (běžte TS-006)

**Testovací data:**
- Vytvořit scénáře s různými hodnotami obsazenosti

**Kroky:**
1. Ručně upravit data zakázek tak, aby vznikly různé scénáře:
   - 0 zakázek v daný den → Očekávám bílou
   - 1 zakázka pro Přípravu → Očekávám oranžovou (kapacita = 1)
   - 2 zakázky pro Přípravu → Očekávám červenou (překročení)
2. Spustit `SumarizaceBoduProVsechnySloupce`
3. Zkontrolovat barvy

**Očekávaný výsledek:**

| Hodnota | Kapacita | Barva | RGB |
|---------|----------|-------|-----|
| 0 | 1 | Bílá | (255, 255, 255) |
| 1 | 2 | Zelená | (218, 242, 208) |
| 2 | 2 | Oranžová | (255, 225, 129) |
| 3 | 2 | Červená | (255, 217, 217) |

**Validace:**
```vba
Sub ValidateTS007()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Gantt")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    Dim cell As Range
    Set cell = ws.Cells(lastRow + 1, 16) ' Příprava, sloupec O

    Dim color As Long
    color = cell.Interior.Color

    Dim value As Integer
    value = cell.Value

    Debug.Print "Hodnota: " & value
    Debug.Print "Barva RGB: " & RGB( _
        color Mod 256, _
        (color \ 256) Mod 256, _
        (color \ 65536) Mod 256)

    ' Kontrola podle hodnoty
    If value = 0 Then
        If color = RGB(255, 255, 255) Then
            Debug.Print "✅ Bílá OK"
        Else
            Debug.Print "❌ Očekávám bílou"
        End If
    ElseIf value < 1 Then ' LidiPriprava = 1
        If color = RGB(218, 242, 208) Then
            Debug.Print "✅ Zelená OK"
        Else
            Debug.Print "❌ Očekávám zelenou"
        End If
    ElseIf value = 1 Then
        If color = RGB(255, 225, 129) Then
            Debug.Print "✅ Oranžová OK"
        Else
            Debug.Print "❌ Očekávám oranžovou"
        End If
    Else
        If color = RGB(255, 217, 217) Then
            Debug.Print "✅ Červená OK"
        Else
            Debug.Print "❌ Očekávám červenou"
        End If
    End If
End Sub
```

---

### TS-008: Víkendy a svátky

**Popis:** Testování, že víkendy a svátky nejsou započítány do kapacit.

**Prerekvizity:**
- List "Svátky" obsahuje alespoň 1 svátek v rozsahu časové osy

**Kroky:**
1. Identifikovat sloupec s víkendem (So nebo Ne)
2. Identifikovat sloupec se svátkem
3. Spustit `SumarizaceBoduProVsechnySloupce`
4. Zkontrolovat hodnoty v těchto sloupcích

**Očekávaný výsledek:**
- ✅ Sloupce s víkendy mají hodnotu 0 nebo jsou prázdné
- ✅ Sloupce se svátky mají hodnotu 0 nebo jsou prázdné
- ✅ Barva je bílá

---

### TS-009: Zrušení filtrů

**Popis:** Testování funkce `ZrusitVsechnyFiltry`.

**Kroky:**
1. Na listu "Gantt" aplikovat AutoFilter
2. Filtrovat sloupec B (zakázky) - vybrat jen některé zakázky
3. Spustit makro `ZrusitVsechnyFiltry`

**Očekávaný výsledek:**
- ✅ Všechny řádky jsou viditelné
- ✅ Filtr zůstane zapnutý, ale žádné podmínky nejsou aktivní

---

### TS-010: Minimalistický režim

**Popis:** Testování zobrazení bez mřížky a záhlaví.

**Kroky:**
1. Zavřít Gantt.xlsm
2. Otevřít Gantt.xlsm
3. Přihlásit se

**Očekávaný výsledek:**
- ✅ Mřížka je skrytá (`DisplayGridlines = False`)
- ✅ Záhlaví řádků/sloupců je skryté (`DisplayHeadings = False`)
- ✅ Panel vzorců je skrytý (`DisplayFormulaBar = False`)

**Cleanup test:**
1. Zavřít Gantt.xlsm

**Očekávaný výsledek:**
- ✅ Před zavřením se mřížka, záhlaví a panel vzorců obnoví

---

### TS-011: Soubor otevřen jiným uživatelem

**Popis:** Testování zamykání souboru.

**Kroky:**
1. Otevřít Gantt.xlsm na počítači A (jako User1)
2. Pokusit se otevřít Gantt.xlsm na počítači B (jako User2)

**Očekávaný výsledek:**
- ✅ Počítač B zobrazí MsgBox: "Tento soubor je již otevřen uživatelem: DOMAIN\User1. Aplikace bude ukončena."
- ✅ Aplikace se zavře

---

## Unit testy

### Test: XOREncryptDecrypt

**Účel:** Ověřit, že šifrování a dešifrování vrací originální text.

```vba
Sub Test_XOREncryptDecrypt()
    Dim original As String
    Dim key As String
    Dim encrypted As String
    Dim decrypted As String

    original = "TestPassword123"
    key = ENCRYPTION_KEY

    ' Šifrování
    encrypted = XOREncryptDecrypt(original, key)

    ' Dešifrování
    decrypted = XOREncryptDecrypt(encrypted, key)

    ' Assert
    If decrypted = original Then
        Debug.Print "✅ Test_XOREncryptDecrypt PASSED"
    Else
        Debug.Print "❌ Test_XOREncryptDecrypt FAILED"
        Debug.Print "Expected: " & original
        Debug.Print "Got: " & decrypted
    End If
End Sub
```

---

### Test: VerifyCredentials

**Účel:** Ověřit autentizaci s platnými a neplatnými údaji.

```vba
Sub Test_VerifyCredentials()
    Dim result As Boolean

    ' Test 1: Platné údaje
    result = VerifyCredentials("testuser", "testpass", "TESTSERVER", "TEST_ERP")
    If result = True Then
        Debug.Print "✅ Test_VerifyCredentials (valid) PASSED"
    Else
        Debug.Print "❌ Test_VerifyCredentials (valid) FAILED"
    End If

    ' Test 2: Neplatné údaje
    result = VerifyCredentials("invaliduser", "wrongpass", "TESTSERVER", "TEST_ERP")
    If result = False Then
        Debug.Print "✅ Test_VerifyCredentials (invalid) PASSED"
    Else
        Debug.Print "❌ Test_VerifyCredentials (invalid) FAILED"
    End If
End Sub
```

---

### Test: CreateConnection

**Účel:** Ověřit vytvoření databázového připojení.

```vba
Sub Test_CreateConnection()
    ' Setup
    Dim creds As New Collection
    creds.Add XOREncryptDecrypt("testuser", ENCRYPTION_KEY), "login"
    creds.Add XOREncryptDecrypt("testpass", ENCRYPTION_KEY), "passw"
    Set loginCredentials = creds

    ' Test
    Dim conn As Object
    Set conn = CreateConnection()

    ' Assert
    If Not conn Is Nothing Then
        If conn.State = 1 Then ' adStateOpen
            Debug.Print "✅ Test_CreateConnection PASSED"
            conn.Close
        Else
            Debug.Print "❌ Test_CreateConnection FAILED - Connection not open"
        End If
    Else
        Debug.Print "❌ Test_CreateConnection FAILED - Connection is Nothing"
    End If

    Set conn = Nothing
End Sub
```

---

### Test: IsInArray

**Účel:** Ověřit funkci vyhledávání v poli.

```vba
Sub Test_IsInArray()
    Dim arr() As Variant
    arr = Array("A", "B", "C", "D")

    ' Test 1: Hodnota v poli
    If IsInArray("B", arr) = True Then
        Debug.Print "✅ Test_IsInArray (found) PASSED"
    Else
        Debug.Print "❌ Test_IsInArray (found) FAILED"
    End If

    ' Test 2: Hodnota není v poli
    If IsInArray("Z", arr) = False Then
        Debug.Print "✅ Test_IsInArray (not found) PASSED"
    Else
        Debug.Print "❌ Test_IsInArray (not found) FAILED"
    End If
End Sub
```

---

### Test: BubbleSort

**Účel:** Ověřit řazení pole.

```vba
Sub Test_BubbleSort()
    Dim arr() As Variant
    arr = Array(5, 2, 8, 1, 9, 3)

    BubbleSort arr

    ' Assert
    Dim expected() As Variant
    expected = Array(1, 2, 3, 5, 8, 9)

    Dim i As Long
    Dim passed As Boolean
    passed = True

    For i = LBound(arr) To UBound(arr)
        If arr(i) <> expected(i) Then
            passed = False
            Exit For
        End If
    Next i

    If passed Then
        Debug.Print "✅ Test_BubbleSort PASSED"
    Else
        Debug.Print "❌ Test_BubbleSort FAILED"
        Debug.Print "Expected: " & Join(expected, ", ")
        Debug.Print "Got: " & Join(arr, ", ")
    End If
End Sub
```

---

## Integrační testy

### INT-001: End-to-End workflow

**Popis:** Kompletní průchod aplikací od přihlášení po zobrazení Ganttu.

**Kroky:**
1. Zavřít všechny instance Excelu
2. Otevřít Gantt.xlsm
3. Přihlásit se (platné údaje)
4. Spustit `LoadOrUpdateData`
5. Spustit `AktualizovatSeznamZakazek`
6. Vizuálně zkontrolovat Gantt diagram
7. Zkontrolovat barevné kódování kapacit

**Očekávaný výsledek:**
- ✅ Všechny kroky proběhnou bez chyby
- ✅ Gantt diagram zobrazuje aktuální data
- ✅ Kapacity jsou správně barevně označeny

**Čas trvání:** ~2-5 minut

---

### INT-002: Databázová integrace

**Popis:** Test připojení a načítání dat z různých SQL instancí.

**Testovací případy:**
1. SQL Server 2012
2. SQL Server 2016
3. SQL Server 2019
4. SQL Server 2022

**Pro každou instanci:**
1. Nakonfigurovat připojení
2. Přihlásit se
3. Načíst data
4. Ověřit, že data jsou správná

---

### INT-003: Výkonnostní test s velkým množstvím dat

**Popis:** Test s velkým objemem zakázek.

**Testovací data:**
- 10 zakázek
- 50 zakázek
- 100 zakázek
- 500 zakázek

**Měření:**
1. Čas načtení dat (`LoadOrUpdateData`)
2. Čas aktualizace Ganttu (`AktualizovatSeznamZakazek`)
3. Čas výpočtu kapacit (`SumarizaceBoduProVsechnySloupce`)

**Příklad:**
```vba
Sub PerformanceTest_LoadData()
    Dim startTime As Double
    Dim endTime As Double

    startTime = Timer
    Call LoadOrUpdateData
    endTime = Timer

    Debug.Print "Čas načtení dat: " & Format(endTime - startTime, "0.00") & " s"
End Sub
```

**Akceptační kritéria:**
- 10 zakázek: < 5 sekund
- 50 zakázek: < 15 sekund
- 100 zakázek: < 30 sekund
- 500 zakázek: < 120 sekund

---

## Testování výkonu

### Benchmark: SumarizaceBoduProVsechnySloupce

```vba
Sub Benchmark_Sumarizace()
    Dim startTime As Double
    Dim endTime As Double
    Dim iterations As Integer
    Dim i As Integer

    iterations = 10

    startTime = Timer
    For i = 1 To iterations
        x = 1 ' Potlačit progress bar
        Call VymazaniSouctuPodGrafem
        Call SumarizaceBoduProVsechnySloupce
    Next i
    endTime = Timer

    Debug.Print "Průměrný čas sumarizace: " & _
                Format((endTime - startTime) / iterations, "0.00") & " s"
End Sub
```

### Profiling: Identifikace bottlenecks

```vba
Sub Profile_SumarizaceBoduProVsechnySloupce()
    Dim t1 As Double, t2 As Double, t3 As Double

    t1 = Timer
    ' Načtení svátků
    Dim holidays As Variant
    Set holidayRange = Sheets("Svátky").Range("B2", Sheets("Svátky").Cells(Sheets("Svátky").Rows.Count, "B").End(xlUp))
    holidays = holidayRange.Value
    t2 = Timer
    Debug.Print "Načtení svátků: " & Format(t2 - t1, "0.000") & " s"

    ' Hlavní smyčka
    ' ... (zde by byl zbytek kódu) ...

    t3 = Timer
    Debug.Print "Hlavní smyčka: " & Format(t3 - t2, "0.000") & " s"
End Sub
```

---

## Bezpečnostní testování

### SEC-001: SQL Injection test

**Popis:** Ověřit, že aplikace není zranitelná SQL injection.

**⚠️ Poznámka:** Aktuální implementace používá parametrizované dotazy přes ADODB, což je bezpečné.

**Test:**
```vba
' Pokus o SQL injection přes přihlašovací formulář
' Username: admin' OR '1'='1
' Password: ' OR '1'='1
```

**Očekávaný výsledek:**
- ✅ Přihlášení selže
- ✅ Žádný SQL kód není vykonán

---

### SEC-002: Šifrování hesel

**Test:**
```vba
Sub Test_PasswordEncryption()
    Dim password As String
    Dim encrypted As String

    password = "SuperTajneHeslo123"
    encrypted = XOREncryptDecrypt(password, ENCRYPTION_KEY)

    ' Ověřit, že šifrovaný text není čitelný
    If encrypted <> password And Len(encrypted) = Len(password) Then
        Debug.Print "✅ Heslo je zašifrované"
        Debug.Print "Original: " & password
        Debug.Print "Encrypted: " & encrypted
    Else
        Debug.Print "❌ Šifrování selhalo"
    End If
End Sub
```

---

### SEC-003: Connection string security

**Test:** Ověřit, že connection string není uložen v plain text.

**Očekávaný výsledek:**
- ✅ Connection string je vytvářen dynamicky
- ✅ Heslo není nikde uloženo nešifrované
- ✅ Connection string není logován (zkontrolovat Debug.Print statements)

---

## Známé problémy

### Issue #1: XOR šifrování není kryptograficky bezpečné

**Popis:** XOR cipher je lehce prolomitelný.

**Řešení:** Pro produkční použití zvážit AES šifrování.

**Workaround:** Používat Windows NT autentizaci, která hesla neukládá.

---

### Issue #2: Progress bar blokuje UI

**Popis:** Progress bar je zobrazen modeless, ale stále blokuje některé akce.

**Řešení:** Používat DoEvents častěji.

---

### Issue #3: Výkon při velkém množství dat

**Popis:** S 500+ zakázkami může `SumarizaceBoduProVsechnySloupce` trvat > 2 minuty.

**Řešení:**
- Optimalizovat algoritmus (např. indexování)
- Použít pole místo Range objektů

---

## Debug nástroje

### Immediate Window

```vba
' Kontrola přihlášení
? isLoggedIn

' Kontrola credentials
? loginCredentials("login")

' Test připojení
? CreateConnection Is Nothing

' Počet zakázek
? Sheets("Gantt").Cells(Rows.Count, "B").End(xlUp).Row - 3
```

### Locals Window

Užitečné pro krokování kódem a sledování hodnot proměnných.

### Watch Window

Přidejte do Watch:
- `loginCredentials`
- `isLoggedIn`
- `Application.ScreenUpdating`
- `Application.Calculation`

### Debug.Print statements

```vba
' V Connection.bas
Debug.Print "Connection String: " & connectionString

' V Advanced.bas
Debug.Print "Posledn sloupec: " & lastCol
Debug.Print "NextRow: " & nextRow
```

### Breakpoints

Nastavte breakpoint na klíčových místech:
- `Connection.bas:40` - před otevřením připojení
- `Data.bas:40` - před načtením dat
- `Advanced.bas:135` - ve smyčce výpočtu kapacit

---

## Test Suite

### Spuštění všech testů

```vba
Sub RunAllTests()
    Debug.Print "====== UNIT TESTS ======"
    Test_XOREncryptDecrypt
    Test_VerifyCredentials
    Test_CreateConnection
    Test_IsInArray
    Test_BubbleSort

    Debug.Print ""
    Debug.Print "====== VALIDATION TESTS ======"
    ValidateTS004
    ValidateTS005
    ValidateTS006
    ValidateTS007

    Debug.Print ""
    Debug.Print "====== PERFORMANCE TESTS ======"
    PerformanceTest_LoadData
    Benchmark_Sumarizace

    Debug.Print ""
    Debug.Print "====== SECURITY TESTS ======"
    Test_PasswordEncryption

    Debug.Print ""
    Debug.Print "====== ALL TESTS COMPLETE ======"
End Sub
```

---

## Testovací checklist

Před nasazením do produkce:

- [ ] TS-001: První přihlášení
- [ ] TS-002: Neplatné přihlášení
- [ ] TS-003: NT Autentizace
- [ ] TS-004: Načtení dat
- [ ] TS-005: Aktualizace Ganttu
- [ ] TS-006: Výpočet kapacit
- [ ] TS-007: Barevné formátování
- [ ] TS-008: Víkendy a svátky
- [ ] TS-009: Zrušení filtrů
- [ ] TS-010: Minimalistický režim
- [ ] TS-011: Zamykání souboru
- [ ] INT-001: End-to-End workflow
- [ ] INT-003: Výkonnostní test
- [ ] SEC-001: SQL Injection
- [ ] SEC-002: Šifrování hesel
- [ ] SEC-003: Connection string security

---

## Reporting chyb

Při nalezení chyby uveďte:

1. **Číslo testu:** např. TS-005
2. **Popis chyby:** Co se stalo
3. **Očekávané chování:** Co mělo nastat
4. **Kroky k reprodukci:** Jak chybu vyvolat
5. **Prostředí:** Excel verze, Windows verze, SQL Server verze
6. **Screenshots:** Pokud relevantní
7. **Log:** Debug.Print výstupy z Immediate Window

**Příklad:**
```
Bug Report: TS-006 Failed

Popis: Sumarizace kapacit nezohledňuje svátky
Očekávané: Svátky by měly mít hodnotu 0
Aktuální: Svátky obsahují nenulové hodnoty
Kroky:
1. Přidat svátek 2024-01-01 do listu Svátky
2. Spustit SumarizaceBoduProVsechnySloupce
3. Zkontrolovat hodnotu pro 2024-01-01

Prostředí:
- Excel 2019
- Windows 10
- SQL Server 2016

Screenshot: [attached]
```

---

## Revize

**Verze:** 1.0
**Datum:** 2026-01-16
**Autor:** IN-EKO VBA Development Team
