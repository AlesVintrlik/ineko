Attribute VB_Name = "Data"

Sub LoadDataFromQueries()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim startTime As Double
    Dim tableName As String
    Dim i As Integer
    Dim progress As Double
    Dim headerRange As Range
    
    'Debug.Print "Je v bìhu naèítání hlavních dat..."
    
    ' Vypnutí vykreslování a aktualizace
    ToggleScreenUpdating True
    
    ' Definice listù a SQL dotazù
    Dim queries As Variant
    queries = Array( _
        Array("Rozpoèet", "SELECT * FROM hvw_ReportRozpocetPlneni ORDER BY Obdobi, Skupina ,Ucet", "Rozpoèet"))
    
    ' Zobrazit UserForm s pokrokovým ukazatelem
    Load frmProgress
    frmProgress.Show vbModeless
    
    startTime = Timer ' Start timer for progress indication
    
    On Error GoTo ErrorHandler ' Pøidání globálního ošetøení chyb
    
    For i = LBound(queries) To UBound(queries)
        ' Aktualizace pokrokového ukazatele
        progress = (i - LBound(queries)) / (UBound(queries) - LBound(queries) + 1)
        frmProgress.UpdateProgressBar progress
        
        ' Zobrazit prùbìh naèítání
        Application.StatusBar = "Naèítání dat pro " & queries(i)(0) & "..."
        
        ' Definice listu, SQL dotazu a názvu tabulky
        tableName = queries(i)(2)
        Set ws = Nothing
        
        ' Zkontrolovat, zda list existuje
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tableName)
        On Error GoTo 0
        
        ' Pokud list neexistuje, vytvoø ho
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = tableName
        End If
        
        ' Vyèištìní existujících dat kromì hlavièek
        ws.Cells.Clear
        
        sql = queries(i)(1)
        
        ' Vytvoøení pøipojení pomocí funkce CreateConnection
        Set conn = CreateConnection()
        
        ' Vytvoøení nového Recordsetu pro naètení hlavièek
        Set rs = CreateObject("ADODB.Recordset")
        
        ' Otevøení Recordsetu pro naètení hlavièek
        rs.Open sql, conn, 1, 1 ' Použití 1, 1 pro naètení hlavièek (jen pro ètení, pohyb dopøedu)
        
        ' Naètení hlavièek
        Dim col As Integer
        For col = 0 To rs.Fields.Count - 1
            ws.Cells(1, col + 1).Value = rs.Fields(col).Name
        Next col
        
        rs.Close
        
        ' Vytvoøení nového Recordsetu pro naètení dat
        Set rs = CreateObject("ADODB.Recordset")
        
        ' Otevøení Recordsetu pro naètení dat
        rs.Open sql, conn, 1, 1 ' Použití 1, 1 pro naètení dat (jen pro ètení, pohyb dopøedu)
        
        ' Naètení dat do listu od druhého øádku
        If Not rs.EOF Then
            ws.Range("A2").CopyFromRecordset rs
        End If
        
        ' Vytvoøení tabulky z dat
        If ws.ListObjects.Count = 0 Then
            Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
            tbl.Name = tableName
        Else
            Set tbl = ws.ListObjects(1)
            tbl.Resize ws.Range("A1").CurrentRegion
        End If
        
        ' Uzavøení Recordsetu
        rs.Close
        Set rs = Nothing
        
        ' Uzavøení pøipojení
        conn.Close
        Set conn = Nothing
        
        Set ws = Nothing
        
        ' Zobrazit èas naèítání pro aktuální dotaz
        Application.StatusBar = "Naèítání dat pro " & queries(i)(0) & " dokonèeno za " & Format(Timer - startTime, "0.00") & " sekund."
        ' Reset timer for next query
        ' startTime = Timer
    
        ' Aktualizace pokrokového ukazatele po naètení dat
        progress = (i + 1 - LBound(queries)) / (UBound(queries) - LBound(queries) + 1)
        frmProgress.UpdateProgressBar progress
        DoEvents ' Umožní aktualizaci UserFormu
    Next i
    
    ToggleScreenUpdating False
    
    ' Povolení událostí
    Application.EnableEvents = True
    
    ' Nastavení výpoèetního režimu na automatický
    Application.Calculation = xlCalculationAutomatic
    
    ' Vynucení pøepoètu celého sešitu
    Application.CalculateFull
    
    Call DocasnaZmenaBunky
           
    ' Kód, který se provede vždy, i když dojde k chybì
        
Cleanup:
        
        ToggleScreenUpdating False
        
        ' Skrytí UserForm
        Unload frmProgress
        
        ' MsgBox "Naèítání dat dokonèeno za " & Format(Timer - startTime, "0.00") & " sekund.", vbInformation
        
        Exit Sub

    ' Ošetøení chyb
ErrorHandler:
        If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
        If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
        MsgBox "Došlo k chybì pøi naèítání dat: " & Err.Description, vbExclamation
        Resume Cleanup
End Sub

Sub LoadAccountDetails()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim tableName As String
    Dim currentCell As Range
    Dim groupName As String
    Dim year As Integer
    Dim month As Integer
    
    ' Vypnutí vykreslování a aktualizace
    ToggleScreenUpdating True
    
    ' Získání aktuální buòky
    Set currentCell = ActiveCell
    
    ' Zkontrolovat, zda jsou sloupce v rozmezí S až AD (19 až 30)
    If currentCell.Column < 19 Or currentCell.Column > 30 Then
        MsgBox "Data mimo povolený rozsah sloupcù S až AD.", vbExclamation
        Exit Sub
    End If
    
    ' Získání názvu skupiny úètù
    groupName = currentCell.Offset(0, -currentCell.Column + 2).Value
    
    ' Získání roku a mìsíce
    year = Cells(4, currentCell.Column).Value
    month = Cells(5, currentCell.Column).Value
    
    ' Vytvoøení pøipojení pomocí funkce CreateConnection
    Set conn = CreateConnection()
    
    ' Definice SQL dotazu
    sql = "SELECT Datum, Firma, Zamestnanec, Ucet, Nazev, ISNULL(CastkaMD,0) AS CastkaMD, ISNULL(CastkaDAL,0) AS CastkaDAL, Popis " & _
          "FROM hvw_ReportRozpocetPlneniDetail " & _
          "WHERE SkupinaUctu = '" & groupName & "' AND Rok = " & year & " AND Mesic = " & month
    
    ' Vytvoøení nového Recordsetu
    Set rs = CreateObject("ADODB.Recordset")
    
    rs.Open sql, conn, 1, 1
    
    ' Zkontrolovat, zda existuje list "Detail" a pøípadnì ho smazat
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("Detail")
    
    If Not newSheet Is Nothing Then
        Application.DisplayAlerts = False
        newSheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Vytvoøení nového listu "Detail"
    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    newSheet.Name = "Detail"
    
    ' Naètení hlavièek
    Dim col As Integer
    For col = 0 To rs.Fields.Count - 1
        newSheet.Cells(1, col + 1).Value = rs.Fields(col).Name
        newSheet.Cells(1, 1).AutoFilter
    Next col
    
    ' Naètení dat do listu od druhého øádku
    If Not rs.EOF Then
        newSheet.Range("A2").CopyFromRecordset rs
    End If
    
    ' Uzavøení Recordsetu a pøipojení
    rs.Close
    conn.Close
    
    ' Vyèištìní
    Set rs = Nothing
    Set conn = Nothing
    
    ' Vytvoøení tabulky z dat
    Dim tbl As ListObject
    Set tbl = newSheet.ListObjects.Add(xlSrcRange, newSheet.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "DetailTable"
 
    tbl.Range.Columns.AutoFit
    ' Nastavení filtru
    tbl.TableStyle = "TableStyleLight1" ' Mùžeš zmìnit na jiný styl dle potøeby

    ' Nastavení formátu sloupcù
    With tbl
        .ListColumns("Datum").Range.NumberFormat = "dd.mm.yyyy" ' Formát datumu
        .ListColumns("CastkaMD").Range.NumberFormat = "#,##0.00" ' Formát èísla s dvìma desetinnými místy
        .ListColumns("CastkaDAL").Range.NumberFormat = "#,##0.00" ' Formát èísla s dvìma desetinnými místy
    End With
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    
    ' Zapnutí vykreslování a aktualizace
    ToggleScreenUpdating False
    
End Sub

Sub KumulujVysledovkuPodleCheckboxu()
Attribute KumulujVysledovkuPodleCheckboxu.VB_ProcData.VB_Invoke_Func = "K\n14"
    Dim wsAplikace As Worksheet
    Dim wsKumulace As Worksheet
    Dim posledniSloupec As Long
    Dim posledniRadek As Long
    Dim kumulativniHodnota As Double
    Dim i As Long
    Dim radek As Long
    Dim aktualniRok As Integer
    Dim aktualniMesic As Integer
    Dim rokSloupce As Variant
    Dim mesicSloupce As Variant
    Dim datumUzaverka As Date
    Dim prvniSloupecPlan As Long
    Dim posledniSloupecPlan As Long
    Dim prvniSloupecSkutecnost As Long
    Dim posledniSloupecSkutecnost As Long
    Dim prvniSloupecRozdil As Long
    Dim sloupecNeprazdny1 As Long
    Dim sloupecNeprazdny2 As Long
    Dim celkovyPocetRadku As Long
    
    'Debug.Print "Je v bìhu nápoèet pro kumulace..."

    ' Odkazy na listy
    Set wsAplikace = ThisWorkbook.Sheets("Aplikace")
    Set wsKumulace = ThisWorkbook.Sheets("Kumulace")

    ' Urèení sloupcù pro Plán, Skuteènost a rozdíl
    prvniSloupecPlan = 6
    posledniSloupecPlan = 17
    prvniSloupecSkutecnost = 19
    posledniSloupecSkutecnost = 30
    prvniSloupecRozdil = 32
    posledniSloupec = 43

    ' Definování sloupcù, které nemají být smazány (napø. sloupec 18 a 31)
    sloupecNeprazdny1 = 18
    sloupecNeprazdny2 = 31

    ' Urèení posledního sloupce (mìsíce) v hlavièce na øádku 5
    ' posledniSloupec = wsAplikace.Cells(5, wsAplikace.Columns.Count).End(xlToLeft).Column

    ' Urèení posledního øádku s daty na listu Kumulace
    posledniRadek = wsKumulace.Cells(wsKumulace.Rows.Count, 4).End(xlUp).Row ' Pøedpokládáme, že poslední øádek s checkboxy je urèen ve sloupci D

    ' Urèení aktuálního roku a mìsíce (pøedchozí uzavøený mìsíc)
    datumUzaverka = DateSerial(year(Now), month(Now), 1) - 1 ' Urèíme poslední den pøedchozího mìsíce
    aktualniRok = year(datumUzaverka)
    aktualniMesic = month(datumUzaverka)

    ' Optimalizace - vypnutí pøepoèítávání a pøekreslování
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Inicializace a zobrazení formuláøe pro prùbìh
    frmProgress.lblProgress.Caption = "Naèítání dat, prosím èekejte..."
    frmProgress.lblProgressBar.Width = 0
    frmProgress.Show vbModeless
    DoEvents

    ' Celkový poèet øádkù pro zpracování
    celkovyPocetRadku = posledniRadek - 5

    ' Procházení každého øádku od 6 do posledního øádku s checkboxem
    For radek = 6 To posledniRadek
        ' Aktualizace ukazatele prùbìhu
        frmProgress.UpdateProgressBar (radek - 5) / celkovyPocetRadku

        ' Kontrola, zda je checkbox (pravda/true) ve sloupci D
        If wsKumulace.Cells(radek, 4).Value = True Then
            kumulativniHodnota = 0 ' Reset kumulativní hodnoty pro nový øádek

            ' Kumulace a zápis pro plán
            For i = prvniSloupecPlan To posledniSloupecPlan
                ' Urèení roku a mìsíce ze sloupce
                rokSloupce = wsAplikace.Cells(4, i).Value
                mesicSloupce = wsAplikace.Cells(5, i).Value

                ' Kontrola, zda jsou rok a mìsíc èíselné hodnoty
                If IsNumeric(rokSloupce) And IsNumeric(mesicSloupce) Then
                    ' Kontrola, zda je mìsíc uzavøen (rok i mìsíc jsou menší nebo rovné aktuálnímu)
                    If (rokSloupce < aktualniRok) Or (rokSloupce = aktualniRok And mesicSloupce <= aktualniMesic) Then
                        ' Pøiètení hodnoty do kumulativní promìnné (jen pokud je hodnota èíselná)
                        If IsNumeric(wsAplikace.Cells(radek, i).Value) Then
                            kumulativniHodnota = kumulativniHodnota + wsAplikace.Cells(radek, i).Value
                        End If

                        ' Zapsání kumulativní hodnoty do stejného sloupce na list Kumulace
                        wsKumulace.Cells(radek, i).Value = kumulativniHodnota
                    End If
                End If
            Next i
            
            ' Reset kumulativní hodnoty pro skuteènost
            kumulativniHodnota = 0
            
            ' Kumulace a zápis pro skuteènost
            For i = prvniSloupecSkutecnost To posledniSloupecSkutecnost
                rokSloupce = wsAplikace.Cells(4, i).Value
                mesicSloupce = wsAplikace.Cells(5, i).Value

                If IsNumeric(rokSloupce) And IsNumeric(mesicSloupce) Then
                    If (rokSloupce < aktualniRok) Or (rokSloupce = aktualniRok And mesicSloupce <= aktualniMesic) Then
                        If IsNumeric(wsAplikace.Cells(radek, i).Value) Then
                            kumulativniHodnota = kumulativniHodnota + wsAplikace.Cells(radek, i).Value
                        End If

                        wsKumulace.Cells(radek, i).Value = kumulativniHodnota
                    End If
                End If
            Next i
            
            ' Reset kumulativní hodnoty pro rozdíly
            kumulativniHodnota = 0
            
            ' Kumulace a zápis pro rozdíly
            For i = prvniSloupecRozdil To posledniSloupec
                rokSloupce = wsAplikace.Cells(4, i).Value
                mesicSloupce = wsAplikace.Cells(5, i).Value

                If IsNumeric(rokSloupce) And IsNumeric(mesicSloupce) Then
                    If (rokSloupce < aktualniRok) Or (rokSloupce = aktualniRok And mesicSloupce <= aktualniMesic) Then
                        If IsNumeric(wsAplikace.Cells(radek, i).Value) Then
                            kumulativniHodnota = kumulativniHodnota + wsAplikace.Cells(radek, i).Value
                        End If

                        wsKumulace.Cells(radek, i).Value = kumulativniHodnota
                    End If
                End If
            Next i

        ElseIf wsKumulace.Cells(radek, 4).Value = False Then ' Pouze pokud je hodnota checkboxu Explicitnì False
            For i = prvniSloupecPlan To posledniSloupec
                ' Vynechání sloupcù, které nesmí být smazány
                If i <> sloupecNeprazdny1 And i <> sloupecNeprazdny2 Then
                    wsKumulace.Cells(radek, i).ClearContents
                End If
            Next i
        End If
    Next radek

    ' Zavøení formuláøe pro prùbìh
    Unload frmProgress

    ' Obnovení nastavení pøepoèítávání a pøekreslování
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    'MsgBox "Kumulace byla úspìšnì provedena pro vybrané øádky.", vbInformation
End Sub

Sub CheckAll()
Attribute CheckAll.VB_ProcData.VB_Invoke_Func = "a\n14"

    Dim wsKumulace As Worksheet
    Dim posledniRadek As Long
    Dim radek As Long
    Dim hodnotaBuòky As Variant

    ' Odkaz na list Kumulace
    Set wsKumulace = ThisWorkbook.Sheets("Kumulace")

    ' Urèení posledního øádku s daty ve sloupci D na listu Kumulace
    posledniRadek = wsKumulace.Cells(wsKumulace.Rows.Count, 4).End(xlUp).Row

    ' Procházení každého øádku od 7 do posledního øádku s checkboxem
    For radek = 6 To posledniRadek
        hodnotaBuòky = wsKumulace.Cells(radek, 4).Value

        ' Kontrola, zda buòka obsahuje hodnotu True nebo False
        If IsEmpty(hodnotaBuòky) Then
            '
        Else: hodnotaBuòky = True Or hodnotaBuòky = False
            wsKumulace.Cells(radek, 4).Value = True
        End If

    Next radek
    
End Sub

Sub UncheckAll()

    Dim wsKumulace As Worksheet
    Dim posledniRadek As Long
    Dim radek As Long
    Dim hodnotaBuòky As Variant

    ' Odkaz na list Kumulace
    Set wsKumulace = ThisWorkbook.Sheets("Kumulace")

    ' Urèení posledního øádku s daty ve sloupci D na listu Kumulace
    posledniRadek = wsKumulace.Cells(wsKumulace.Rows.Count, 4).End(xlUp).Row

    ' Procházení každého øádku od 7 do posledního øádku s checkboxem
    For radek = 6 To posledniRadek
        hodnotaBuòky = wsKumulace.Cells(radek, 4).Value

        ' Kontrola, zda buòka obsahuje hodnotu True nebo False
        If IsEmpty(hodnotaBuòky) Then
            '
        Else: hodnotaBuòky = True Or hodnotaBuòky = False
            wsKumulace.Cells(radek, 4).Value = False
        End If
        
    Next radek
    
End Sub

Sub ToggleSingleCheckBox(ByVal radek As Long)
    Dim ws As Worksheet
    Dim posledniRadek As Long
    Dim i As Long

    ' Odkaz na list Kumulace
    Set ws = ThisWorkbook.Sheets("Kumulace")
    
    ' Vypnutí vykreslování a aktualizace
    Application.ScreenUpdating = False

    ' Urèení posledního øádku s daty ve sloupci C na listu Kumulace
    posledniRadek = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row

    ' Nejprve zrušíme všechny checkboxy ve sloupci C
    For i = 7 To posledniRadek
        If VarType(ws.Cells(i, 3).Value) = vbBoolean And ws.Cells(i, 3).Value = True Then
            ws.Cells(i, 3).Value = False
        End If
    Next i

    ' Nastavení hodnoty True pro vybraný øádek
    ws.Cells(radek, 3).Value = True
    
    ' Pokud je sloupec D False, nastavíme True a zavoláme proceduru
    If ws.Cells(radek, 4).Value = False Then
        ws.Cells(radek, 4).Value = True
        'Debug.Print "Zavolám KumulujVysledovkuPodleCheckboxu"
        Call KumulujVysledovkuPodleCheckboxu
    End If
    
    ' Aktualizace grafu a uložení obrázku
    'Debug.Print "Zavolám CopyDataToDataProGraf(radek)"
    Call CopyDataToDataProGraf(radek)
    
    ' Zapnutí vykreslování a aktualizace
    Application.ScreenUpdating = True
    
    'Debug.Print "Zavolám Call SaveChartAsImage"
    Call SaveChartAsImage
    
End Sub


Sub CopyDataToDataProGraf(ByVal radek As Long)
    Dim wsAplikace As Worksheet
    Dim wsKumulace As Worksheet
    Dim wsDataProGraf As Worksheet
    
    ' Odkazy na jednotlivé listy
    Set wsAplikace = ThisWorkbook.Sheets("Aplikace")
    Set wsKumulace = ThisWorkbook.Sheets("Kumulace")
    Set wsDataProGraf = ThisWorkbook.Sheets("DataProGraf")
    
    ' Kopírování z listu Aplikace do DataProGraf
    ' 1. Zkopírujeme B4:AQ4 na øádek 1
    wsAplikace.Range("B4:AQ4").Copy Destination:=wsDataProGraf.Range("B1")
    
    ' 2. Zkopírujeme B5:AQ5 na øádek 2
    wsAplikace.Range("B5:AQ5").Copy Destination:=wsDataProGraf.Range("B2")
    
    ' 3. Zkopírujeme B(radek):AQ(radek) na øádek 3 jako hodnoty
    wsDataProGraf.Range("B3:AQ3").Value = wsAplikace.Range(wsAplikace.Cells(radek, 2), wsAplikace.Cells(radek, 43)).Value
    
    ' Kopírování z listu Kumulace do DataProGraf
    ' 4. Zkopírujeme B(radek):AQ(radek) na øádek 4 jako hodnoty
    wsDataProGraf.Range("B4:AQ4").Value = wsKumulace.Range(wsKumulace.Cells(radek, 2), wsKumulace.Cells(radek, 43)).Value
    
    wsDataProGraf.Range("A4").Formula = "=DataProGraf!$B$4 & "" - kumulativnì"""
    
End Sub



