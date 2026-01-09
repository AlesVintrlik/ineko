Attribute VB_Name = "Data"

Sub LoadOrUpdateData()
    Dim adoConn As Object
    Dim ws As Worksheet
    Dim tableName As String
    Dim sql As String
    Dim adoRS As Object
    Dim i As Integer
    Dim j As Integer
    Dim progress As Double
    Dim targetRange As Range
    
    On Error GoTo ErrorHandler
    WriteLog "LoadOrUpdateData started."

    ' Vypnout aktualizaci obrazovky pro zrychlení
    Application.ScreenUpdating = False
    
    ' Zobraz formuláø pouze tehdy, pokud ještì nebìží
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = "frmProgress" Then GoTo FormAlreadyLoaded
    Next frm
    
    ' Pokud formuláø nebìží, naèti ho a zobraz
    Load frmProgress
    frmProgress.Show vbModeless

FormAlreadyLoaded:
    
    ' Definice listù a SQL dotazù
    Dim queries As Variant
    queries = Array( _
        Array("Zakazky", "SELECT * FROM hvw_TerminyZakazekProPlanovani", "Zakazky"), _
        Array("Operace", "SELECT * FROM hvw_VyrobniPozadavkyProPlanovani", "Operace"), _
        Array("Kapacity", "SELECT * FROM hvw_KapacityProPlanovani", "Kapacity"))
    WriteLog "SQL queries defined."
    
    ' Vytvoøení ADO pøipojení
    Set adoConn = CreateConnection()
    If adoConn Is Nothing Then
        WriteLog "Error: Nepodaøilo se otevøít pøipojení."
        MsgBox "Nepodaøilo se otevøít pøipojení.", vbCritical
        Unload frmProgress
        GoTo Cleanup
    End If
    WriteLog "ADO connection established."
    
    For i = LBound(queries) To UBound(queries)
        tableName = queries(i)(2)
        sql = queries(i)(1)
        
        ' Aktualizace prùbìhu
        progress = (i + 1) / (UBound(queries) + 1)
        frmProgress.UpdateProgressBar progress
        WriteLog "Processing sheet: " & tableName & " (" & progress * 100 & "%)."
        
        ' Najít nebo vytvoøit list
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(tableName)
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            ws.Name = tableName
            WriteLog "Created sheet: " & tableName
        Else
            WriteLog "Found sheet: " & tableName
        End If
        
        ' Vymazání pouze oblasti dat (nikoli celého listu)
        If ws.UsedRange.Cells.Count > 1 Then
            ws.UsedRange.ClearContents
            WriteLog "Cleared UsedRange on sheet: " & tableName
        Else
            WriteLog "No data to clear on sheet: " & tableName
        End If
        
        ' Naèíst data z SQL dotazu pomocí ADO
        Set adoRS = CreateObject("ADODB.Recordset")
        WriteLog "Executing SQL for sheet " & tableName & ": " & sql
        adoRS.Open sql, adoConn
        
        ' Kontrola, zda dotaz vrátil nìjaká data
        If Not adoRS.EOF Then
            ' Vložit názvy sloupcù
            For j = 1 To adoRS.Fields.Count
                ws.Cells(1, j).Value = adoRS.Fields(j - 1).Name
            Next j
            WriteLog "Column names written on sheet: " & tableName
            
            ' Vložit data do listu
            ws.Cells(2, 1).CopyFromRecordset adoRS
            WriteLog "Data copied to sheet: " & tableName
        Else
            WriteLog "No data returned for query on sheet: " & tableName
        End If
        
        ' Zavøít Recordset
        adoRS.Close
        Set adoRS = Nothing
        
        ' Automatické pøizpùsobení šíøky sloupcù
        ws.Columns.AutoFit
        WriteLog "AutoFit applied on sheet: " & tableName
    Next i
    
Cleanup:
    ' Uzavøení ADO pøipojení
    If Not adoConn Is Nothing Then
        adoConn.Close
        Set adoConn = Nothing
        WriteLog "ADO connection closed."
    End If

    ' Obnovení aktualizace obrazovky
    Application.ScreenUpdating = True
    WriteLog "Screen updating restored."
    
    Call AktualizovatSeznamZakazek
    WriteLog "Call AktualizovatSeznamZakazek executed."
    
    ' Zavolání funkce PridejNazevUseku, pokud existuje
    On Error Resume Next
        Call PridejNazevUseku
        WriteLog "Call PridejNazevUseku executed."
    On Error GoTo 0

    WriteLog "LoadOrUpdateData finished."
    Exit Sub

ErrorHandler:
    WriteLog "Error: " & Err.Description
    MsgBox "Došlo k chybì: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Sub PridejNazevUseku()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim j As Long
    
    WriteLog "Executing PridejNazevUseku."
    ' Ujistìte se, že list Kapacity je správnì nastaven
    Set ws = ThisWorkbook.Sheets("Kapacity")
    
    ' Najdìte poslední použitý øádek ve sloupci A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Nastavte název sloupce v prvním øádku sloupce H
    ws.Cells(1, "H").Value = "NazevUseku"
    
    ' Pøidání vzorce do sloupce H od druhého øádku po poslední øádek
    For j = 2 To lastRow
        ws.Cells(j, "H").Formula = "=IFS(A" & j & "=0,""Ostatní"",A" & j & "=1,""Pøíprava"",A" & j & "=2,""Montáž"",A" & j & "=3,""Elektro"",A" & j & "=4,""Svaøování"",A" & j & "=5,""Segmenty"",A" & j & "=7,""Rozvadìèe"",A" & j & "=8,""Kontrola"")"
    Next j
    WriteLog "PridejNazevUseku executed for " & (lastRow - 1) & " rows."
End Sub

Sub CreateNamedRanges()
    Dim ws As Worksheet
    Dim namedRanges As Variant
    Dim addresses As Variant
    Dim i As Integer

    WriteLog "Executing CreateNamedRanges."
    ' Definice pojmenovaných oblastí a jejich adres
    namedRanges = Array("Kapacity.Usek", _
                        "DenniKapacitaHod", _
                        "Kapacity.Pracoviste", _
                        "PracovisteNazev", _
                        "NazevUseku", _
                        "TydenniSkluz", _
                        "PocetDniVtydnu", _
                        "DenniSkluz", _
                        "SkluzDoLonska", _
                        "Korekce", _
                        "Usek", _
                        "Zakazka", _
                        "Pracoviste", _
                        "PotrebnyCasCelkemHod", _
                        "TerminVyrobyRok", _
                        "TerminVyrobyTyden", _
                        "Zakazky.Zakazka", _
                        "Zakazky.Firma", _
                        "Zakazky.DatumZahajeni", _
                        "Zakazky.DatumExpedice", _
                        "MnozstviPozadovane", _
                        "MnozstviZive")
    addresses = Array("=Kapacity!$A:$A", _
                        "=Kapacity!$G:$G", _
                        "=Kapacity!$B:$B", _
                        "=Kapacity!$C:$C", _
                        "=Kapacity!$H:$H", _
                        "=Konfigurace!$C$2", _
                        "=Konfigurace!$C$3", _
                        "=Konstanty!$C$2", _
                        "=Konstanty!$C$3", _
                        "=Kontrola!$E:$E", _
                        "=Operace!$A:$A", _
                        "=Operace!$B:$B", _
                        "=Operace!$G:$G", _
                        "=Operace!$R:$R", _
                        "=Operace!$T:$T", _
                        "=Operace!$U:$U", _
                        "=Zakazky!$A:$A", _
                        "=Zakazky!$B:$B", _
                        "=Zakazky!$C:$C", _
                        "=Zakazky!$D:$D", _
                        "=Zakazky!$E:$E", _
                        "=Zakazky!$F:$F")
    
    For i = LBound(namedRanges) To UBound(namedRanges)
        On Error Resume Next
        Dim nm As Name
        Set nm = ThisWorkbook.Names(namedRanges(i))
        On Error GoTo 0
        
        If nm Is Nothing Then
            ThisWorkbook.Names.Add Name:=namedRanges(i), RefersTo:=addresses(i)
            Debug.Print "Vytvoøeno: " & namedRanges(i) & " s adresou " & addresses(i)
            WriteLog "Created named range: " & namedRanges(i) & " with address " & addresses(i)
        Else
            Debug.Print "Existuje: " & namedRanges(i) & " s adresou " & nm.RefersTo
            WriteLog "Named range already exists: " & namedRanges(i) & " with address " & nm.RefersTo
        End If
        
        ' Resetování promìnné nm na Nothing pro další iterace
        Set nm = Nothing
    Next i
    
    MsgBox "Pojmenované oblasti byly úspìšnì vytvoøeny.", vbInformation
    WriteLog "CreateNamedRanges finished."
End Sub

Sub RefreshAllConnections()
    WriteLog "Executing RefreshAllConnections."
    ' Aktualizace všech pøipojení v sešitu
    ThisWorkbook.RefreshAll
    MsgBox "Data byla úspìšnì aktualizována.", vbInformation
    WriteLog "RefreshAllConnections finished."
End Sub

' Logovací funkce pro zápis zpráv do souboru PlanovaniLog.txt s kontrolou velikosti a s UTF-8 kódováním
Sub WriteLog(msg As String)
    Dim filePath As String
    Dim maxSize As Long
    Dim newFilePath As String
    Dim stream As Object
    
    filePath = ThisWorkbook.Path & "\PlanovaniLog.txt"
    maxSize = 1000000 ' Limit 1 MB – uprav dle potøeby

    ' Pokud soubor existuje, zkontroluj jeho velikost
    If Dir(filePath) <> "" Then
        If FileLen(filePath) > maxSize Then
            ' Archivuj starý log – nový název obsahuje datum a èas
            newFilePath = ThisWorkbook.Path & "\PlanovaniLog_backup_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
            Name filePath As newFilePath
        End If
    End If

    ' Použij ADODB.Stream pro zápis s UTF-8 kódováním
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 2 ' Text
        .Charset = "utf-8"
        
        ' Pokud soubor už existuje, naèti jeho obsah a pøejdi na konec, jinak vytvoø nový stream
        .Open
        If Dir(filePath) <> "" Then
            .LoadFromFile filePath
            .Position = .Size ' nastavení na konec pro doplnìní textu
        End If
        ' Zapsání øádku s èasovým razítkem a novou zprávou
        .WriteText Format(Now, "yyyy-mm-dd hh:nn:ss") & " - " & msg, 1
        ' Uložení zmìn do souboru (2 = adSaveCreateOverWrite)
        .SaveToFile filePath, 2
        .Close
    End With
    Set stream = Nothing
End Sub

Sub DumpRangeFormulas(rng As Range)
    Dim arr As Variant
    Dim i As Long, j As Long
    Dim rowStr As String
    
    arr = rng.Formula
    For i = 1 To UBound(arr, 1)
        rowStr = ""
        For j = 1 To UBound(arr, 2)
            rowStr = rowStr & arr(i, j) & vbTab
        Next j
        ' Logování s zalomením: nejprve hlavièka s øádkem, poté obsah na nový øádek.
        WriteLog "DumpRangeFormulas - Row " & i & ":" & vbCrLf & rowStr
    Next i
End Sub



