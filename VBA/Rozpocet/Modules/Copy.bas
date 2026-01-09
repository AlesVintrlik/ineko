Attribute VB_Name = "Copy"

Sub KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty()
Attribute KopirovatRozsahyDoNovehoSesituAPrevestNaHodnoty.VB_ProcData.VB_Invoke_Func = "k\n14"
    Dim wsAplikace As Worksheet
    Dim wsKumulace As Worksheet
    Dim wbNovy As Workbook
    Dim posledniRadekAplikace As Long
    Dim posledniRadekKumulace As Long
    Dim rozsahAplikace As Range
    Dim rozsahKumulace As Range
    Dim sloupec As Long
    Dim tempRadek As Long
    Dim wsNovyAplikace As Worksheet
    Dim wsNovyKumulace As Worksheet
    Dim ws As Worksheet
    
    ' Nastavení odkazù na listy
    Set wsAplikace = ThisWorkbook.Sheets("Aplikace")
    Set wsKumulace = ThisWorkbook.Sheets("Kumulace")
    
    ' Zjištìní posledního øádku s daty ve sloupcích B:AQ na listu Aplikace
    posledniRadekAplikace = 4 ' Zaèínáme od øádku 4
    For sloupec = wsAplikace.Range("B1").Column To wsAplikace.Range("AQ1").Column
        tempRadek = wsAplikace.Cells(wsAplikace.Rows.Count, sloupec).End(xlUp).Row
        If tempRadek > posledniRadekAplikace Then posledniRadekAplikace = tempRadek
    Next sloupec
    
    ' Nastavení rozsahu ke kopírování na listu Aplikace
    Set rozsahAplikace = wsAplikace.Range(wsAplikace.Cells(4, "B"), wsAplikace.Cells(posledniRadekAplikace, "AQ"))
    
    ' Zjištìní posledního øádku s daty ve sloupcích B:AQ na listu Kumulace
    posledniRadekKumulace = 4 ' Zaèínáme od øádku 4
    For sloupec = wsKumulace.Range("B1").Column To wsKumulace.Range("AQ1").Column
        tempRadek = wsKumulace.Cells(wsKumulace.Rows.Count, sloupec).End(xlUp).Row
        If tempRadek > posledniRadekKumulace Then posledniRadekKumulace = tempRadek
    Next sloupec
    
    ' Nastavení rozsahu ke kopírování na listu Kumulace
    Set rozsahKumulace = wsKumulace.Range(wsKumulace.Cells(4, "B"), wsKumulace.Cells(posledniRadekKumulace, "AQ"))
    
    ' Vytvoøení nového sešitu s jedním listem
    Set wbNovy = Workbooks.Add(xlWBATWorksheet)
    
    ' Pøejmenování výchozího listu na "Aplikace"
    Set wsNovyAplikace = wbNovy.Sheets(1)
    wsNovyAplikace.Name = "Aplikace"
    
    ' Kopírování šíøek sloupcù z "Aplikace"
    wsAplikace.Range("B:AQ").Copy
    wsNovyAplikace.Range("B1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    
    ' Kopírování dat z "Aplikace"
    rozsahAplikace.Copy Destination:=wsNovyAplikace.Cells(4, "B")
    
    ' Pøevod vzorcù na hodnoty v "Aplikace"
    With wsNovyAplikace
        .Range(.Cells(4, "B"), .Cells(posledniRadekAplikace, "AQ")).Value = .Range(.Cells(4, "B"), .Cells(posledniRadekAplikace, "AQ")).Value
    End With
    
    ' Pøidání nového listu pro "Kumulace"
    Set wsNovyKumulace = wbNovy.Sheets.Add(After:=wbNovy.Sheets(wbNovy.Sheets.Count))
    wsNovyKumulace.Name = "Kumulace"
    
    ' Kopírování šíøek sloupcù z "Kumulace"
    wsKumulace.Range("B:AQ").Copy
    wsNovyKumulace.Range("B1").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False
    
    ' Kopírování dat z "Kumulace"
    rozsahKumulace.Copy Destination:=wsNovyKumulace.Cells(4, "B")
    
    ' Pøevod vzorcù na hodnoty v "Kumulace"
    With wsNovyKumulace
        .Range(.Cells(4, "B"), .Cells(posledniRadekKumulace, "AQ")).Value = .Range(.Cells(4, "B"), .Cells(posledniRadekKumulace, "AQ")).Value
    End With
    
    ' Odstranìní výchozího listu, pokud zùstal
    Application.DisplayAlerts = False
    For Each ws In wbNovy.Worksheets
        If ws.Name <> "Aplikace" And ws.Name <> "Kumulace" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ' Uložení nového sešitu s dialogem pro výbìr umístìní
    Dim ulozitJako As Variant
    ulozitJako = Application.GetSaveAsFilename(InitialFileName:="Kopie rozpoètu ze dne " & Format(Date, "yyyymmdd") & ".xlsx", FileFilter:="Excel soubory (*.xlsx), *.xlsx")
    
    If ulozitJako <> False Then
        wbNovy.SaveAs Filename:=ulozitJako, FileFormat:=xlOpenXMLWorkbook
        MsgBox "Nový sešit byl úspìšnì vytvoøen a uložen.", vbInformation
    Else
        ' Zavøení nového sešitu bez uložení, pokud uživatel zruší uložení
        wbNovy.Close SaveChanges:=False
        MsgBox "Uživatel zrušil uložení nového sešitu.", vbExclamation
    End If
    
End Sub

