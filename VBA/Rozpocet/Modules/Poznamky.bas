Attribute VB_Name = "Poznamky"

Option Explicit

Sub SaveToINI()
    Dim iniPath As String
    Dim lastRow As Long
    Dim cell As Range
    Dim section As String
    Dim outputLine As String
    Dim stream As Object
    
    ' Cesta k INI souboru
    iniPath = ThisWorkbook.Path & "\data.ini"
    section = "Poznamky"
    
    ' Najít poslední øádek ve sloupci AS, ale zaèít minimálnì od øádku 6
    lastRow = Application.WorksheetFunction.Max(6, Cells(Rows.Count, "AS").End(xlUp).Row)

    
    ' Vytvoøení ADODB.Stream pro zápis v UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Textový stream
    stream.Charset = "utf-8" ' Nastavení kódování na UTF-8
    stream.Open
    
    ' Zápis sekce
    stream.WriteText "[" & section & "]" & vbCrLf
    
    ' Zápis hodnot od buòky AS6 dál
    For Each cell In Range("AS6:AS" & lastRow)
        If cell.Value <> "" Then
            outputLine = "R" & cell.Row & "=" & cell.Value
            stream.WriteText outputLine & vbCrLf
        End If
    Next cell
    
    ' Uložení do souboru
    stream.SaveToFile iniPath, 2 ' AdSaveCreateOverWrite
    stream.Close
    
End Sub

Sub LoadFromINI()
    Dim iniPath As String
    Dim stream As Object
    Dim fileContent As String
    Dim lines() As String
    Dim lineData() As String
    Dim i As Long
    Dim sectionFound As Boolean
    Dim rowNumber As Long

    ' Cesta k INI souboru
    iniPath = ThisWorkbook.Path & "\data.ini"

    ' Ovìøení existence INI souboru
    If Dir(iniPath) = "" Then
        MsgBox "INI soubor nebyl nalezen: " & iniPath, vbExclamation
        Exit Sub
    End If

    ' Použití ADODB.Stream pro naètení souboru v UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Textový stream
    stream.Charset = "utf-8" ' Nastavení na UTF-8
    stream.Open
    stream.LoadFromFile iniPath

    ' Naètení obsahu jako text
    fileContent = stream.ReadText
    stream.Close
    Set stream = Nothing

    ' Rozdìlení obsahu na jednotlivé øádky
    lines = Split(fileContent, vbCrLf)
    sectionFound = False

    ' Zpracování øádkù
    For i = LBound(lines) To UBound(lines)
        If Trim(lines(i)) = "[Poznamky]" Then
            sectionFound = True
        ElseIf sectionFound And Trim(lines(i)) <> "" And Left(Trim(lines(i)), 1) <> "[" Then
            lineData = Split(lines(i), "=")
            If UBound(lineData) = 1 Then
                rowNumber = Val(Replace(lineData(0), "R", ""))
                If rowNumber > 0 Then
                    ' Naèíst hodnoty do sloupce AS na listu "Aplikace"
                    ThisWorkbook.Sheets("Aplikace").Cells(rowNumber, "AS").Value = lineData(1)
                    Debug.Print "Naèteno: Adresa R" & rowNumber & ", Hodnota: " & lineData(1)
                Else
                    Debug.Print "Neplatná adresa v øádku: " & lines(i)
                End If
            Else
                Debug.Print "Neplatný formát øádku: " & lines(i)
            End If
        ElseIf sectionFound And Left(Trim(lines(i)), 1) = "[" Then
            Exit For ' Konec sekce
        End If
    Next i
End Sub

