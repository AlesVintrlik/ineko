Attribute VB_Name = "Temp"
Option Explicit

Sub ShowTempFolder()

    MsgBox Environ("TEMP")
    
End Sub

Sub TestAssignMacroToShapes(targetSheet As Worksheet)
    Dim shp As Shape
    Dim remainingName As String
    Dim groupShape As Shape
    Dim i As Integer
    
    ' Hledáme skupinu "navigace" na zadaném listu
    On Error Resume Next
    Set groupShape = targetSheet.Shapes("navigace")
    On Error GoTo 0
    
    ' Pokud skupina neexistuje
    If groupShape Is Nothing Then
        'Debug.Print "Skupina 'navigace' nebyla nalezena na listu " & targetSheet.Name
        Exit Sub
    End If
    
    ' Kontrola, zda je to skupina
    If groupShape.Type = msoGroup Then
        ' Iterace pøes všechny objekty ve skupinì "navigace"
        For i = 1 To groupShape.GroupItems.Count
            Set shp = groupShape.GroupItems(i)
            
            ' Ovìøíme, že název objektu zaèíná na "ico_"
            If Left(shp.Name, 4) = "ico_" Then
                ' Zjistíme zbývající název objektu (od 5. znaku dále)
                remainingName = Mid(shp.Name, 5)
                ''Debug.Print "remainingName z TestAssignMacroToShapes je " & remainingName
                
                ' Simulace pøiøazení makra na základì názvu objektu a aktuálního listu
                Select Case LCase(remainingName) ' Používáme LCase pro ignorování velikosti písmen
                    Case "aplikace"
                        If targetSheet.Name = "Aplikace" Or targetSheet.Name = "Kumulace" Or targetSheet.Name = "Kontingenèní tabulka" Then
                            'Debug.Print "Pøiøadil bych NavigateSheet pro " & shp.Name
                        End If
                        
                    Case "load"
                        Select Case targetSheet.Name
                            Case "Aplikace"
                                'Debug.Print "Pøiøadil bych LoadDataFromQueries pro " & shp.Name
                            Case "Kumulace"
                                'Debug.Print "Pøiøadil bych KumulujVysledovkuPodleCheckboxu pro " & shp.Name
                            Case "Kontingenèní tabulka"
                                'Debug.Print "Pøiøadil bych TotoMakroNicNedela pro " & shp.Name
                        End Select
                        
                    Case "load_detail"
                        Select Case targetSheet.Name
                            Case "Aplikace"
                                'Debug.Print "Pøiøadil bych LoadAccountDetails pro " & shp.Name
                            Case "Kumulace", "Kontingenèní tabulka"
                                'Debug.Print "Pøiøadil bych TotoMakroNicNedela pro " & shp.Name
                        End Select
                        
                    Case "kontingenèní tabulka", "kumulace"
                        If targetSheet.Name = "Aplikace" Or targetSheet.Name = "Kumulace" Or targetSheet.Name = "Kontingenèní tabulka" Then
                            'Debug.Print "Pøiøadil bych NavigateSheet pro " & shp.Name
                        End If
                        
                End Select
            End If
        Next i
    Else
        'Debug.Print "'navigace' není skupina."
    End If
End Sub


Sub ZjistitSkryteSloupce()
    Dim ws As Worksheet
    Dim sloupec As Range
    Dim skryteSloupce As String
    
    ' Nastavení aktivního listu nebo mùžete specifikovat konkrétní list
    Set ws = ActiveSheet
    
    ' Inicializace promìnné pro uložení skrytých sloupcù
    skryteSloupce = ""
    
    ' Procházení všech sloupcù v použité oblasti (UsedRange)
    For Each sloupec In ws.UsedRange.Columns
        If sloupec.EntireColumn.Hidden = True Then
            skryteSloupce = skryteSloupce & sloupec.Column & ", "
        End If
    Next sloupec
    
    ' Kontrola, zda byly nalezeny skryté sloupce
    If skryteSloupce <> "" Then
        ' Odstranìní poslední èárky a mezery
        skryteSloupce = Left(skryteSloupce, Len(skryteSloupce) - 2)
        MsgBox "Skryté sloupce: " & skryteSloupce
    Else
        MsgBox "Nebyly nalezeny žádné skryté sloupce."
    End If
End Sub

Sub ListSheetProtectionProperties()
    Dim ws As Worksheet
    Dim protectionMsg As String
    
    For Each ws In ThisWorkbook.Worksheets
        protectionMsg = "List: " & ws.Name & vbNewLine
        
        If ws.ProtectContents Then
            protectionMsg = protectionMsg & " - Obsah je chránìn" & vbNewLine
        Else
            protectionMsg = protectionMsg & " - Obsah není chránìn" & vbNewLine
        End If
        
        If ws.ProtectDrawingObjects Then
            protectionMsg = protectionMsg & " - Kreslící objekty jsou chránìny" & vbNewLine
        Else
            protectionMsg = protectionMsg & " - Kreslící objekty nejsou chránìny" & vbNewLine
        End If
        
        If ws.ProtectScenarios Then
            protectionMsg = protectionMsg & " - Scénáøe jsou chránìny" & vbNewLine
        Else
            protectionMsg = protectionMsg & " - Scénáøe nejsou chránìny" & vbNewLine
        End If

        ' Výpis výsledku do okna Immediate
        ' 'Debug.Print
        MsgBox protectionMsg
    Next ws
End Sub

Sub VypisProceduryAFunkceFinal()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim vbMod As Object
    Dim line As String
    Dim i As Long
    Dim totalLines As Long
    Dim procName As String
    
    ' Nastavíme promìnnou pro VBA projekt
    Set vbProj = ThisWorkbook.VBProject
    
    ' Projdeme všechny komponenty v projektu (moduly)
    For Each vbComp In vbProj.VBComponents
        ' Kontrola, zda se jedná o standardní modul nebo tøídu
        If vbComp.Type = 1 Or vbComp.Type = 2 Then
            'Debug.Print "Modul: " & vbComp.Name
            ' Získáme kód modulu
            Set vbMod = vbComp.CodeModule
            totalLines = vbMod.CountOfLines
            
            ' Projdeme všechny øádky modulu a hledáme procedury a funkce
            For i = 1 To totalLines
                line = vbMod.lines(i, 1)
                
                ' Zjistíme, zda øádek obsahuje klíèové slovo "Sub" nebo "Function" a vyhneme se prázdným øádkùm nebo øádkùm s podmínkami
                If InStr(1, line, "Sub ", vbTextCompare) > 0 Or InStr(1, line, "Function ", vbTextCompare) > 0 Then
                    'Debug.Print "  " & line
                End If
            Next i
        End If
    Next vbComp
End Sub


