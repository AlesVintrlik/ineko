Attribute VB_Name = "Navigace"
' Definice modulové promìnné pro barvy ikonek
Public ICO_DISABLE_COLOR As Long
Public ICO_ENABLE_COLOR As Long

' Inicializace barev (mùžete zavolat tuto proceduru pøi spuštìní)
Sub InitializeColors()

    ICO_DISABLE_COLOR = RGB(250, 250, 250) ' Svìtle šedá
    ICO_ENABLE_COLOR = RGB(134, 134, 134) ' Støednì šedá
    
End Sub

' Funkce pro kontrolu existence listu
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If Not ws Is Nothing Then
        SheetExists = True
    Else
        SheetExists = False
    End If
    On Error GoTo 0
End Function

Sub Plus()
    'ActiveSheet.PivotTables("Rozpoèet").PivotCache.Refresh
    ActiveSheet.PivotTables("Rozpoèet").PivotFields( _
        "[Rozpoèet].[Skupina].[Skupina]").DrilledDown = True
End Sub

Sub Minus()
    'ActiveSheet.PivotTables("Rozpoèet").PivotCache.Refresh
    ActiveSheet.PivotTables("Rozpoèet").PivotFields( _
        "[Rozpoèet].[Skupina].[Skupina]").DrilledDown = False
End Sub

Sub TotoMakroNicNedela()
    'Toto makro nic nedìlá
    MsgBox "Toto tlaèítko není na listu " & ActiveSheet.Name & " aktivní"
End Sub

Sub KontingencniTabulka()

    Dim wsRozpocet As Worksheet
    Set wsRozpocet = ThisWorkbook.Sheets("Kontingenèní tabulka")
    
    wsRozpocet.Visible = True
    wsRozpocet.Activate
    ActiveSheet.PivotTables("Rozpoèet").PivotCache.Refresh

End Sub

Sub AssignMacroToShapes(targetSheet As Worksheet)
    Dim shp As Shape
    Dim remainingName As String
    Dim groupShape As Shape
    Dim i As Integer
    
    'Inicializace barev
    Call InitializeColors
    
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
                'Debug.Print "remainingName z AssignMacroToShapes je " & remainingName
                
                ' Pøiøazení makra na základì názvu objektu a aktuálního listu
                Select Case LCase(remainingName) ' Používáme LCase pro ignorování velikosti písmen
                    Case "aplikace"
                        If targetSheet.Name = "Aplikace" Or targetSheet.Name = "Kumulace" Or targetSheet.Name = "Kontingenèní tabulka" Then
                            shp.OnAction = "NavigateSheet"
                            shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR ' Zmìna barvy pro aktivní tlaèítko
                            'Debug.Print "Pøiøadil jsem NavigateSheet pro " & shp.Name
                        End If
                        
                    Case "load"
                        Select Case targetSheet.Name
                            Case "Aplikace"
                                shp.OnAction = "LoadDataFromQueries"
                                shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
                                'Debug.Print "Pøiøadil jsem LoadDataFromQueries pro " & shp.Name
                            Case "Kumulace"
                                shp.OnAction = "KumulujVysledovkuPodleCheckboxu"
                                shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
                                'Debug.Print "Pøiøadil jsem KumulujVysledovkuPodleCheckboxu pro " & shp.Name
                            Case "Kontingenèní tabulka"
                                shp.OnAction = "TotoMakroNicNedela"
                                shp.Fill.ForeColor.RGB = ICO_DISABLE_COLOR
                                'Debug.Print "Pøiøadil jsem TotoMakroNicNedela pro " & shp.Name
                        End Select
                        
                    Case "load_detail"
                        Select Case targetSheet.Name
                            Case "Aplikace"
                                shp.OnAction = "LoadAccountDetails"
                                shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
                                'Debug.Print "Pøiøadil jsem LoadAccountDetails pro " & shp.Name
                            Case "Kumulace", "Kontingenèní tabulka"
                                shp.OnAction = "TotoMakroNicNedela"
                                shp.Fill.ForeColor.RGB = ICO_DISABLE_COLOR
                                'Debug.Print "Pøiøadil jsem TotoMakroNicNedela pro " & shp.Name
                        End Select
                        
                    Case "kontingenèní tabulka", "kumulace"
                        If targetSheet.Name = "Aplikace" Or targetSheet.Name = "Kumulace" Or targetSheet.Name = "Kontingenèní tabulka" Then
                            shp.OnAction = "NavigateSheet"
                            shp.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
                            'Debug.Print "Pøiøadil jsem NavigateSheet pro " & shp.Name
                        End If
                        
                End Select
            End If
        Next i
    Else
        'Debug.Print "'navigace' není skupina."
    End If
    
    'Debug.Print "Aktivní" & ICO_ENABLE_COLOR
    'Debug.Print "Neaktivní" & ICO_DISABLE_COLOR
    
End Sub

Sub NavigateSheet()
    Dim shpName As String
    Dim targetSheetName As String
    Dim sheetToShow As Worksheet
    Dim ws As Worksheet

    ' Zjistíme název objektu, který volal makro
    shpName = LCase(Application.Caller)

    ' Kontrola, zda název objektu zaèíná na "ico_"
    If Left(shpName, 4) = "ico_" Then
        ' Extrahujeme název listu odstranìním prefixu "ico_"
        targetSheetName = Mid(shpName, 5)

        'Debug.Print "Cílový list extrahován z názvu objektu: " & targetSheetName

        ' Nastavíme sheetToShow na odpovídající list
        On Error Resume Next
        Set sheetToShow = ThisWorkbook.Sheets(targetSheetName)
        On Error GoTo 0

        ' Pokud list existuje
        If Not sheetToShow Is Nothing Then
            'Debug.Print "List '" & targetSheetName & "' existuje."

            ' Nastavení viditelnosti cílového listu
            If sheetToShow.Visible <> xlSheetVisible Then
                sheetToShow.Visible = xlSheetVisible
            End If

            ' Aktivace cílového listu
            'Debug.Print "Aktivuji cílový list: " & sheetToShow.Name
            sheetToShow.Activate

            ' Skrývání ostatních listù
            Application.ScreenUpdating = False
            For Each ws In ThisWorkbook.Worksheets
                If ws.Name <> sheetToShow.Name And ws.Visible = xlSheetVisible Then
                    ws.Visible = xlSheetHidden
                    'Debug.Print "List '" & ws.Name & "' byl skryt."
                End If
            Next ws
            Application.ScreenUpdating = True
        Else
            'Debug.Print "List '" & targetSheetName & "' neexistuje."
        End If
    Else
        'Debug.Print "Neznámý objekt: " & shpName
    End If
End Sub

Sub CopyNavigationGroup(targetSheet As Worksheet)
    Dim masterSheet As Worksheet
    Dim navigationGroup As Shape
    Dim newShape As Shape
    Dim leftPos As Single, topPos As Single
    Dim originalSheet As Worksheet ' Pro uložení pùvodního aktivního listu
    Dim allowedSheets As Variant ' Pole povolených listù
    Dim i As Integer
    Dim isAllowed As Boolean ' Pøíznak, zda je list povolen
    
    Call UnlockAllSheets
    DoEvents
    
    ' Seznam listù, na které chcete kopírovat navigaci
    allowedSheets = Array("Aplikace", "Kumulace", "Kontingenèní tabulka") ' Zmìòte podle potøeby
    
    ' Zkontrolujte, zda je cílový list v povoleném seznamu
    isAllowed = False
    For i = LBound(allowedSheets) To UBound(allowedSheets)
        If targetSheet.Name = allowedSheets(i) Then
            isAllowed = True
            Exit For
        End If
    Next i
    
    ' Pokud není list povolen, ukonèíme makro
    If Not isAllowed Then
        'Debug.Print "Navigace nebude kopírována na list: " & targetSheet.Name
        Exit Sub
    End If
    
    ' Nastavení názvu master listu
    Set masterSheet = ThisWorkbook.Sheets("Konfigurace")
    
    ' Získání skupiny "navigace" z master listu
    On Error Resume Next
    Set navigationGroup = masterSheet.Shapes("navigace")
    On Error GoTo 0
    
    If navigationGroup Is Nothing Then
        MsgBox "Skupina 'navigace' nebyla nalezena na listu 'Konfigurace'.", vbExclamation
        Exit Sub
    End If
    
    ' Kontrola typu skupiny
    If navigationGroup.Type <> msoGroup Then
        MsgBox "Skupina 'navigace' není typu msoGroup.", vbExclamation
        Exit Sub
    End If
    
    ' Odstranìní existující skupiny "navigace" na cílovém listu, pokud existuje
    On Error Resume Next
    targetSheet.Shapes("navigace").Delete
    On Error GoTo 0
    
    ' Uložení pùvodního aktivního listu
    Set originalSheet = ActiveSheet
    
    ' Aktivace cílového listu pøed vložením
    targetSheet.Activate
    
    ' Kopírování skupiny "navigace"
    navigationGroup.Copy
    
    ' Pøidání DoEvents pro zajištìní pøipravenosti Clipboardu
    DoEvents
    
    ' Nastavení pozice na cílovém listu (napø. horní levý roh)
    leftPos = 10 ' Nastavte podle potøeby
    topPos = 10  ' Nastavte podle potøeby
    
    ' Vložení skupiny na cílový list pomocí Worksheet.Paste
    On Error Resume Next
    ActiveSheet.Paste
    If Err.Number <> 0 Then
        'Debug.Print "Chyba pøi vkládání skupiny 'navigace' na list: " & targetSheet.Name & ". Chyba: " & Err.Description
        MsgBox "Chyba pøi vkládání skupiny 'navigace' na list: " & targetSheet.Name & ".", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Získání vložené skupiny (poslední vložený objekt)
    Set newShape = targetSheet.Shapes(targetSheet.Shapes.Count)
    newShape.Name = "navigace"
    newShape.Left = leftPos
    newShape.Top = topPos
    
    ' Pøiøazení makra ke všem relevantním objektùm ve skupinì
    Call AssignMacroToShapes(targetSheet)
    
    'Debug.Print "Skupina 'navigace' byla úspìšnì vložena na list: " & targetSheet.Name & " a byla pøiøazena makra."
    
    ' Aktivace pùvodního listu
    originalSheet.Activate
End Sub

Sub PrepnoutSkrytiSloupcu(sloupce As String, shapeVisible As String, shapeHidden As String)
    Call InitializeColors
    
    Dim shpVisible As Shape
    Dim shpHidden As Shape
    
    ' Nastavení grafických prvkù pro zobrazení/skrytí
    Set shpVisible = ActiveSheet.Shapes(shapeVisible)
    Set shpHidden = ActiveSheet.Shapes(shapeHidden)
    
    ' Pøepnutí viditelnosti sloupcù
    If Columns(sloupce).EntireColumn.Hidden = True Then
        ' Zobrazení sloupcù a nastavení barev
        Columns(sloupce).EntireColumn.Hidden = False
        shpVisible.Fill.ForeColor.RGB = ICO_DISABLE_COLOR
        shpHidden.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
    Else
        ' Skrytí sloupcù a nastavení barev
        Columns(sloupce).EntireColumn.Hidden = True
        shpVisible.Fill.ForeColor.RGB = ICO_ENABLE_COLOR
        shpHidden.Fill.ForeColor.RGB = ICO_DISABLE_COLOR
    End If
End Sub

Sub PrepnoutSkrytiSloupcuPlan()
    PrepnoutSkrytiSloupcu "F:Q", "RozbalitPlan", "SbalitPlan"
End Sub

Sub PrepnoutSkrytiSloupcuSkutecnost()
    PrepnoutSkrytiSloupcu "S:AD", "RozbalitSkutecnost", "SbalitSkutecnost"
End Sub

Sub PrepnoutSkrytiSloupcuRozdil()
    PrepnoutSkrytiSloupcu "AF:AQ", "RozbalitRozdil", "SbalitRozdil"
End Sub




