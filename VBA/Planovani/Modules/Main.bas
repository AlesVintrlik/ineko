Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As LongPtr, lpLogFont As LOGFONT, ByVal lpEnumFontFamExProc As LongPtr, ByVal lParam As LongPtr, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Function EnumFontFamiliesEx Lib "gdi32" Alias "EnumFontFamiliesExA" (ByVal hdc As Long, lpLogFont As LOGFONT, ByVal lpEnumFontFamExProc As Long, ByVal lParam As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

' Struktura LOGFONT
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Dim m_targetFont As String
Dim m_fontFound As Boolean

Sub AktualizovatSeznamZakazek()
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    frmProgress.lblProgress.Caption = "Aktualizuji zakázky."
    DoEvents

    Dim wsZakazky As Worksheet, wsPlan As Worksheet
    Set wsZakazky = ThisWorkbook.Sheets("Zakazky")
    Set wsPlan = ThisWorkbook.Sheets("Plan")

    wsPlan.Unprotect password:="MrkevNeniOvoce123"

    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    Dim zakazka As Range, elektroKonec As Range, firma As Range
    Dim key As String
    Dim elektroKonecCol As Integer, nazevFirmyCol As Integer

    ' Najdi sloupce
    On Error Resume Next
    elektroKonecCol = wsZakazky.Rows(1).Find(What:="ElektroKonec", LookIn:=xlValues, LookAt:=xlWhole).Column
    nazevFirmyCol = wsZakazky.Rows(1).Find(What:="Firma", LookIn:=xlValues, LookAt:=xlWhole).Column
    On Error GoTo ErrorHandler

    If elektroKonecCol = 0 Then MsgBox "Sloupec 'ElektroKonec' nebyl nalezen.": GoTo Cleanup
    If nazevFirmyCol = 0 Then MsgBox "Sloupec 'Firma' nebyl nalezen.": GoTo Cleanup

    ' Sestavení dictionary
    For Each zakazka In wsZakazky.Range("A2:A" & wsZakazky.Cells(wsZakazky.Rows.Count, "A").End(xlUp).Row)
        If zakazka.Value <> "" Then
            Set elektroKonec = zakazka.Offset(0, elektroKonecCol - zakazka.Column)
            Set firma = zakazka.Offset(0, nazevFirmyCol - zakazka.Column)

            key = zakazka.Value
            If Not dic.exists(key) Then
                Dim firmaVal As Variant
                firmaVal = ""
                If Not IsError(firma.Value) Then firmaVal = firma.Value
                dic.Add key, Array(zakazka.Value, firmaVal, elektroKonec.Value)
            End If
        End If
    Next zakazka

    frmProgress.lblProgress.Caption = "Aktualizuji zakázky.."
    DoEvents

    Dim zakazky() As Variant
    zakazky = dic.Items()

    Call BubbleSortZakazky(zakazky)

    wsPlan.Range("F15:G" & dic.Count + 14).NumberFormat = "0"

    Dim i As Long
    For i = 0 To UBound(zakazky)
        wsPlan.Cells(i + 15, "B").Value = zakazky(i)(0)
        wsPlan.Cells(i + 15, "C").Value = zakazky(i)(1)
    Next i

    If dic.Count > 0 Then
        wsPlan.Range("B15:C15").Copy
        wsPlan.Range("B16:C" & dic.Count + 14).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False

    Dim targetRange As Range
    
    ' Zaznamenej zaèátek procedury
    WriteLog "Start processing the target range."
    
    ' Nastavení cílového rozsahu – logujeme, jaký rozsah se má nastavit
    WriteLog "Setting range D15:AB" & (dic.Count + 14)
    Set targetRange = wsPlan.Range("D15:AB" & dic.Count + 14)
    
    ' Zkopíruj rozsah pro vzorce a formátování
    wsPlan.Range("D15:AB15").Copy
    
    ' Aplikuj vzorce
    WriteLog "Inserting formulas into the target range."
    targetRange.PasteSpecial Paste:=xlPasteFormulas
    
    ' Aplikuj formátování
    WriteLog "Inserting formatting into the target range."
    targetRange.PasteSpecial Paste:=xlPasteFormats
    
    ' Zavoláme funkci, která vypíše obsah targetRange do logu
    ' DumpRangeFormulas targetRange
    
    ' Vyèisti režim kopírování
    Application.CutCopyMode = False
    WriteLog "Ending processing of the target range."
        Application.CutCopyMode = False
    End If

    frmProgress.lblProgress.Caption = "Aktualizuji zakázky..."
    DoEvents

    Dim lastRow As Long
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp).Row
    If lastRow < 15 Then lastRow = 15

    With wsPlan
        If lastRow >= 15 Then
            .Range("M15:M" & lastRow).ClearContents
            .Range("P15:P" & lastRow).ClearContents
            .Range("S15:S" & lastRow).ClearContents
            .Range("V15:V" & lastRow).ClearContents
            .Range("Y15:Y" & lastRow).ClearContents
            .Range("AB15:AB" & lastRow).ClearContents
        End If
    End With

    If dic.Count < lastRow - 14 Then
        wsPlan.Range("B" & dic.Count + 15 & ":B" & lastRow).EntireRow.Delete
    End If
    
    Application.Calculate

    wsPlan.Range("I3").Select

Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Calculate

    If wsPlan.AutoFilterMode Then wsPlan.AutoFilterMode = False
    wsPlan.Range("B14:E" & wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp).Row).AutoFilter

    wsPlan.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True

    WriteLog "Inserting formatting into the target range."
    Call NastavFontPlusFormatDatumu

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Došlo k chybì pøi aktualizaci zakázek: " & Err.Description & " (Èíslo chyby: " & Err.Number & ")"
    Resume Cleanup
End Sub

' Funkce BubbleSort pro seøazení podle DatumExpedice
Sub BubbleSortZakazky(arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    Dim val1 As Variant, val2 As Variant
    
    WriteLog "Starting BubbleSortOrders"

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            val1 = arr(i)(2)
            val2 = arr(j)(2)
            
            ' Posuò prázdné hodnoty dopøedu
            If IsEmpty(val1) Or val1 = "" Then
                ' val1 je prázdné › nic nedìlej
            ElseIf IsEmpty(val2) Or val2 = "" Then
                ' val2 je prázdné › prohodit
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            ElseIf val1 > val2 Then
                ' Standardní porovnání dat
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Sub NaplnitKontrolu()

    ' Potlaèení pøekreslování a aktualizace
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsPlan As Worksheet
    Dim wsKontrola As Worksheet
    Dim wsOperace As Worksheet
    Set wsPlan = ThisWorkbook.Sheets("Plan")
    Set wsKontrola = ThisWorkbook.Sheets("Kontrola")
    Set wsOperace = ThisWorkbook.Sheets("Operace") ' Zdroj dat z Heliosu
    
    wsPlan.Unprotect password:="MrkevNeniOvoce123"

    ' Smaž hodnoty v listu Kontrola a zachovej pouze hlavièku
    wsKontrola.Rows("2:" & wsKontrola.Rows.Count).Delete

    Dim lastRow As Long
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp).Row

    Dim i As Long
    Dim kontrolaRow As Long
    kontrolaRow = 2 ' zaèátek na druhém øádku, protože první øádek je hlavièka

    Dim columnCodes As Variant
    columnCodes = Array("M", "P", "S", "V", "Y", "AB")
    Dim usekValues As Variant
    usekValues = Array(1, 4, 2, 3, 5, 8)

    Dim j As Long
    For i = 13 To lastRow ' prochází øádky od 13 do posledního øádku
        For j = 0 To 5 ' prochází sloupce M,P,S,V,Y,AB
            If wsPlan.Range(columnCodes(j) & i).Value <> "" Then ' pokud buòka obsahuje datum
                wsKontrola.Cells(kontrolaRow, "B").Value = wsPlan.Cells(i, "B").Value ' Zakazka
                wsKontrola.Cells(kontrolaRow, "C").Value = wsPlan.Cells(i, Columns(columnCodes(j)).Column - 1).Value ' PuvodniTermin
                wsKontrola.Cells(kontrolaRow, "D").Value = wsPlan.Cells(i, columnCodes(j)).Value ' TerminVyrobyDo
                wsKontrola.Cells(kontrolaRow, "E").Value = WorksheetFunction.SumIfs(wsOperace.Range("PotrebnyCasCelkemHod"), wsOperace.Range("Zakazka"), wsPlan.Cells(i, "B").Value, wsOperace.Range("Usek"), usekValues(j)) ' PotrebnyCasCelkemHod
                wsKontrola.Cells(kontrolaRow, "A").Value = usekValues(j) ' Usek
                wsKontrola.Cells(kontrolaRow, "F").Value = Application.WorksheetFunction.IsoWeekNum(wsKontrola.Cells(kontrolaRow, "C").Value)
                wsKontrola.Cells(kontrolaRow, "G").Value = Application.WorksheetFunction.IsoWeekNum(wsKontrola.Cells(kontrolaRow, "D").Value)
                kontrolaRow = kontrolaRow + 1
            End If
        Next j
    Next i
    
    wsKontrola.Visible = xlSheetVisible
    wsKontrola.Activate

    ' Zapnutí pøekreslování a aktualizace
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    wsPlan.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True

End Sub

Sub PlusTyden()
Attribute PlusTyden.VB_ProcData.VB_Invoke_Func = "n\n14"

    ' Získání sloupce a øádku aktivní buòky
    Dim activeCellColumn As Integer
    activeCellColumn = ActiveCell.Column
    Dim activeCellRow As Integer
    activeCellRow = ActiveCell.Row

    ' Definice sloupcù, které máme na starosti (M, P, S, V, Y)
    Dim columnCodes As Variant
    columnCodes = Array(13, 16, 19, 22, 25)

    ' Kontrola, zda je aktivní buòka v požadovaných sloupcích a øádek je >= 13
    If IsInArray(activeCellColumn, columnCodes) And activeCellRow >= 13 Then
        Dim leftCell As Range
        Set leftCell = Cells(activeCellRow, activeCellColumn - 1)

        ' Kontrola, zda je hodnota v buòce vlevo datum
        If IsDate(leftCell.Value) Then
            ' Pokud aktivní buòka není prázdná a rozdíl mezi daty je násobek 7
            If Not IsEmpty(ActiveCell.Value) And IsDate(ActiveCell.Value) And ((ActiveCell.Value - leftCell.Value) Mod 7 = 0) Then
                ' Pøidání sedmi dní k aktuálnímu datu aktivní buòky
                ActiveCell.Value = DateAdd("d", 7, CDate(ActiveCell.Value))
            Else
                ' Pøidání sedmi dní k datu v buòce vlevo a vložení do aktivní buòky
                ActiveCell.Value = DateAdd("d", 7, CDate(leftCell.Value))
            End If
        End If
    End If

End Sub

Sub MinusTyden()
Attribute MinusTyden.VB_ProcData.VB_Invoke_Func = "d\n14"

    ' Získání sloupce a øádku aktivní buòky
    Dim activeCellColumn As Integer
    activeCellColumn = ActiveCell.Column
    Dim activeCellRow As Integer
    activeCellRow = ActiveCell.Row

    ' Definice sloupcù, které máme na starosti (M, P, S, V, Y)
    Dim columnCodes As Variant
    columnCodes = Array(13, 16, 19, 22, 25)

    ' Kontrola, zda je aktivní buòka v požadovaných sloupcích a øádek je >= 13
    If IsInArray(activeCellColumn, columnCodes) And activeCellRow >= 13 Then
        Dim leftCell As Range
        Set leftCell = Cells(activeCellRow, activeCellColumn - 1)

        ' Kontrola, zda je hodnota v buòce vlevo datum
        If IsDate(leftCell.Value) Then
            ' Pokud aktivní buòka není prázdná a rozdíl mezi daty je násobek 7
            If Not IsEmpty(ActiveCell.Value) And IsDate(ActiveCell.Value) And ((ActiveCell.Value - leftCell.Value) Mod 7 = 0) Then
                ' Odebrání sedmi dní od aktuálního data aktivní buòky
                ActiveCell.Value = DateAdd("d", -7, CDate(ActiveCell.Value))
            Else
                ' Odebrání sedmi dní od data v buòce vlevo a vložení do aktivní buòky
                ActiveCell.Value = DateAdd("d", -7, CDate(leftCell.Value))
            End If
        End If
    End If

End Sub

Sub ObsazenostUseku()

    Dim wsUseky As Worksheet
    Set wsUseky = ThisWorkbook.Sheets("Obsazenost úsekù")
    
    wsUseky.Visible = True
    wsUseky.Activate

End Sub

Sub ObsazenostPracovist()

    Dim wsPracoviste As Worksheet
    Set wsPracoviste = ThisWorkbook.Sheets("Obsazenost pracoviš")
    
    wsPracoviste.Visible = True
    wsPracoviste.Activate

End Sub

Sub Konfigurace()

    Dim wsKonfigurace As Worksheet
    Set wsKonfigurace = ThisWorkbook.Sheets("Konfigurace")
    
    wsKonfigurace.Visible = True
    wsKonfigurace.Activate

End Sub

' Pomocná funkce pro kontrolu, zda je hodnota v poli
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, valToBeFound)) > -1)
End Function

Sub Ribbon()
Attribute Ribbon.VB_ProcData.VB_Invoke_Func = "l\n14"

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"

End Sub

' Callback funkce, volaná pøi enumeraci fontù
Function EnumFontsProc(ByVal lpelfe As LongPtr, ByVal lpntme As LongPtr, ByVal FontType As Long, ByVal lParam As LongPtr) As Long
    Dim lf As LOGFONT
    CopyMemory lf, ByVal lpelfe, Len(lf)
    Dim currentFont As String
    currentFont = Left$(lf.lfFaceName, InStr(lf.lfFaceName, vbNullChar) - 1)
    
    If StrComp(currentFont, m_targetFont, vbTextCompare) = 0 Then
        m_fontFound = True
        EnumFontsProc = 0 ' zastav enumeraci
    Else
        EnumFontsProc = 1 ' pokraèuj
    End If
End Function

' Funkce ovìøuje, zda je zadaný font nainstalován
Function IsFontInstalled(ByVal FontName As String) As Boolean
    Dim hdc As LongPtr
    Dim lf As LOGFONT
    Dim retVal As Long
    
    m_targetFont = FontName
    m_fontFound = False
    
    hdc = GetDC(0)
    
    lf.lfFaceName = FontName & vbNullChar
    lf.lfCharSet = 0 ' výchozí znaková sada
    
    retVal = EnumFontFamiliesEx(hdc, lf, AddressOf EnumFontsProc, 0, 0)
    
    ReleaseDC 0, hdc
    
    IsFontInstalled = m_fontFound
End Function

' Funkce kontroluje, zda font podporuje èeskou diakritiku.
' Zde ovìøujeme podle seznamu známých fontù – uprav si jej dle potøeby.
Function SupportsCzech(ByVal FontName As String) As Boolean
    Select Case LCase(FontName)
        Case "segoe ui light", "segoe ui semilight", "calibri light", "arial narrow", "arial narrow", "arial"
            SupportsCzech = True
        Case Else
            SupportsCzech = False
    End Select
End Function

Sub NastavFontPlusFormatDatumu()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim col As Range
    Dim maxWidth As Double
    Dim favFonts As Variant
    Dim chosenFont As String
    Dim i As Long
    Dim totalSheets As Long, currentSheet As Long
    
    Set ws = ThisWorkbook.Sheets("Plan")
    
    ' Odemkneme list "Plan"
    ws.Unprotect password:="MrkevNeniOvoce123"
    
    ' Seznam oblíbených písem seøazených podle priority
    favFonts = Array("Segoe UI Semilight", "Segoe UI Light", "Calibri Light", "Arial Narrow", "Arial")
    
    chosenFont = ""
    For i = LBound(favFonts) To UBound(favFonts)
        ' Aktualizace formuláøe – vypíše se právì kontrolovaný font
        frmProgress.lblProgress.Caption = "Kontroluji font: " & favFonts(i)
        DoEvents
        Delay 0.5 ' Prodleva 0,5 sekundy
            
        If IsFontInstalled(favFonts(i)) And SupportsCzech(favFonts(i)) Then
            chosenFont = favFonts(i)
            frmProgress.lblProgress.Caption = "Zkoumám font: " & chosenFont
            frmProgress.UpdateProgressBar 0.2
            DoEvents
            Debug.Print chosenFont
            Delay 0.5 ' Prodleva 0,5 sekundy
            Exit For
        End If
    Next i
    
    ' Pokud žádný z oblíbených není dostupný, použijeme default
    If chosenFont = "" Then
        chosenFont = "Arial"
        frmProgress.lblProgress.Caption = "Aplikuji font " & chosenFont
        DoEvents
    End If

    frmProgress.lblProgress.Caption = "Aplikuji font " & chosenFont
    frmProgress.UpdateProgressBar 0.4
    DoEvents
    
    ' Aplikuj font na celý list
    ws.Cells.Font.Name = chosenFont
    ws.Cells.Font.Size = 10
    
    frmProgress.lblProgress.Caption = "Nastavuji formát datumu."
    frmProgress.UpdateProgressBar 0.5
    DoEvents
    
    ' Nastav formát datumu pro všechny buòky v rozsahu H:AB
    Set rng = Intersect(ws.Range("H:AB"), ws.UsedRange)
    If Not rng Is Nothing Then
        For Each cell In rng
            If IsDate(cell.Value) Then
                cell.NumberFormat = "dd.mm.yy"
            End If
        Next cell
    End If
    
    frmProgress.lblProgress.Caption = "Nastavuji šíøku sloupcù."
    frmProgress.UpdateProgressBar 0.6
    DoEvents
    
    ' 1. Automaticky pøizpùsob šíøku sloupcù H:AA
    ws.Range("H:AA").Columns.AutoFit
    
    ' 2. Zjisti nejširší sloupec v rozsahu H:AA
    maxWidth = 0
    For Each col In ws.Range("H:AA").Columns
        If col.ColumnWidth > maxWidth Then
            maxWidth = col.ColumnWidth
        End If
    Next col
    
    frmProgress.UpdateProgressBar 0.7
    
    ' Pøidáme okraj 1 jednotku
    maxWidth = maxWidth + 1
    
    ' 3. Nastav šíøku sloupcù H:AB podle nejširšího sloupce
    ws.Range("H:AB").ColumnWidth = maxWidth
    
    frmProgress.UpdateProgressBar 0.8
    
    ' 4. Nastav šíøku sloupce J na 1 jednotku
    ws.Columns("J").ColumnWidth = 1
    
    ' 5. Nastav formát èísla pro oblast K5:Y11
    ws.Range("K5:Y11").NumberFormat = "0"
    
    frmProgress.lblProgress.Caption = "Bìhem chvíle dojde k otevøení aplikace."
    frmProgress.UpdateProgressBar 1
    DoEvents
    
    ' Zavøi formuláø, jen když bìží
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = "frmProgress" Then
            Unload frmProgress
            ' Voláme formuláø se zprávou
            ZobrazFormularZpozdene "frmReady"
            frmReady.Show vbModal
            Exit For
        End If
    Next frm
    
    ' Zamkni list "Plan" zpìt
    ws.Protect password:="MrkevNeniOvoce123", AllowFiltering:=True
End Sub

Sub Delay(Seconds As Single)
    Dim endTime As Single
    endTime = Timer + Seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

Sub Ladeni()
Attribute Ladeni.VB_ProcData.VB_Invoke_Func = "l\n14"
    Static ribbonHidden As Boolean
    Dim formulaVisible As Boolean
    Dim statusVisible As Boolean
    
    ' Zjisti, stav prvkù
    formulaVisible = Application.DisplayFormulaBar
    statusVisible = Application.DisplayStatusBar
    
    ' Pøepni stav ribbonu podle jeho aktuálního stavu
    If ribbonHidden Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
        ribbonHidden = False
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)"
        ribbonHidden = True
    End If
    
    If formulaVisible = False Then
        Application.DisplayFormulaBar = True
    Else
        Application.DisplayFormulaBar = False
    End If
    
    If statusVisible = False Then
        Application.DisplayStatusBar = True
    Else
        Application.DisplayStatusBar = False
    End If
    
End Sub

Public Sub ToggleLabel()
    On Error Resume Next

    ' Ovìøíme, že formuláø existuje a že blikání má pokraèovat
    If frmReady.Visible Then
        With frmReady
            If .Label2.ForeColor = vbRed Then
                .Label2.ForeColor = vbBlack
            Else
                .Label2.ForeColor = vbRed
            End If
        End With
        ' Znovu naplánujeme bliknutí
        Application.OnTime Now + TimeValue("00:00:01"), "ToggleLabel"
    End If
End Sub

Public Sub ZobrazFormularZpozdene(nazevFormulare As String, Optional sekundyZpozdeni As Double = 1)
    Dim cas As Date
    cas = Now + TimeSerial(0, 0, sekundyZpozdeni)
    Application.OnTime cas, "'ZobrazFormular """ & nazevFormulare & """'"
End Sub

Public Sub ZobrazFormular(formName As String)
    Select Case formName
        Case "frmReady"
            If Not frmReady.Visible Then
                frmReady.Show vbModal
            End If
    End Select
End Sub


