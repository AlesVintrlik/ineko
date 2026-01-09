Attribute VB_Name = "Data"

Sub LoadOrUpdateData()
    Dim adoConn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim i As Integer
    
    ' Nastav list, do kterého chceš naèíst data
    Set ws = ThisWorkbook.Sheets("Zakazky")
    ws.Cells.ClearContents ' pøípadnì vymaž stará data
    
    ' Vytvoøení ADO pøipojení pomocí funkce CreateConnection
    Set adoConn = CreateConnection()
    If adoConn Is Nothing Then
        MsgBox "Nepodaøilo se otevøít pøipojení.", vbCritical
        Exit Sub
    End If
    
    ' Definice SQL dotazu
    sql = "SELECT * FROM hvw_TerminyZakazekProPlanovani ORDER BY DatumUkonceni"
    
    ' Otevøení recordsetu
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, adoConn, 0, 1
    
    ' Vytvoøení instance formuláøe frmProgress
    Load frmProgress
    frmProgress.Show vbModeless
    
    ' Zapsání názvù sloupcù (hlavièky) do první øádky
    If Not rs.EOF Then
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(1, i + 1).Value = rs.Fields(i).Name
        Next i
    End If
    
    ' Naètení dat do listu poèínaje øádkem 2
    ws.Range("A2").CopyFromRecordset rs
    
    ' Uzavøení recordsetu a pøipojení
    rs.Close
    Set rs = Nothing
    adoConn.Close
    Set adoConn = Nothing
    
    MsgBox "Data byla naètena do listu s hlavièkou.", vbInformation
End Sub

Sub AktualizovatSeznamZakazek()
    On Error GoTo ErrorHandler ' Pøesmìrování chyb na error handler

    ' Potlaèení pøekreslování a aktualizace
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Call ZrusitVsechnyFiltry
    Call VymazaniSouctuPodGrafem

    ' Nastavení listù
    Dim wsZakazky As Worksheet
    Dim wsPlan As Worksheet
    Set wsZakazky = ThisWorkbook.Sheets("Zakazky")
    Set wsPlan = ThisWorkbook.Sheets("Gantt")

    ' Získání posledního øádku s daty na listu "Gantt" pøed aktualizací
    Dim lastRowPlan As Long
    lastRowPlan = wsPlan.Cells(wsPlan.Rows.Count, "B").End(xlUp).row

    ' Získání unikátního seznamu zakázek (bez hlavièky)
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    Dim zakazka As Range
    For Each zakazka In wsZakazky.Range("A2:A" & wsZakazky.Cells(wsZakazky.Rows.Count, "A").End(xlUp).row)
        If Not dic.exists(zakazka.Value) Then
            dic(zakazka.Value) = 1
        End If
    Next zakazka
    
    ' Pøevod dictionary na pole
    Dim zakazky() As Variant
    zakazky = dic.Keys()
    
    ' Seøazení pole zakázek
    ' Call BubbleSort(zakazky)
    
    ' Vložení unikátních zakázek a kopírování vzorcù
    Dim i As Long
    For i = 0 To UBound(zakazky)
        wsPlan.Cells(i + 4, "B").Value = zakazky(i)
    Next i

    ' Kopírování vzorcù z øádku 4 na nové øádky
    If dic.Count > 0 Then
        wsPlan.Range("C4:O4").Copy
        wsPlan.Range("C5:O" & dic.Count + 3).PasteSpecial Paste:=xlPasteFormulas
        Application.CutCopyMode = False
    End If

    ' Kopírování formátu z øádku 4 na nové øádky
    If dic.Count > 0 Then
        wsPlan.Rows("4").Copy
        wsPlan.Rows("5:" & dic.Count + 3).PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
    End If

    ' Odstranìní pøebyteèných øádkù
    If dic.Count + 3 < lastRowPlan Then
        wsPlan.Range("B" & dic.Count + 4 & ":B" & lastRowPlan).EntireRow.Delete
    End If
    
    ' Zapnutí pøekreslování a aktualizace
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    x = 0
    Call SumarizaceBoduProVsechnySloupce

    MsgBox "Naèetl jsem aktuální zakázky."
    Calculate
    
    Exit Sub ' Pøeskoèení error handleru po úspìšném provedení kódu

ErrorHandler: ' Zpracování chyby
    ' V pøípadì chyby zapnout pøekreslování a aktualizaci a informovat o chybì
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Došlo k chybì pøi aktualizaci zakázek: " & Err.Description & " (Èíslo chyby: " & Err.Number & ")"
    
    Exit Sub
End Sub

Sub RefreshAllConnections()
    ' Aktualizace všech pøipojení v sešitu
    ThisWorkbook.RefreshAll
    MsgBox "Data byla úspìšnì aktualizována.", vbInformation
End Sub

' Pomocná funkce pro kontrolu, zda je hodnota v poli
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, valToBeFound)) > -1)
End Function

' BubbleSort funkce
Sub BubbleSort(arr As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

