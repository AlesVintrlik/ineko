Attribute VB_Name = "VypocetPlanu"
Option Explicit

Sub OtevritFormularZakazky()
    Dim zakazka As String
    
    ' Kontrola výbìru jedné buòky
    If Selection.Rows.Count > 1 Then
        MsgBox "Oznaète pouze jednu buòku.", vbCritical
        Exit Sub
    End If

    ' Získání èísla zakázky ze sloupce B
    zakazka = ActiveSheet.Cells(ActiveCell.Row, "B").Value
    If zakazka = "" Then
        MsgBox "Na oznaèeném øádku není žádné èíslo zakázky.", vbCritical
        Exit Sub
    End If

    ' Pøedání hodnoty do formuláøe
    ZobrazFormularZpozdene "frmZakazka"
    frmZakazka.txtZakazka.Value = zakazka
    frmZakazka.Show
End Sub

Sub CallPlanTerminyVyrobyDoZakazek(zakazkaID As Long, Optional datumUkonceni As Variant)
    Dim conn As Object
    Dim cmd As Object
    Dim debugCommand As String ' Pro sestavení debug výpisu
    Dim datumSQL As String

    On Error GoTo ErrorHandler

    ' Vytvoøení pøipojení
    Set conn = CreateConnection()

    ' Vytvoøení pøíkazu
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 4 ' adCmdStoredProc
    cmd.CommandText = "dbo.EP_PlanTerminyVyrobyDoZakazek"

    ' Nastavení parametrù
    cmd.Parameters.Append cmd.CreateParameter("@ID", 3, 1, , zakazkaID) ' adInteger, adParamInput

    ' Zpracování datumu
    If IsMissing(datumUkonceni) Or IsNull(datumUkonceni) Or Trim(datumUkonceni) = "" Then
        ' Pokud není datum zadáno, pøedá se NULL
        cmd.Parameters.Append cmd.CreateParameter("@DatumUkonceni", 7, 1, , Null) ' adDate, adParamInput
        debugCommand = "EXEC dbo.EP_PlanTerminyVyrobyDoZakazek @ID = " & zakazkaID & ", @DatumUkonceni = NULL"
    Else
        ' Validace datumu
        If Not IsDate(datumUkonceni) Then
            MsgBox "Chyba: Zadané datum není platné.", vbCritical
            Exit Sub
        End If
        
        ' Pøevedení datumu na formát YYYY-MM-DD
        datumSQL = Format(CDate(datumUkonceni), "YYYY-MM-DD")
        
        ' Pøedání datumu jako parametr
        cmd.Parameters.Append cmd.CreateParameter("@DatumUkonceni", 7, 1, , datumSQL) ' adDate, adParamInput
        debugCommand = "EXEC dbo.EP_PlanTerminyVyrobyDoZakazek @ID = " & zakazkaID & ", @DatumUkonceni = '" & datumSQL & "'"
    End If

    ' Debug výpis pøíkazu
    Debug.Print debugCommand
    'MsgBox "Debug: " & debugCommand, vbInformation

    ' Zakomentováno pro ladìní
    cmd.Execute

    ' Uzavøení pøipojení
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing

    ' Pøípadné další akce
    'MsgBox "Procedura byla ladìna. SQL pøíkaz byl vypsán do Debug Console.", vbInformation
    
    WriteLog "Starting dbo.EP_PlanTerminyVyrobyDoZakazek procedures."

    Exit Sub

ErrorHandler:
    ' Chybové hlášení
    MsgBox "Došlo k chybì: " & Err.Description, vbCritical
    WriteLog "Error dbo.EP_PlanTerminyVyrobyDoZakazek: " & Err.Description
    
    ' Uvolnìní pøipojení a objektù
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
End Sub

Function GetZakazkaID(cisloZakazky As String) As Long
    Dim conn As Object
    Dim rs As Object
    Dim sql As String

    On Error GoTo ErrorHandler

    ' Vytvoøení pøipojení k databázi
    Set conn = CreateConnection()

    ' SQL dotaz pro získání ID zakázky
    sql = "SELECT ID FROM TabZakazka WHERE CisloZakazky = '" & cisloZakazky & "'"
    
    Debug.Print sql

    ' Vytvoøení a spuštìní recordsetu
    Set rs = conn.Execute(sql)
    If Not rs.EOF Then
        GetZakazkaID = rs.Fields(0).Value
    Else
        GetZakazkaID = 0 ' Zakázka nenalezena
    End If

    ' Uzavøení pøipojení
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Chyba pøi získávání ID zakázky: " & Err.Description, vbCritical
    WriteLog "Error while retrieving order ID: " & Err.Description
    GetZakazkaID = 0
    If Not rs Is Nothing Then rs.Close
    If Not conn Is Nothing Then conn.Close
End Function





