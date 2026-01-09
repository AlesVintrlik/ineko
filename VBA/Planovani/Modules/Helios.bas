Attribute VB_Name = "Helios"
Option Explicit

Sub AktualizovatData()

    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim i As Long
    Dim dateDiff As Integer
    Dim ID As Long

    ' Vytvoøení objektu pro pøipojení a otevøení pøipojení
    Set conn = CreateConnection()

    ' Nastavení listu se zdrojovými daty
    Set ws = ThisWorkbook.Sheets("Kontrola")

    ' Projití všech øádkù se zdrojovými daty
    For i = 2 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
              
 '       If dateDiff = 0 Then
            If IsNumeric(ws.Cells(i, "H").Value) Then
                If ws.Cells(i, "H").Value = Int(ws.Cells(i, "H").Value) Then
                    dateDiff = ws.Cells(i, "H").Value
                End If
            End If
 '       End If
    
        ' Aktualizace úsekù
        sql = "UPDATE TabZakazka_EXT SET _U" & ws.Cells(i, "A").Value & "Konec = '" & Format(ws.Cells(i, "D").Value, "yyyymmdd") & "' "
        sql = sql & "FROM TabZakazka_EXT AS TZE "
        sql = sql & "JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID "
        sql = sql & "WHERE TZ.CisloZakazky = '" & ws.Cells(i, "B").Value & "'"
        conn.Execute sql
        
        sql = "UPDATE TabZakazka_EXT SET _U" & ws.Cells(i, "A").Value & "Start = DATEADD(day, " & dateDiff * -1 & ", CONVERT(datetime, '" & Format(ws.Cells(i, "D").Value, "yyyymmdd") & "', 102)) "
        sql = sql & "FROM TabZakazka_EXT AS TZE "
        sql = sql & "JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID "
        sql = sql & "WHERE TZ.CisloZakazky = '" & ws.Cells(i, "B").Value & "'"
        conn.Execute sql
        
        Application.Wait Now + TimeValue("00:00:01")
        dateDiff = GetDateDiffAndID(ws, ws.Cells(i, "A").Value, ws.Cells(i, "B").Value, ID)
        Application.Wait Now + TimeValue("00:00:01")
        
        Call CallStoredProcedure(ID)
        
    Next i

    ' Uzavøení pøipojení
    conn.Close

    ' Uvolnìní objektù
    Set conn = Nothing
    
    ActiveWorkbook.RefreshAll
    
    MsgBox "Nastavil jsem termíny zakázek v Heliosu."
    WriteLog "Deadlines for orders set in Helios."

End Sub

Function GetDateDiffAndID(ws As Worksheet, usek As String, zakazka As String, ByRef ID As Long) As Long
On Error GoTo ErrorHandler

    Dim conn As Object
    Set conn = CreateConnection()
    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    ' SQL dotaz pro získání rozdílu ve dnech mezi zaèátkem a koncem a ID
    Dim sql As String
    sql = "SELECT TZ.ID, DATEDIFF(day, _U" & usek & "Start, _U" & usek & "Konec) as DateDiff "
    sql = sql & "FROM TabZakazka_EXT AS TZE "
    sql = sql & "JOIN TabZakazka AS TZ ON TZE.ID = TZ.ID "
    sql = sql & "WHERE TZ.CisloZakazky = '" & zakazka & "'"
    
    ' Provedení dotazu a získání výsledku
    rs.Open sql, conn
    If Not rs.EOF Then
        GetDateDiffAndID = rs.Fields("DateDiff").Value
        ID = rs.Fields("ID").Value
    Else
        GetDateDiffAndID = 0
        ID = 0
    End If
    
    ' Uzavøení pøipojení
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
ErrorHandler:

End Function

Sub CallStoredProcedure(ID As Long)
    Dim conn As Object
    Set conn = CreateConnection()
    
    ' Vytvoøení pøíkazu
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = conn
    cmd.CommandType = 4 'adCmdStoredProc
    cmd.CommandText = "dbo.ep_DoplnTerminOperacePodleUseku"
    
    ' Pøidání parametru
    Dim param As Object
    Set param = cmd.CreateParameter("@ID", 3, 1, , ID) 'adInteger
    cmd.Parameters.Append param
    
    ' Spuštìní pøíkazu
    cmd.Execute
    
    ' Uzavøení pøipojení
    conn.Close
    
    Dim wsKontrola As Worksheet
    Dim wsPlan As Worksheet

    ' Nastavení listu se zdrojovými daty
    Set wsKontrola = ThisWorkbook.Sheets("Kontrola")
    Set wsPlan = ThisWorkbook.Sheets("Plan")
    
    wsPlan.Activate
    wsKontrola.Visible = False
    
End Sub





