VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Worksheet_Activate()
    Dim listNames As Variant, wsName As Variant
    Dim ws As Worksheet

    listNames = Array("Obsazenost ˙sek˘", "Obsazenost pracoviöù", "Konfigurace", "Kontrola")

    Application.ScreenUpdating = False

    ' Vûdy nejprve zviditelnÌme list "Plan"
    On Error Resume Next
    ThisWorkbook.Sheets("Plan").Visible = xlSheetVisible
    On Error GoTo 0

    ' Skr˝v·nÌ dalöÌch list˘
    For Each wsName In listNames
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(wsName)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            If ws.Name <> Me.Name And ws.Name <> "Plan" And ws.Visible = xlSheetVisible Then
                ws.Visible = xlSheetHidden
            End If
        End If
        
        Set ws = Nothing
    Next

    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    ' Kontrola, zda hypertextov˝ odkaz smÏ¯uje na buÚku B5
    If Target.Range.Address = "$B$13" Then
        ' ZobrazenÌ UserFormu
        ZobrazFormularZpozdene "frmChangelog"
        frmChangelog.Show
    End If
End Sub
