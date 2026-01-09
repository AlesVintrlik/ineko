VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Sheets("Detail").Visible = 0
End Sub
    
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next ' Prevence chyb (napø. u propojených listù)
    
    ' Nastavení pevné šíøky sloupce AR
    Me.Columns("AR").ColumnWidth = 0.2
    
    ' Kontrola, zda zmìna nastala ve sloupci AS
    If Not Intersect(Target, Me.Columns("AS")) Is Nothing Then
        Me.Columns("AS").AutoFit ' Automatická úprava šíøky sloupce AS
        Call SaveToINI
    End If

    On Error GoTo 0 ' Obnovení standardního zpracování chyb
End Sub

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    ' Kontrola, zda hypertextový odkaz smìøuje na buòku B5
    If Target.Range.Address = "$B$5" Then
        ' Zobrazení UserFormu
        frmChangelog.Show
    End If
End Sub


