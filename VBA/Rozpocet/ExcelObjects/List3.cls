VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    Call KumulujVysledovkuPodleCheckboxu
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Kumulace")
    
    ' Ovìøíme, zda byla zmìna provedena ve sloupci C na listu Kumulace
    If Not Intersect(Target, ws.Range("C:C")) Is Nothing Then
        ' Ovìøíme, zda se hodnota zmìnila na True nebo False (tedy že jde o checkbox)
        If Target.Value = True Or Target.Value = False Then
            ' Zabráníme zacyklení tím, že doèasnì vypneme události
            Application.EnableEvents = False
                       
            ' Voláme makro a pøedáme èíslo øádku, kde byl zmìnìn checkbox
            Call ToggleSingleCheckBox(Target.Row)
            
            ' Znovu povolíme události
            Application.EnableEvents = True
        End If
    End If
End Sub




