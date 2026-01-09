Attribute VB_Name = "OldVersion"
Option Explicit

Sub RunToggleSingleCheckBox()
Attribute RunToggleSingleCheckBox.VB_ProcData.VB_Invoke_Func = "g\n14"
    ' Zkontrolujte, zda je aktivní list "Kumulace"
    If ActiveSheet.Name <> "Kumulace" Then
        MsgBox "Toto makro lze spustit pouze na listu 'Kumulace'.", vbExclamation
        Exit Sub
    End If
    
    ' Zkontrolujte, zda je vybrán pouze jeden øádek
    If Selection.Rows.Count <> 1 Then
        MsgBox "Nejednoznaèné zadání. Prosím, oznaète pouze jeden øádek s daty.", vbCritical
        Exit Sub
    End If
    
    ' Nastavte hodnotu True do sloupce C aktuálního øádku
    Cells(Selection.Row, "C").Value = True
End Sub




