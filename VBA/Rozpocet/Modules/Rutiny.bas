Attribute VB_Name = "Rutiny"
Option Explicit

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

Sub LockSpecificSheets(ByVal Sh As Object)
    
    If Not Sh.ProtectContents Then
        Sh.Protect UserInterfaceOnly:=True, DrawingObjects:=True, Contents:=True, AllowUsingPivotTables:=True, AllowFormattingColumns:=True
        Sh.EnableSelection = xlNoRestrictions
    End If
       
End Sub

Sub UnlockAllSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        ' Odemknout pouze, pokud je list zamknutý
        If ws.ProtectContents Then
            ws.Unprotect
            'Debug.Print "List '" & ws.Name & "' byl odemknut."
        End If
    Next ws
    
End Sub

Sub ToggleScreenUpdating(turnOff As Boolean)
    If turnOff Then
        ' Vypnutí vykreslování a aktualizace
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
    Else
        ' Zapnutí vykreslování a aktualizace
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

Sub DocasnaZmenaBunky()
    'Tato akce se spouští kvùli vydráždìní aktualizace a naètení dat po importu, protože to po kompilaci nefungovalo správnì
    
    Dim puvodníHodnota As Variant
    Dim ws As Worksheet

    ' Nastavení odkazu na list "Rozpoèet"
    Set ws = ThisWorkbook.Worksheets("Rozpoèet")
   
    puvodníHodnota = ws.Range("B2").Value   ' Uložení pùvodní hodnoty buòky B2 do promìnné
    ws.Range("B2").Value = 13               ' Zmìna hodnoty buòky B2 na 13 (neexistující mìsíc)
    ws.Range("B2").Value = puvodníHodnota   ' Obnovení pùvodní hodnoty do buòky B2

End Sub

Sub FollowHyperlink(ByVal Target As Range)
    If Target.Address = "$B$5" Then
        ' Naètení historie zmìn do textboxu
        frmChangelog.txtChangelog.Value = "Historie verzí aplikace:" & vbCrLf & _
                                          "v1.0 - První verze aplikace." & vbCrLf & _
                                          "v1.1 - Oprava chyb a vylepšení výkonu." & vbCrLf & _
                                          "v2.0 - Pøidána podpora vícenásobných pøipojení." & vbCrLf & _
                                          "v2.1 - Nový design a možnosti exportu."
        ' Zobrazení UserFormu
        frmChangelog.Show
    End If
End Sub


