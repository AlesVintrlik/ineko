VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    ' Uložení poznámek do externího souboru data.ini
    Call SaveToINI
    
    ' Obnovení viditelnosti prvkù
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", True)" ' Zobrazit Ribbon
    Application.DisplayFormulaBar = True ' Zobrazit lištu vzorcù
    Application.DisplayStatusBar = True ' Zobrazit stavový øádek
    
    ' Potvrzení od uživatele, zda opravdu chce zavøít aplikaci
    If MsgBox("Opravdu chcete zavøít aplikaci?", vbQuestion + vbYesNo) = vbNo Then
        Cancel = True ' Zabránit zavøení
        Exit Sub
    End If
    
End Sub

Private Sub Workbook_Open()
    Dim userName As String
    Dim currentUser As String
    Dim cleanedUserName As String

    ' Základní skrytí prvkù aplikace (nikoli celé aplikace)
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"", False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = False

    ' Získání aktuálního uživatele
    currentUser = Environ("Username")

    ' Pokus o zmìnu pøístupového režimu na ètení/zápis
    On Error GoTo OpenError
    Me.ChangeFileAccess Mode:=xlReadWrite
    On Error GoTo 0

    ' Zobrazení pøihlašovacího formuláøe
    frmLogin.Show vbModal ' Ujistìte se, že uživatel nemùže zavøít formuláø mimo tok aplikace
    Exit Sub

OpenError:
    ' Získání informace o uživateli, který má soubor otevøen
    userName = GetFileInUseUserName(ThisWorkbook.FullName)
    cleanedUserName = Split(userName, "\")(UBound(Split(userName, "\")))

    If cleanedUserName = currentUser Then
        frmLogin.Show vbModal ' i v pøípadì selhání režimu zápisu pokraèovat
    Else
        MsgBox "Soubor je otevøen jiným uživatelem: " & userName & ".", vbExclamation
        Application.Quit
    End If
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Nastavení titulku okna
    ActiveWindow.Caption = "Rozpoèet"

    ' Skrytí Ribbonu (použít místo Excel4Macro, pokud je dostupné)
    Application.CommandBars("Ribbon").Visible = False

    ' Nastavení aktivní buòky podle listu
    Select Case Sh.Name
        Case "Aplikace", "Kumulace"
            Sh.Range("E4").Value = Sh.Range("E4").Value ' Efektivní nastavení místo Select
        Case "Kontingenèní tabulka"
            Sh.Range("D4").Value = Sh.Range("D4").Value
        Case Else
            Sh.Range("A1").Value = Sh.Range("A1").Value
    End Select

    ' Správa posuvníkù pouze pøi zmìnì
    With ActiveWindow
        .DisplayHorizontalScrollBar = (Sh.Name = "Aplikace" Or Sh.Name = "Kumulace")
        .DisplayVerticalScrollBar = Not .DisplayHorizontalScrollBar
    End With
End Sub


Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    ' Odemknout listy pouze pøi pøepnutí z listu "Aplikace" nebo "Kumulace"
    If Sh.Name = "Aplikace" Or Sh.Name = "Kumulace" Then
        UnlockAllSheets
        DoEvents
    End If
End Sub

Function GetFileInUseUserName(filePath As String) As String
    Dim objShell As Object
    Dim objFolder As Object
    Dim objFolderItem As Object
    Dim userName As String

    On Error GoTo ErrorHandler
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(Left(filePath, InStrRev(filePath, "\") - 1))
    Set objFolderItem = objFolder.ParseName(Mid(filePath, InStrRev(filePath, "\") + 1))

    userName = objFolderItem.ExtendedProperty("Owner")
    GetFileInUseUserName = userName
    Exit Function

ErrorHandler:
    GetFileInUseUserName = "neznámý uživatel"
End Function

