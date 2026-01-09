VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
    ActiveWindow.DisplayGridlines = True
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    
End Sub
    
Private Sub Workbook_Open()
    Dim userName As String
    Dim currentUser As String
    Dim cleanedUserName As String
    
    Application.Visible = False
    
    ActiveWindow.DisplayGridlines = False
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayHeadings = False
          
    ' Získání aktuálního uživatele
    currentUser = Environ("Username")
    
    ' Pokus o zmìnu pøístupového režimu na pro ètení a zápis
    On Error GoTo OpenError
    Me.ChangeFileAccess Mode:=xlReadWrite
    On Error GoTo 0
    
    ' Pokud úspìšnì zmìní pøístupový režim, zobrazí pøihlašovací formuláø
    frmLogin.Show
    Exit Sub
    
OpenError:
    ' Pokud se nepodaøí zmìnit pøístupový režim, pokus o získání jména uživatele, který má soubor otevøený
    userName = GetFileInUseUserName(ThisWorkbook.FullName)
    
    ' Odstranìní doménové èásti z userName
    cleanedUserName = Split(userName, "\")(UBound(Split(userName, "\")))

    ' Kontrola, zda je soubor otevøen aktuálním uživatelem
    If cleanedUserName = currentUser Then
        ' Pokud soubor otevøel aktuální uživatel, pokraèuje normálnì
        ' MsgBox "Soubor je již otevøen vámi. Pokraèuji...", vbInformation
        frmLogin.Show
    Else
        ' Pokud soubor otevøel jiný uživatel, zobrazí zprávu a ukonèí aplikaci
        If userName <> "neznámý uživatel" Then
            MsgBox "Tento soubor je již otevøen uživatelem: " & userName & ". Aplikace bude ukonèena.", vbExclamation
        Else
            MsgBox "Tento soubor je již otevøen jiným uživatelem. Aplikace bude ukonèena.", vbExclamation
        End If
        'Application.Quit
    End If
    
    ' Nastavení hodnoty buòky P1 na 28
    Sheets("Gantt").Range("P1").Value = 28
      
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
