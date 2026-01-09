VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Helios"
   ClientHeight    =   3883
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   4128
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnLogin_Click()
    Dim userName As String
    Dim password As String
    Dim ServerName As String
    Dim databaseName As String
    Dim key As String
    Dim credentials As New Collection
    Dim wsKonfig As Worksheet
    
    ' Nastavení listu se zdrojovými daty
    Set wsKonfig = ThisWorkbook.Sheets("Konfigurace")
    ServerName = wsKonfig.Range("serverName").Value
    databaseName = wsKonfig.Range("databaseName").Value

    ' Klíè pro šifrování a dešifrování
    key = ENCRYPTION_KEY
    
    ' Naètení údajù z formuláøe
    userName = txtUsername.Value
    password = txtPassword.Value
    ServerName = txtServerName.Value
    databaseName = txtDBName.Value
    
    ' Uložení hodnot z formuláøe do pojmenovaných oblastí
    wsKonfig.Range("serverName").Value = ServerName
    wsKonfig.Range("databaseName").Value = databaseName
    wsKonfig.Range("login").Value = userName
    
    ' Ovìøení pøihlašovacích údajù
    If VerifyCredentials(userName, password, ServerName, databaseName) Then
        ' Pokud jsou údaje správné, šifruj a ulož je do globální promìnné
        credentials.Add XOREncryptDecrypt(userName, key), "login"
        credentials.Add XOREncryptDecrypt(password, key), "passw"
        Set loginCredentials = credentials
        
        ' Nastavení globální promìnné pro sledování pøihlášení
        isLoggedIn = True
        
        ' Skrytí formuláøe
        Unload Me
        
        ' Maximalizace a zviditelnìní aplikace Excel
        Application.WindowState = xlMaximized
        Application.Visible = True
        ActiveWindow.Caption = "Plánování"
        
        ' Pøesmìrování na konkrétní list po úspìšném pøihlášení
        Application.ScreenUpdating = False ' Potlaèení pøekreslování
    
        If Sheets("Gantt").Visible <> xlSheetVisible Then
            Sheets("Gantt").Visible = xlSheetVisible
            
            Sheets("Gantt").Activate
        End If
        
        Application.ScreenUpdating = True ' Obnovení pøekreslování
    Else
        ' Pokud jsou údaje nesprávné, zobraz chybovou zprávu
        MsgBox "Nesprávné pøihlašovací údaje. Zkuste to prosím znovu.", vbExclamation
        txtUsername.SetFocus
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim wsKonfig As Worksheet
    Set wsKonfig = ThisWorkbook.Sheets("Konfigurace")
    
    ' Naètení hodnot z pojmenovaných oblastí do textových polí
    txtServerName.Value = wsKonfig.Range("serverName").Value
    txtDBName.Value = wsKonfig.Range("databaseName").Value
    txtUsername.Value = wsKonfig.Range("login").Value
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not isLoggedIn Then
        ' Pokud není uživatel pøihlášen, zavøi sešit
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

