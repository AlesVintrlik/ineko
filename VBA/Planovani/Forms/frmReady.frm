VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReady 
   Caption         =   "Plánování"
   ClientHeight    =   2057
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   5784
   OleObjectBlob   =   "frmReady.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReady"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blinking As Boolean

Private Sub cmdOK_Click()

    blinking = False
    
    Me.Hide 'nejprve skryj tento form
    
    ' Teprve nyní zviditelni Excel (tady bude chvíli èekat uživatel)
    Application.Visible = True
    Application.WindowState = xlMaximized
    
    ' Aktivuj cílový list
    ThisWorkbook.Sheets("Plan").Activate
    ActiveWindow.Caption = "Plánování"
    
    Unload Me 'nakonec zcela zavøi formuláø
    
End Sub

Private Sub UserForm_Activate()
    blinking = True
    StartBlinking
End Sub

Private Sub StartBlinking()
    On Error Resume Next
    Application.OnTime Now + TimeValue("00:00:01"), "ToggleLabel"
End Sub

