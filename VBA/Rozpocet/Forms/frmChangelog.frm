VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangelog 
   Caption         =   "Historie zmìn"
   ClientHeight    =   8280.001
   ClientLeft      =   96
   ClientTop       =   456
   ClientWidth     =   9384.001
   OleObjectBlob   =   "frmChangelog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmChangelog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With txtChangelog
        ' První èást
        .Value = ""
        ' Pøipojení poslední èásti
        .Value = .Value & vbCrLf & _
                 "Verze 1.0.4 – 2025-09-27" & vbCrLf & _
                 "-----------------" & vbCrLf & _
                 "- Opraveny chyby ve vzorcích (skuteènost variabilních nákladù)." & vbCrLf & vbCrLf
        ' Pøipojení pøedchozí èásti
        .Value = .Value & vbCrLf & _
                 "Verze 1.0.3 – 2025-06-10" & vbCrLf & _
                 "-----------------" & vbCrLf & _
                 "- Opraven nápoèet kumulativních dat." & vbCrLf & _
                 "- Pøizpùsobení velikosti textu šíøce souètových sloupcù." & vbCrLf & vbCrLf
        
        ' Pøipojení pøedchozí èásti
        .Value = .Value & _
                 "Verze 1.0.2 – 2025-04-08" & vbCrLf & _
                 "-----------------" & vbCrLf & _
                 "- Zmìnìný instalátor (možnost výbìru instalaèního adresáøe)." & vbCrLf & _
                 "- Informace o novinkách verze již bìhem instalace." & vbCrLf & _
                 "- Oprava drobných chyb." & vbCrLf & vbCrLf
                 
        ' Pøipojení pøedchozí èásti ...
        .Value = .Value & _
                 "Verze 1.0.1 – 2025-03-19" & vbCrLf & _
                 "-----------------" & vbCrLf & _
                 "- Pøidán changelog pro výpis zmìn v aplikaci." & vbCrLf & _
                 "- Pøidána možnost vložení vlastní poznámky ke každému øádku rozpoètu." & vbCrLf & vbCrLf
        
        .ScrollBars = fmScrollBarsVertical
        .Enabled = True
        .Locked = True
    End With
End Sub
