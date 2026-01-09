VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "Helios"
   ClientHeight    =   1551
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   4356
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    lblProgress.Caption = "Naèítání dat, prosím èekejte..."
    lblProgressBar.Width = 0
End Sub

Public Sub UpdateProgressBar(progress As Double)
    lblProgressBar.Width = progress * (Me.FrameProgress.Width - 10)
    DoEvents
End Sub

