VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZakazka 
   Caption         =   "Plánování termínu pro zakázku"
   ClientHeight    =   2860
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   5376
   OleObjectBlob   =   "frmZakazka.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZakazka"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Spustí se po stisknutí tlaèítka Potvrdit
Private Sub cmdPotvrdit_Click()
    Dim datumUkonceni As Variant
    Dim zakazkaID As String
    Dim sqlID As Long

    ' Získání èísla zakázky
    zakazkaID = Me.txtZakazka.Value
    
    ' Kontrola vyplnìní èísla zakázky
    If zakazkaID = "" Then
        MsgBox "Chyba: Èíslo zakázky není vyplnìno.", vbCritical
        Exit Sub
    End If
    
    ' Kontrola data ukonèení
    datumUkonceni = Me.txtDatumUkonceni.Value
    If Trim(datumUkonceni) <> "" Then
        If Not IsDate(datumUkonceni) Then
            MsgBox "Chyba: Zadané datum není platné. Platný formát je DD.MM.YYYY", vbCritical
            Exit Sub
        End If
    Else
        datumUkonceni = Null ' Pokud pole je prázdné
    End If
    
    ' Získání ID zakázky z databáze
    sqlID = GetZakazkaID(zakazkaID)
    If sqlID = 0 Then
        MsgBox "Chyba: Zakázka s èíslem " & zakazkaID & " nebyla nalezena v databázi.", vbCritical
        Exit Sub
    End If
    
    ' Volání procedury s parametry
    Call CallPlanTerminyVyrobyDoZakazek(sqlID, datumUkonceni)
    
    ' Zavøení formuláøe
    Unload Me
    
    ' Aktualizuje data po naètení
    Call LoadOrUpdateData
    
End Sub

' Zavøe formuláø
Private Sub cmdZavrit_Click()
    Unload Me
End Sub

