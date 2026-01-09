VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHodiny 
   Caption         =   "Nastavení hodin pro zakázku"
   ClientHeight    =   6281
   ClientLeft      =   96
   ClientTop       =   444
   ClientWidth     =   6576
   OleObjectBlob   =   "frmHodiny.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHodiny"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim zakazkaID As Long
    Dim cisloZakazky As String

    ' Naètení èísla zakázky pøedaného do formuláøe
    cisloZakazky = Me.txtCisloZakazky.Value
    Debug.Print "Èíslo zakázky z formuláøe: " & cisloZakazky

    ' Získání zakazkaID zakázky pomocí funkce GetZakazkaID
    zakazkaID = GetZakazkaID(cisloZakazky)
    Debug.Print "Získané ID zakázky: " & IIf(zakazkaID = 0, "Nenalezeno", zakazkaID)

    ' Kontrola, zda bylo nalezeno ID
    If zakazkaID = 0 Then
        MsgBox "Chyba: ID zakázky pro èíslo '" & cisloZakazky & "' nebylo nalezeno!", vbCritical
        Unload Me
        Exit Sub
    End If

    ' Vytvoøení pøipojení
    Set conn = CreateConnection()

    ' SQL dotaz pro naètení dat
    sql = "SELECT _HodSkPrac1, _HodSkPrac4, _HodSkPrac2, _HodSkPrac3, _HodSkPrac5, _HodKoop, _HodCelkem " & _
          "FROM TabZakazka_EXT WHERE ID = " & zakazkaID
    Debug.Print "SQL dotaz pro naètení hodnot: " & sql

    ' Spuštìní SQL dotazu
    Set rs = conn.Execute(sql)

    ' Naètení dat do formuláøe, pokud existují
    If Not rs.EOF Then
        Debug.Print "Naètené hodnoty z databáze:"
        Debug.Print "_HodSkPrac1: " & (rs.Fields(0).Value)
        Debug.Print "_HodSkPrac4: " & (rs.Fields(1).Value)
        Debug.Print "_HodSkPrac2: " & (rs.Fields(2).Value)
        Debug.Print "_HodSkPrac3: " & (rs.Fields(3).Value)
        Debug.Print "_HodSkPrac5: " & (rs.Fields(4).Value)
        Debug.Print "_HodKoop: " & (rs.Fields(5).Value)
        Debug.Print "_HodCelkem: " & (rs.Fields(6).Value)

        ' Pøiøazení hodnot do formuláøe
        Me.Controls("HodSkPrac1").Value = (rs.Fields(0).Value)
        Me.Controls("HodSkPrac4").Value = (rs.Fields(1).Value)
        Me.Controls("HodSkPrac2").Value = (rs.Fields(2).Value)
        Me.Controls("HodSkPrac3").Value = (rs.Fields(3).Value)
        Me.Controls("HodSkPrac5").Value = (rs.Fields(4).Value)
        Me.Controls("HodKoop").Value = (rs.Fields(5).Value)
        Me.Controls("HodCelkem").Value = (rs.Fields(6).Value)
    Else
        Debug.Print "Žádné záznamy nebyly nalezeny pro ID: " & zakazkaID
        MsgBox "Pro tuto zakázku neexistují záznamy, vyplòte nová data.", vbInformation
    End If

    ' Uzavøení pøipojení
    rs.Close
    conn.Close
End Sub

Private Sub cmdPotvrdit_Click()
    Dim zakazkaID As Long
    Dim HodCelkem As Long
    Dim HodSkPrac1 As Long, HodSkPrac2 As Long, HodSkPrac3 As Long
    Dim HodSkPrac4 As Long, HodSkPrac5 As Long, HodKoop As Long
    Dim soucet As Long

    ' Naètení hodnot z formuláøe
    zakazkaID = GetZakazkaID(Me.Controls("txtCisloZakazky").Value)
    HodSkPrac1 = Val(Me.Controls("HodSkPrac1").Value) ' Pøíprava
    HodSkPrac4 = Val(Me.Controls("HodSkPrac4").Value) ' Svaøování
    HodSkPrac2 = Val(Me.Controls("HodSkPrac2").Value) ' Montáž
    HodSkPrac3 = Val(Me.Controls("HodSkPrac3").Value) ' Elektro
    HodSkPrac5 = Val(Me.Controls("HodSkPrac5").Value) ' Segmenty
    HodKoop = Val(Me.Controls("HodKoop").Value)       ' Kooperace
    HodCelkem = Val(Me.Controls("HodCelkem").Value)   ' Celkem

    ' Výpoèet souètu všech hodnot mimo celkem
    soucet = HodSkPrac1 + HodSkPrac4 + HodSkPrac2 + HodSkPrac3 + HodSkPrac5 + HodKoop

    ' Kontrola: Hodnota celkem musí být vìtší nebo rovna souètu ostatních
    If Not IsNumeric(HodCelkem) Or HodCelkem < soucet Then
        MsgBox "Hodnota 'Celkem' musí být vìtší nebo rovna souètu ostatních.", vbCritical
        Exit Sub
    End If

    ' Volání procedury pro uložení hodnot
    Call UlozitHodiny(zakazkaID, HodCelkem, HodSkPrac1, HodSkPrac2, HodSkPrac3, HodSkPrac4, HodSkPrac5, HodKoop)

    ' Zavøení formuláøe
    Unload Me
End Sub

Private Sub cmdZavrit_Click()
    Unload Me
End Sub
