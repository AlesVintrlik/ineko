Attribute VB_Name = "PracnostZakazek"
Option Explicit

Sub OtevritFormularHodiny()
    Dim zakazka As String
    
    ' Kontrola výbìru jedné buòky
    If Selection.Rows.Count > 1 Then
        MsgBox "Oznaète pouze jednu buòku.", vbCritical
        Exit Sub
    End If

    ' Získání èísla zakázky ze sloupce B
    zakazka = ActiveSheet.Cells(ActiveCell.Row, "B").Value
    If zakazka = "" Then
        MsgBox "Na oznaèeném øádku není žádné èíslo zakázky.", vbCritical
        Exit Sub
    End If
    
    Call GetZakazkaID(zakazka)
    
    ' Inicializace formuláøe
    frmHodiny.txtCisloZakazky.Value = zakazka
    
    ' Zobrazení formuláøe
    ZobrazFormularZpozdene "frmHodiny"
    frmHodiny.Show
    
End Sub

Sub UlozitHodiny(zakazkaID As Long, HodCelkem As Long, HodSkPrac1 As Long, HodSkPrac2 As Long, HodSkPrac3 As Long, HodSkPrac4 As Long, HodSkPrac5 As Long, HodKoop As Long)
    Dim conn As Object
    Dim sqlInsert As String
    Dim sqlUpdate As String

    On Error GoTo ErrorHandler

    ' Vytvoøení pøipojení
    Set conn = CreateConnection()

    ' SQL pøíkaz pro kontrolu existence a pøípadné založení ID
    sqlInsert = "IF NOT EXISTS (SELECT 1 FROM TabZakazka_EXT WHERE ID = " & zakazkaID & ") " & _
                "INSERT INTO TabZakazka_EXT (ID) VALUES (" & zakazkaID & ")"

    ' SQL pøíkaz pro UPDATE všech hodnot
    sqlUpdate = "UPDATE TabZakazka_EXT SET " & _
                "_HodCelkem = " & HodCelkem & ", " & _
                "_HodSkPrac1 = " & HodSkPrac1 & ", " & _
                "_HodSkPrac2 = " & HodSkPrac2 & ", " & _
                "_HodSkPrac3 = " & HodSkPrac3 & ", " & _
                "_HodSkPrac4 = " & HodSkPrac4 & ", " & _
                "_HodSkPrac5 = " & HodSkPrac5 & ", " & _
                "_HodKoop = " & HodKoop & " " & _
                "WHERE ID = " & zakazkaID

    ' Spuštìní pøíkazù
    conn.Execute sqlInsert ' Nejprve kontrola a založení ID
    conn.Execute sqlUpdate ' Poté UPDATE všech hodnot

    ' Uzavøení pøipojení
    conn.Close

    ' Úspìšná zpráva
    MsgBox "Hodiny byly úspìšnì uloženy.", vbInformation
    WriteLog "Starting UlozitHodiny procedures."

    Exit Sub

ErrorHandler:
    MsgBox "Došlo k chybì: " & Err.Description, vbCritical
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set conn = Nothing
End Sub



