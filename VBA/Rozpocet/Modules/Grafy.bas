Attribute VB_Name = "Grafy"
Option Explicit
Public BarvyGrafu(1 To 3) As Long

Sub InicializujBarvyGrafu()
    ' Nastavení barev pro tøi datové øady
    BarvyGrafu(1) = ThisWorkbook.Sheets("Konfigurace").Range("C7").Value ' Hlavní barva RGB(35, 176, 160)
    BarvyGrafu(2) = ThisWorkbook.Sheets("Konfigurace").Range("C8").Value ' Doplòková barva RGB(209, 209, 209)
End Sub

Sub VyberBarvyGrafu()
    Dim Barva(1 To 2) As Long
    Dim i As Integer
    Dim result As Boolean
    
    MsgBox "Vyber postupnì hlavní a doplòkovou barvu pro graf"

    ' Pro každou datovou øadu umožni uživateli vybrat barvu
    For i = 1 To 2
        
        ' Vyvolání dialogu pro výbìr barvy
        result = Application.Dialogs(xlDialogEditColor).Show(i)
        
        ' Pokud uživatel vybral barvu, uložíme její RGB hodnotu
        If result Then
            Barva(i) = ActiveWorkbook.Colors(i)
        Else
            MsgBox "Nebyla vybrána žádná barva pro sérii " & i
        End If
    Next i

    ' Uložení barev do skrytého listu
     Call UlozBarvy(Barva)
     Call NastavBarvyGrafu
    
End Sub

Sub UlozBarvy(Barva() As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Konfigurace")
    
    ' Uložení barev do skrytého listu
    ws.Range("C7").Value = Barva(1)
    ws.Range("C8").Value = Barva(2)
End Sub

Sub NastavBarvyGrafu()
    Dim ws As Worksheet
    Dim graf As ChartObject
    Dim graf_k As ChartObject
    
    ' Nastaví list "Grafy"
    Set ws = ThisWorkbook.Sheets("Grafy")
    
    ' Nastaví grafy na listu "Grafy"
    Set graf = ws.ChartObjects("GrafKategorie")
    Set graf_k = ws.ChartObjects("GrafKategorieKumulativni")
    
    ' Ujisti se, že barvy byly inicializovány
    Call InicializujBarvyGrafu
    
    DoEvents
    
    ThisWorkbook.Sheets("Kumulace").Shapes("Graphic 8").Fill.ForeColor.RGB = BarvyGrafu(1)
    
    ' Nastavení barev pro graf GrafKategorie
    graf.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = BarvyGrafu(2)
    graf.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = BarvyGrafu(1)
    graf.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = BarvyGrafu(2)
    graf.Chart.SeriesCollection(3).Format.line.ForeColor.RGB = BarvyGrafu(1)
    
    ' Nastavení barev pro graf GrafKategorieKumulativni
    graf_k.Chart.SeriesCollection(1).Format.Fill.ForeColor.RGB = BarvyGrafu(2)
    graf_k.Chart.SeriesCollection(2).Format.Fill.ForeColor.RGB = BarvyGrafu(1)
    graf_k.Chart.SeriesCollection(3).Format.Fill.ForeColor.RGB = BarvyGrafu(2)
    graf_k.Chart.SeriesCollection(3).Format.line.ForeColor.RGB = BarvyGrafu(1)
        
End Sub

Sub SaveChartAsImage()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim chartObjKumulativni As ChartObject
    Dim savePath As String
    Dim savePathKumulativni As String
    
    ' Odkaz na list Kumulace
    Set ws = ThisWorkbook.Sheets("Grafy")
    
    ' Odkaz na graf s názvem GrafKategorie
    Set chartObj = ws.ChartObjects("GrafKategorie")
    Set chartObjKumulativni = ws.ChartObjects("GrafKategorieKumulativni")
    
    ' Nastavíme barvy a popisky v grafu
    Call NastavitPopiskyDatoveRady
    Call NastavitPopiskyDatoveRadyKumulativni
    
    ' Zajistíme, že Excel dokonèí všechny zmìny v grafu
    DoEvents
    
    ' Definuj cestu a název souboru, kam se obrázek uloží
    ' savePath = "C:\Users\vintr\Downloads\GrafKategorie.jpg"
    savePath = Environ("TEMP") & "\GrafKategorie.jpg"
    savePathKumulativni = Environ("TEMP") & "\GrafKategorieKumulativni.jpg"
    
    ' Kontrola, zda cesta existuje
    If Dir(Environ("TEMP"), vbDirectory) = "" Then
        ' Pokud složka neexistuje, vyvolá se dialog pro výbìr nové složky
        MsgBox "Složka neexistuje. Vyberte platnou složku.", vbExclamation
        
        ' Vyvolání dialogu pro výbìr složky
        With Application.FileDialog(msoFileDialogFolderPicker)
            .Title = "Vyberte složku pro uložení grafu"
            If .Show = -1 Then ' Pokud uživatel vybral složku
                MsgBox "Tady by došlo k uložení cesty " & .SelectedItems(1)
                ' Uložení nové cesty
                'Call UlozNovouCestu(nováCesta)
            Else
                MsgBox "Nebyla vybrána žádná složka.", vbExclamation
                Exit Sub
            End If
        End With
    Else
        ' MsgBox "Složka existuje: ", vbInformation
    End If
    
    ' Zajistíme, že Excel dokonèí všechny zmìny v grafu
    DoEvents
    
    ' Ulož graf jako obrázek ve formátu JPG
    chartObj.Chart.Export Filename:=savePath, FilterName:="JPG"
    chartObjKumulativni.Chart.Export Filename:=savePathKumulativni, FilterName:="JPG"
    
    ' Naèti uložený obrázek do imgGraf na formuláøi frmGraf
    frmGraf.imgGraf.Picture = LoadPicture(savePath)
    frmGraf.imgGrafKumulativni.Picture = LoadPicture(savePathKumulativni)
    
    ' Zobraz formuláø
    frmGraf.Show
    
End Sub

Sub OptimalizovatGraf()

Dim cht As ChartObject
Dim visibleWidth As Double
Set cht = Worksheets("Grafy").ChartObjects("GrafKategorie")
        
' Získání šíøky viditelné oblasti listu
    visibleWidth = (ActiveWindow.VisibleRange.Width - 20) - 10
    ' Nastavení pozice grafu na (0,0)
    With cht

        .Width = visibleWidth / 2
        .Height = (visibleWidth / 4) - 100 ' Mùžete upravit hodnotu podle potøeby
        .Left = 10
        .Top = 10
        
    End With
    
Set cht = Worksheets("Grafy").ChartObjects("GrafKategorieKumulativni")
        
' Získání šíøky viditelné oblasti listu
    visibleWidth = (ActiveWindow.VisibleRange.Width - 20) - 10
    ' Nastavení pozice grafu na (0,0)
    With cht

        .Width = visibleWidth / 2
        .Height = (visibleWidth / 4) - 100 ' Mùžete upravit hodnotu podle potøeby
        .Left = 20 + visibleWidth / 2
        .Top = 10
        
    End With
    
End Sub

Sub NastavitPopiskyDatoveRady()
    Dim ws As Worksheet
    Dim graf As ChartObject
    Dim serie As Series
    Dim i As Long
    Dim hodnota As Double
    Dim popis As String
    Dim zaporny As Boolean
    
    ' Nastavení pracovního listu, kde graf existuje
    Set ws = ThisWorkbook.Sheets("Grafy") ' Zmìò na název svého listu, pokud je jiný
    
    ' Najdeme graf podle jeho názvu
    On Error Resume Next
    Set graf = ws.ChartObjects("GrafKategorie")
    On Error GoTo 0
    
    ' Zkontrolujeme, jestli graf existuje
    If graf Is Nothing Then
        Exit Sub
    End If
    
    ' Pøedpokládáme, že chceme upravit popisky pro 3. datovou øadu
    On Error Resume Next
    Set serie = graf.Chart.SeriesCollection(3)
    On Error GoTo 0
    
    ' Zkontrolujeme, jestli tøetí øada dat existuje
    If serie Is Nothing Then
        Exit Sub
    End If
    
    ' Ujistíme se, že popisky jsou aktivovány pro všechny body
    serie.HasDataLabels = True
    
    ' Projdeme všechny hodnoty v øadì a nastavíme formát popiskù
    For i = 1 To serie.Points.Count
        hodnota = serie.Values(i)
        
        ' Urèíme, jestli je hodnota negativní
        zaporny = hodnota < 0
        
        ' Urèení formátu podle absolutní hodnoty
        If Abs(hodnota) >= 1000000 Then
            If zaporny Then
                popis = "-" & Format(Abs(hodnota) / 1000000, "0.0") & " M"
            Else
                popis = Format(hodnota / 1000000, "0.0") & " M"
            End If
        ElseIf Abs(hodnota) >= 1000 Then
            If zaporny Then
                popis = "-" & Format(Abs(hodnota) / 1000, "0.0") & " tis."
            Else
                popis = Format(hodnota / 1000, "0.0") & " tis."
            End If
        Else
            If zaporny Then
                popis = "-" & Format(Abs(hodnota), "0")
            Else
                popis = Format(hodnota, "0")
            End If
        End If
        
        ' Nastavení popisku pro konkrétní bod
        serie.Points(i).DataLabel.text = popis
    Next i
End Sub

Sub NastavitPopiskyDatoveRadyKumulativni()
    Dim ws As Worksheet
    Dim graf As ChartObject
    Dim serie As Series
    Dim i As Long
    Dim hodnota As Double
    Dim popis As String
    Dim zaporny As Boolean
    
    ' Nastavení pracovního listu, kde graf existuje
    Set ws = ThisWorkbook.Sheets("Grafy") ' Zmìò na název svého listu, pokud je jiný
    
    ' Najdeme graf podle jeho názvu
    On Error Resume Next
    Set graf = ws.ChartObjects("GrafKategorieKumulativni")
    On Error GoTo 0
    
    ' Zkontrolujeme, jestli graf existuje
    If graf Is Nothing Then
        Exit Sub
    End If
    
    ' Pøedpokládáme, že chceme upravit popisky pro 3. datovou øadu
    On Error Resume Next
    Set serie = graf.Chart.SeriesCollection(3)
    On Error GoTo 0
    
    ' Zkontrolujeme, jestli tøetí øada dat existuje
    If serie Is Nothing Then
        Exit Sub
    End If
    
    ' Ujistíme se, že popisky jsou aktivovány pro všechny body
    serie.HasDataLabels = True
    
    ' Projdeme všechny hodnoty v øadì a nastavíme formát popiskù
    For i = 1 To serie.Points.Count
        hodnota = serie.Values(i)
        
        ' Urèíme, jestli je hodnota negativní
        zaporny = hodnota < 0
        
        ' Urèení formátu podle absolutní hodnoty
        If Abs(hodnota) >= 1000000 Then
            If zaporny Then
                popis = "-" & Format(Abs(hodnota) / 1000000, "0.0") & " M"
            Else
                popis = Format(hodnota / 1000000, "0.0") & " M"
            End If
        ElseIf Abs(hodnota) >= 1000 Then
            If zaporny Then
                popis = "-" & Format(Abs(hodnota) / 1000, "0.0") & " tis."
            Else
                popis = Format(hodnota / 1000, "0.0") & " tis."
            End If
        Else
            If zaporny Then
                popis = "-" & Format(Abs(hodnota), "0")
            Else
                popis = Format(hodnota, "0")
            End If
        End If
        
        ' Nastavení popisku pro konkrétní bod
        serie.Points(i).DataLabel.text = popis
    Next i
End Sub


