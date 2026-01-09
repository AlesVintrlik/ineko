VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmChangelog 
   Caption         =   "Historie zmìn"
   ClientHeight    =   8280.001
   ClientLeft      =   96
   ClientTop       =   444
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

Dim changelog As String
Dim logPart1_3 As String
Dim logPart1_4 As String
Dim logPart1_5 As String


logPart1_5 = vbCrLf & "Verze 1.5 – 2025-04-14" & vbCrLf & _
           "-----------------------------" & vbCrLf & _
           "** NOVÌ: **" & vbCrLf & _
           "- Pøidáno logování do souboru PlanovaniLog.txt" & vbCrLf & vbCrLf & _
           "** ZMÌNY: **" & vbCrLf & " " & vbCrLf & vbCrLf & _
           "** OPRAVENÉ CHYBY: **" & vbCrLf & vbCrLf

logPart1_4 = vbCrLf & "Verze 1.4 – 2025-04-08" & vbCrLf & _
           "-----------------------------" & vbCrLf & _
           "** NOVÌ: **" & vbCrLf & _
           "- Zakázky bez vyplnìného termínu 'ElektroKonec' se nyní øadí na zaèátek seznamu zakázek." & vbCrLf & _
           "- Formátování sloupcù B a C (Zakázka, Firma) se automaticky kopíruje z øádku 15 na novì vložené øádky." & vbCrLf & vbCrLf & _
           "** ZMÌNY: **" & vbCrLf & _
           "- Øazení zakázek probíhá podle hodnoty 'ElektroKonec', pøièemž prázdné hodnoty mají prioritu." & vbCrLf & _
           "- Záznam do listu 'Plan' nyní používá kombinaci hodnoty 'Zakázka' a 'ElektroKonec' jako unikátní klíè pro eliminaci duplicit." & vbCrLf & _
           "- Ošetøeno, pokud je hodnota 'Firma' prázdná – nedojde k chybì." & vbCrLf & vbCrLf & _
           "** OPRAVENÉ CHYBY: **" & vbCrLf & _
           "- Zakázky s prázdným datem se døíve zobrazovaly na konci seznamu." & vbCrLf & _
           "- V nìkterých pøípadech chybìlo formátování novì vložených hodnot ve sloupcích B a C." & vbCrLf & vbCrLf

logPart1_3 = vbCrLf & "Verze 1.3 – 2025-03-21" & vbCrLf & _
           "-----------------------------" & vbCrLf & _
           "** NOVÌ: **" & vbCrLf & _
           "- Pøidán changelog pro výpis zmìn v aplikaci." & vbCrLf & vbCrLf & _
           "** ZMÌNY: **" & vbCrLf & _
           "- Bylo upraveno pøihlašování do aplikace. Nyní nemusí uživatel mazat uživatelské heslo, pokud se pøihlašuje prostøednictvím Windows." & vbCrLf & _
           "- Byl zmìnìn formát u datumù z DD.MM.RRRR na DD.MM.RR." & vbCrLf & _
           "- Byl zmìnìn výchozí font. Nyní je použit font 'Segoe UI Semilight', který by mìl být k dispozici i ve starších verzích Windows. Pokud ho nenalezne, zkouší pøi spuštìní dohledat vhodnou alternativu a na konci použitému fontu pøizpùsobí šíøku sloupcù u datumových formátù." & vbCrLf & _
           "- Bylo upraveno naèítání dat. Nyní nemusí uživatel explicitnì naèítat data, ale k naètení dojde bezprostøednì po pøihlášení do aplikace. Následnì se aktualizují i zakázky." & vbCrLf & vbCrLf & _
           "** OPRAVENÉ CHYBY: **" & vbCrLf & _
           "- Množství u zakázky se nezobrazovalo správnì." & vbCrLf & _
           "- V nìkterých pøípadech se nezobrazoval správnì datum (byly úzké sloupce)." & vbCrLf

changelog = logPart1_5 & logPart1_4 & logPart1_3

    With txtChangelog
        .Value = changelog
        .ScrollBars = fmScrollBarsVertical ' Povolit svislé posuvníky pro dlouhý text
        .Enabled = True ' Zamezit úpravám textu
        .Locked = True  ' Uzamèení proti úpravám, ale aktivní text
    End With
End Sub

