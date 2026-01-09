Attribute VB_Name = "Connection"
' Deklarace globální promìnné pro uchování pøihlašovacích údajù
Public loginCredentials As Collection

' Deklarace klíèe pro šifrování a dešifrování
Public Const ENCRYPTION_KEY As String = "a7$D9!pR#3fG5&Z1"

' Deklarace globální promìnné pro sledování pøihlášení
Public isLoggedIn As Boolean

' Funkce pro šifrování a dešifrování
Function XOREncryptDecrypt(text As String, key As String) As String
    Dim i As Integer
    Dim result As String
    result = ""
    For i = 1 To Len(text)
        result = result & Chr(Asc(Mid(text, i, 1)) Xor Asc(Mid(key, (i Mod Len(key)) + 1, 1)))
    Next i
    XOREncryptDecrypt = result
End Function

Function VerifyCredentials(userName As String, password As String, ServerName As String, databaseName As String) As Boolean
    Dim conn As Object
    Dim connectionString As String
    On Error GoTo ErrorHandler

    If userName = "" And password = "" Then
        ' NT ovìøení
        connectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
    Else
        ' Pøihlášení pomocí uivatelského jména a hesla
        connectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & databaseName & ";User ID=" & userName & ";Password=" & password & ";"
    End If
    
    ' Kontrolní vıpis pøipojovacího øetìzce
    ' MsgBox "Connection String pro VerifyCredentials: " & connectionString
    
    ' Vytvoøení objektu pro pøipojení a otevøení pøipojení
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connectionString
    
    ' Pokud se pøipojení podaøilo, zavøi pøipojení a vra True
    If conn.State = 1 Then ' 1 = adStateOpen
        conn.Close
        VerifyCredentials = True
    Else
        VerifyCredentials = False
    End If
    Exit Function

ErrorHandler:
    ' Pokud došlo k chybì, vra False
    MsgBox "Connection Error pro VerifyCredentials: " & Err.Description
    VerifyCredentials = False
End Function

' Funkce pro naèítání dešifrovanıch pøihlašovacích údajù
Function GetDecryptedCredentials() As Collection
    If loginCredentials Is Nothing Then
        MsgBox "Pøihlašovací údaje nebyly nalezeny. Pøihlaste se prosím znovu.", vbExclamation
        frmLogin.Show
    Else
        Set GetDecryptedCredentials = loginCredentials
    End If
End Function

' Funkce pro vytvoøení pøipojení k databázi
Function CreateConnection() As Object
    Dim conn As Object
    Dim connectionString As String
    Dim ServerName As String
    Dim databaseName As String
    Dim login As String
    Dim passw As String
    Dim key As String
    Dim localServer As String
    Dim wsKonfig As Worksheet
    Dim credentials As Collection
    
    ' Nastavení listu se zdrojovımi daty
    Set wsKonfig = ThisWorkbook.Sheets("Konfigurace")
    
    ' Naètení nešifrovanıch údajù
    ServerName = wsKonfig.Range("serverName").Value
    databaseName = wsKonfig.Range("databaseName").Value
    localServer = wsKonfig.Range("localServer").Value
    
    ' Naètení dešifrovanıch pøihlašovacích údajù z globální promìnné
    Set credentials = GetDecryptedCredentials()
    
    ' Pokud pøihlašovací údaje nebyly nalezeny, ukonèi funkci
    If credentials Is Nothing Then
        MsgBox "Nepodaøilo se naèíst pøihlašovací údaje.", vbCritical
        Exit Function
    End If
    
    ' Klíè pro šifrování a dešifrování
    key = ENCRYPTION_KEY
           
    login = XOREncryptDecrypt(credentials("login"), key)
    passw = XOREncryptDecrypt(credentials("passw"), key)
    
    ' Kontrolní vıpis dešifrovanıch hodnot pøihlašovacích údajù
    ' MsgBox "Decrypted Login: " & login & vbCrLf & "Decrypted Password: " & passw
    
    If login = "" And passw = "" Then
        ' NT ovìøení
        connectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & databaseName & ";Integrated Security=SSPI;"
    Else
        ' Pøihlášení pomocí uivatelského jména a hesla
        connectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & databaseName & ";User ID=" & login & ";Password=" & passw & ";"
    End If

    ' Kontrolní vıpis pøipojovacího øetìzce
    Debug.Print "Connection String: " & connectionString
    
    ' Vytvoøení objektu pro pøipojení a otevøení pøipojení
    Set conn = CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open connectionString
    On Error GoTo 0
    
    ' Pokud se pøipojení podaøilo, vra objekt pøipojení
    If conn.State = 1 Then ' 1 = adStateOpen
        Set CreateConnection = conn
    Else
        Set CreateConnection = Nothing
        MsgBox "Nepodaøilo se otevøít pøipojení.", vbCritical
    End If
End Function




