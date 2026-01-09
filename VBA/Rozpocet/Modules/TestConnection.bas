Attribute VB_Name = "TestConnection"
Option Explicit

Sub TestConnection()
    Dim ServerName As String
    Dim databaseName As String
    Dim userName As String
    Dim password As String
    Dim conn As Object
    Dim connectionString As String
    
    ' Nastavte zde své hodnoty
    ServerName = "192.168.25.5\HELIOS"
    databaseName = "iNuvio001"
    userName = "vintrlik"
    password = ""
    
    connectionString = "Provider=SQLOLEDB;Data Source=" & ServerName & ";Initial Catalog=" & databaseName & ";User ID=" & userName & ";Password=" & password & ";"
    
    ' Kontrolní výpis pøipojovacího øetìzce
    MsgBox "Test Connection String: " & connectionString
    
    On Error GoTo ErrorHandler
    Set conn = CreateObject("ADODB.Connection")
    conn.Open connectionString
    
    If conn.State = 1 Then ' 1 = adStateOpen
        MsgBox "Pøipojení bylo úspìšné!", vbInformation
        conn.Close
    Else
        MsgBox "Pøipojení se nezdaøilo.", vbCritical
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Chyba pøipojení: " & Err.Description, vbCritical
End Sub

Sub TestEncryption()
    Dim originalText As String
    Dim encryptedText As String
    Dim decryptedText As String
    Dim key As String
    
    originalText = "Testkryptování"
    key = ENCRYPTION_KEY
    
    ' Šifrování
    encryptedText = XOREncryptDecrypt(originalText, key)
    MsgBox "Encrypted Text: " & encryptedText
    
    ' Dešifrování
    decryptedText = XOREncryptDecrypt(encryptedText, key)
    MsgBox "Decrypted Text: " & decryptedText
End Sub


