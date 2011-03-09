'Option Explicit

Dim objRootDSE, strDNSDomain, adoConnection, adoCommand, strQuery
Dim adoRecordset, strComputerDN, strOS
Dim WshShell, curDir, wShell, file

Set wShell = WScript.CreateObject("Shell.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")
Set OutPutFile = FileSystem.OpenTextFile("Enumhosts.txt",2,True)

' Determine DNS domain name from RootDSE object.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")

' Use ADO to search Active Directory for all computers.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
curDir = WshShell.CurrentDirectory
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

strQuery = "<LDAP://" & strDNSDomain & ">;(objectCategory=computer);" & "sAMAccountName,operatingSystem;subtree"

adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

Set adoRecordset = adoCommand.Execute

' Enumerate computer objects with server operating systems.
Do Until adoRecordset.EOF
    strOS = adoRecordset.Fields("operatingSystem").Value
    If (InStr(UCase(strOS), "XP") > 0) Then
        
        strComputerDN = adoRecordset.Fields("sAMAccountName").Value

OutPutFile.WriteLine Left(strComputerDN,Len(strComputerDN)-1)


    End If
    adoRecordset.MoveNext
Loop
adoRecordset.Close
OutPutFile.Close

' Clean up.
adoConnection.Close
Set objRootDSE = Nothing
Set adoCommand = Nothing
Set adoConnection = Nothing
Set adoRecordset = Nothing

Set wShell = Nothing
Set WshShell = Nothing
Set FileSystem = Nothing
Set OutPutFile = Nothing
