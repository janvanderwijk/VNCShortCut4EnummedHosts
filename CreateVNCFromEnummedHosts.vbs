'Option Explicit
On Error Resume Next
Dim sComputerName, sInput, oWMIService, colComputers, colRunningServices, sArgument, oService,oFSO, oFile, aComputers, WshShell, curDir, wShell, file, filename

Set wShell = WScript.CreateObject("Shell.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")
Set FileSystem = WScript.CreateObject("Scripting.FileSystemObject")





'If no arguments are provided, run the script on the local computer.
'
If Wscript.Arguments.Count = 0 Then
	sComputerName = "."
	Call ServStat
	Wscript.Quit
End If

sInput = lcase(Wscript.Arguments(0))
Select Case sInput

	Case "file"
		'BEGIN COMMENT LINE
		'Give the INPUT_FILE_NAME constant a value and specify that you'll be reading the file.
		
		Const INPUT_FILE_NAME = "Enumhosts.txt"
		Const FOR_READING = 1
		'BEGIN COMMENT LINE
		'Create an object for the file.
		
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		'BEGIN COMMENT LINE
		'Open the file.
		
		Set oFile = oFSO.OpenTextFile(INPUT_FILE_NAME, FOR_READING)
		'BEGIN COMMENT LINE
		'Read the file and assign its contents to the sComputerName variable.
		
		sComputerName = oFile.ReadAll
		'BEGIN COMMENT LINE
		'Close the file.
		
		oFile.Close
		aComputers = Split(sComputerName, vbCrLf)
		For Each sComputerName in aComputers
			'WScript.Echo vcCrLf & "Results From"& vbTab & sComputerName & vbCrLf
			Call ServStat
		Next
		Wscript.Quit
    
	Case "all"
		'Wscript.Echo "Script will check all computers running WMI."
	'Set colComputers = GetObject(LDAP://CN=Computers, DC=labrynth, DC=com).
		For Each oComputer in colComputers
			sComputerName = oComputer.CN
			Call ServStat
		Next

	Case "/?"
		Wscript.Echo "Instructions"
		Wscript.Echo "This script checks service status on the selected computer(s)"
		Wscript.Echo  "so long as they're running WMI."
		Wscript.Echo "To check service status for all hosts in the Active Directory"
		Wscript.Echo "type 'all'."
		Wscript.Echo "To read server names from a file, type 'file'."
		Wscript.Echo "To check status on the local computer provide no arguments."

	Case Else
		For Each sArgument in WScript.Arguments
			sComputerName = sArgument
			'WScript.Echo vcCrLf & "Results From"& vbTab & sComputerName & vbCrLf
			Call ServStat
		Next

End Select

Sub ServStat
'BEGIN COMMENT LINE
'Connect to the WMI service on the target computer.

Set oWMIService  = GetObject("winmgmts:" _
	& "{impersonationLevel=impersonate}!\\" & sComputerName& "\root\cimv2")
'BEGIN COMMENT LINE
'Get a list of available services.

Set colRunningServices = oWMIService.ExecQuery _
	("Select * from Win32_Service")
'BEGIN COMMENT LINE
'Run through the list and report the service's name and status.

For Each oService in colRunningServices
	    If InStr (oService.DisplayName, "Terminal") Then
			If InStr (oService.State, "Running") Then
			 	'Wscript.Echo sComputerName & ": " & oService.DisplayName  & VbTab & oService.State

				Set OutPutFile = FileSystem.OpenTextFile(sComputerName & ".vnc",2,True)
				OutPutFile.WriteLine "[Connection]"
				OutPutFile.WriteLine "Host=" & sComputerName
				OutPutFile.WriteLine "[Options]"
				OutPutFile.WriteLine "UseLocalCursor=1"
				OutPutFile.WriteLine "UseDesktopResize=1"
				OutPutFile.WriteLine "FullScreen=0"
				OutPutFile.WriteLine "FullColour=0"
				OutPutFile.WriteLine "LowColourLevel=1"
				OutPutFile.WriteLine "PreferredEncoding=ZRLE"
				OutPutFile.WriteLine "AutoSelect=1"
				OutPutFile.WriteLine "Shared=0"
				OutPutFile.WriteLine "SendPtrEvents=1"
				OutPutFile.WriteLine "SendKeyEvents=1"
				OutPutFile.WriteLine "SendCutText=1"
				OutPutFile.WriteLine "AcceptCutText=1"
				OutPutFile.WriteLine "DisableWinKeys=1"
				OutPutFile.WriteLine "Emulate3=0"
				OutPutFile.WriteLine "PointerEventInterval=0"
				OutPutFile.WriteLine "Monitor="
				OutPutFile.WriteLine "MenuKey=F8"
				OutPutFile.WriteLine "AutoReconnect=1"
			End If	
		End If
		Next
End Sub
curDir = WshShell.CurrentDirectory
OutPutFile.Close
Set wShell = Nothing
Set WshShell = Nothing
Set FileSystem = Nothing
Set OutPutFile = Nothing
WScript.Quit(0)
