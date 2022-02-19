Set objShell = CreateObject("WScript.Shell")

strProcessCounter = "0"
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Set colProcessList = objWMIService.ExecQuery ("SELECT * FROM Win32_Process Where Name LIKE 'notepad.exe'")

For Each objProcess in colProcessList
	colProperties = objProcess.GetOwner(strNameOfUser,strUserDomain)
	If (strNameOfUser = objShell.ExpandEnvironmentStrings("%UserName%")) then
		strProcessCounter = strProcessCounter + 1
	End if
Next
If (strProcessCounter >= 1) then
      MsgBox "Приложение уже запущено, одного экземпляра достаточно.", 16, "WTF???"
	Wscript.Quit
End if

objShell.Run("""C:\Windows\system32\notepad.exe""")

