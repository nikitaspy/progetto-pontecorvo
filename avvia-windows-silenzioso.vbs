Set shell = CreateObject("WScript.Shell")
Set files = CreateObject("Scripting.FileSystemObject")
folder = files.GetParentFolderName(WScript.ScriptFullName)
command = Chr(34) & folder & "\avvia-windows-silenzioso.bat" & Chr(34)
shell.Run command, 0, False
