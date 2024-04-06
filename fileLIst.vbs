Option Explicit
Dim i, fso, ts, ws, currentPath, cnt, fileList
Const ForWriting = 2
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("scripting.FileSystemObject")
currentPath = fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path
fileList = currentPath + "\" + "fileLIst.txt"
Set ts = fso.OpenTextFile(fileList, ForWriting, True)
cnt = WScript.Arguments.Count
if Not fso.FileExists(fileList) Then
	fso.CreateTextFile(fileList)
end if
For i = 0 To cnt - 1
	ts.WriteLine "file" + " " + "'" + WScript.Arguments(i) + "'"
Next