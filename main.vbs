Option Explicit
Dim i, file(),cnt,ws,cmdstr,n,fso,CMDlog, currentpath
Set fso = CreateObject("Scripting.FileSystemObject")
currentpath = fso.GetFile(Wscript.ScriptFullName).ParentFolder.Path
cnt=WScript.Arguments.Count
If Not fso.folderExists(currentpath + "\logs") Then fso.createfolder(currentpath + "\logs")
If Not fso.folderExists(currentpath + "\out") Then fso.createfolder(currentpath + "\out")
redim file(cnt)
For i = 0 To WScript.Arguments.Count - 1
    file(i+1) = WScript.Arguments(i)
    'WScript.Echo """" + file(i+1) + """"
Next
'ffmpeg -i input.mp4 -c:v hevc_nvenc -b:v 7000k -preset slow -c:a libopus -b:a 128k -map 0 -threads 1 output.mp4
set ws=CreateObject("WScript.Shell")
i=0
for i=1 to WScript.Arguments.Count
    cmdstr="ffmpeg -y -hide_banner -i " + """" + file(i) + """" + " -c:v hevc_nvenc -b:v 2500k -preset slow -c:a libopus -b:a 128k -map 0 -threads 1  -map_chapters 0 -map_metadata 0  " + """" + currentpath + "\out\" + fso.GetFileName(file(i)) + """"
    CMDlog="cmd /c set FFREPORT=file=" + replace(replace(currentpath,"\","\\"),":","\:") + "\\logs\\%p-%t.log:level=48 && "
    cmdstr=CMDlog + cmdstr
    'msgbox cmdstr
    ws.Run cmdstr,3,true
next