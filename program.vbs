'env
'Path ติดตั้ง webapp สามารถเปลี่ยนได้
webapp_name = "webapp"
webapp_port = "8080"
webapp_path = "C:\webapp\"
'รูป icon ของwebapp สามารถเปลี่ยนได้
icon_file ="favicon.ico"
'---------------------------------------------------------------------------------
'copy icon
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(webapp_path) Then
	Set objFolder = objFSO.GetFolder(webapp_path)
Else
  objFolder = objFSO.CreateFolder(webapp_path)
End If

'Dim FSO
'Set FSO = CreateObject("Scripting.FileSystemObject")
'FSO.CopyFile "favicon.ico", webapp_path

objFSO.CopyFile "favicon.ico", webapp_path

'--------------------------------------------------------------------------------
'get IP
ip = "localhost"
strMsg = ""
strComputer = "."
strMsg = ""
strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'")

For Each IPConfig in IPConfigSet
 If Not IsNull(IPConfig.IPAddress) Then
 For i = LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
  If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
  strMsg = strMsg & IPConfig.IPAddress(i) & vbcrlf
  End If
 Next
 End If
Next

ip=strMsg
'----------------------------------------------------------------------------------
'create shortcut
Set sh = CreateObject("WScript.Shell")
Set shortcut = sh.CreateShortcut(webapp_name+".lnk")
'แก้ไข IE หรือโปรแกรมอื่นได้
shortcut.TargetPath = "C:\Program Files\Internet Explorer\IEXPLORE.EXE"
'แก้ไข Protocal ได้
shortcut.Arguments = "http://"+ip+":"+webapp_port+"/"+webapp_name
shortcut.IconLocation= webapp_path +icon_file
shortcut.Save

wscript.echo "Completed create shotcut:"+shortcut.Arguments
'----------------------------------------------------

