

'''######################################################################
'''''http://www.excelvbscript.com/
'''######################################################################


' Variables 
Dim Web_Site_Link,DesktopPath

'Create Shell object
Set Shell = CreateObject("WScript.Shell")

'Get the Desktop
DesktopPath = Shell.SpecialFolders("Desktop")

'Update the Web Site Link 
Web_Site_Link = "http://www.excelvbscript.com/"



' Add Web Site Link to the desktop
Set Xlvbs_Link = Shell.CreateShortcut(DesktopPath & "\xlvbs.lnk")

Xlvbs_Link.Description = "Excelvbscript.com"
Xlvbs_Link.IconLocation = ("%SystemRoot%\system32\SHELL32.dll,14")	
Xlvbs_Link.TargetPath = Web_Site_Link
Xlvbs_Link.HotKey = "CTRL+SHIFT+X"
Xlvbs_Link.WindowStyle = 3
Xlvbs_Link.Save


'End of coding
' If you need more information on VB scripting drop a mail to "excelvbscript@gmail.com" 
'''######################################################################
