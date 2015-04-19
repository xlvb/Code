
Dim Web_Site_Link,DesktopPath


Set Shell = CreateObject("WScript.Shell")
DesktopPath = Shell.SpecialFolders("Desktop")
Web_Site_Link = "http://www.excelvbscript.com/"



' Add Web Site Link to the desktop
Set Xlvbs_Link = Shell.CreateShortcut(DesktopPath & "\xlvbs.lnk")

Xlvbs_Link.Description = "Excelvbscript.com"
Xlvbs_Link.IconLocation = ("%SystemRoot%\system32\SHELL32.dll,14")	
Xlvbs_Link.TargetPath = Web_Site_Link
Xlvbs_Link.HotKey = "CTRL+SHIFT+X"
Xlvbs_Link.WindowStyle = 3
Xlvbs_Link.Save

