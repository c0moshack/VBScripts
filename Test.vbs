set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
set oShellLink = WshShell.CreateShortcut(strDesktop _
  & "\MyExcel.lnk")
oShellLink.TargetPath = _
  "C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE"
oShellLink.WindowStyle = 1
oShellLink.Hotkey = "CTRL+SHIFT+F"
oShellLink.IconLocation = _
  "C:\Program Files\Microsoft Office\OFFICE11\EXCEL.EXE, 0"
oShellLink.Description = "My Excel Shortcut"
oShellLink.WorkingDirectory = strDesktop
oShellLink.Save