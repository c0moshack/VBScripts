' Proxy Switcher VB Script for Windows
' Neil Summers 30 Jan 2013
' Recommend copy code to file Proxy Switcher.vbs and save to Desktop
' No Copyright - Freeware - Open Source - do what you want with it
' Maybe buy me a beer down my local if it's THAT good for you!
Option Explicit 
Dim WSHShell, strSetting
Set WSHShell = WScript.CreateObject("WScript.Shell")

'Determine current proxy setting and toggle to opposite setting
strSetting = wshshell.regread("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
If strSetting = 1 Then 
NoProxy
 Else Proxy
End If

'Subroutine to Toggle Proxy Setting to ON
Sub Proxy 
WSHShell.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 1, "REG_DWORD"
Set WshShell = Wscript.CreateObject("Wscript.Shell")
WshShell.Popup "Proxy On",1,"Proxy Status",64
End Sub

'Subroutine to Toggle Proxy Setting to OFF
Sub NoProxy 
WSHShell.regwrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable", 0, "REG_DWORD"
Set WshShell = Wscript.CreateObject("Wscript.Shell")
WshShell.Popup "Proxy Off",1,"Proxy Status",64
End Sub