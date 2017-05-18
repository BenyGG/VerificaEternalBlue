' Checks if the computer is vulnerable to EternalBlue
' Comprueba si el equipo es vulnerable a EternalBlue
'
'	Author: Cassius Puodzius
'   Re-Fix: Bastian Novoa
'   Traduccion: Spanish

Dim KBList
Set KBList = CreateObject("System.Collections.ArrayList")

KBList.Add "4012212"
KBList.Add "4012213"
KBList.Add "4012214"
KBList.Add "4012215"
KBList.Add "4012216"
KBList.Add "4012217"
KBList.Add "4012598"
KBList.Add "4012606"
KBList.Add "4013198"
KBList.Add "4013429"

Set WshShell = CreateObject("WScript.Shell")

CreateObject("WScript.Shell").Popup "Obteniendo la lista de archivos KBs Instalados " & vbcrlf & vbcrlf & " Esto puede tomar algunos minutos...", 5, " Vulnerable a EternalBlue?", vbOKOnly

'Dim HotFixList : Set HotFixList = WshShell.Exec("wmic qfe list")
WshShell.Run "cmd /c wmic qfe list | clip", 0, True

For Each kb In KBList
	If InStr(CreateObject("htmlfile").ParentWindow.ClipboardData.GetData("text"), "KB" + kb) > 0 Then
		MsgBox("Su computadora está protegida contra EternalBlue (KB" + kb + ")")
		WScript.Quit
	End If
Next

MsgBox("Su Computadora NO ESTA PROTEGIDA ANTE EternalBlue, Por favor Actualice su equipo" & vbcrlf & vbcrlf & "Se Forzara al Sistema para Instalar Actualizaciones Ahora!")

' Ejecucion de Windows Update
' Opcion para buscar Actualizaciones!
'WshShell.run "control.exe /name Microsoft.WindowsUpdate",0,True  //Esta Opcion Ejecutaba la GUI de Windows Update
'Forzar al Sistema a Actualizar!
WshShell.run ("wuauclt /detectnow")
' Desactivar SMBv1
CreateObject("WScript.Shell").Popup "Desactivando SMBV1 de tu Sistema! ", 5, " Desactivando!", vbOKOnly
WshShell.run "sc.exe config lanmanworkstation depend= bowser/mrxsmb20/nsi",0,True
WshShell.run "sc.exe config mrxsmb10 start= disabled",0,True
'Forzamos al sistema por powershell! a desactivar el SMBV1
WshShell.run ("powershell Set-ItemProperty -Path “HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters” SMB1 -Type DWORD -Value 0 -Force")
	'Fueron Desactivas estas Opciones ya que puede que alguna oficina o entidad pueda ocupar este Script ,
	'Pero si estas en tu casa puedes Activarlas!
' Desactivar SMBV2
'CreateObject("WScript.Shell").Popup "Desactivando SMBV2 de tu Sistema! ", 5, " Desactivando!", vbOKOnly
'WshShell.run "sc.exe config lanmanworkstation depend= bowser/mrxsmb10/nsi  ",0,True
'WshShell.run "sc.exe config mrxsmb20 start= disabled",0,True
' Desactivar SMBV3
'CreateObject("WScript.Shell").Popup "Desactivando SMBV3 de tu Sistema! ", 5, " Desactivando!", vbOKOnly
'WshShell.run "sc.exe config lanmanworkstation depend= bowser/mrxsmb10/mrxsmb20/nsi  ",0,True
'WshShell.run "sc.exe config mrxsmb20 start= auto",0,True
