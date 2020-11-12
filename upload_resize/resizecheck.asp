<%
'	Author: Danilo Cicognani
'	Script: resizecheck.asp
'	Version: 1.00
'	Date: 13/08/2009
'	Use: Controlla se una delle librerie AspImage or Microsoft Office Web Components e' presente sul server
'	Copyright (c) 2009 Danilo Cicognani
'	https://www.ciconet.it
'	You can use, modify and redistribute this file, but you must mantain this copyright notice

'Check resizing libraries
Function CheckResizeLib()
	Dim CheckImageObj
	On Error Resume Next
	Err = 0
	Set CheckImageObj = Server.CreateObject("AspImage.Image")
	If Err = 0 Then
		CheckResizeLib = true
	Else
		Err.Clear
		Set CheckImageObj = CreateObject("OWC10.ChartSpace")
		If Err = 0 Then
			CheckResizeLib = true
		Else
			CheckResizeLib = false
		End If
	End If
	'CheckResizeLib = false
End Function
%>