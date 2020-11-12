<%Function FormatMessage(strMessage)
	'Smilies
	'DEL FORUM SONO QUA !!!
	%>
	 
    <!-- #include file = "replace_codici_img.asp" -->
    
	<%
	
	strMessage = Replace(strMessage, "[B]", "<strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/B]", "</strong>", 1, -1, 1)
	strMessage = Replace(strMessage, "[I]", "<em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/I]", "</em>", 1, -1, 1)
	strMessage = Replace(strMessage, "[U]", "<u>", 1, -1, 1)
	strMessage = Replace(strMessage, "[/U]", "</u>", 1, -1, 1)

	'Loop through the message till all font colour codes are turned into fonts colours
	Do While InStr(1, strMessage, "[color=", 1) > 0  AND InStr(1, strMessage, "[/color]", 1) > 0
		Dim lngStartPos
		Dim lngEndPos
		Dim strMessageLink
		Dim strTempMessage

		'Find the start position in the message of the [COLOR= code
		lngStartPos = InStr(1, strMessage, "[color=", 1)

		'Find the position in the message for the [/COLOR] closing code
		lngEndPos = InStr(lngStartPos, strMessage, "[/color]", 1) + 8

		'Make sure the end position is not in error
		If lngEndPos < lngStartPos Then lngEndPos = lngStartPos + 9

		'Read in the code to be converted into a font colour from the message
		strMessageLink = Trim(Mid(strMessage, lngStartPos, (lngEndPos - lngStartPos)))

		'Place the message colour into the tempoary message variable
		strTempMessage = strMessageLink

		'Format the link into an font colour HTML tag
		strTempMessage = Replace(strTempMessage, "[color=", "<font color=", 1, -1, 1)
		'If there is no tag shut off place a > at the end
		If InStr(1, strTempMessage, "[/COLOR]", 1) Then
			strTempMessage = Replace(strTempMessage, "[/color]", "</font>", 1, -1, 1)
			strTempMessage = Replace(strTempMessage, "]", ">", 1, -1, 1)
		Else
			strTempMessage = strTempMessage & ">"
		End If

		'Place the new fromatted colour HTML tag into the message string body
		strMessage = Replace(strMessage, strMessageLink, strTempMessage, 1, -1, 1)
	Loop

	FormatMessage = strMessage
End Function
%>

