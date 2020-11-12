<%Function FormatMessage(strMessage)
' della CHAT	
   
  ' QuerySQL="Select * from TUTTESMILES where ID_Categoria=1 order by Posizione;"
'   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
'   rsTabellaS.movefirst
'   do while not rsTabellaS.eof 
'        strMessage = Replace(strMessage,rsTabellaS("Codice"), "<img src=" & rsTabellaS("Cartella_Cat")&"/"&rsTabellaS("Url")&" align=absmiddle>")
'  	    rsTabellaS.movenext
'   loop	
' 
	'Smilies
	'strMessage = Replace(strMessage, ":huh?", "<img src=smilies/on_1.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":s", "<img src=smilies/on_2.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":P", "<img src=smilies/on_3.gif align=absmiddle>")
'	strMessage = Replace(strMessage, "}:)", "<img src=smilies/on_4.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":D", "<img src=smilies/on_5.gif align=absmiddle>")
'	strMessage = Replace(strMessage, "}:|", "<img src=smilies/on_6.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":)", "<img src=smilies/on_7.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":oops", "<img src=smilies/on_8.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ";)", "<img src=smilies/on_9.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":pff", "<img src=smilies/on_10.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":\\", "<img src=smilies/on_11.gif align=absmiddle>")
'	strMessage = Replace(strMessage, ":0", "<img src=smilies/on_12.gif align=absmiddle>")
'	
	strMessage = Replace(strMessage, ":b;", "<img src=smilies/on_13.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":xx", "<img src=smilies/on_14.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":gg", "<img src=smilies/on_15.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":nn", "<img src=smilies/on_16.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":pp", "<img src=smilies/on_17.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":kk", "<img src=smilies/on_18.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":yy", "<img src=smilies/on_19.gif align=absmiddle>")
	strMessage = Replace(strMessage, ":zz", "<img src=smilies/on_20.gif align=absmiddle>")
	
	
	  
	 ' Percezioni
 
 QuerySQL="Select * from TUTTESMILES where ID_Categoria<>1 order by Posizione;"
   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
    rsTabellaS.movefirst
   do while not rsTabellaS.eof 
        strMessage = Replace(strMessage,rsTabellaS("Codice"), "<img class='imground_shadow' src=../img_social/" & rsTabellaS("Cartella_Cat")&"/"&rsTabellaS("Url")&" align=absmiddle>")
  	    rsTabellaS.movenext
   loop	
 
 
 
	'strMessage = Replace(strMessage, ":;0_00", "<img src=../img_social/connessioni_percezioni/0_0_MatrixOmino_2.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_01", "<img src=../img_social/connessioni_percezioni/0_1_incredibile_solo.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_02", "<img src=../img_social/connessioni_percezioni/0_2_vedoilsole.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_03", "<img src=../img_social/connessioni_percezioni/0_3_NavigaOcchioSi.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_04", "<img src=../img_social/connessioni_percezioni/0_4__MondoCoerenzaVerde.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_05", "<img src=../img_social/connessioni_percezioni/0_5_LampadinaAccesa.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_06", "<img src=../img_social/connessioni_percezioni/0_6_TestaAccesa.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_07", "<img src=../img_social/connessioni_percezioni/0_7_vedopioggia.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_08", "<img src=../img_social/connessioni_percezioni/0_8_NavigaOcchioNo.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_09", "<img src=../img_social/connessioni_percezioni/0_9_MondoParadossoRosso.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_10", "<img src=../img_social/connessioni_percezioni/0_10_LampadinaSpenta.jpg align=absmiddle>")
'	strMessage = Replace(strMessage, ":;0_11", "<img src=../img_social/connessioni_percezioni/0_11_TestaSpenta.jpg align=absmiddle>")
'	 
	    
set rsTabellaS=nothing
 







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

