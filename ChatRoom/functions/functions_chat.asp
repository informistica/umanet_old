
<!-- #include file = "../../extra/test_server.asp" -->
<%
'Const AdminPassword = "pass1234"


'**********************************************************
'**********************************************************
'**********************************************************
'**********************************************************
' DONT EDIT BELOW UNLESS YOU KNOW WHAT YOU ARE DOING
'**********************************************************
'**********************************************************




Const ApplicationMsg = "ChatRoomMsg"
Dim esecuzione
set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito
   
Function lastMessageID()
	Dim saryMessages

	'Get the array
	If IsArray(Application(ApplicationMsg)) Then
		saryMessages = Application(ApplicationMsg)
	Else
		ReDim saryMessages(6, 0)

		Application.Lock
		Application(ApplicationMsg) = saryMessages
		Application.UnLock
	End If

	lastMessageID = saryMessages(5, UBound(saryMessages, 2))
End Function

Function PostMessage(strUsername, strMessage, intType)
	Application.Lock

	Dim saryMessages
	Dim saryTempArray
	Dim intArrayPass

	'Get the array
	If IsArray(Application(ApplicationMsg)) Then
		saryMessages = Application(ApplicationMsg)
	Else
		ReDim saryMessages(5, 0)

		Application(ApplicationMsg) = saryMessages
	End If

	Dim strColor
	Dim strFormat
	Dim intLastMessageID

	'Double up slahes so it pashes through the javascript and displays
	strMessage = Replace(strMessage, "\", "\\")

	'Remove HTML if any
	strMessage = removeAllTags(strMessage)

	'If a link, format it
	strMessage = CheckForLink(strMessage)

	'format the message
	strMessage = FormatMessage(strMessage)

	If strMessage <> "" Then
		'Array Legend
		'0 = Author
		'1 = Message
		'2 = Date
		'3 = Type
		'4 = User ID, 0 = All
		'5 = Message ID

		Dim intTempSize

		intTempSize = UBound(saryMessages, 2)

		If intTempSize = 0 Then
			intLastMessageID = 0
		Else
			intLastMessageID = Clng(saryMessages(5, intTempSize))
		End If

		intTempSize = intTempSize + 1

		ReDim Preserve saryMessages(5, intTempSize)

		saryMessages(0, intTempSize) = strUsername
		saryMessages(1, intTempSize) = strMessage
		saryMessages(2, intTempSize) = CDbl(Now())
		saryMessages(3, intTempSize) = intType
		saryMessages(4, intTempSize) = 0
		saryMessages(5, intTempSize) = (intLastMessageID + 1)

		Application(ApplicationMsg) = saryMessages

		'******************************************
		'***   	Trim array if over 40 messages	***
		'******************************************
		If UBound(saryMessages, 2) => 20 Then
			'put array in a temp array so we can update it
			ReDim saryTempArray(5, 0)

			'cut the array in half
			For intArrayPass = 10 TO UBound(saryMessages, 2)
				ReDim Preserve saryTempArray(5, UBound(saryTempArray, 2) + 1)

				saryTempArray(0, UBound(saryTempArray, 2)) = saryMessages(0, intArrayPass)
				saryTempArray(1, UBound(saryTempArray, 2)) = saryMessages(1, intArrayPass)
				saryTempArray(2, UBound(saryTempArray, 2)) = saryMessages(2, intArrayPass)
				saryTempArray(3, UBound(saryTempArray, 2)) = saryMessages(3, intArrayPass)
				saryTempArray(4, UBound(saryTempArray, 2)) = saryMessages(4, intArrayPass)
				saryTempArray(5, UBound(saryTempArray, 2)) = saryMessages(5, intArrayPass)
			Next

			'Transfer array to update
			saryMessages = saryTempArray

			Application(ApplicationMsg) = saryMessages
		End If
	End If

	Application.UnLock
End Function

Function FormatMessage(strMessage)

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

'*********************************************
'***  	Check and Format Links	 	*****
'*********************************************

Function CheckForLink(strMessage)

	'Dim regEx

	'Set regEx = New RegExp

	'regEx.IgnoreCase = True
	'regEx.Global = True

	'regEx.Pattern = "((https|ftp)://\S+)"            'https://
	'strMessage = regEx.Replace(strMessage, "<a href=\""$1\"" target=\""_new\"">$1</a>")

	'regEx.Pattern = "([^/])(www\.\S+)"            'www.xxx.xxx
	'strMessage = regEx.Replace(strMessage, "$1<a href=\""https://$2\"" target=\""_new\"">$2</a>")

	'regEx.Pattern = "(\S+@\S+\.\S+)"            ' match email addresses
	'strMessage = regEx.Replace(strMessage, "<a href=\""mailto:$1\"">$1</a>")   ' Make replacement.

' MIO CODICE
'if inStr(strMessage,"/img_social/img/")=0 then
' faccio sostituire solo se non è già presente il link perchè è stato messo durante la creazione dell'immaginario, vedi problema in note_personali.txt 
 if inStr(strMessage,"<a target=_blank href='https://www.youtube.com/watch")=0 then

	if inStr(strMessage,"www.youtube.com")<>0 then
	 ' strMessage1= "<a href='" & strMessage & "' target=blank >" 
	     strMessage1=strMessage
	     strMessage = Replace(strMessage,strMessage,"<a target=_blank href=' " & strMessage1 & "'><img title='vai youtube' src=images/play.jpg width=50 height=36 /></a>")
	else if inStr(strMessage,"Immagini/Rete_di_Immagini/")=0 and inStr(strMessage,"img_social/img/")=0 and inStr(strMessage,"www.")<>0 then
	         strMessage1=strMessage
			
			 ' if instr(strMessage,"https")=0 then
'	         strMessage = Replace(strMessage,strMessage,"<a target=_blank href=' https://" & strMessage1 & "'>" &strMessage1 &"</a>")
'			 else
'			 strMessage = Replace(strMessage,strMessage,"<a target=_blank href='" & strMessage1 & "'>" &strMessage1 &"</a>")
'			 end if		 
	    end if
	end if
end if


'end if	




 



	CheckForLink=strMessage



End Function

'*********************************************
'***  		Strip all tags		 *****
'*********************************************

'Remove all tags for text only display (mainly for subject lines)
Private Function removeAllTags(ByVal strInputEntry)

	'Remove all HTML scripting tags etc. for plain text output
	strInputEntry = Replace(strInputEntry, "&", "&amp;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "<", "&lt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ">", "&gt;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "'", "&#146;", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, """", "&quot;", 1, -1, 1)

	'Return
	removeAllTags = strInputEntry
End Function

Function CleanMessage(strMessage)
	strMessage = Replace(strMessage, "[B]", "", 1, -1, 1)
	strMessage = Replace(strMessage, "[/B]", "", 1, -1, 1)
	strMessage = Replace(strMessage, "[I]", "", 1, -1, 1)
	strMessage = Replace(strMessage, "[/I]", "", 1, -1, 1)
	strMessage = Replace(strMessage, "[U]", "", 1, -1, 1)
	strMessage = Replace(strMessage, "[/U]", "", 1, -1, 1)

	Dim lngMessagePosition
	Dim intHTMLTagLength
	Dim strHTMLMessage
	Dim strTempMessageText

	strTempMessageText = strMessage

	For lngMessagePosition = 1 to CLng(Len(strMessage))

		'If this is the end of the message then save some process time and jump out the loop
		If Mid(strMessage, lngMessagePosition, 1) = "" Then Exit For

		'If an BBCode tag is found then jump to the end so we can strip it
		If Mid(strMessage, lngMessagePosition, 6) = "[color" OR Mid(strMessage, lngMessagePosition, 7) = "[/color" Then

			'Get the length of the BBCode tag
			intHTMLTagLength = (InStr(lngMessagePosition, strMessage, "]", 1) - lngMessagePosition)

			'If the end of the BBCode string is in error then set it to the number of characters being passed
			If intHTMLTagLength < 0 Then intHTMLTagLength = CLng(Len(strTempMessageText))

			'Place the BBCode tag back into the temporary message store
			strHTMLMessage = Mid(strMessage, lngMessagePosition, intHTMLTagLength + 1)

			'Strip the BBCode
			strTempMessageText = Replace(strTempMessageText, strHTMLMessage, " ", 1, -1, 0)
		End If
	Next

	CleanMessage = strTempMessageText
End Function
%>