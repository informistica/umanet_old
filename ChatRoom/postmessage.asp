<% 'Option Explicit %>
<!--#include file="functions/functions_chat.asp"-->
<!--#include file="functions/functions_users.asp"-->
<!--#include file="functions/functions_ban.asp"-->
 <!-- #include file = "../var_globali.inc" -->
 
  <%Set ConnessioneDB = Server.CreateObject("ADODB.Connection") ' per il forum%>
 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
Application.Lock
'on error resume next
Dim saryMessages
Dim saryTempArray
Dim intArrayPass

Dim blnAdmin
Dim blnFormatText

strUsername = Session("Username")
blnAdmin = CBool(Session("Admin"))
blnFormatText = CBool(Session("FormatText"))

'If strUsername = "" OR CheckIfBanned(getIP()) Then Response.End

'Get the array
If IsArray(Application(ApplicationMsg)) Then
	saryMessages = Application(ApplicationMsg)
Else
	ReDim saryMessages(5, 0)

	Application(ApplicationMsg) = saryMessages
End If

Dim strMessage
Dim strColor
Dim strFormat
Dim intLastMessageID
Dim intType
Dim strCommand
Dim blnPrivateMessage
Dim saryUserMsgTo
Dim nomeFileChat, url1,origine,rsTabella,maxIdChat,titoloChat
Const ForReading = 1, ForWriting = 2, ForAppending = 8
  
' file che atesta la registrazione in corso  
origine=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Session("cartella") & "/Chatlog/registra_chat_in.txt"
origine=replace(origine,"\","/")				

blnPrivateMessage = False
intType = 0
saryUserMsgTo = 0

strMessage = Request.Form("message")
strColor = Request.Form("color")
strFormat = Request.Form("format")
' nel file scrivo per ogni messaggio queste tre righe 
' 
'function inHTML(sReadAll)
'   sReadAll=replace(sReadAll,"[color","<font color")
'   sReadAll=replace(sReadAll,"[/color]","</font>")  
'   sReadAll=replace(sReadAll,"[i]","<i>")
'   sReadAll=replace(sReadAll,"[/i]","</i>")
'   sReadAll=replace(sReadAll,"[b]","<b>")
'   sReadAll=replace(sReadAll,"[/b]","</b>")
'   sReadAll=replace(sReadAll,"]",">")
'   inHTML=sReadAll 
'
'end function
	
 
 
 
	
	            'dim objFSO,objCreatedFile
				'Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'Dim sRead, sReadLine, sReadAll, objTextFile
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logChat.txt"
'				Set objCreatedFile = objFSO.OpenTextFile(url, ForAppending,True)
'				objCreatedFile.WriteLine(strUsername &":"&strMessage & "-" & strColor &"-" &strFormat)
'				objCreatedFile.Close		
'end if


If Mid(strMessage, 1, 1) = "/" Then
	If Mid(strMessage, 1, 10) = "/password " Then
		strCommand = Trim(Mid(CleanMessage(strMessage), 10, Len(CleanMessage(strMessage))))

		If strCommand = AdminPassword Then
			Session("Admin") = True
			strMessage = "Logged in as admin."
		Else
			strMessage = "Failed to login. Password given: " & strCommand
		End If

		saryUserMsgTo = Array(strUsername)
		intType = 1
	ElseIf Mid(strMessage, 1, 7) = "/logout" Then
		Session("Admin") = False

		If blnAdmin Then
			strMessage = "Logged out."
		Else
			strMessage = "You where never logged in."
		End If

		saryUserMsgTo = Array(strUsername)
		intType = 1
	ElseIf Mid(strMessage, 1, 7) = "/alert " Then
		strCommand = Trim(Mid(strMessage, 7, Len(strMessage)))

		If blnAdmin Then
			strMessage = "alert('" & strCommand & "');"

			saryUserMsgTo = 0
			intType = 2
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 8) = "/kickall" Then
		If blnAdmin Then
			strMessage = "alert('You have been kicked!');parent.location=""../../home.asp"";"

			saryUserMsgTo = 0
			intType = 2
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 6) = "/kick " Then
		strCommand = Trim(Mid(strMessage, 6, Len(strMessage)))

		If blnAdmin Then
			If KickUser(strCommand) Then
				strMessage = strCommand & " has been kicked."
			Else
				strMessage = strCommand & " username not found."
			End If

			saryUserMsgTo = Array(strUsername)
			intType = 1
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 5) = "/ban " Then
		strCommand = Trim(Mid(strMessage, 5, Len(strMessage)))

		If blnAdmin Then
			If CheckIfBanned(strCommand) Then
				strMessage = strCommand & " is already in the ban list."
			Else
				Call BanUser(strCommand)
				strMessage = strCommand & " has been added to the ban list."
			End If

			saryUserMsgTo = Array(strUsername)
			intType = 1
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 7) = "/unban " Then
		strCommand = Trim(Mid(strMessage, 7, Len(strMessage)))

		If blnAdmin Then
			If CheckIfBanned(strCommand) Then
				Call UnBanUser(strCommand)
				strMessage = strCommand & " was removed from the ban list."
			Else
				strMessage = strCommand & " was not found in the ban list."
			End If

			saryUserMsgTo = Array(strUsername)
			intType = 1
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 8) = "/banlist" Then
		If blnAdmin Then
			Dim intBanLoop

			saryBanList = GetBanList()

			strMessage = "<hr size=1><b><u>Ban List (" & UBound(saryBanList) & "):</u></b><br>"


			If UBound(saryBanList) = 0 Then
				strMessage = strMessage & "No bans."
			Else
				For intBanLoop = 0 TO UBound(saryBanList) - 1
					strMessage = strMessage & saryBanList(intArrayPass) & " - <a onclick=\""insertText('/unban " & saryBanList(intArrayPass) & "')\"" style=\""cursor: hand\"">Unban</a><br>"
				Next
			End If

			strMessage = strMessage & "<hr size=1>"


			saryUserMsgTo = Array(strUsername)
			intType = 1
		Else
			strMessage = "You do not have permission to use this command."

			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	ElseIf Mid(strMessage, 1, 6) = "/name " Then
		strCommand = Trim(Mid(strMessage, 6, Len(strMessage)))

		Session("Username") = strCommand

		strMessage = strUsername & " changed username to " & strCommand

		saryUserMsgTo = 0
		intType = 1
	ElseIf Mid(strMessage, 1, 7) = "/format" Then
		If blnFormatText Then
			Session("FormatText") = False
			strMessage = "Text formating turned off."
		Else
			Session("FormatText") = True
			strMessage = "Text formating turned on."
		End If

		saryUserMsgTo = Array(strUsername)
		intType = 1
	
	ElseIf Mid(strMessage, 1, 9) = "/rec" and Session("Admin")=True Then ' abilito la registrazione della chat
		 
		        dim objFSO,objCreatedFile,registra
			'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
			'	Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				Set fso = CreateObject("Scripting.FileSystemObject")
				session("registra")=true
				nomeChat=year(date()) &"_"& month(date()) &"_" & day(date())&"_"& left(FormatDateTime(now(),4),2)&"_"&right(FormatDateTime(now(),4),2) &".txt" 
				nomeChat=Replace(nomeChat,":","_")
				url1=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Session("cartella") & "/Chatlog/" & nomeChat  
				url1=Replace(url1,"\","/")
				session("urlreg")=url1
		        ' inserisco nel file registra_chat_in il nome del file della chat corrente url1
				Set objCreatedFile = fso.CreateTextFile(origine, True)
		 	    objCreatedFile.WriteLine(url1)	  
		 	    objCreatedFile.WriteLine(strMessage)
		        objCreatedFile.Close
		        Set objCreatedFile = fso.CreateTextFile(url1, True)
		        strMessage="Inizio Chat registrata " & now() 
		       ' objCreatedFile.WriteLine(strMessage)
		        objCreatedFile.Close 
				 
				 
				 
				 
				
 

				'inserisco nel database la registrazione
				 	
  QuerySQL="INSERT INTO CHAT_SESSION (Id_Classe, Cartella, Nome,Inizio,Fine) " &_
  " SELECT '" & Session("Id_Classe") & "','" & Session("cartella") & "', '" & nomeChat & "','" & now() & "','" & now() &"';" 
			     ConnessioneDB.Execute(QuerySQL)
	
		 
	ElseIf Mid(strMessage, 1, 6) = "/norec" and Session("Admin")=True Then ' abilito la registrazione della chat
	            titoloChat= trim(right(strMessage, len(strMessage)-6) )
				if titoloChat="" then
				  titoloChat="Senza titolo"
				end if
	 	        Session("registra")=false
				Set fso = CreateObject("Scripting.FileSystemObject")
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				Set objCreatedFile = fso.OpenTextFile(session("urlreg"), ForAppending,True)
		         
		        strMessage="Fine Chat registrata " & now()  
		        objCreatedFile.WriteLine(strMessage)
		        objCreatedFile.Close 
				fso.DeleteFile origine
				QuerySQL="select max(ID_Chat) from CHAT_SESSION where Id_Classe='" & Session("Id_Classe") &"';"
				set rsTabella=ConnessioneDB.Execute(QuerySQL)
				maxIdChat=rsTabella(0)
				QuerySQL ="UPDATE CHAT_SESSION SET Titolo='" & titoloChat &"', Fine = '" & now() & "' WHERE ID_Chat =" &maxIdChat&";"
				 ConnessioneDB.Execute(QuerySQL)	 
				 set ConnessioneDB = nothing
				
				
				
	ElseIf Mid(strMessage, 1, 9) = "/commands" Then

		strMessage = "<hr size=1><table><tr><td colspan=2><b><u>Admin Commands:</u></b></td></tr>" & _
			"<tr><td><b>Registra:</b></td><td>/rec</td></tr>" & _
			"<tr><td><b>Termina registra:</b></td><td>/norec</td></tr>" & _
			"<tr><td><b>Login:</b></td><td>/password [Password]</td></tr>" & _
			"<tr><td><b>Logout:</b></td><td>/logout</td></tr>" & _
			"<tr><td><b>Kick:</b></td><td>/kick [Username]</td></tr>" & _
			"<tr><td><b>Ban List:</b></td><td>/banlist</td></tr>" & _
			"<tr><td><b>Ban:</b></td><td>/ban [IP]</td></tr>" & _
			"<tr><td><b>Unban:</b></td><td>/unban [IP]</td></tr>" & _
			"<tr><td><b>Alert:</b></td><td>/alert [Message]</td></tr>" & _
			"<tr><td colspan2><br><b><u>User Commands:</u></b></td></tr>" & _
			"<tr><td><b>Private Message:</b></td><td>/[Username] [Message]</td></tr>" & _
			"<tr><td><b>Change Name:</b></td><td>/name [Username]</td></tr>" & _
			"<tr><td><b>Text Formating on/off:</b></td><td>/format</td></tr>" & _
			"<tr><td colspan=2><br><b><u>Smilies:</u></b></td></tr>" & _
			"<tr><td colspan=2><table cellspacing=1 cellpadding=4 bgcolor=\""#EAEAEA\"">" & _
			"<tr><td bgcolor=white><img src=smilies/on_1.gif align=absmiddle></td><td bgcolor=white align=center>:huh?</td>" & _
			"<td bgcolor=white><img src=smilies/on_2.gif align=absmiddle></td><td bgcolor=white align=center>:s</td>" & _
			"<td bgcolor=white><img src=smilies/on_3.gif align=absmiddle></td><td bgcolor=white align=center>:P</td>" & _
			"<td bgcolor=white><img src=smilies/on_4.gif align=absmiddle></td><td bgcolor=white align=center>}:)</td>" & _
			"<tr><td bgcolor=white><img src=smilies/on_5.gif align=absmiddle></td><td bgcolor=white align=center>:D</td>" & _
			"<td bgcolor=white><img src=smilies/on_6.gif align=absmiddle></td><td bgcolor=white align=center>}:|</td>" & _
			"<td bgcolor=white><img src=smilies/on_7.gif align=absmiddle></td><td bgcolor=white align=center>:)</td>" & _
			"<td bgcolor=white><img src=smilies/on_8.gif align=absmiddle></td><td bgcolor=white align=center>:oops</td>" & _
			"<tr><td bgcolor=white><img src=smilies/on_9.gif align=absmiddle></td><td bgcolor=white align=center>;)</td>" & _
			"<td bgcolor=white><img src=smilies/on_10.gif align=absmiddle></td><td bgcolor=white align=center>:pff</td>" & _
			"<td bgcolor=white><img src=smilies/on_11.gif align=absmiddle></td><td bgcolor=white align=center>:/</td>" & _
			"<td bgcolor=white><img src=smilies/on_12.gif align=absmiddle></td><td bgcolor=white align=center>:0</td></tr></table></td></tr></table><hr size=1>"

		saryUserMsgTo = Array(strUsername)
		intType = 1
	ElseIf Mid(strMessage, 1, 7) = "/color " Then
		strCommand = Replace(strMessage, "/color ", "")

		Response.Redirect("message.asp?Color=" & strCommand & "&Format=" & strFormat)
	Else
		'Dim saryActiveUsers
		'Dim blnPrivateMessage

		blnPrivateMessage = False

		'Get the array
		If IsArray(Application(ApplicationUsers)) Then
			saryActiveUsers = Application(ApplicationUsers)

			For intArrayPass = 1 TO UBound(saryActiveUsers, 2)
				If Instr(strMessage, "/" & saryActiveUsers(1, intArrayPass) & " ") <> 0 Then

					strMessage = Replace(strMessage, "/" & saryActiveUsers(1, intArrayPass) & " ", "")

					saryUserMsgTo = Array(saryActiveUsers(1, intArrayPass), strUsername)
					intType = 3

					blnPrivateMessage = True
					Exit For
				End If
			Next
		End If

		If blnPrivateMessage = False Then
			strMessage = "Hai digitato un comando sconosciuto. Digita /commands per aiuto."
			saryUserMsgTo = Array(strUsername)
			intType = 1
		End If
	End If
Else
	If strColor <> "" Then strMessage = "[color=" & strColor & "]" & strMessage & "[/color]"
	If strFormat <> "" Then strMessage = "[" & strFormat & "]" & strMessage & "[/" & strFormat & "]"
End If

strMessage = Trim(strMessage)


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

	Response.Write(vbCrLf & intTempSize)

	ReDim Preserve saryMessages(5, intTempSize)

	saryMessages(0, intTempSize) = strUsername
	saryMessages(1, intTempSize) = strMessage
	saryMessages(2, intTempSize) = CDbl(Now())
	saryMessages(3, intTempSize) = intType
	saryMessages(4, intTempSize) = saryUserMsgTo
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







 ' scrivo nella registrazione
' se esiste vuol dire che sto registrando, non uso session("registra") = true (visto solo da admin) ma dovrei usare var di applicazione, invece testo l'esistenza del file che punta al logchat
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FileExists (origine) then
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = fso.OpenTextFile(origine, ForReading)
	url1 = objTextFile.ReadLine
	Set objCreatedFile = objFSO.OpenTextFile(url1, ForAppending,True)
	objCreatedFile.WriteLine("<b>"&Session("Username") &" : </b> "& CheckForLink(strMessage) & "<br>")		 
	objCreatedFile.Close		


end if

Application.UnLock
Response.Redirect("message.asp?Color=" & strColor & "&Format=" & strFormat)

Function KickUser(strUsername)
	KickUser = False

	Dim intArrayPass
	Dim saryActiveUsers

	Application.Lock

	'Get the array
	If IsArray(Application(ApplicationUsers)) Then
		saryActiveUsers = Application(ApplicationUsers)
	Else
		ReDim saryActiveUsers(6, 0)
	End If

	For intArrayPass = 1 TO UBound(saryActiveUsers, 2)
		If saryActiveUsers(1, intArrayPass) = strUsername AND saryActiveUsers(3, intArrayPass) <> getIP() Then
			saryActiveUsers(6, intArrayPass) = "kick"

			KickUser = True

			Application(ApplicationUsers) = saryActiveUsers

			Exit For
		End If
	Next

	Application.UnLock
End Function
%>