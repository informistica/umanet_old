<%@ Language=VBScript %>
<!doctype html>

<% dim maxID
 scegli=request.QueryString("scegli") ' 0 = forum 1=bacheca 2=diario
 Reply=request("Reply") ' è settato se sto rispondendo serve per redirect ad unzip
 Azione=Request("AZIONE1")

' response.write("Azione="&Azione)
select case scegli
 case "0"
     session("social")="forum"
	 fileZip="file_forum"

 case "1"

    session("social")="bacheca"
	fileZip="file_lavagna"
  case "2"
    session("social")="diario"
	fileZip="file_diario"
  case "2"
    session("social")="diario"
  fileZip="file_diario"
  case "3"
    session("social")="interrogazioni"
  fileZip="file_interrogazioni"

 end select %>

<%
 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 	  dim objFSO,objCreatedFile
	  Const ForReading = 1, ForWriting = 2, ForAppending = 8
	 Dim sRead, sReadLine, sReadAll, objTextFile
	 Set objFSO = CreateObject("Scripting.FileSystemObject")

%>

<!--#include file = "../service/controllo_sessione.asp"-->
  <!-- #include file = "../var_globali.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    <!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->
   <!--#include file = "include/format_message.asp"-->
<!-- #include file = "../cAdmin/include_mail.asp" -->


<%
divid=Session("divid")
id_classe= Session("Id_Classe")
codBacheca=request.form("codBacheca")
if codBacheca="" then
codBacheca=Session("bacheca")
else
Session("bacheca")=codBacheca
end if
cognome=request.QueryString("cognome")
nome=request.QueryString("nome")
cartella=request.form("cartella")
if cartella="" then
cartella=Session("Cartella")
else
Session("Cartella")=cartella
end if
id_categoria=session("id_categoria")
categoria=session("categoria")
if id_categoria="" and request.QueryString("id_categoria")="" then
  id_categoria=0
else
    if request.QueryString("id_categoria")<>"" then
	id_categoria=request.QueryString("id_categoria")
	categoria=request.QueryString("categoria")
	end if

end if


'byChiamante=request.QueryString("byChiamante")
 'if byChiamante<>"" then

    cbEmail0=request.Form("cbEmail0")
	cbEmail1=request.Form("cbEmail1")
    cbEmail2=request.Form("cbEmail2")
    cbEmailProf=request.Form("cbEmailProf")
	cbNascosto=request.Form("cbNascosto")
	cbAnonimo=request.Form("cbAnonimo")
	 
 

	cbCompito=request.Form("cbCompito")
	cbImg=request.Form("cbImg")
	cbFile=request.Form("cbFile")
	date3=request.Form("date3")
	Scadenza=request.Form("Date3")
	Rispondi=request.Form("Rispondi")
	cbZip=request.form("cbZip") ' ha valore se devo creare la cartella delle consegne html in zip

	'ATTENZIONE SE VENGO DA REPLY NON é SETTATO cbZip
	if (cbZip<>"") or (Session("Zip")=1) then
	Session("zipFile")=1


	else
	Session("zipFile")=0

	end if

if (Request("cbNascosto")<>"") or strcomp(Request("cbNascosto"),"on")=0 then
	    Session("visibile")=0
		 
	 else
	   Session("visibile")=1
	 end if
	 if (Request("cbAnonimo")<>"") or strcomp(Request("cbAnonimo"),"on")=0 then
	    Session("anonimo")=1
		 
	 else
	   Session("anonimo")=0
	 end if

Set objFSO = CreateObject("Scripting.FileSystemObject")
			    
				'url="C:\Inetpub\umanetroot\expo2015Server\logNascosto.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(request.form("cbNascosto") & "<br>"&Session("visibile") & "<br>"& sSQL)
				'objCreatedFile.Close

' else
   '  cbEmail0=request.QueryString("cbEmail0")
   '  cbEmail=request.QueryString("cbEmail")
   '  cbEmailProf=request.QueryString("cbEmailProf")
 ' end if
 RCount=request.QueryString("RCount")
 if RCount="" then RCount="0"
RCount=cint(RCount)
'response.write("<br>byChiamante="&byChiamante)
'response.write("<br>cbEmail0="&cbEmail0)
'response.write("<br>cbEmail1="&cbEmail1)
'response.write("<br>cbEmail2="&cbEmail2)
'response.write("<br>cbEmailProf="&cbEmailProf)
'
' lo tolgo
 if request("INCORPORA")<>"" then
	      messaggio="<iframe src="& Request("INCORPORA") &" width= 98%  height= 480 ></iframe><BR>" & Request("Message") & "<br>"
		  messaggio=messaggio & " <a target=blank href=" & Request("INCORPORA") & "> Apri a TUTTO schermo</a>"
	      sComments = ReplaceComments(messaggio)
	   else
	      sComments=Request("Message")
	      sComments = ReplaceComments(sComments)
	   end if

    if strcomp(sComments&"","")=0 then
	     sComments="..."
	  end if


if (session("CodiceAllievo")="") or (session("Id_Classe")="") then response.Redirect("../../home.asp")

Function isBlank(Value)

if isNull(Value) then
	bAns = true
else
	bAns = trim(Value) = ""
end if
isBlank = bAns

end function

Function FixNull(Value)
if isNull(Value) then
	sAns = ""
else
	sAns = trim(Value)
end if
FixNull = sAns
end function

Function prepStringForSQL(sValue)

Dim sAns

'if inStr(sValue,"www.youtube.com")<>0 then

'sostituisce ' con quello storto
 sAns=Replace(sValue, Chr(39), Chr(96))
sAns=Replace(sAns, Chr(34), "&nbsp;")

   sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")

'else
	'sAns = Replace(sValue, Chr(39), Chr(96))
'end if
sAns = "'" & sAns & "'"
prepStringForSQL = sAns
End Function

function ReplaceComments(sInput)
dim sAns
'sAns = replace(sInput, "  ", "&nbsp; ")
'if inStr(sValue,"www.youtube.com")=0 then
   sAns = replace(sInput, chr(34), "")
'end if
sAns = replace(sAns, "<!--", "&lt;!--")
sAns = replace(sAns, "-->", "--&gt;")

ReplaceComments = sAns
end function

function YOUTUBEFormat(strMessage)

	if inStr(strMessage,"www.youtube.com")<>0 then
	 ' strMessage1= "<a href='" & strMessage & "' target=blank >"
	     strMessage1=strMessage
	     strMessage = Replace(strMessage,strMessage,"<a target=_blank href=' " & strMessage1 & "'><img title='vai youtube' src=img/play.jpg width=50 height=36 /></a>")

	else

	end if
	YOUTUBEFormat=strMessage

end function


function HTMLFormat(sInput)
dim sAns
sAns = replace(sInput, "  ", "&nbsp; ")
'sAns = replace(sAns, chr(34), "&quot;")
sIllStart = "<" & chr(37)
sIllEnd = chr(37)  & ">"
if instr(sAns, sIllStart) > 0 or instr(sAns, sIllEnd) > 0 then
  sAns = replace(sAns, "<" & chr(37), "")
  sAns = replace(sAns, chr(37)  & ">", "")
  bIllegal = true
end if
sAns = replace(sAns, ">", "&gt;")
sAns = replace(sAns, "<", "&lt;")
sAns = replace(sAns, vbcrlf, "<BR>")
HTMLFormat = sAns
end function

'------------
if Request("SubmitMessage") <> "" then bNew = true
if request("SubmitReply") <> "" or request("Reply") <> "" then bReply = true
if request("ApplyMessage") <> "" then bApply = true
bValid = bNew or bReply or bApply



' response.write("<br>bNew="&bNew)
' response.write("<br>bAddNew="&bAddNew)
' response.write("<br>bReply="&bReply)
' response.write("<br>bApply="&bApply)
' response.write("<br>bValid="&bValid)




if bApply then

sName = request("AuthorName")
sEmail = request("AuthorEmail")


'bAddNew =  request("MessageType") = "New"
bAddNew =  request("MessageType") = "New"

if Session("Caricata")=true then
		sUrlimg = "'" & Session("NomeImgForum")  & "',"
	  else
	    sUrlimg= "'" & ""  & "',"
end if
 if Session("CaricatoFile")=true then ' viene settato da db-file-to-disk è serve per vedere
		sUrlfile = "'" & Session("NomeFileForum")  & "',"
		Session("NomeFileForum")=""
	  else
	    sUrlfile= "'" & ""  & "',"

end if


	if bAddNew then

		' policy condivisione
		 if codBacheca<>"" then
			' Request.form("txtNUMREC") &"--" &  Request.querystring("numStud")
			  numCond=cint(Request("txtNUMREC"))
			  numStud=cint(Request("numStud"))
			  if numCond<>numStud then ' se non condivido con tutta la classe vuol dire che è privato=1
				 privato=1
			  else
				 privato=0
			  end if
			else
				 privato=0
			end if

	  sTopic = prepStringForSQL(Request("Topic")) & ","

	   ' sAbstract = prepStringForSQL(ReplaceComments(Request("Abstract")))
	   sAbstract = Request("Breve")
	  'sAbstract ="stronzo"
	   

	  
	  sName = prepStringForSQL(sName) & ","
	  'sEmail= prepStringForSQL(sEmail) & ","
	  scodAllievo=prepStringForSQL(Session("CodiceAllievo")) & ","
	  sIdClasse=prepStringForSQL(Session("Id_Classe")) & ","
	  'sBacheca=prepStringForSQL(Session("CodAdmin"))  & ","
	  sBacheca=prepStringForSQL(Request.form("CodBacheca"))  & ","
	  if Request.form("CodBacheca")="" then
	      sBacheca=prepStringForSQL(Session("CodAdmin"))  & ","
	  end if
	  sComments = ReplaceComments(sComments)

	 ' sComments=YOUTUBEFormat(sComments) Devo fare una inputbox dedidcata per i link
	  ' se ho caricato l'immagine aggiungo url altrimento vuoto
	 ' se devo gestire cartelle compresso
	 if Request("cbZip")<>"" then
	    Zip=1
		' per il problema di inserire url nel messaggio cosi da modificarlo da forum, sospeso!
		'Session("UrlZip")=Session("NomeFileForum")
		'sComments=sComments&"<br><br><a target=blank href="&Session("UrlZip")&">Index.html</a>"

	 else
	   Zip=0
	 end if
	  sComments = prepStringForSQL(sComments)
	  
	if instr(sComments,"<script>")<>0 then
		sComments=Replace(sComments,"<script>","")
		sComments=Replace(sComments,"</script>","")
		sComments=Replace(sComments,"onerror","")
	end if

	  sSQL = "INSERT INTO FORUM_MESSAGES (AUTHORNAME,CODICEALLIEVO,ID_CLASSE,TOPIC,URLIMG,URLFILE,BACHECA,COMMENTS,Privato,Zip,Id_Social,DATEPOSTED,Id_Categoria,Abstract,Punti, Visibile,Anonimo) VALUES (" & sName & scodAllievo & sIdClasse & sTopic & sUrlimg & sUrlfile & sBacheca & sComments & ","&privato&","&Zip&","&scegli&",'"&now&"',"&id_categoria&",'"& sAbstract&"',0,"&session("visibile")&","&session("anonimo")&");"

	  response.write("sAbstract="&sAbstract)
	' Questa riga sostiyuirà la precedente, perchè l'inserimento del post vuoto viene fatto da upload per l'immagine
	'  sSQL = "UPDATE FORUM_MESSAGES SET AUTHORNAME = " &sName & ",CODICEALLIEVO =" &scodAllievo &", ID_CLASSE = " & sIdClasse & ", TOPIC = " & sTopic & " COMMENTS=" & sComments & " WHERE ID = " & ID

 RCount=RCount+1


	  sSQL1=sSQL





				'Set objFSO = CreateObject("Scripting.FileSystemObject")
			    
				'	url="C:\Inetpub\umanetroot\expo2015Server\logNascosto.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(request.form("cbNascosto") & "<br>"&Session("Visibile") & "<br>"& sSQL)
				'objCreatedFile.Close

	  response.write(sSQL &"<br>")

	  conn.execute sSQL
	  session("visibile")=1
	 
	  sSQL="select max(ID) from FORUM_MESSAGES"
		 set rs=conn.Execute(sSQL)







		 maxID=rs(0)
		  ' per tornare alla discussione a cui ho risposto
		 ID=maxID


	   	rs.close
	 ' sSQL = "UPDATE FORUM_MESSAGES SET THREADPARENT = "& ID &" WHERE THREADPARENT = 0"
	 sSQL = "UPDATE FORUM_MESSAGES SET THREADPARENT = "& ID &",PARENTMESSAGE=0,LASTTHREADPOST='"&now&"' WHERE ID = "&ID


	  conn.execute sSQL
	  response.write(sSQL &"<br>")

	  ' se il messaggio non è privato devo inserire nella tabella Condividi il MAXID e CodiceAllievo
	  ' devo controllare se numStud==tutti quindi messasggio pubblico altrimenti devo settare i privilegi per i selezionati.
	  '
     ' soloMio=cint(Request("soloMio"))
	' response.write("numCond="&numCond)
	     if privato=1 then
            for k=1 to numCond
			     codstud=Session(codBacheca&"-"&k)
				 studente=Session(Studente&"-"&k)
		     	 sSQL = "INSERT INTO CONDIVIDI (Id_Post,CodiceAllievo) VALUES (" & maxID & ",'"&codstud &"');"
				 conn.execute sSQL
		    next
		 end if

	 ' se inserisco una nuova discussione verifico se devo creare cartella per consegne in zip
	 if Zip=1 and Session("Admin")=true then


			Set fso = CreateObject("Scripting.FileSystemObject")
			url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella") &"/"&fileZip&"/"&maxID
			url=Replace(url,"\","/")
			if fso.FolderExists (url) then
				 response.Write( "La cartella " & url & " esiste già.<br>")
			else
				fso.CreateFolder (url)
			end if
			' adesso creo le sottocartelle per gli studenti

			 sSQL="select CodiceAllievo from Allievi where Classe='"&Session("Cartella")&"' or Classe='Admin';"
			 set rs=conn.Execute(sSQL)
			 do while not rs.eof
				 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella") &"/"&fileZip&"/"&maxID&"/"&rs("CodiceAllievo")
				 url=Replace(url,"\","/")
				if fso.FolderExists (url) then
				 response.Write( "La cartella " & url & " esiste già.<br>")
				else
					fso.CreateFolder (url)
				end if


				 rs.movenext
			 loop


		 end if



	 else ' addNew quindi Reply
		sOrigAuthor = Request("OrigAuthor")
		if sOrigAuthor = "" then
		   sOrigAuthor = Request.QueryString("OrigAuthor")
	    end if
		iThread = Request("ThreadID")
		iParent = Request("ParentID")
		CodiceAllievoOrig=Request("CodiceAllievo")

		sName = prepStringForSQL(sName) & ","
		sEmail= prepStringForSQL(sEmail) & ","
		scodAllievo=prepStringForSQL(Session("CodiceAllievo")) & ","
		sIdClasse=prepStringForSQL(Session("Id_Classe")) & ","

		sTopic = prepStringForSQL(Request("Topic")) & ","
		if Request("Breve")<>"" then
		 sAbstract = prepStringForSQL(ReplaceComments(Request("Breve")))
		end if
	  ' sAbstract = Request("Breve")
	  'sAbstract ="stronzo"

		'if Session("Zip")=1 then
'		urlF=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella") &"/"&file_zip&"/"&maxID&"/"&Session("CodiceAllievo")
'           sTopic = prepStringForSQL(Request("Topic")&"<br>"&"<a href="&urlF&"/index.html target=blank>Apri</a>") & ","
'		else
'		   sTopic = prepStringForSQL(Request("Topic")) & ","
'		end if

		sBacheca=prepStringForSQL(Request.form("CodBacheca"))  & ","
		 if Request.form("CodBacheca")="" then
	      sBacheca=prepStringForSQL(Session("CodAdmin"))  & ","
	    end if

		sComments = prepStringForSQL(sComments)
		if iThread = 0 then iThread = iParent




		sSQL = "INSERT INTO FORUM_MESSAGES (PARENTMESSAGE,THREADPARENT,AUTHORNAME,CODICEALLIEVO,ID_CLASSE,TOPIC,URLIMG,URLFILE,BACHECA,COMMENTS,Id_Social,DatePosted,Id_Categoria,Abstract,Punti,Visibile, Anonimo) VALUES (" & iParent & "," &  iThread & "," & sName & scodAllievo & sIdClasse &  sTopic & sUrlimg & sUrlfile & sBacheca & sComments & ","&scegli&",'"&now&"',"&id_categoria&",'"& sAbstract&"',0,"&Session("visibile")&","&Session("anonimo")&");"



		response.write(sSQL &"<br>")

		RCount=RCount+1

				'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logPREW2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(sSQL &"<br>"&now)
'				objCreatedFile.Close

		conn.execute sSQL




		Session("visibile")=1
		Session("CaricatoFile")=""
		Session("CaricatoFileForum")=""
		Session("Caricata")=""
		Session("NomeImgForum")=""
'	Session("zipFile")=""
'        Session("IDTHREAD")=""


		cmd.CommandText = "LASTMESSAGE"
		cmd.CommandType = 4
		set rs = cmd.Execute
		sID = rs("ID")
		'rs.close



		 sSQL="select max(ID) from FORUM_MESSAGES"
		 set rs=conn.Execute(sSQL)
		 maxID=rs(0)
		' per tornare alla discussione a cui ho risposto
		 ID=iThread
	   	 session("iThread")=iThread
		rs.close



		sSQL = "UPDATE FORUM_MESSAGES SET REPLYCOUNT = REPLYCOUNT + 1, LASTTHREADPOST ='"& NOW &"' WHERE ID = " & iThread
			'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logPREW3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(sSQL &"<br>"&now)
'				objCreatedFile.Close

		conn.execute sSQL

	'scegli=1&bacheca=&Rispondi=1&ID=3204&Zip=0&CodiceAllievo=Margherita&divid=&id_classe=14COM&RCount=29&categoria=Programma&id_categoria=101

	 'Testo="Ha risposto ad un tuo post !"
	 'Azione="<a  target=blank href=../cSocial/ShowMessage.asp?scegli="&scegli&"&ID="&maxID&">Ho risposto ad un tuo post !</a>"
	 parametriurlnotifica = "scegli="&scegli&"&bacheca="&codBacheca&"&Rispondi="&Rispondi&"&ID="&maxID&"&Zip="&session("Zip")&"&CodiceAllievo="&session("CodiceAllievo")&"&id_classe="&id_classe&"&RCount="&RCount&"&categoria="&categoria&"&id_categoria="&id_categoria
	 Azione="<a  target=blank href=../cSocial/ShowMessage.asp?"&parametriurlnotifica&">Ho risposto ad un tuo post !</a>"
	 Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."
	 QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievoOrig & "','" & Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"

	 'response.write(parametriurlnotifica)


	 response.write("<br>504:"&sSQL)

	if (strcomp(Session("CodiceAllievo"),CodiceAllievoOrig)<>0) then
		 ConnessioneDB.Execute(QuerySQL)
     	end if

	' lo metto prima end if perchè se è un nuovo messaggio non faccio nulla

	end if 'bAddNew
	Session("Caricata")=false

	if (strcomp(cbCompito,"on") = 0)  then
' inserisco il compito come frase
response.write("517")
	QuerySQL="Select count(*) from preFrasi where Id_Paragrafo='"&Session("ID_ParSel")&"';"
	set rsTabella=ConnessioneDB.Execute (QuerySQL)
	if rsTabella(0)>0 then
		QuerySQL="Select max(Posizione) from preFrasi where Id_Paragrafo='"&Session("ID_ParSel")&"';"
		set rsTabella=ConnessioneDB.Execute (QuerySQL)
		contPos=rsTabella(0)
	else
		 contPos=0
	end if
	  if cbImg<>"" then
	     img=1
	  else
	     img=0
	  end if
	  if cbFile<>"" then
	     cFile=1
	  else
	  	 cFile=0
	  end if


	  Frase=Request("Topic") ' tolgo la parentesi ed il numero 12)
      preFrase=right(Frase,len(Frase)-instr(Frase,")"))
	 ' per inserire il compito nel libro
	 if Scadenza <>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
				   QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files)  SELECT '" & Session("IdMod") & "','" &  Session("ID_ParSel") & "', '" &preFrase & "'," & 0 & "," & contPos+1 & ",'" & Scadenza & "'," & img & "," & cFile & ";"
				   else
				   QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Img,Files)  SELECT '" & Session("IdMod") & "','" & Session("ID_ParSel") & "', '" & preFrase & "'," & 0 & "," & contPos+1 & "," & img & "," & cFile & ";"
				   end if
	 response.write(QuerySQL)
	  ConnessioneDB.execute(QuerySQL)
	   Session("ID_ParSel")=""
	   Session("IdMod")=""


end if


response.write("<br>556sono qui")



mes = ""
IsSuccess = false
sFrom = "Umanet Expo <noreply@iisvittuone.it>"
sMailServer = "mail.iisvittuone.it"
'sMailServer="95.141.32.176"
sBody = Trim(Request.Form("txtBody"))
sSubject = Request("Topic")





Sub TestEMail2()



  Set objMail = Server.CreateObject("CDO.Message")
  Set objConf = Server.CreateObject("CDO.Configuration")
  Set objFields = objConf.Fields
SMTP_SERVER_PICKUP_DIRECTORY="C:\inetpub\mailroot\Pickup"

sch = "http://schemas.microsoft.com/cdo/configuration/"
with objFields
 .item(sch & "sendusing") = 2 ' cdoSendUsingPort
 .item(sch & "smtpserver") = "mail.iisvittuone.it" 'application("smtpserver")
 .item(sch & "smtpserverport") = 587 'application("smtpserverport")
 .item(sch & "smtpauthenticate") = 1 'basic auth
 .item(sch & "sendusername") = "umanet" 'application("sendusername")
 .item(sch & "sendpassword") = "Inform1stic@" 'application("sendpassword")
 .update
end with


  With objMail
    Set .Configuration = objConf
	'.Charset = "UTF-8"
	.BodyPart.Charset = "utf-8"
    .From = sFrom
    .To = sTo
    .Subject = sSubject
    .HTMLBody = sBody
  End With

    Err.Clear
  on error resume next
 response.write("Provo ad inviare mail")
    objMail.Send

  if len(Err.Description) = 0 then
        mes = " MESSAGGIO INVIATO a " + sTo
     '   mes = mes + " TESTS COMPLETED SUCCESSFULLY!"
        IsSuccess = true
    else
    mes = " " + Err.Description + " INVIO NON RIUSCITO!"
  end if
  Set objFields = Nothing
  Set objConf = Nothing
  Set objMail = Nothing
End sub



if cbEmail2<>"" then
	 '  response.write("Email alla classe")
        QuerySQL="Select CodiceAllievo,Email,PasswordSHA256 from Allievi where Id_Classe='"&id_classe&"' and Email<>'' and Attivo=1;"
		' response.write("<br>"&QuerySQL)
	     set rsTabella=ConnessioneDB.Execute(QuerySQL)
    else
		if (strcomp(cbEmail0,"on") = 0)  then ' solo al mittente
		'response.write("<br>VERO0")
		'Da completare, devo per ogni codiceallievo della discussione leggere la sua email ed inviare
			'  sSQL="Select distinct(CodiceAllievo)  from FORUM_MESSAGES where ParentMessage='"&iParent&"';"

			sSQL="SELECT distinct(FORUM_MESSAGES.CodiceAllievo), Allievi.Email, FORUM_MESSAGES.ThreadParent, Allievi.Id_Classe,Allievi.PasswordSHA256 " &_
" FROM Allievi INNER JOIN FORUM_MESSAGES ON Allievi.CodiceAllievo = FORUM_MESSAGES.CodiceAllievo " &_
" WHERE FORUM_MESSAGES.ID="&Request("ParentId")&" and Allievi.Attivo=1;"

		   set rsTabella = conn.execute (sSQL)


		end if
		if (strcomp(cbEmail1,"on") = 0)  then ' solo ai commentatori
	'	response.write("<br>VERO1")
		'Da completare, devo per ogni codiceallievo della discussione leggere la sua email ed inviare

		sSQL="SELECT distinct(FORUM_MESSAGES.CodiceAllievo), Allievi.Email, FORUM_MESSAGES.ThreadParent, Allievi.Id_Classe,Allievi.PasswordSHA256 " &_
" FROM Allievi INNER JOIN FORUM_MESSAGES ON Allievi.CodiceAllievo = FORUM_MESSAGES.CodiceAllievo " &_
" WHERE FORUM_MESSAGES.ThreadParent="&iParent&" and Allievi.Attivo=1;"
 set rsTabella = conn.execute (sSQL)

        end if
		 response.write("<br>651:"&QuerySQL)

     end if

	 'response.write("<br>0:"&cbEmail0)
	 'response.write("<br>1:"&cbEmail1)
	 'response.write("<br>2:"&cbEmail2&"<br>")
	  response.write("<br>658:")
	  response.write(Session("Cartella") &"-<br>" )
	   response.write(InStr(Session("Cartella"),"$") &"-<br>" )
	   if (InStr(Session("Cartella"),"$"))<> 0 then
			messaggioda=left(Session("Cartella"),InStr(Session("Cartella"),"$")-1)
		else
		  messaggioda=Session("Cartella")
		end if
	 response.write("<br>660:")
 if (cbEmail0<>"") or (cbEmail1<>"") or (cbEmail2<>"") then
  do while not rsTabella.eof
    sBody = "Messaggio da: " & ucase(session("social")) &":"& messaggioda & "<br>Autore: "&request("AuthorName") &"<br><br>"& Request("Message")
	'sBody = Server.HTMLEncode(sBody)
    linkAvviso=dominio&homesito&"/script/cSocial/ShowMessage.asp?scegli="&scegli&"&ID="&maxID&"&RCount=0&TParent="&maxID&"&id_classe="&Id_Classe&"&Classe="&Session("Cartella")&"&Cartella="&Session("Cartella")&"&hash="&rsTabella("PasswordSHA256")&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&id_categoria="&Session("id_categoria")&"&categoria="&Session("categoria")
 sBody = sBody &"  <br> <a title 'Vai ad Umanet' href='"& linkAvviso&"'> Entra in Umanet Evolution 3.0</a> <img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' /> "

	   sTo=rsTabella("Email")
	   'sTo="mauro.spinarelli@gmail.com"
	   response.write("<br>Invi mail a " & sTo)
	   TestEMail2()

	rsTabella.movenext
   loop

end if

	if cbEmailProf<>"" then
	sBody="Messaggio da: " & ucase(session("social")) &":"& messaggioda & "<br>Autore: "&request("AuthorName") &"<br><br>" & Request("Message")

	    ' response.write("Email prof")
		  linkAvviso=dominio&homesito&"/script/cSocial/ShowMessage.asp?scegli="&scegli&"&ID="&maxID&"&RCount=0&TParent="&maxID&"&divid="&divid&"&id_classe="&Id_Classe&"&Classe="&Session("Cartella")&"&hash="&pwdAdmin&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&id_categoria="&Session("id_categoria")&"&categoria="&Session("categoria")&"&Cartella="&Session("Cartella")
		 linkAvviso2=dominio&homesito&"/script/cSocial/unsubscribe.asp?scegli="&scegli&"&ID="&maxID&"&RCount=0&TParent="&maxID&"&divid="&divid&"&id_classe="&Id_Classe&"&Classe="&Session("Cartella")&"&CodiceAllievo="&codAdmin&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&id_categoria="&Session("id_categoria")&"&categoria="&Session("categoria")&"&Cartella="&Session("Cartella")

		sBody = sBody  &"  <br> <a title='Vai ad Umanet' href='"& linkAvviso&"'>Entra in Umanet Evolution 3.0</a> <img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' /> "
		sBody = sBody  &"  <br> <a title='Disiscriviti' href='"& linkAvviso2&"'>Unsubscribe</a>"

		'sBody = Server.HTMLEncode(sBody)
	   sTo=eMailAdmin
	   'sTo="mauro.spinarelli@gmail.com"
	   response.write("<br>Invio mail a " & sTo)
	   TestEMail2()



	end if






%>
<!--#include file = "database_cleanup.inc"-->
<%

 response.write(sSQL)

' response.write("cbEmailProf="&cbEmailProf)

   ' response.redirect "default.asp?scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&divid="&Session("divid")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")
   'response.redirect "ShowMessage.asp?scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cartella="&Session("cartella")&"&ID="&maxID&"&RCount="&RCount

' provo questa per tornare alla discussione e non all'ultimo post

' adesso se devo decomprimere distinguo il redirect
 if Session("Zip")=1 and Reply<>"" then
'if Session("Zip")=1  then
'    '  se è zippato e sono in risposta allora chiamo index.asp passandogli i parametri, la pagina avrà jquery per cliccare chiamare subito .php , la quale riceve sia i parametri per url cartella da decompimere sia parametri per fare il redirect a showmessage.
    response.redirect "../Unzip/unzip.asp?scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cartella="&Session("cartella")&"&ID="&ID&"&RCount="&RCount&"&Destinazione="&Session("Nomefilezip")&"&homeserver="&homeserver&"&homesito="&homesito&"&Materia="&Session("ID_Materia")&"&IDPARENT="&ID&"&Social=file_"&Session("Social")&"&CodiceAllievo="&Session("CodiceAllievo")&"&categoria="&session("categoria")&"&id_categoria="&session("id_categoria")
'

 else
	response.redirect "ShowMessage.asp?cognome="&cognome&"&nome="&nome&"&Zip="&Session("zipFile")&"&scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cartella="&Session("cartella")&"&ID="&ID&"&RCount="&RCount&"&categoria="&session("categoria")&"&id_categoria="&session("id_categoria")
	end if


end if 'bApply



%>

<html>
<head>

   <title>Preview message</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	  <meta charset="UTF-8">

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>


	<!-- jQuery UI -->
	  <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>

	<!-- Theme framework -->
  <script src="../../js/eak_app_dem.min.js"></script>


	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />




   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />-->



<style>
.loader {
display: block;
position: fixed;
left: 0px;
top: 0px;
width: 100%;
height: 100%;
z-index: 9999;
background: #fafafa url(../image/page-loader.gif) no-repeat center center;
text-align: center;
color: #999;
}
</style>

</head>

<body class='theme-<%=session("stile")%>'>
					  <div class="loader"></div>


	<div id="navigation">

        <%



		%>


  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 %>


	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Prewiew</h1>

					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>
						<li>
							<a href="#more-login.html">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-files.html">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html"></a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>







				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  </h3>
			          </div>
				      <div class="box-content">



				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">


		    <div class="box-content">
                     <br>


                     <%

if not bValid then
  response.write "You cannot navigate to this page without entering a forum message.  Please "
  response.write "return to the <A HREF = 'default.asp'>forum index</A> and try again."
  response.end
end if
'Write to db and redirect home.
'response.write "<HR>"

if bReply then
ParentID = Request("ParentID")
ThreadID = request("ThreadID")
 '+*****provo
		 'ID=ThreadID
		 ID=ParentID
		'ID=maxID
sOrigAuthor = request("OrigAuthor")
CodiceAllievoOrig=request("CodiceAllievo")
else
  ParentID=maxID
end if
sTopic = request("Topic")


if sOrigAuthor = "" then sOrigAuthor = request.QueryString("OrigAuthor")
sOrigMessage = HTMLFormat(Request("Message"))
sOrigMessageGrezzo=Request("Message")

'sOrigMessage=SMILEFormat(sOrigMessage)
sOrigMessage=FormatMessage(sOrigMessage)

'sOrigMessage=CONNESSIONIFormat(sOrigMessage)


%>
<CENTER><FONT SIZE = +2 COLOR=RED><B>Invio in corso...<br></B></FONT></CENTER><P><br><i>Se hai richiesto l'invio delle email, il server potrebbe impiegare qualche minuto ad eseguire la richiesta. Sei pregato di attendere e non interrompere il caricamento o chiudere la pagina. Grazie.</i> <center>

</P>
<%
 	   if Session("Caricata")=true then


'imgPathDir=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_"&Session("social")&"/img"
		   url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/img_"&Session("social")&"/img"  ' vuole il percorso relativo della cartella
		   url=Replace(url,"\","/")
		   urlimg=url&"/"& session("NomeImgForum")
		  ' response.write(urlimg)
		   'Session("Caricata")=false
		   'response.write(urlimg)
		   %>
           <img  class="imground" src="<%=urlimg%>" title="<%=urlimg%>" border="1">
   		  <%
	   end if
%>

</center></P><br>
<FORM ACTION = "PreviewMessage.asp?cognome=<%=request.QueryString("cognome")%>&nome=<%=request.QueryString("nome")%>&Zip=<%=Session("Zip")%>&scegli=<%=scegli%>&cbEmail=<%=cbEmail%>&cbEmailProf=<%=cbEmailProf%>&TParent=<%=ParentID%>&RCount=<%=RCount%>&txtNUMREC=<%=Request.form("txtNUMREC")%>&cbNascosto=<%=cbNascosto%>&cbAnonimo=<%=cbAnonimo%>" METHOD = "POST">

<% if bReply Then %>
<INPUT TYPE="HIDDEN" NAME="ParentID" VALUE="<%= ParentID %>">
<INPUT TYPE="HIDDEN" NAME="ThreadID" VALUE="<%= ThreadID %>">
<INPUT TYPE="HIDDEN" NAME="OrigAuthor" VALUE="<%= sOrigAuthor %>">
<INPUT TYPE="HIDDEN" NAME="Rispondi" VALUE=1>
<INPUT TYPE="HIDDEN" NAME="Reply" VALUE=1>

<%else%>
<INPUT TYPE="HIDDEN" NAME="AZIONE1" VALUE="<%=Azione%>">



<% end if %>
<INPUT TYPE="HIDDEN" NAME="codBacheca" VALUE="<%= codBacheca %>">

<INPUT TYPE="HIDDEN" NAME="cartella" VALUE="<%= session("Cartella") %>">
<INPUT TYPE="HIDDEN" NAME="Topic" VALUE="<%=sTopic %>">
<INPUT TYPE="HIDDEN" NAME="Breve" VALUE="<%=sAbstract %>">
<INPUT TYPE="HIDDEN" NAME="Message" VALUE="<%= sComments %>">
<INPUT TYPE="HIDDEN" NAME="AuthorName" VALUE="<%= Request("Name") %>">
<INPUT TYPE="HIDDEN" NAME="AuthorEmail" VALUE="<%= Request("Email") %>">
<INPUT TYPE="HIDDEN"  NAME="CodiceAllievo" VALUE="<%= CodiceAllievoOrig %>">
<INPUT TYPE="HIDDEN"  NAME="numStud" VALUE="<%=Request.QueryString("numStud") %>">
<INPUT TYPE="HIDDEN"  NAME="txtNUMREC" VALUE="<%=Request.form("txtNUMREC") %>">
<INPUT TYPE="HIDDEN" NAME="cbEmail0" VALUE="<%= cbEmail0 %>">
<INPUT TYPE="HIDDEN" NAME="cbEmail1" VALUE="<%= cbEmail1 %>">
<INPUT TYPE="HIDDEN" NAME="cbEmail2" VALUE="<%= cbEmail2 %>">
<INPUT TYPE="HIDDEN" NAME="cbEmailProf" VALUE="<%= cbEmailProf %>">
<INPUT TYPE="HIDDEN" NAME="cbNascosto" VALUE="<%= cbNascosto %>">
<INPUT TYPE="HIDDEN" NAME="cbAnonimo" VALUE="<%= cbAnonimo %>">
<INPUT TYPE="HIDDEN" NAME="cbCompito" VALUE="<%= cbCompito %>">
<INPUT TYPE="HIDDEN" NAME="cbImg" VALUE="<%= cbImg %>">
<INPUT TYPE="HIDDEN" NAME="cbFile" VALUE="<%= cbFile %>">
<INPUT TYPE="HIDDEN" NAME="date3" VALUE="<%= date3 %>">
<INPUT TYPE="HIDDEN" NAME="cbZip" VALUE="<%= cbZip %>">



<%

'response.write("Session zip ="&Session("Zip"))
'response.write("<br>Request(numStud)="& Request("numStud"))
'response.write("<br>Request(txtNUMREC)="& Request("txtNUMREC"))
if (Request("numStud")<>"") and (Request("txtNUMREC")<>"") then
if cint(Request("numStud")) <> cint(Request("txtNUMREC")) then
	if cint(Request.form("txtNUMREC"))<>0 then
	%>

	<center><b>Condividi con :   </b><br>
	<%

	' se non è pubblico creo il file con l'elenco degli abilitati che avrà il nome del codiceallievo che poi leggerò e cancellero nell'altra parte della stessa pagina
	' serve perchè i parametri sono contenuti in request.form (del new_post) ma poichè adesso richiamo previewmessage esso si perde
	' quindi li conservo nel file di test
	       contCond=1
		   for i=1 to cint(Request.QueryString("numStud"))
			   if cint(Request.Form("cbCondividi"&i)<>0) then ' devo condividere con quello stud
			      Session(codBacheca&"-"&contCond)=Request.Form("txtStud"&i)
				  Session(Studente&"-"&contCond)=Request.Form("txtStud2"&i)
			  Response.write("<br>" & left (Request.Form("txtStud2"&i),instr(Request.Form("txtStud2"&i),".")))
			'	  Response.write("<br>" & Request.Form("txtStud2"&i))
				  contCond=contCond+1
			   end if
		   next
	else ' solo per il proprietario, ad esempio archivio lavori parziali
	%>

	<%

	end if
end if
end if
%>
<INPUT TYPE = "HIDDEN" NAME = "MessageType"
<%
if bReply then
	Response.Write "VALUE = 'REPLY'"
else
	Response.Write "VALUE = 'New'"
end if
%>
><center>

 <%

 'response.write("ParentID="&ParentID & " " &"ThreadID="&ThreadID)
' response.write("<br>bNew="&bNew)
' response.write("<br>bAddNew="&bAddNew)
' response.write("<br>bReply="&bReply)
' response.write("<br>bApply="&bApply)
' response.write("<br>bValid="&bValid)



 %>
 </center>
<CENTER><INPUT TYPE="Submit" style="display:none"  VALUE = "Invia" name="ApplyMessage" id="ApplyMessage">

</CENTER>
</FORM>




<%
if bIllegal then %>
<FONT COLOR = "RED" SIZE = = -1>Your message was altered to delete the ASP delimiters &lt;<%= chr(37) %> and <%= chr(37) %>&gt;
</FONT><P>
<% end if %>
</P></div>







               <!--<h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> -->
                      </div>
			        </div>
			      </div>
			    </div>











                      </div>
			        </div>
			      </div>
			    </div>
			</div>


		</div> <!--fine main-->
        </div>




	</body>

    <script type="text/javascript">


$(window).load(function () {

	  $('#ApplyMessage').click();


	    event.stopPropagation();

	});

</script>

 </html>
