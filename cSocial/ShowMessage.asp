<%@ Language=VBScript %>
<%
'Response.charset="utf-8"
'Response.AddHeader "Refresh", "600"
	by_email=request.QueryString("by_email")
 CodiceAllievo=request.QueryString("CodiceAllievo")
 PasswordSHA256=request.QueryString("hash")
 cartella=request.QueryString("cartella")
 if cartella="" then
   cartella=session("Cartella")
 end if
 %>




 <%if by_email<>"" then

	Session("DBCopiatestonline")="OK" ' sarà meglio togliere il controllo sessione sul test sul suo valore

	Session("Id_Classe")=request.querystring("id_classe")
	Session("Cartella")=request.QueryString("Classe")
	Session("cartella")=request.QueryString("Classe")
	'Response.Cookies("Dati")("id_classe_img")=Request.QueryString("Classe")
	Session("id_classe_img")=Request.QueryString("Classe")
	Session.Timeout=60

	'response.write("Id: "&Request.Cookies("Dati")("id_classe_img"))

 end if

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>

 	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     	<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->
        <!-- #include file = "../var_globali.inc" -->

<%  if by_email<>"" then%>
<%else%>
<!--#include file = "../service/controllo_sessione.asp"-->
<%end if%>
<!--#include file = "include/format_message.asp"-->

 <% if by_email<>"" then

			' if strcomp(PasswordSHA256,"")=0  Then
			'     QuerySQL="Select * from Allievi where CodiceAllievo='"&CodiceAllievo&"';"
			' else
			' ''	QuerySQL="Select * from Allievi where CodiceAllievo='"&CodiceAllievo&"';"
			' 	QuerySQL="Select * from Allievi where PasswordSHA256='"&PasswordSHA256&"';"
			' end if

	QuerySQL="Select * from Allievi where PasswordSHA256='"&PasswordSHA256&"';"   
 set rsTabella=ConnessioneDB.Execute(QuerySQL)


 		  Session("Loggato") = True
		  Session("Cognome") = rsTabella("Cognome")
		  Session("Nome") = rsTabella("Nome")
		  Session("CodiceAllievo") =CodiceAllievo
		  Session("Username")= CodiceAllievo ' per la chat dopo disastro
		  Session("DataTest") = DataTest
		  Session("stile")=rsTabella("Stile")

		  Session("cartella")=rsTabella("Classe")
		  Session("CodAdmin")=codAdmin
		 ' response.Write("Session=IdCla="&Session("Id_Classe"))
		'' if (rsTabella("PasswordSHA256")=pwdAdmin) and (rsTabella("CodiceAllievo")=codAdmin) then
		     if (rsTabella("PasswordSHA256")=pwdAdmin) and (rsTabella("CodiceAllievo")=codAdmin) then
		        Session("Admin")=True
		      else
		         Session("Admin")=False
		      end if
			  Session("CartellaAdmin")="Admin"
			  session("ID_Materia")=request.querystring("id_materia")
			   session("DB")=request.querystring("DB")
			  session("codAdmin")=codAdmin

		'		if strcomp(cod,codAdmin)<>0 then
'
'			   idcla=rsTabella("Id_Classe")
'			'idcla="4COM"
'			    QuerySQL="Select * from [dbo].[3PERIODI] where Id_Classe='"&idcla&"' and Iniziale=1;"
' 				set rsTabella1=ConnessioneDB.Execute(QuerySQL)
'
'			    response.write("data="&rsTabella1(0))
'            rsTabella.close
'			set rsTabella=nothing
'
'
'		 else

			   Session("DataCla")="10/09/2013"
		   Session("DataCla2")="10/09/2015"
		    Session("DataClaq")="10/09/2013"
		   Session("DataClaq2")="10/09/2015"

			'  end if
			if Session("Loggato") <> True Then
			'response.redirect "https://elexpo.net"
			end if

 end if%>


 <%




 ' on error resume next



 categoria=request.QueryString("categoria")
 Response.Cookies("Dati")("categoria")=categoria

 id_categoria=request.QueryString("id_categoria")
 Response.Cookies("Dati")("id_categoria")=id_categoria

 if categoria="" then
    categoria=session("categoria")

 end if
 if id_categoria="" then
    id_categoria=session("id_categoria")
 end if

  maxIDAvviso=request.QueryString("maxIDAvviso") ' per segnare l'avviso come letto
 byNotifiche=request.QueryString("byNotifiche")  ' se provengo dal centro notifiche-messaggi aggiorno il campo visto
 scegli=request.QueryString("scegli") '01 = forum 1=lavagna 2=diario

select case scegli
 case "0"
     session("social")="forum"
	 pt="PS"

 case "1"

    session("social")="bacheca"
	pt="PL"
  case "2"
    session("social")="diario"
     pt="P"
		 case "3"
	     session("social")="interrogazioni"
	      pt="PI"
 end select  %>









  <%if (session("CodiceAllievo")="") or (session("Id_Classe")="") then response.Redirect("../../home.asp")
%>
<%
'on error resume next
'sCaption = request.QueryString("Caption")
RCount= request.QueryString("RCount") ' numero di risposte della discussione serve per decrementare in update in delete
TParent=request.QueryString("TParent") ' IDdel post per aggiornare ReplyCount

ID= request.QueryString("ID")
if ID="" then
  ID= session("iThread")
else
  ID=cint(request.QueryString("ID"))
end if


q="select ReplyCount from FORUM_MESSAGES where ID="&ID&";"
'response.write(q)
set RsCount= ConnessioneDB.execute(q)
RCount=RsCount(0)

if TParent="" then
  TParent= ID
end if

'if session("Admin")=true and TParent=ID then
if session("Admin")=true  then
	q="SELECT [CodiceAllievo] FROM [FORUM_MESSAGES] where ThreadParent="&ID
	'response.write(q&"<br>")
	consegnato=""
	set rsTab=ConnessioneDB.execute(q)
	do while not rsTab.eof
	consegnato=consegnato&"'"&rsTab("CodiceAllievo")&"'"&","
		rsTab.movenext
	loop
	if consegnato<>"" then
	consegnato=left(consegnato,len(consegnato)-1)
	else 
	consegnato="CodiceAllievo"
	end if

	q="select Cognome,Nome,CodiceAllievo,Attivo from Allievi where Id_Classe='"&request.querystring("id_classe")&"' and CodiceAllievo not in ("&consegnato&") and Attivo=1;"
	'response.write(q)
	set rsTabellaNC= ConnessioneDB.execute(q)

	'response.write(" <hr> <b>Mancata consegna</b> ")
	nonconsegnato=""
	do while not rsTabellaNC.eof
	nonconsegnato=nonconsegnato&rsTabellaNC("Cognome")&" "&left(rsTabellaNC("Nome"),1)&".; "
	rsTabellaNC.movenext
	loop

end if
'nonconsegnato="TUTTI 207"
' aggiorno il numero di visualizzazioni
' QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

if session("Admin")=False then
	if session("Visto")<>ID then
	QuerySQL="UPDATE FORUM_MESSAGES SET Visualizzazioni = Visualizzazioni + 1 WHERE ID="&ID
		'response.write(QuerySQL)
	session("Visto")=ID
	ConnessioneDB.Execute(QuerySQL)
	end if
end if

by_search=request.QueryString("by_search") ' se sono stato chiamato da forum_searc
Zip=cint(request.QueryString("Zip"))
Session("Zip")=Zip
visibile=request.querystring("visibile")
'privato=request.querystring("privato")
IDPOST=cint(request.QueryString("ID")) ' IDdel post per verificare privacy

if (ID=TParent) then
   session("discussione")=ID

end if
 divid=request.querystring("divid")
  cartella=request.querystring("cartella")
  id_classe=request.querystring("id_classe")
 bacheca= request.QueryString("bacheca")
  if id_classe="" then
      divid=Session("divid")
       cartella=Session("cartella")
     id_classe=Session("Id_Classe")
  end if

function ReplaceCar(sInput)
dim sAns
  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,chr(138),"&egrave;")
  sAns=  Replace(sAns,chr(130),"&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
 sAns=  Replace(sAns,"&nbsp;"," ")
ReplaceCar = sAns
'ReplaceCar = sInput

end function




  '   QuerySQL="Select count(*) from Allievi where Id_Classe='" & Session("Id_Classe")&"';"
'	'response.write(QuerySQL)
'	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'	numStud=rsTabella(0)

    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	'response.write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	VotoPalese=rsTabella("VotoPalese")
	'VotoPalese è globale per tutti i post, viene sovrascritto con il valore Anonimo della discussione locale
	MaxStelline=rsTabella("MaxStelline")
	rsTabella.close
	if byNotifiche<>"" then
	  QuerySQL="UPDATE AVVISI SET Visto = 1 "&_
  " where  ID_Avviso="&maxIDAvviso&";"
 ' response.write(QuerySQL)
 ' disabilito archivizione automatica delle notifiche
 'set rsTabella=ConnessioneDB.Execute(QuerySQL)
 end if


function HTMLFormat(sInput)
	dim sAns
	'sAns = replace(sInput, "  ", "&nbsp; ")
	'sAns = replace(sAns, chr(34), "&quot;")lo commento perchè altrimenti non funzionano i video embeded nella tabella superiore della pagina
	sAns = replace(sAns, "<!--", "&lt;!--")
	sAns = replace(sAns, "-->", "--&gt;")

	HTMLFormat = sAns
end function

function formattaData(DataCla)
  giornoD=DatePart("d",DataCla)
 if len(giornoD)=1 then
    giornoD= "0" & giornoD
 end if
 meseD=DatePart("m",DataCla)
  if len(meseD)=1 then
    meseD= "0" & meseD
 end if
 annoD=DatePart("yyyy",DataCla)
 formattaData=giornoD&"/"&meseD&"/"&annoD
end function
function formattaOra(DataCla)
  oraD=DatePart("h",DataCla)
 if len(oraD)=1 then
    oraD= "0" & oraD
 end if
 minD=DatePart("n",DataCla)
  if len(minD)=1 then
    minD= "0" & minD
 end if
 'secD=DatePart("s",secD)
'   if len(secD)=1 then
'    secD= "0" & secD
 'end if

 'formattaOra=oraD&"."&minD&"."&secD
 formattaOra=oraD&"."&minD
end function

cont=0
Function MessageChildren(ID, IndentLevel, iCurrentMessage,i)
	dim oRs,oRs1, oCmd, sSQL, sAns
	'FIRST GET MESSAGE, TEXT, CLOSE
    cont=cont+1
	dim oParam
	set oCmd = Server.CreateObject("ADODB.Command")
	set oCmd.ActiveConnection = conn
	oCmd.CommandText = "FORUM_MESSAGE"
	oCmd.CommandType = 4
	set oParam = cmd.CreateParameter("MESSAGEID", 3, 1)
	oCmd.parameters.append oParam
	oParam.value = cint(ID)
	'set oParam1 = cmd.CreateParameter("Id_Social", 3, 1)
'	oCmd.parameters.append oParam1
'	oParam1.value = cint(scegli)


	set oRs = oCmd.execute
	set oParam = nothing

	'sSQL = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"';"
'cmd.CommandText = sSQL
 'cmd.CommandText = "MESSAGETHREADS"
'cmd.CommandType = 4
'rs.open cmd, , 1, 3




	iIndent = IndentLevel
	set oRs = oCmd.execute

	if oRs.eof then
		oRs.close
		set oRs = Nothing
		set oCmd = nothing
		MessageChildren = ""
		exit function
	end if
	' if (i mod 2) = 0  then
	'    classe_riga="zebra-dispari"
	'  else
	'    classe_riga=""
    ' end if
	 ' devo contare i voti per il messaggio corrente ID
	 idConta=oRs("ID")
	 IdStud=oRs("CodiceAllievo")
	 Zip=oRs("Zip")



	 QuerySQL="select count(*) from Voti WHERE ThreadQuote="& idConta &" and Voto=-1;"


		 set oRs1=conn.Execute(QuerySQL)
		 if oRs1.eof then
	       votineg=0
		   else
		   votineg=oRs1(0)
		 end if

		  QuerySQL="select count(*) from Voti WHERE ThreadQuote="& idConta &" and Voto=1;"
		 set oRs11=conn.Execute(QuerySQL)
		 if oRs1.eof then
	       votipos=0
		   else
		   votipos=oRs11(0)
		 end if


		'  QuerySQL="select * from Voti WHERE ThreadQuote="& idConta &";"
		'  QuerySQL=" SELECT count(*) AS numVotanti  from Voti WHERE ThreadQuote="& idConta &" GROUP BY CodiceAllievo ;"
		' set oRs1=conn.Execute(QuerySQL)
		 'numVotanti=oRs1(0)
		'response.Write(QuerySQL)

	  	   QuerySQL=" SELECT Sum(Voto) AS SommaDiVoto, CodiceAllievo,Cognome,Nome from Voti WHERE ThreadQuote="& idConta &" GROUP BY CodiceAllievo,Cognome,Nome ;"

		' dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logVota2.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
''
			 'response.write(QuerySQL &"<br>")
		 set oRs1=conn.Execute(QuerySQL)
		 if oRs1.eof then
	       voti=0
		   else
		   voti=oRs1(0)
		 end if


data=oRs("DATEPOSTED")
	if isnull(data) then data="12/12/2112"
	'sAns2 = sAns & "</TD><font color='black'><TD class='hidden-480'>" & formattaOra(data) & "</TD></font>"
	if Session("Admin") = true then


			'compongo select
			isel = -10
			seleziona = "<select id='sel"&oRs("ID")&"' onchange='cambiavoto("&oRs("ID")&")' style='width:60px'>"
			do while isel < 10

				if oRs("Punti")= (isel+1) then
					tipo = "selected"
				else
					tipo = ""
				end if

				seleziona = seleziona&"<option "&tipo&" value='"&(isel+1)&"'>"&(isel+1)&"</option>"
				isel = isel+1
			loop
			seleziona = seleziona & "</select>"
			sAns1 = "<span data-placement='bottom'   rel='tooltip' title='Punti assegnati dal docente'>"&pt & ". "&seleziona&"</span>    "



	else
			sAns1 = "<span data-placement='bottom'   rel='tooltip' title='Punti assegnati dal docente'>"&pt & "." & oRs("Punti")&"</span>    "



	end if
'	sAns1="<center>"
'for k=1 to voti ' aggiungo le stelline dei voti
	mipiace=0
	nonpiace=0
	numVotanti=0
	voto=0
	titolopos=""
	titoloneg=""
	titolo1="Piace a "
	titolo2="Non piace a "

	while not  oRs1.eof
	   if oRs1("SommaDiVoto")>0 then
			 if (VotoPalese=1) or (Session("Admin")=true) then
			     titolopos=titolopos&" "&titolo1&" "&oRs1("Cognome") &" " & left(oRs1("Nome"),1) &"." & "(Feedback(+) =" & votipos &")"
			 else
				 'sAns1=sAns1&"<img src='img/icon_star_red.gif' width='13' height='12' ><br>"
			 end if
	      mipiace=votipos
	   else
	         if (VotoPalese=1) or (Session("Admin")=true) then
			    titoloneg=titoloneg&" "&titolo2&" "&oRs1("Cognome") &" " & left(oRs1("Nome"),1) &"." & "(Feedback(-) =" & votineg &")"

	         else
			 			    ' sAns1=sAns1&"<img src='img/icon_star_black.gif' width='13' height='12' ><br>"
			 end if
	     nonpiace=nonpiace+oRs1("SommaDiVoto")
	   end if
'next
	if oRs1("SommaDiVoto")>0 then
		voto=voto + (5 + oRs1("SommaDiVoto"))
	else
	   voto=voto + (6 + oRs1("SommaDiVoto"))
	end if
	
	numVotanti=numVotanti+1


	oRs1.movenext
	wend

	set orS1=nothing


if ((strcomp(Session("CodiceAllievo"),oRs("CodiceAllievo"))=0) and strcomp(categoria,"Feedback")<>0) or (session("Admin")=true)  then

	  ' testo="Testo del post..."
			sAns1 ="   "& sAns1 & "&nbsp;<a href='#modal-1' onClick=""modifica('"&ID&"');"" data-toggle='modal'><i style='text-decoration:none' class='icon-pencil' title='Modifica post'></i></a>&nbsp;&nbsp;      "

	   end if


	if mipiace=0 and nonpiace=0 then
	  ' sAns = sAns & "<TD  class='hidden-480'>" & sAns1 &"</TD>"
	else
	    ' sAns = sAns & "<TD  class='hidden-480'>"
		 'if (nonpiace+mipiace)>=0 then
		   sAns1 =  sAns1 &" <img id=img"&id&" src='img/facebook2.jpg' style='border:1px solid #000;' width='21' height='19' title='"&titolopos&"'><span id=voto"&id&" title='Feedback (+) = "&votipos&"'> " & votipos &"</span>"
		 'else
		   sAns1 =  sAns1 &" <img id=img"&id&" src='img/facebook8_nonpiace_small.jpg' style='border:1px solid #000;' width='21' height='19' title='"&titoloneg&"'><span id=voto"&id&" title='Feedback (-) = "&votineg&"'> " &  votineg &"</span>"
		' end if
 	    ' sAns = sAns & "</TD>"
	end if
'
'sAns1=sAns1&" <img src='img/facebook2.jpg' width='21' height='19' align='bottom'>&nbsp;Ci provo&nbsp;<img src='img/icon_star_red.gif' width='13' height='12'>"



	 

		  QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			VotoAttivo=rsTabella("VotoAttivo")
		  
		 if VotoAttivo=1 then
		   sAns1 = sAns1 &"&nbsp;&nbsp; <A onClick=""vota_post('"&ID&"','"&TParent&"',1,"&scegli&");"" ><img style='border:1px dotted #000;' title='Mi piace' src=img/facebook2.jpg width=21 height=19 align=bottom title='"&titolopos&"'></a>"
       sAns1 = sAns1 &"&nbsp;&nbsp; <A onClick=""vota_post('"&ID&"','"&TParent&"',0,"&scegli&");"" ><img style='border:1px dotted #000;' title='Non mi piace' src=img/facebook8_nonpiace_small.jpg width=21 height=19 align=bottom title='"&titoloneg&"'></a>"

		 end if

  if ((strcomp(Session("CodiceAllievo"),oRs("CodiceAllievo"))=0) and strcomp(categoria,"Feedback")<>0) or (session("Admin")=true)  then
 
  sAns1 = sAns1 &"&nbsp;&nbsp; <A onClick=""elimina_post('"&ID&"','"&TParent&"');"" ><i title=Elimina class='icon-trash' ></i></a>"

	       else
	    ' sAns1 = sAns1 &"<A HREF='DeleteMessage.asp?zip="&Session("Zip")&"&scegli="&scegli&"&CodiceAllievo="&oRs("CodiceAllievo")&"&ID=" & oRs("ID") & "&RCount="& RCount  & "&TParent=" & TParent &"'>&nbsp;&nbsp; <i title=Elimina class='icon-trash' onClick=return window.confirm('Vuoi veramente cancellare questa discussione ?');></i></a>"
	   end if



		querySQL="Select Url_img, Cognome,Nome,CodiceAllievo from Allievi where CodiceAllievo='"&IdStud&"';"
		 set oRs2=ConnessioneDB.Execute(QuerySQL)
		 if not oRs2.eof then
		 Url_img=oRs2(0)
		 else
		  Url_img=""
		 end if
		 Cognome=oRs2("Cognome")
		 Nome=oRs2("Nome")
		 'CodiceAll=oRs2("CodiceAllievo")
		 CodiceAll=IdStud
		' response.write(querySQL & "-----" & oRs2(0) &"<br>")
		' if oRs1.eof then
'	       voti=0
'		   else
'		   voti=oRs1(0)
'		 end if


                  if (strcomp(Url_img&"","")=0) or (anonimo=1) then
					urlimmagine="../../img/no-avatar.jpg"

                  else
					   if IdStud= Session("CodAdmin") then
					      url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/Admin/Profili/thumb"
					   else
					       url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&cartella&"/Profili/thumb" ' vuole il percorso relativo della cartella
    				   end if
					  url=Replace(url,"\","/")
					  urlimmagine=url&"/"& Url_img

		           end if










if oRs("ParentMessage")<>0 then

 sAns="<div id=post_"&id&" class='media'><a class='pull-left' href='#'><img title='"&autore&"."&"' src='"&urlimmagine&"' class='img-rounded' style='width:40px; height:40px'></a><div class='media-body'>"


	'tentativo fallito  di separare ogni risposta di primo livello da 1 riga
	'if iIndent=0 then
	'   sAns= "<hr>"&sAns
	'end if

	'for i = 0 to iIndent - 1
	'**** qua aprirò i div per la nidificazione
		' sAns = sAns & "&nbsp;&nbsp;&nbsp;&nbsp;<img src='img/icon_pencil.gif' > "
	'Next
	if oRs("ID") <> cLng(iCurrentMessage) then
        
 sAns = sAns & " <font color='black'><A style='text-decoration:none;'  HREF='ShowMessage.asp?scegli="&scegli&"&bacheca="&bacheca&"&Rispondi=1&ID=" & oRs("ID") & "&Zip="&Session("Zip")& "&divid="&divid& "&id_classe="&id_classe& "&RCount="&RCount&"&categoria="&categoria&"&id_categoria="&id_categoria&"'> "

		'sAns = sAns & " <font color='black'><A style='text-decoration:none;'  HREF='ShowMessage.asp?scegli="&scegli&"&bacheca="&bacheca&"&Rispondi=1&ID=" & oRs("ID") & "&Zip="&Session("Zip")& "&CodiceAllievo="&oRs("CodiceAllievo")& "&divid="&divid& "&id_classe="&id_classe& "&RCount="&RCount&"&categoria="&categoria&"&id_categoria="&id_categoria&"'> "

	 ' divid=request("divi




	       url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/img_"&session("social")&"/thumb" ' vuole il percorso relativo della cartella
		   url=Replace(url,"\","/")
		   urlimg=url&"/"& oRs("Urlimg")
			 if Anonimo=0 then
      autore= Cognome&" "&left(Nome,1)
			else
			autore="anonimo"
			end if
	  '  response.write "CurMsg=" & iCurrentMessage & " ID = " & oRs("ID")
		' qua ho fatto delle modifiche
		'sAns = sAns & "<img src='img/icon_group.gif' > " & " "& ucase(oRs("Topic"))  &"<br>"& left(oRs("comments"),250)& "...</A>"
	'sAns = sAns & "<img src='img/icon_group.gif' > " & " "& ucase(oRs("Topic"))  &"<br>"& oRs("comments")& "...</A>"
		' vedo se devo aggiungere il thumb
		if strcomp(oRs("Urlimg")&"","")<>0 then


			sAns = sAns & "<img src='img/icon_group.gif' > " & "<img class=imground title='" & urlimg&  "'  src='" & urlimg&  "' ><span class='post-title' id=titolo"&id&"><b>  "& oRs("Topic")  &"</b></a></span><br><small><div class='post-meta'><span class='date'><i class='icon-user'></i>&nbsp;" &autore& ".&nbsp;<i class='icon-calendar'></i>&nbsp;" & oRs("DatePosted")    & "</span>&nbsp;&nbsp; "&sAns1&"</span><br></small><br><span id="&id&"> "& oRs("comments") &"</span>"

	     else
		    sAns = sAns & "<img src='img/icon_group.gif' > " &  "<span class='post-title' id=titolo"&id&"><b>  "& oRs("Topic")  &"</b></a></span><br><small><div class='post-meta'><span class='date'><i class='icon-user'></i>&nbsp;" &autore& ".&nbsp;<i class='icon-calendar'></i>&nbsp;" & oRs("DatePosted")  & "</span>&nbsp;&nbsp;"&sAns1&"</span><br></small><br><span id="&id&"> "& oRs("comments") &"</span>"


		 end if
		 if strcomp(oRs("Urlfile")&"","")<>0 then

					 key = cint(left(oRs("Urlfile"),len(oRs("Urlfile"))-4)) ' tolgo il suffisso
					QuerySQL = "select Nome from file_forum where ID_Smile="&key&";"
					set oRs4=conn.Execute(QuerySQL)
					NomeFile=oRs4(0)
					CartellaFile=left(NomeFile,len(NomeFile)-4)

		     if Session("Zip")=1 then
			   url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/file_"&session("social")&"/"&oRs("ParentMessage")&"/"&CodiceAll
			 else
		      url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/file_"&session("social") ' vuole il percorso relativo della cartella url=Replace(url,"\","/")
			  end if

	     urlfile=url&"/"& oRs("Urlfile")


		 end if
		  if Session("Zip")=1 then
		  if  ((strcomp(ucase(Session("CodiceAllievo")),ucase(oRs("CodiceAllievo")))=0) or session("Admin")=true) or (Session("Privato1")=0) then

			   url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/file_"&session("social")&"/"&oRs("ParentMessage")&"/"&replace(CodiceAll," ","%20")
			    sAns = sAns & " <a target='_blank' href="&url& "/"&replace(CartellaFile," ","%20") &"/index.html title='"&url & "/"&replace(CartellaFile," ","%20") &"/index.html'><i class='icon-search'>index.html</i></a>"
				 sAns = sAns & "<br> <a href="&urlfile &" title='"& oRs("Urlfile")&"'><i class='icon-cloud-download'>  "&oRs("Urlfile")&"</i></a><br><br>"

			   ' sAns = sAns & " <a target='_blank' href="&url& "/index.html title='"& oRs("Urlfile")&"'><i class='icon-search'>index.html</i></a>"
			end if
		end if

		 sAns = sAns & "</div> "


	else
		'sans = sans & "<B>" & ucase(oRS("Topic"))  &"<br>"& left(oRs("comments"),250) & "...</B>"
	sans = sans & " <h5 class='media-heading'>" &  oRS("Topic") &"</a><br><div class='post-meta'><span class='date'><i class='icon-calendar'></i>&nbsp;"  & oRs("DatePosted")  &  "</span>&nbsp;&nbsp; <img src='img/facebook2.jpg' width='21' height='19' align='bottom'>&nbsp;Ci provo&nbsp;<img src='img/icon_star_red.gif' width='13' height='12'>"&sAns1&"</span><br><span id=com"& cont&"> "&oRs("comments") & "</span></h5>"





	end if

end if ' if oRs("ParentMessage")<>0 then




	 sans = sans & "<div class='media-actions'></div>"


	oRs.close

	if (privato=0) or (session("admin")=true) then
	sSQL = "SELECT ID FROM FORUM_MESSAGES WHERE PARENTMESSAGE = " & ID & " order by DatePosted asc;"
	oCmd.CommandText = sSQL
	oCmd.CommandType = 1
	set oRs = oCmd.execute
		if ors.eof and iIndent = 0 then
			sAns = ""
		else
		    i=0
			do while not oRs.eof
				'if oRs("ParentMessage")<>0 then
				' sAns = sAns & MessageChildren(oRs("ID"), iIndent + 1, iCurrentMessage,i+1)&"</div></div>"
				'else
				   sAns = sAns &"&nbsp;"& MessageChildren(oRs("ID"), iIndent + 1, iCurrentMessage,i+1)&"</div></div>"
				'end if
					 oRs.MoveNext
					 i=i+1

			Loop
		end if
		oRs.Close
	else
	sAns= "Discussione privata, non puoi visualizzare le risposte finchè l'amministratore non la rende pubblica"
	end if


			set oRs = nothing
			set oCmd = nothing
	

		'MessageChildren = sAns
		'MessageChildren = SMILEFormat(sAns)
        MessageChildren = FormatMessage(ReplaceCar(sAns))
End Function




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

iMessageId = request.QueryString("ID")
bValid = isNumeric(iMessageId) and iMessageId <> ""

if bValid then
cmd.CommandText = "Forum_Message"
cmd.CommandType = 4
set Param = cmd.CreateParameter("MESSAGEID", 3, 1)
cmd.parameters.append Param
Param.value = iMessageId
set rs = cmd.Execute
'on error resume next
iThreadParent  = rs("ThreadParent")
iParMes  = rs("ParentMessage")
'sMsg = CONNESSIONIFormat(HTMLFormat(rs("Comments")))

'sMsg = replace(sMsg, "  ", "&nbsp; ")
sMsg = replace(sMsg, vbcrlf, "<BR>")

sMsg =  rs("Comments")
Urlimg2=rs("Urlimg")
Urlfile2=rs("Urlfile")
Privato=rs("Privato")
Privato1=rs("PrivatoLab")


QuerySQL = "select Anonimo,Privato from FORUM_MESSAGES where ID="&iThreadParent&";"
set oRsAnonimo=conn.Execute(QuerySQL)

If (rs("Anonimo")=1) or (oRsAnonimo(0)=1) then
Anonimo=1
VotoPalese=0
end if
If (rs("Privato")=1) or (oRsAnonimo(1)=1) then
Privato=1
VotoPalese=1
end if

if strcomp(request.QueryString("ID"), request.QueryString("TParent"))=0 then ' la radice setta lo stato per tutti i figli
Session("Privato1")=rs("PrivatoLab")
end if

if strcomp(request.QueryString("ID"), request.QueryString("TParent"))=0 then ' la radice setta lo stato per tutti i figli
Session("Privato1")=rs("PrivatoLab")
end if

if Privato1="" then
Privato1=0 ' di default sono visibili i link
	if TParent=ID then
	Session("Privato1")=0
	end if
end if
' per leggere il nome del file partendo dall'indice
if  Urlfile2<>"" then
	key = cint(left(Urlfile2,len(Urlfile2)-4)) ' tolgo il suffisso
	QuerySQL = "select Nome from file_forum where ID_Smile="&key&";"
	set oRs1=conn.Execute(QuerySQL)
	NomeFile=oRs1(0)
	if Session("Zip")=1 then
	  CartellaFile=left(NomeFile,len(NomeFile)-4)
	end if
end if

end if

%>


<!doctype html>

<html>
<head>
<!-- <meta charset="utf-8"> -->
<style>
a {
	text-decoration:none;
	color:#000;
}
a:hover {
	text-decoration:none;

}
</style>
   <meta charset="utf-8">
   <title>Showmessage <%

                          if strcomp(ucase(session("social")),"LAVAGNA")=0 then
                            response.write("BACHECA")
                          else
                          response.write(ucase(session("social")))
                          end if
   %></title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />

	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />


	<!-- Bootstrap -->

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <!-- dataTables -->
	<link rel="stylesheet" href="../../css/plugins/datatable/TableTools.css">
<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">

     <link rel="stylesheet" href="../../css/style-themes.css">
        <link rel="stylesheet" href="../../css/docs.css">

	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
     <!-- jQuery UI -->
    <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>


	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
 <!-- dataTables -->
	<script src="../../js/plugins/datatable/megaDatatable.min.js"></script>

<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>

	 	<link rel="stylesheet" href="../../css/plugins/xeditable/bootstrap-editable.css">

	<script src="../../js/plugins/xeditable/bootstrap-editable.min.js"></script>
	<!--<script src="../../js/plugins/xeditable/demo.js"></script>-->
	<script src="../../js/plugins/xeditable/address.js"></script>
<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>

<!-- Theme framework -->
    <script src="../../js/eak_app_dem.min.js"></script>

<!-- CKEditor -->
	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>
	
<!--
	<script src="https://cdn.ckeditor.com/ckeditor5/16.0.0/classic/ckeditor.js"></script>

-->


	<!--[if lte IE 9]>
		<script src="../../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../social/img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../social/img/apple-touch-icon-precomposed.png" />


    <script src="_assets/js/jquery.zclip.js"></script>
<script src="include/copiaincolla.js"></script>

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



    <script type="text/javascript">


	function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!");

location.href="studente_domande.asp"
//location.href=window.history.back();
 }

function showText2() {window.alert("Discussione privata, non puoi vedere le risposte degli altri e una volta risposto non potrai modificare !");
//location.href="default.asp?scegli=<%=scegli%>&bacheca=<%=bacheca%>&nome=<%=request.QueryString("nome")%>&cognome=<%=request.QueryString("cognome")%>&id_classe=<%=id_classe%>&cartella=<%=session("Cartella")%>"
 }


	function addsmile(codice) {
	 alert (codice);
		with (document.InputForm) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina

		   // Name.value= Name.value + codice;

		  messaggio.value= messaggio.value + codice;

	    }
}


	</script>


   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />-->





</head>


<%   ' controllo prima se è privato, quindi con chi devo condividere
		     QuerySQL="select Privato from FORUM_MESSAGES WHERE ID="& IDPOST &";"
			 set oRs1=conn.Execute(QuerySQL)
			 privacy=oRs1(0)
			 condivisoCon=""
			 if privacy=1 then
				 QuerySQL="select * from Condividi WHERE Id_Post="& IDPOST &" and CodiceAllievo='" & Session("CodiceAllievo") &"';"
				 set oRs1=conn.Execute(QuerySQL)
				 flag1=1
				 if oRs1.eof  then
					flag1=0
					'condivisoCon="Nessuno"
				 end if
				  'response.write(QuerySQL)
				 QuerySQL="select * from FORUM_MESSAGES WHERE ID="& IDPOST &" and CodiceAllievo='" & Session("CodiceAllievo") &"';"
				 set oRs1=conn.Execute(QuerySQL)
				 flag2=1
				 if oRs1.eof then
					flag2=0
				 end if
				 QuerySQL="select * from Condividi WHERE Id_Post="& IDPOST
				 set oRs3=conn.Execute(QuerySQL)
				 if oRs3.eof then
				   condivisoCon="nessuno"
				 else
					   i=0
					   while not oRs3.eof
						  condivisoCon=condivisoCon &  oRs3(1) &", "
						  oRs3.movenext
						  i=i+1
					   wend
				 end if


				 if flag1=0 and flag2=0 and Session("Admin")=false then
				 %>
					 <body onLoad="showText2();">
				<% end if
			else
		        condivisoCon="tutti"
			end if
%>



<%QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato=rsTabellaCI("CIAbilitato")
	rsTabellaCI.close

	'response.write("CIAAbilitato="&CIAbilitato)

%>

<body class='theme-<%=session("stile")%>' data-layout-sidebar='fixed' data-layout-topbar='fixed'>

	<div id="navigation">

        <%

  Cartella=Request.QueryString("Cartella")
  TitoloCapitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")





'	 response.write("<br>Session(codAdmin)"&Session("codAdmin"))
	 ' response.write("<br>Session(DB)"&Session("DB"))
		%>


		<!-- #include file = "../include/navigation.asp" -->



	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
                    <%if by_search<>"" then %>
                      <h3> <i class="icon-twitter"></i><%=Cartella%>&nbsp;
					   <% if strcomp(ucase(session("social")),"LAVAGNA")=0 then
                            response.write("BACHECA")
                          else
                          response.write(ucase(session("social")))
                          end if  %>
                      &nbsp;(<%=categoria%>) </h3>

                    <%else%>
						<h3> <i class="icon-twitter"></i>
						<% if strcomp(ucase(session("social")),"LAVAGNA")=0 then
                            response.write("BACHECA")
                          else
                          response.write(ucase(session("social")))
                          end if  %>
                        &nbsp;(<%=categoria%>) </h3>
					<%end if%>

                      <% if session("DB")=1 and strcomp(session("Username"),"ospite")<>0 then%>
                        <a title="Condividi link alla pagina" href="#" onClick="javascript:PopUpWindow(600,400,<%=scegli%>);return false;"><i class="glyphicon-share_alt"> </i> <small>Condividi</small> </a>
                      <% end if%>
                    </div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->

                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>


                         <%select case scegli
							 case "0"
								 session("social")="forum"
							 %>
                             <li>
							<a href="#">Umanet</a>
							<i class="icon-angle-right"></i>
						</li>
							<li>
							<a   href="../cSocial/default0.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Forum</a>
						    </li>



							 <%
							 case "1"
							 %>
                             <li>
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>
							 <li>
							<a  href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>">&nbsp;Bacheca</a>
						    </li>
							 <%
							  case "2"
							 %>
                             <li>
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>
							 <li>
							<a   href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Diario</a>
						    </li>

							 <%
							 case "3"
							%>
														<li>
						 <a href="#">Classe</a>
						 <i class="icon-angle-right"></i>
					 </li>
							<li>
						 <a   href="../cSocial/default0.asp?scegli=3&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Interrogazioni</a>
							 </li>

							<%
							 end select %>


                        <li> <i class="icon-angle-right"></i>

<%
 response.write" <TD><A HREF='ShowMessage.asp?categoria=" & categoria &"&id_categoria=" & id_categoria &"&nome=" & request.QueryString("nome") &"&cognome="&request.QueryString("cognome")&"&scegli="&scegli&"&bacheca="&bacheca&"&ID=" & rs("ID") & "&Zip=" & rs("Zip")&"&RCount=" & rs("ReplyCount")& "&TParent=" & rs("ID")& "&divid=" & divid & "&id_classe=" & id_classe & "&visibile=" & rs("Visibile") & "&privato=" & rs("Privato") & "'>"  & rs("Topic") & "</A></FONT></TD>"
	
%>
                       <%select case scegli
							 case "0"
								 session("social")="forum"
							 %>



							<a title="Torna alle Discussioni" href="default0.asp?scegli=0&id_classe=<%=id_classe%>&cartella=<%=cartella%>&bacheca=<%=bacheca%>&nome=<%=request("nome")%>&cognome=<%=request("nome")%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>
							 <%
							 case "1"
							 %>


							 <li>
							<a title="Torna alle Discussioni" href="default0.asp?scegli=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>
							 <%

							  case "2"

							 %>
							 <li>
							<a title="Torna alle Discussioni" href="default0.asp?scegli=2&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>

							 <%
							 case "3"

							%>
							<li>
						 <a title="Torna alle Discussioni" href="default0.asp?scegli=3&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
													 <i class="icon-angle-right"></i>
							 </li>

							<%
							 end select %>


						 <li>
							<a title="Torna alla Discussione" href="ShowMessage.asp?scegli=<%=scegli%>&id_classe=<%=Session("Id_Classe")%>&cartella=<%=session("cartella")%>&bacheca=<%=bacheca%>&ID=<%=Session("discussione")%>&Zip=<%=Session("Zip")%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>">Discussione</a>
                           <i class="icon-angle-right"></i>
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


                        <div class="bs-docs-example">

                       <div class="box box-bordered box-color">
							<div class="box-title">
                            <%if strcomp(bacheca,codAdmin)=0 then%>
								<h3><i class="icon-th-list"></i> <%=rs("Topic")%></h3>
							<%else%>
                             <h3>
                             <% if bacheca<>"" then%>
                             (Bacheca di <%=bacheca%>) <i class="icon-th-list"></i> <%=rs("Topic")%>
                              <%else%>
                                      <i class="icon-th-list"></i> <%=rs("Topic")%>
                              <%end if%>
                              </h3>
                            <%end if%>
                            </div>
							<div class="box-content nopadding">

							<% if Anonimo=1 Then
							    autore="anonimo"
									else
									autore= rs("AuthorName")
								 end if


							%>

								<form action="#" method="POST" class='form-horizontal form-bordered'>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Pubblicato da</b></label>
										<div class="controls">
                                     <span class='author'>
							 <i class='icon-user'></i>  <%=autore%> </span>&nbsp;

							<% if session("Zip")=1 then%>
                             <span class='author'>
                            <i title="Prevede la consegna di siti in .zip " class='icon-paper-clip'></i></span> &nbsp;
                            <% end if%>
							 <span class='date'>
                             <%DataUltimoPost=rs("LastThreadPost")%>
							 <i class='icon-calendar'></i> Il  <%= rs("DatePosted")%>
							 </span> &nbsp;
                              <span class='comments'>
                             <i  class="glyphicon-eye_open"></i> <%= rs("Visualizzazioni")%>
                          Visualizzazioni</span>&nbsp;

                              <span class='comments'>
							 <i class='icon-comments'></i> <%= rs("ReplyCount")%> commenti </span> &nbsp;
                              <span class='date'>
                             <%
							' qua lalogica per recuperare l'ultimo post in modo da mettere il link diretto

							 %>


							 <i class='icon-calendar'></i> Ultimo  <%=DataUltimoPost%>
							 </span> &nbsp; <span class='comments'>
                                  <i class="glyphicon-user_add" title="Punti aggiunti in classifica"></i>&nbsp;<span title="Punti aggiunti in classifica"><%=pt%>=<%=rs("Punti")%></span>&nbsp;
							   </span>

										</div>
									</div>

                                    <% if strcomp(scegli,"0")=0 then %>
 									<div class="control-group">
										<label   class="control-label"><b>Nella bacheca di</b> </label>
										<div class="controls">
											<%=ucase(rs("Bacheca"))%>; Condiviso con <%=condivisoCon%>
										</div>
									</div>

                                    <% end if%>





                                <%
								CodiceAllievo=rs("CodiceAllievo")
								if Session("Cartella")="" then
		QuerySQLR="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
		Set rsTabellaR = ConnessioneDB.Execute(QuerySQLR)
		Session("Cartella")=rsTabellaR("Cartella")
end if

				url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/file_"&session("social") ' vuole il percorso relativo della cartella url=Replace(url,"\","/")
			   urlfile=url&"/"& Urlfile2
			   'response.write(urlimg)
			   'Session("Caricata")=false
if NomeFile<>"" then
			   %>

                <div class="control-group">
                <%if Session("Privato1")=0 then %>
										<label   class="control-label"><b>Risorse <%=rs("PrivatoLab")%></b></label>
										<div class="controls">
											<B><a target="_blank" href="<%=urlfile%>"><%=NomeFile%></a></B>
										</div>
                  					  <%if (instr(Urlfile2,".pdf")<>0) or (instr(Urlfile2,".txt")<>0) or (instr(Urlfile2,".cpp")<>0) or (instr(Urlfile2,".java")<>0)  then
								'  url2=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/file_lavagna"  'per il server on line
								  %>
								  <iframe src=<%=urlfile%>  width=100%  height= 480px ></iframe>
								  <%end if%>
               <% end if%>
                </div>



<% end if %>



<%


    ' tolgo suffiso e ricavo chiave
	if strcomp(Urlimg2&"","")<>0 then
		keyimg=left(Urlimg2, (len(Urlimg2)-4)  )

		QuerySQL = "select Href_O from IMG_FORUM where ID_Smile="&keyimg&";"

			'response.write(QuerySQL)
			 set oRs1=conn.Execute(QuerySQL)

			 if not oRs1.eof then
			  Azione=oRs1("Href_O")
			 else
			  Azione="#"
			 end if
			' response.write("Azione="&Azione)


	  ' per l'immagine eventuale
				url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/img_"&session("social")&"/img" ' vuole il percorso relativo della cartella
			   url=Replace(url,"\","/")
			   urlimg=url&"/"& Urlimg2
			   'response.write(urlimg)
			   'Session("Caricata")=false
			   %>
			  <br> <center>
			  <% if strcomp(Azione&"","#")= 0  then %>
				  <img class="imground" src="<%=urlimg%>" title="<%=urlimg%>" border="1"></center>
			  <% else%>
			  <a target="_blank" href="<%=Azione%>"><img class="imground" title="<%=urlimg%>" src="<%=urlimg%>" border="1"> </a></center>
			  <%
			     end if%>
  <% end if %>



        			<div class="control-group">

										<label for="textarea" class="control-label"><b>Post</b></label>
                                        <div class="controls">

                                        <center>
                                        <% ' per scrivere fuori da ckeditor il iframe con il video
										  k=instr(sMsg,"</iframe>")
										  if (k <>0) then%>

										  <%iFrameMsg=left(sMsg,k+8)
										'  response.write(iFrameMsg)
										 sMsgDx =right(sMsg,len(sMsg)-(instr(sMsg,"</iframe>")+8))
										 else
										 sMsgDx=sMsg
										  end if
										%>

                                        </center>
                                        <% 
										
										 
										sMsg2= ReplaceCar(sMsg)
										
										if instr(sMsg2,"<script>")<>0 then
										sMsg2=Replace(sMsg2,"<script>","")
										sMsg2=Replace(sMsg2,"</script>","")
										end if
										
										%>
                                    <!--     <textarea class='ckeditor span12' rows="5" name="messaggio" cols="40"  >-->
										 <%=FormatMessage(sMsg2)%>
                                        <%'= sMsgDx
										%>
                                      <!--   </textarea>-->



										</div>
									</div>

                                    <div class="control-group">
									 	<table class="table table-hover table-nomargin  table-striped">

                                        <tr><td colspan="2"><center>
<a title="Fai da 1 a 5 click per esprimere quanto ti piace (Voto da 6 a 10); 5 click = ti piace 10" href="vota_messaggio.asp?scegli=<%=scegli%>&ID=<%=iMessageId%>&Topic=<%=rs("Topic")%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&MaxStelline=<%=MaxStelline%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>&Zip=<%=Zip%>"><img src="img/facebook2.jpg" width="21" height="19" align="bottom">&nbsp;OK&nbsp;<img src="img/OCCHI_pupilla_bianca_small.png" width="20" height="15" ></a>
<br><br>
      


<a title="Fai da 1 a 5 click per esprimere quanto non ti piace (Voto da 5 a 0); 5 click = ti piace 0" href="vota_messaggio.asp?scegli=<%=scegli%>&revoca=1&ID=<%=iMessageId%>&Topic=<%=rs("Topic")%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&MaxStelline=<%=MaxStelline%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>&Zip=<%=Zip%>">
<img src="img/facebook8_nonpiace_small.jpg" width="20" height="17">&nbsp;KO&nbsp;<img
src="img/OCCHI_pupilla_nera_small.png" width="20" height="15"  ></a>
</center></td></tr>

</TABLE>
</form>


    <div class="accordion" id="accordion2">
<%if  (Session("Admin")=true) then %>



									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapseOne">
												<center>Gestisci voti</center>
											</a>
										</div>
										<div id="collapseOne" class="accordion-body collapse">
											<div class="accordion-inner"><center>
												   <a title="Fai click per resettare le votazioni degli studenti" href="vota_messaggio.asp?scegli=<%=scegli%>&revocatutti=1&ID=<%=iMessageId%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&Zip=<%=Zip%>">
 &nbsp;<i class=" icon-remove-sign"></i></a>&nbsp; Resetta Votazioni <hr>
  <a title="Fai click per resettare i punti social assegnati dal docente" href="vota_messaggio.asp?scegli=<%=scegli%>&revocaPS=1&ID=<%=iMessageId%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&Zip=<%=Zip%>">
 &nbsp;<i class="icon-remove"></i></a>&nbsp; Resetta Punti Social<hr>


     <form  action="valuta_post.asp">
     Valuta tutti :
     <input type="text" value ="" size="1" name="txtVoto">
     <INPUT TYPE="HIDDEN" NAME="MessageIDV" VALUE="<%= iMessageID %>">
	 <INPUT TYPE="HIDDEN" NAME="ThreadIDV" VALUE="<%= iThreadParent %>">
      <INPUT TYPE="HIDDEN" NAME="categoria" VALUE="<%= categoria %>">
       <INPUT TYPE="HIDDEN" NAME="id_categoria" VALUE="<%= id_categoria %>">
     <INPUT TYPE="HIDDEN" NAME="scegli" VALUE="<%= scegli %>">
     <input type="submit"   value="Valuta..">
     </form>
     </center>


											</div>
										</div>
									</div>

     <%end if%>

   <%if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then %>

									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapseTwo">
												<center>Aggiorna</center>
											</a>
										</div>
										<div id="collapseTwo" class="accordion-body collapse">
											<div class="accordion-inner">


          <form class="form-vertical" name="InputForm" action="aggiorna_messaggio.asp?scegli=<%=scegli%>&ID_Smile=<%=keyimg%>&ID=<%=iMessageId%>&CodiceAllievo=<%=CodiceAllievo%>&bacheca=<%=bacheca%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>" METHOD = "POST"  >


    <div class="control-group">
				<label for="textfield" class="control-label"><b>Argomento: </b></label>
				<div class="controls">
	 			<input type="text" value ="<%=rs("Topic")%>" name="txtArg" id="textfield" class="input-xxlarge" >
				</div>
	</div>
     <div class="control-group">
				<label for="textfield" class="control-label"><b>In breve: </b></label>
				<div class="controls">
	 			<input type="text" value ="<%=rs("Abstract")%>" name="txtAbstract" id="textfield" class="input-xxlarge" maxlength="198" >
				</div>
	</div>

     <div class="control-group">
				<label for="textfield" class="control-label"><b>Messaggio :</b></label>
				<div class="controls">
	 		  <textarea class='ckeditor span12' rows="5" name="messaggio" cols="40" ><%=sMsg%></textarea>
              <hr>
           <center>   <input type="submit"  value="Aggiorna" class="btn"></center><br>
		    <input type="checkbox"  name="cbEmail2" id="cbEmail1" title="Selezionare per inviare un email alla classe">   Notifica per email a tutta la classe &nbsp;&nbsp;&nbsp;<br><br/><br>



				</div>
	</div>

    <p>
    <% if session("Admin")=true then %>



    <div class="control-group">
				<label for="textfield" class="control-label"><b>Azione : </b></label>
				<div class="controls">
	 		   <textarea name="txtAzione" placeholder="Incolla URL Vai al compito" rows="5" class="input-block-level">
			   <%=rs("Azione")%>
               </textarea><br>
				</div>
	</div>

      <div class="control-group">
				<label for="textfield" class="control-label"><b>Nome autore :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtUser" value="<%=rs("AuthorName")%>" class="input-xlarge">
				</div>
	</div>

      <div class="control-group">
				<label for="textfield" class="control-label"><b> Codice autore :  </b></label>
				<div class="controls">
	 			<input type="text"  name="txtCodiceAllievo" value="<%=rs("CodiceAllievo")%>" class="input-xlarge">
				</div>
	</div>
     <div class="control-group">
				<label for="textfield" class="control-label"><b> Cambia autore </b></label>
				<div class="controls">

<div class="accordion" id="accordionStud">
									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle"  data-toggle="collapse" data-parent="#accordionStud" href="#collapseStud">
												(+) Scegli nuovo
											</a>
										</div>
										<div id="collapseStud" class="accordion-body collapse">
											<div class="accordion-inner">


                <table class="table table-hover table-nomargin table-condensed">
				<%  QuerySQL="SELECT Cognome,Nome,CodiceAllievo" &_
                " FROM Allievi  " &_
                " WHERE Id_Classe ='" & Session("Id_Classe") & "' and Attivo=1" &_
                " ORDER BY Allievi.Cognome Asc; "
                Set rsTabella = ConnessioneDB.Execute(QuerySQL) %>
                <tr><td><b>Studente</b></td><td><b>Codice</b></td><tr>
                <%
                   i=1
                   do while not rsTabella.eof %>
                       <tr>
                           <td><%=rsTabella.fields("Cognome") & " " & left(rsTabella.fields("Nome"),1) &"." %></td>
                           <td>  <%=rsTabella.fields("CodiceAllievo")%></td>
                       </tr>
                   <%  rsTabella.movenext
                   i=i+1
                   loop
                   rsTabella.close
                %>
                </table>


											</div>
										</div>
									</div>
								</div>

    </div>
	</div>

    <div class="control-group">

				<label for="textfield" class="control-label"><b> ID:  </b></label>
				<div class="controls">
	 			<input type="text"  name="txtID" value="<%=rs("ID")%>" class="input-mini">
				</div>
	</div>

    <div class="control-group">
				<label for="textfield" class="control-label"><b> ThreadParent :  </b></label>
				<div class="controls">
	 			<input type="text"  name="txtThreadParent" value="<%=rs("ThreadParent")%>" class="input-mini">
				</div>
	</div>
     <div class="control-group">
				<label for="textfield" class="control-label"><b> Parent Message :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtParentMessage" value="<%=rs("ParentMessage")%>"  class="input-mini">
				</div>
	</div>
    <div class="control-group">
				<label for="textfield" class="control-label"><b> Data :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtDATA" value="<%=rs("DatePosted")%>" class="input-xsmall">
				</div>
	</div>
	<% Scadenza = rs("ScadenzaEvent")
    if Scadenza<>"" then
      DateIta = Split(Scadenza, "-")
	  ToEng = DateIta(0) & "-" & DateIta(1) & "-" & DateIta(2)

		Scadenza = cDate(ToEng)
	end if
	%>
	<div class="control-group">
				<label for="textfield" class="control-label"><b> Scadenza (come compiti) :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtScadenza" id="txtScadenza" value="<%=Scadenza%>" class="input-xsmall">
				</div>
	</div>
    <div class="control-group">
				<label for="textfield" class="control-label"><b> Valutazione :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtVAl" value="<%=rs("Punti")%>" class="input-mini">
				</div>
	</div>
    <div class="control-group">
				<label for="textfield" class="control-label"><b> Risposte :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtRC" value="<%=rs("ReplyCount")%>" class="input-mini">
				</div>
	</div>
      <div class="control-group">
				<label for="textfield" class="control-label"><b> Visibile :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtVisibile" value="<%=rs("Visibile")%>" class="input-mini">
				</div>
			</div>
     <div class="control-group">
				<label for="textfield" class="control-label"><b> Privato :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtPrivato" value="<%=rs("Privato")%>" class="input-mini">
				</div>
			</div>

     	<div class="control-group">
				<label for="textfield" class="control-label"><b> Private Lab :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtPrivato1" value="<%=rs("PrivatoLab")%>" class="input-mini">
				</div>
			</div>

        <div class="control-group">
				<label for="textfield" class="control-label"><b>Visualizzazioni :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtVisualizzazioni" value="<%=rs("Visualizzazioni")%>" class="input-mini">
				</div>
				</div>

				<div class="control-group">
				<label title="Nome autore anonimo" for="textfield" class="control-label"><b>Anonimo :  </b></label>
				<div class="controls">
	 			<input type="text" name="txtAnonimo" value="<%=rs("Anonimo")%>" class="input-mini">
				</div>
				</div>

				 <div class="control-group">
				<label for="textfield" class="control-label"><b>Non hanno risposto :  </b></label>
				<div class="controls">
				 <%=nonconsegnato%>
				</div>
			</div>




	<%else%>



    <input  TYPE="HIDDEN" name="txtUser" value="<%=rs("AuthorName")%>">&nbsp;&nbsp;&nbsp;

    <input  TYPE="HIDDEN" name="txtCodiceAllievo" value="<%=rs("CodiceAllievo")%>"><br>
     <input  TYPE="HIDDEN" name="txtID" value="<%=rs("ID")%>" size="2">
      <input  TYPE="HIDDEN" name="txtThreadParent" value="<%=rs("ThreadParent")%>" size="2">
    <input  TYPE="HIDDEN" name="txtParentMessage" value="<%=rs("ParentMessage")%>" size="2">
     <input  TYPE="HIDDEN" name="txtDATA" value="<%=rs("DatePosted")%>" size="12">
    <input  TYPE="HIDDEN" name="txtVAl" value="<%=rs("Punti")%>" size="1"><br>
    <input  TYPE="HIDDEN" name="txtRC" value="<%=rs("ReplyCount")%>" size="1">

	<%end if%>












<center>


    </form>



											</div>
										</div>
									</div>

								</div>

<%end if%>




                                        </table>

									</div>

                                  <!--

									<div class="form-actions">
										<button type="submit" class="btn btn-primary">Save changes</button>
										<button type="button" class="btn">Cancel</button>
                                        -->
									</div>
								</form>
							</div>
						</div>







				<!--      <div class="box-content">-->








<%'else



		 %>


				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">




					<%  i=0 %>
		<!--    <div class="box-content"> -->










<%


if not bValid then
  response.write "You cannot navigate to this page without selecting a forum message.  Please "
  response.write "return to the <A HREF = 'default.asp?scegli="&scegli&"&'>forum index</A> and try again."
  response.end
end if
sTopic = rs("Topic") ' & rs("Comments") io
' ??? CodiceAllievo=rs("Topic")
CodiceAllievo=rs("CodiceAllievo")
' se sono stato richiamata dalla stessa pagina vuol dire che voglio rispondere, quindi vado diretto senza chiedere
' di premere il bottone rispondi
Rispondi=request.QueryString("Rispondi")
if Rispondi<> "" then
 '  response.redirect "reply.asp?MessageID="&rs("ID")&"&ThreadID="&rs("ThreadParent")&"&OrigAuthor="&rs("AuthorName")
end if


%>
<% if sCaption <> "" then
response.write "<B><FONT COLOR = RED> " & sCaption & "</B></FONT><P>"
end if
%>






<!-- </div>-->







<FORM ACTION = "reply.asp?scegli=<%=scegli%>&bacheca=<%=bacheca%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>" METHOD = "POST">
<INPUT TYPE="HIDDEN" NAME="MessageID" VALUE="<%= iMessageID %>">
<INPUT TYPE="HIDDEN" NAME="ThreadID" VALUE="<%= iThreadParent %>">
<INPUT TYPE="HIDDEN" NAME="Topic" VALUE="<%= rs("Topic") %>">
<INPUT TYPE="HIDDEN" NAME="OrigAuthor" VALUE="<%= rs("AuthorName") %>">
<br>
<CENTER>
<% if strcomp(scegli,"2")=0 then %>
 <INPUT TYPE="Submit"  style="width:130px;" NAME = "RequestReply" class="btn" VALUE = "Commenta">
<a target="_blank" href="<%=rs("Azione")%>"><INPUT type="button" class="btn"  style="width:130px;"  NAME = "buttonCompito" VALUE = "Vai al compito"></a>
<%else %>
<INPUT type="submit"  NAME = "RequestReply" class="btn" VALUE = "Rispondi">

<%end if%>
</CENTER>

</FORM>
</Fieldset>

<%

rs.close
cmd.parameters.delete(0)
sThread = MessageChildren(iThreadParent, 0, iMessageID,0)
 if instr(sThread,"<script>")<>0 then
	   sThread=Replace(sThread,"<script>","")
	   sThread=Replace(sThread,"</script>","")
	end if
if sThread <> "" then

	  response.write "<center>  "
	response.write "<div class='hr' style='width:85%; '><hr /></div>"  'HR HEIGHT = 1 NOSHADE

	response.write "</center> "

	 response.write "<div id='d1'><div class='contenuti' style='background-color:#ffffff;padding:10px;'>"

	response.write " <div class='blog-list-post'>"
	'response.write "<div class='preview-img'>img</div>"
    response.write "<div class='post-content'>"
	response.write "<h4 class='post-title'><a href='#'>Commenti</a></h4>"
    response.write "<div class='post-meta'>"
							response.write "<span class='author'>"
							if Anonimo=0 then
							response.write "<i class='icon-user'></i>  Autore: " & CodiceAllievo&" "
							else
							response.write "<i class='icon-user'></i>  Autore: anonimo"
							end if
							response.write "</span>&nbsp;"
							response.write "<span class='comments'>"
							response.write "<i class='icon-comments'></i> Commenti : "& RCount &"  "
							response.write "</span>&nbsp;"
							response.write "<span class='date'>"
							response.write "<i class='icon-calendar'></i> Ultimo  "&DataUltimoPost
								'response.write "<i class='icon-calendar'></i> Voti=  "&Voti
							response.write "</span>&nbsp;"

							'response.write "<span class='tags'>"
							'response.write "<i class='icon-tag'></i>"
							'response.write "<a href='#'>ui</a>"
							'response.write "<a href='#'>flat</a>"
							'response.write "<a href='#'>clean</a>"
							'response.write "</span>"

							response.write "</div>"
							response.write "<div class='post-text'>"
							'response.write sCaption
							response.write "</div>"
							response.write "</div>"
							response.write "<div class='post-comments'>"
							response.write sThread
							response.write "</div>"

							response.write "</div></div>"
end if



%>
 
<!--#include file = "database_cleanup.inc"-->
</div>






                      </div>
			        </div>
			  

			    </div>
	       
			        </div>
			      </div>
			    </div>
			</div>
			<!-- <div id="load" class="loader" style="display:none"></div> -->
		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->

		 <script>

		 function cambiavoto(id){

		 idpost = id;
		 voto = $("#sel"+id).val();

		// document.getElementById("load").style.display="block";

		  $.ajax({
						method: "POST",
						url: "aggiornavoto_post.asp?id="+id+"&voto="+voto,
						dataType: "html",
						data: {  }
					}) /* .ajax */
					.done(function( ans ) {
					/*	var t = setTimeout(function(){
							document.getElementById("load").style.display="none";
							clearTimeout(t);
						},700);*/
					}) /* .done */
					.error(function(jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});

		 }

		 </script>

		 <form id="mod" name="mod" method="post" action="aggiorna_post_ajax.asp">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none; ">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Modifica post</h3><button type="button" id="inviamodifica" class="btn btn-primary" onClick="aggiorna_post(<%=id%>)">Aggiorna</button>

				</div>
				<div class="modal-body">

					<input  name="url" id="url" type="hidden"><input  name="codice" id="codice" type="hidden"> <input  name="idtxt" id="idtxt" type="hidden">
					<input class="input-xlarge"  name="titolopost" id="titolopost" type="text"> <input  name="idtxt" id="idtxt" type="hidden">

				<!--	<textarea    style="width: 97%" name="spiegazione" id="spiegazione"   cols="150" rows="30">
  class='ckeditor span12' -->
					<!--</textarea>-->
<!--
					<div id="editor">
       				 <p>This is some sample content.</p>
    				</div>
-->
				<textarea name="editor1" id="editor1"></textarea>

				</div>
				<div class="modal-footer">
					<button id ="chiudi" class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>

				</div>
			</div>
		</form>

	</body>

 <script>
 /*
    ClassicEditor
    .create( document.querySelector( '#editor' ) )
    .then( newEditor => {
        editor = newEditor;
    } )
    .catch( error => {
        console.error( error );
	} );*/

	 
    CKEDITOR.replace('editor1');
 
	
</script>



<script>
function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

window.open('share.asp?scegli='+s,'share.asp?scegli='+s, 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);

}
var globalpostid;
	function modifica(post){
	     globalpostid=post;
		var url = "carica_post_ajax.asp?id="+post;
		var xhttp = new XMLHttpRequest();
		xhttp.onreadystatechange = function() {
		if (xhttp.readyState == 4 && xhttp.status == 200) {
			var testo=xhttp.responseText;
			var res = testo.split("£££");
  			titolo=res[0];
			testo=res[1];
		//	titolo=testoJSON["topic"]; quando ci sono righe vuote nella risposta il json va in errore quindi uso il trick sopra £££
		//	testo=testoJSON["comments"];
			document.getElementById("titolopost").value=titolo;
			CKEDITOR.instances.editor1.setData(testo);
			//document.getElementById("editor1").value=testo; errato
		}
		};
		xhttp.open("GET", url, true);
		xhttp.send();
		}


function elimina_post(post,tparent) {
	if (window.confirm('Vuoi veramente cancellare il post?')) {
	  	 	var url="../cSocial/DeleteMessage_ajax.asp?ID="+post+"&TParent="+tparent;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								if (risposta=="ok")
								{
								   document.getElementById("post_"+post).style.display="none";
								//	$('#riga_'+riga).remove();
								  //alert("Eliminato");
								}
								else
									alert(risposta);
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 }
}

function calcola_media_voto(post,tparent) {
		 
			
	 
	  	 	var url="voti_messaggio_media_ajax.asp?ID="+post+"&TParent="+tparent;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
							var risposta=String(xhttp.responseText);
							var res=risposta.split("£")
							 document.getElementById("voto"+post).innerText=res[0];
							 document.getElementById("img"+post).setAttribute("title",res[1]);
							  
								 
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 

}


function vota_post(post,tparent,segno,scegli) {
			var stella;
			if (segno==1)
			  stella="(+)";
			else 
			stella="(-)";
			
	if (window.confirm('Vuoi assegnare una stella '+ stella+' al post?')) {
	  	 	var url="vota_messaggio_ajax.asp?ID="+post+"&TParent="+tparent+"&segno="+segno+"&scegli="+scegli;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								if (risposta=="ok")
								{
								   document.getElementById("post_"+post).style.display="none";
								//	$('#riga_'+riga).remove();
								  //alert("Eliminato");
								}
								else{
									calcola_media_voto(post,tparent);
									alert(risposta);
								}
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 }

}




function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
window.open('share.asp?scegli='+s,'share.asp?scegli='+s, 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);
}



</script>

  <script type="text/javascript" src="../js/refresh_session.js"></script>
 </html>
