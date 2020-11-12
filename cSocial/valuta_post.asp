<%@ Language=VBScript %>

  <% Response.Buffer=True
   Dim ConnessioneDB, conn, rsTabella,rsTabella1, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome


    ID=request.querystring("ThreadIDV")
	punti=cint(Request.querystring("txtVoto"))
	RCount= request.QueryString("RCount") ' numero di risposte della discussione serve per decrementare in update in delete
	TParent=cint(request.QueryString("TParent")) ' IDdel post per aggiornare ReplyCount
	scegli=cint(request.QueryString("scegli"))
	bacheca=cint(request.QueryString("bacheca"))
	id_categoria=request.form("id_categoria")
    categoria=request.form("categoria")

	 set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>
<html>
<head>
	<link rel="stylesheet" type="text/css" href="../../stile.css">


</head>


	<!--#include file = "../service/controllo_sessione.asp"-->
    <!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->

<%



 scegli=request("scegli") ' 0 = forum 1=lavagna 2=diario
'on error resume next
select case scegli
 case "0"
     session("social")="forum"
	 response.write("0")

 case "1"

    session("social")="lavagna"
	 response.write("1")
  case "2"
    session("social")="diario"
	 response.write("2")
   case "3"
     session("social")="interrogazioni"
    response.write("3")

 end select  %>
 




<%if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<body>
    <div id="container">
<div class="contenuti_forum">
	<font color=#FF0000 size="4">

<%


' prelevo tutti i codiciallievi di chi ha risposto almeno una volta, poi assegno 1 punto al primo post di ognuno
' quindi di default 1 punto a chi partecipa indipendentemente da quante volte risponde
'QuerySQL="SELECT DISTINCT (FORUM_MESSAGES.CodiceAllievo), FORUM_MESSAGES.ThreadParent" &_
'" FROM FORUM_MESSAGES WHERE  ThreadParent ="&ID &";"
'set rsTabella=conn.execute(QuerySQL)
 'response.write("<br>"&QuerySQL)

i=0
'do while not rsTabella.eof and i<10
'do while i<10
i=i+1
  '  QuerySQL="SELECT * FROM FORUM_MESSAGES " &_
'" WHERE  ThreadParent ="&ID &" and CodiceAllievo='"&rsTabella("CodiceAllievo")& "';"


	'set rsTabella1=conn.execute(QuerySQL)

	' response.write("<br>"&QuerySQL)

	'QuerySQL="UPDATE FORUM_MESSAGES SET punti = (punti + 1) WHERE ThreadParent="&ID&" and CodiceAllievo='"&rsTabella1("CodiceAllievo")&"' and ID="& rsTabella1("ID")& " And Punti=0;"
'QuerySQL="UPDATE FORUM_MESSAGES SET punti = (punti + 1) WHERE ThreadParent="&ID&" and CodiceAllievo='"&rsTabella("CodiceAllievo")&"' And Punti=0;"
	QuerySQL="UPDATE FORUM_MESSAGES SET punti = (punti + " & request("txtVoto") &") WHERE ThreadParent="&ID&" And Punti=0;"
    ConnessioneDB.Execute(QuerySQL)
	'response.write("<br>"&i&" "&QuerySQL)

	'rsTabella.movenext
'loop




'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA


'response.write(url)

 Set ConnessionDB=nothing

On Error Resume Next
If Err.Number = 0 Then
		Response.Write "Aggiornamento avvenuto! "
		Response.Redirect "ShowMessage.asp?ID="&ID&"&scegli="& scegli&"&ThreadParent="& iMessageId&"&RCount="&RCount&"&TParent="&TParent&"&id_classe="&id_classe&"&bacheca="&bacheca&"&categoria="&categoria&"&id_categoria="&id_categoria

Else
		Response.Write Err.Description
		Err.Number = 0
End If
%>
	<center><br><br><font size="3">

</center>

 <!-- se il login � corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>

   <BODY onLoad="showText();">

	<%end if%>
	</body>
	</html>
