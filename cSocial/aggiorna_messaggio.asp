<%@ Language=VBScript %>

  <% Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome

   on error resume next
   'Apertura della connessione al database
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	ID=request.querystring("ID")
	bacheca=request.querystring("bacheca")
    messaggio=Request.Form("messaggio")
	argomento=Request.Form("txtArg")
	data=Request.Form("txtData")
  punti=Request.Form("txtVAl")
  if punti="" Then
  punti=0
  else
	punti=cint(Request.Form("txtVAl"))
  end if
	CodiceAllievo=request.querystring("CodiceAllievo")
	AuthorName=Request.Form("txtUser")
	Azione=Request.Form("txtAzione")
	ID_Smile=Request.querystring("ID_Smile")
	AuthorCode=Request.Form("txtCodiceAllievo")
	ParentMessage=Request.Form("txtParentMessage")
	ThreadParent=Request.Form("txtThreadParent")
	ReplyCount=Request.Form("txtRC")
	Visibile=Request.Form("txtVisibile")
	Privato=Request.Form("txtPrivato")
	Privato1=Request.Form("txtPrivato1")
	 abstract=Request.Form("txtAbstract")
	 scadenza = Request.Form("txtScadenza")
	 cbEmail2=request.Form("cbEmail2")

	if Privato="" then
	  Privato=0
	end if
	if Privato1="" then
	  Privato1=0
	end if
	if Visibile="" then
	  Visibile=1
	end if
	Visualizzazioni=Request.Form("txtVisualizzazioni")
	if Visualizzazioni="" then
	  Visualizzazioni=1
	end if
  Anonimo=Request.Form("txtAnonimo")
  if Anonimo="" then
    Anonimo=0
  end if



       %>
<html>
<head>
	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">

    <!-- Datepicker new-->
	<link rel="stylesheet" href="../../css/plugins/datepicker/datepicker.css">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery UI -->
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.draggable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />


       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>


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


   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
-->

<!-- Datepicker -->

<!-- <script src="../js/plugins/datepicker/bootstrap-datepicker.it.js"></script> -->

  <script src="../../js/jquery-ui.js"></script>
 <script src="../../js/datapicker_it.js"></script>

    <script language="javascript" type="text/javascript">
function showText() {window.alert("Non puoi modificare i dati degli altri studenti!")

location.href="ShowMessage.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>"
//location.href=window.history.back();
 }
 </script>
</head>
 <!--#include file = "../stringhe_connessione/stringa_connessione_social.inc"-->
 <!-- #include file = "../cAdmin/include_mail.asp" -->
<%
scegli=request.QueryString("scegli")
select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="lavagna"
  case "2"
    session("social")="diario"
  case "3"
      session("social")="interrogazioni"

 end select  %>

  <%

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

  %>


	<!--#include file = "../service/controllo_sessione.asp"-->
    <!--#include file = "include/format_message.asp"-->

<%





  QuerySQL = "SELECT * FROM Classi WHERE Id_Classe = '"&Session("Id_Classe")&"'"
 set rsTabella = conn.Execute(QuerySQL)

 idcal = rsTabella("Url_calendar")

 QuerySQL = "SELECT IDEvent FROM FORUM_MESSAGES WHERE ID = '"&ID&"'"
 set rsTabella = conn.Execute(QuerySQL)

 evento = rsTabella("IDEvent")


if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then  %>
<body class='theme-<%=session("stile")%>'>
    <div id="container">
<div class="contenuti_forum">

   <div class="loader"></div>

	<font color=#FF0000 size="4">

<%
' li tolgo per il problema con l'aggiornamento delle imaggini come url
'messaggio=Replace(messaggio, Chr(39), Chr(96))
'argomento=Replace(argomento, Chr(39), Chr(96))

'argomento=Replace(argomento, Chr(34), "") ' toglie "
argomento = Replace(argomento, Chr(39), chr(96)) ' cambia ' in quello storto
'messaggio=Replace(messaggio, Chr(34), "")
'messaggio = Replace(messaggio, Chr(39), chr(96))
	'  messaggio = ReplaceComments(messaggio)
	%><textarea>
    <%=messaggio%>
    </textarea>
	 <%
	  'messaggio = prepStringForSQL(messaggio)

'messaggio=FormatMessage(messaggio)

QuerySQL="UPDATE FORUM_MESSAGES SET AuthorName='" & AuthorName & "', Topic= '" & argomento & "', comments = '" & messaggio &_
		"', DatePosted = '" & data & "', punti = " &punti & ", CodiceAllievo = '" &AuthorCode & "',ParentMessage = " &ParentMessage & ", ThreadParent = " &ThreadParent & ", ReplyCount = " &ReplyCount & ", Azione='"& Trim(Azione) & "', Visibile="& Visibile & ", Privato="& Privato & ", PrivatoLab="& Privato1& ", ScadenzaEvent = '"&Scadenza&"', Visualizzazioni="& Visualizzazioni& ", Anonimo="& Anonimo& ", Abstract="&  prepStringForSQL(ReplaceComments(abstract))&" WHERE ID="&ID&";"

		' dim objFSO,objCreatedFile
				' Const ForReading = 1, ForWriting = 2, ForAppending = 8
				' Dim sRead, sReadLine, sReadAll, objTextFile
				' Set objFSO = CreateObject("Scripting.FileSystemObject")
				' url="C:\inetpub\umanetroot\expo2015Server\logAggiornaForum.txt"
				' Set objCreatedFile = objFSO.CreateTextFile(url, True)
				' objCreatedFile.WriteLine(QuerySQL)
				' objCreatedFile.Close
			      response.write(QuerySQL)

			  conn.Execute(QuerySQL)


if ID_Smile<>"" then
QuerySQL="UPDATE IMG_FORUM SET Href_O='" & Azione & "' WHERE ID_Smile="&cint(ID_Smile)&";"
conn.Execute(QuerySQL)
end if



'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA


'response.write(url)

'On Error Resume Next
If Err.Number = 0 Then
		Response.Write "Aggiornamento avvenuto! "
		'Response.Redirect "ShowMessage.asp?ID="&ID&"&bacheca="&bacheca
		'if Request.ServerVariables("HTTP_REFERER") <>"" then
		'					response.Redirect request.serverVariables("HTTP_REFERER")
		 'end if
		Response.Write (QuerySQL)
Else
		Response.Write Err.Description
		Err.Number = 0
End If
%>
	<center><br><br><font size="3">
<!--#include file = "footer.inc"-->
</center>
<!--#include file = "database_cleanup.inc"-->
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>

   <BODY onLoad="showText();">

	<%end if%>

	<script>
 $.ajax({
						method: "POST",
						url: "../../../../googleapi/updevento.php",
						dataType: "html",
						data: { calendario: "<%=idcal%>", summary: "<%=argomento%>", end: "<%=Scadenza%>", idevento: "<%=evento%>" }
					}) /* .ajax */
					.done(function( ans ) {

						window.location.href = "<%=Request.ServerVariables("HTTP_REFERER")%>";

					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});

 </script>

	</body>
	</html>
