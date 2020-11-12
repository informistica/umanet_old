<%@ Language=VBScript %>


<% ID=Request.QueryString("ID")




%>
<html>
<head>
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />


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
function showText() {window.alert("Non puoi cancellare i dati degli altri studenti!")

location.href="ShowMessage.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&ID=<%=ID%>"
//location.href=window.history.back();
 }

 $(window).ready(function () {
	   $('#msg').click();

	  // event.stopPropagation();

	});

 </script>
</head>


   <% Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome




   'Apertura della connessione al database
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	'id_classe=request.querystring("id_classe")

	'serve per tornare alle bacheche
	cognome=request.querystring("cognome")
	nome=request.querystring("nome")
	bacheca=request.querystring("bacheca")
	'bacheca=session("bacheca")
	CodiceAllievo=request.querystring("CodiceAllievo")
	discussione=request.querystring("discussione")


	categoria=request.querystring("categoria")
	id_categoria=request.querystring("id_categoria")

	id_classe=session("Id_Classe")
	cartella=request.QueryString("cartella")




	RCount=cint(request.QueryString("RCount")) ' numero di risposte della discussione serve per decrementare in update in delete
'*** se RCount="" si pianta qua
'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logDeleteMessage.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine("150")
'				objCreatedFile.Close


     TParent=cint(request.QueryString("TParent")) ' IDdel post per aggiornare ReplyCount
Zip= request.QueryString("Zip")
       %>



   <%
 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="lavagna"
  case "2"
    session("social")="diario"
    case "3"
      session("social")="interrogazioni"

 end select




  %>




    <%



	%>
     <!--#include file = "../service/controllo_sessione.asp"-->
		<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->

<%
                            'Lettura dei dati memorizzati nei cookie.
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")






if (strcomp(ucase(session("CodiceAllievo")),ucase(CodiceAllievo) )=0 )  or (Session("Admin")=true)  or (strcomp(ucase(session("CodiceAllievo")),ucase(bacheca) )=0 )  then  %>
<body class='theme-<%=session("stile")%>'>
    <div id="container">
<div class="contenuti_forum">

 <div class="loader"></div>
	<font color=#FF0000 size="4">


<%



		if discussione<>"" then ' cancello tutti i post della discussioner

		QuerySQL = "SELECT * FROM FORUM_MESSAGES WHERE ID = "&ID&";"
 set rsTabella = conn.Execute(QuerySQL)

 eventopost = rsTabella("IDEvent")

			QuerySQL =" SELECT FORUM_MESSAGES.*, FORUM_MESSAGES.ThreadParent " &_
			" FROM FORUM_MESSAGES " &_
			" WHERE  FORUM_MESSAGES.ThreadParent = " & ID & ";"
			set rs=conn.Execute(QuerySQL)

				do while not rs.eof
				 	 ID=rs("ID")
					  QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE ID =" &ID&";"
					  conn.Execute(QuerySQL)
					  rs.movenext

				loop
				 QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE ID =" &ID&";"
			     conn.Execute(QuerySQL)
				 url = "default0.asp?scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cognome="&cognome&"&nome="&nome&"&categoria="&categoria&"&id_categoria="&id_categoria

				 QuerySQL = "SELECT * FROM Classi WHERE ID_Classe = '"&Session("Id_Classe")&"'"
				 set rsTabella = conn.Execute(QuerySQL)

				 idcal = rsTabella("Url_calendar")

				%>

				<script>
					$.ajax({
						method: "POST",
						url: "../../../../googleapi/delevento.php?calendario=<%=idcal%>",
						dataType: "html",
						data: { calendario: "<%=idcal%>", idevento: "<%=eventopost%>" }
					}) /* .ajax */
					.done(function( ans ) {

					//alert("../../../../googleapi/addevento.php?calendario=<%=idcal%>&summary=<%=Avviso%>&description=<%=Azione%>&end=<%=Scadenza%>");

						//alert(ans);
						window.location.href = "<%=url%>";

					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(textStatus+": "+errorThrown);
					});

 </script>

 <%
		else ' cancello solo il commento



			 QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE ID =" &ID&";"
			 conn.Execute(QuerySQL)
			 response.write(QuerySQL &"<br>")
			 QuerySQL="UPDATE FORUM_MESSAGES SET ReplyCount = ReplyCount-1  " &_
		 " WHERE ID="&TParent&";"
		''	 response.write(QuerySQL &"<br>")
			  conn.Execute(QuerySQL)

			''	Response.write "ShowMessage.asp?zip="&zip&"&scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cartella="&Session("cartella")&"&ID="&Session("discussione")&"&RCount="&RCount-1&"&cognome="&cognome&"&nome="&nome&"&categoria="&categoria&"&id_categoria="&id_categoria&"&TParent="&TParent

Response.redirect "ShowMessage.asp?zip="&zip&"&scegli="&scegli&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&cartella="&Session("cartella")&"&ID="&Session("discussione")&"&RCount="&RCount-1&"&cognome="&cognome&"&nome="&nome&"&categoria="&categoria&"&id_categoria="&id_categoria&"&TParent="&TParent



		end if%>

         <span class="invisible">    <p><a id="msg" href="#modal-1" role="button" class="btn notify" data-notify-title="Valutazione effettuata con successo" data-notify-message="Stai tornando al Libro">Torna al libro</a>
                   </span>
        <%

'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA


'response.write(url)

On Error Resume Next

If Err.Number = 0 Then

		Response.Write "Cancellazione avvenuta! "
Else
	Response.Write Err.Description
	Err.Number = 0
End If


  							' if Request.ServerVariables("HTTP_REFERER") <>"" then
							'		response.Redirect request.serverVariables("HTTP_REFERER")
							'	end if %>

   %>


	<center><br><br><font size="3">
</center>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->
</font>
</div>
	<%else%>

   <BODY onLoad="showText();">

	<%

	end if%>




 </body>
	</html>
