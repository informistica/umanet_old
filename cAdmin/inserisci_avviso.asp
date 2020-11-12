<!-- calcola_risultato_MODBC3.asp -->
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

</head>
<body>
    <div id="container">
	<div class="risultati_test" >

	 <div class="loader"></div>

	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
<!-- #include file = "../extra/test_server.asp" -->

<% Response.Buffer=True %>


<!-- #include file = "include_mail.asp" -->

<%   Dim ConnessioneDB, rsTabella, QuerySQL


   'Apertura della connessione al database
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	        %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

     <!-- #include file = "../var_globali.inc" -->
     <%
ID_Avviso=Request.QueryString("ID_Avviso")
nomeparagrafo=Request.QueryString("par")
'Segnalibro=Request.Form("txtSegnalibro")
'Segnalibro=Session("idBox") ' numero del modulo (da 1 a 9)
Segnalibro = Session("segnalibro")
if Segnalibrio="" then Segnalibrio=1
Segnalibrio=1


on error resume next

'BoxApro=Session("PosMod") ' numero di riga dell'argomento del modulo
BoxApro = Request.QueryString("par")
'BoxApro=Request.Form("txtBoxApro")
Avviso=Request.Form("txtAvviso")
Azione= Request.Form("txtAzione")
Comments=Request.Form("txtDescrizione")
Scadenza=Request.Form("txtScadenza")

      DateIta = Split(Scadenza, "/")
   ToEng = DateIta(2) & "-" & DateIta(1) & "-" & DateIta(0)

	ScadSQL = cDate(ToEng)

	QuerySQL="SELECT Id_Classe, Titolo, TitPar, ID_Mod, ID_Paragrafo,Cartella,URL,URL_OL,Classe,URL_L,URL_O,Posizione from MODULI_NOT_UMANET  where Id_Classe='"&Session("id_Classe")&"' and Visibile=1 order by PosMod, PosPar ;"
    set rsDivid=ConnessioneDB.execute(QuerySQL)
	segnalibro=1
	trovato=0
	capapro=0
	k=1
	'capapro=Request.QueryString("cap")
	capitolo=rsDivid("Titolo")

	 'Set objFSO = CreateObject("Scripting.FileSystemObject")
	'url="C:\inetpub\umanetroot\expo2015Server\logavviso.txt"
	'Set objCreatedFile = objFSO.CreateTextFile(url, True)

	' ' chiudere recrodset del file in fondo alla pagina
			



	do while (not rsDivid.eof) and (trovato=0)
		if (strcomp(rsDivid("TitPar"),nomeparagrafo)=0) then
			trovato=1
		else
			segnalibro=segnalibro+1
			rsDivid.movenext()
			if not rsDivid.eof then
				c=rsDivid("Titolo")
			'  response.write(capitolo & " " & c)
			 'objCreatedFile.WriteLine(capitolo & " " & c )
						if StrComp(capitolo, c) = 0 then
						' Response.Write("Le due stringhe sono uguali")
						else
							capitolo=c 
							k=k+1 ' conta i moduli inseriti mi serve come indice per le ancore al modulo dal quaderno
						' Response.Write("Le due stringhe sono diverse")
						end if
			end if

		end if
	loop
	capapro=k

	
	'objCreatedFile.Close

	'if Azione="" then
	'Azione=Server.MapPath(homesito)&"/home_app.asp?id_classe="&Session("Id_Classe")&"&divid="&Session("divid")
'	Azione="../.."&"/UECDL/home_app.asp?id_classe="&Session("Id_Classe")&"&divid="&Session("divid")
'	Azione = Azione & "&dividApro="& Segnalibro ' in realt� � id del box da aprire (r-f-n)
'	Azione = Azione & "#" &BoxApro ' in realt� � il segnalibro
	      Azione=homesito&"/script/cClasse/home_app.asp?divid="&Session("divid")&"&id_classe="&Session("id_Classe")&"&classe="&Session("Classe")&"&cartella="&Session("Cartella")&"&capApro="&capapro&"&sottoCapApro="&Request.QueryString("sottopar")


		  ' Azione=dominio&homesito&"/home_app.asp?divid="&Session("divid")&"&id_classe="&Session("id_Classe")&"&classe="&Session("Classe")&"&cartella="&Session("Cartella")
		  ' Azione= "https://localhost/anno_2012-2013/UECDL/home_app.asp?divid=quarta&id_classe=5COM&classe=3PC&cartella=3PC"
		   'Azione = Azione & "&dividApro="& BoxApro
		   'Azione = Azione & "#" & Segnalibro
			Azione = Azione & "&dividApro="& Segnalibro& "#"& Segnalibro

	Azione=replace(Azione,"\","/")
	Comments=replace(Comments,"\","/")
'end if

'Id_Classe=Request.QueryString("Id_Classe")
'divid=Request.QueryString("divid")
Id_Classe=Session("Id_Classe")
divid=Session("divid")
Avviso = Replace(Avviso, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
Avviso=  Replace(Avviso,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql

 QuerySQL = "SELECT * FROM Classi WHERE Id_Classe = '"&Id_Classe&"'"
 set rsTabella = ConnessioneDB.Execute(QuerySQL)

 idcal = rsTabella("Url_calendar")

 QuerySQL="select Id_Categoria from CAT_CAT where Id_Classe='"&Id_Classe&"' and Descrizione='Compiti';"
 rsTabella=ConnessioneDB.Execute (QuerySQL)
 id_categoria=rsTabella(0)

Commentatore= Session("Cognome") & " " & left(Session("Nome"),1) &"."

 'datains=  rtrim(left(now(),10))
' datains=FormatDateTime(now, 4) solo ora
datains=now()

   'QuerySQL="INSERT INTO FORUM_MESSAGES (Topic,DatePosted,Azione,AuthorName,Id_Classe,CodiceAllievo,Comments,Id_Social,Id_Categoria)  SELECT '" & Avviso & "','" & FormatDateTime(now(),0)   & "', '" & Azione & "', '" & Commentatore & "', '" & Session("Id_Classe") & "', '" & Session("CodiceAllievo")& "', '" & Comments & "',"&Session("Id_Social")&",1;"
   QuerySQL="INSERT INTO FORUM_MESSAGES (ParentMessage,Topic,DatePosted,Azione,AuthorName,Id_Classe,CodiceAllievo,Comments,Id_Social,Id_Categoria,LASTTHREADPOST,ScadenzaEvent)  SELECT 0,'" & Avviso & "','"& datains&"','" & Azione & "', '" & Commentatore & "', '" & Session("Id_Classe") & "', '" & Session("CodiceAllievo")& "', '" & Comments & "',2,"&id_categoria&",'"&now()&"','"&ScadSQL&"';"

   response.write(QuerySQL)
   ConnessioneDB.Execute (QuerySQL)
   QuerySQL="select max(ID) from FORUM_MESSAGES"
   rsTabella=ConnessioneDB.Execute (QuerySQL)
   maxID=rsTabella(0)
   QuerySQL ="UPDATE FORUM_MESSAGES SET ThreadParent = '" & maxID  &"' WHERE ID =" &maxID&";"
   response.write(QuerySQL)
   ConnessioneDB.Execute (QuerySQL)

   ' DEVO AGGIUNGERE CREAZIONE FILE PER SPIEGAZIONE AVVISO OPPURE CAMPO MEMO IN TABELLA

	'On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento dell'avviso avvenuto! "
		Session("IdxSel")=""
		Session("IdxSelPar")=""
		Session("PosPar")=""
	Else
		Response.Write Err.Description
		Err.Number = 0
	End If


mes = ""
IsSuccess = false

'sFrom = Trim(Request.Form("txtFrom"))
sFrom = "Umanet Expo <noreply@iisvittuone.it>"
sSubject = Avviso
'sMailServer = "127.0.0.1"
sMailServer ="mail.iisvittuone.it"
'sBody = Trim(Request.Form("txtBody"))



if  strcomp("on",Request.Form("cbEmail"))=0 then ' se non lo devo registrare per la media dello scrutinio ma solo per la classifica

  QuerySQL="Select CodiceAllievo,Email from Allievi where Id_Classe='"&Id_Classe&"' and Email<>'' and Attivo = 1;"
 set rsTabella=ConnessioneDB.Execute(QuerySQL)
  ' response.write(QuerySQL)
   k=0
  do while not rsTabella.eof
     sBody= Comments
   hash

'linkAvviso=dominio&homesito&"/script/cSocial/ShowMessage.asp?scegli=2&ID="&maxID&"&RCount=0&TParent="&maxID&"&divid="&divid&"&id_classe="&Id_Classe&"&CodiceAllievo="&rsTabella("CodiceAllievo")&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&Classe="&session("Cartella")&"&id_categoria="&id_categoria&"&categoria=Compiti"
linkAvviso=dominio&homesito&"/script/cSocial/ShowMessage.asp?scegli=2&ID="&maxID&"&RCount=0&TParent="&maxID&"&divid="&divid&"&id_classe="&Id_Classe&"&hash="&rsTabella("PasswordSHA256")&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&Classe="&session("Cartella")&"&id_categoria="&id_categoria&"&categoria=Compiti"


linkAvviso2=dominio&homesito&"/script/cSocial/unsubscribe.asp?scegli=2&ID="&maxID&"&RCount=0&TParent="&maxID&"&divid="&divid&"&id_classe="&Id_Classe&"&CodiceAllievo="&rsTabella("CodiceAllievo")&"&by_email=1&DB="&Session("DB")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")&"&Classe="&session("Cartella")&"&id_categoria="&id_categoria&"&categoria=Compiti"


	    sBody = sBody &"  <br><br> <a title 'Vai ad Umanet' href='"& linkAvviso&"'> Entra in Umanet Evolution 3.0</a> <img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' /> "
	   sBody = sBody  &"  <br> <a title='Disiscriviti' href='"& linkAvviso2&"'>Unsubscribe</a>"


	   sTo=rsTabella("Email")
	  ' sTo="mauro.spinarelli@gmail.com"
	  'if k=0 then
	   TestEMail()

        response.write("<br>Inviata mail a " & sTo)
   'end if
	 ' response.write("<br> " &rsTabella("CodiceAllievo"))
	'response.write("<br>Invio email a " &sTo)
  ' response.write("<br>Body" &sBody)
  ' response.write("<br>"&sSubject)
	rsTabella.movenext()
	k=k+1
   loop




end if

'if Request.ServerVariables("HTTP_REFERER") <>"" then
'							response.Redirect request.serverVariables("HTTP_REFERER")
'		 end if

   %>









<div id=piede_pagina>
				<p><p>

				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../cClasse/home_app.asp?id_classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Apprendimento... </a></h3>


			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->


 <script>

$.ajax({
						method: "POST",
						url: "../../../../googleapi/addevento.php?DB=<%=Session("DB")%>&id=<%=maxID%>&idclasse=<%=Session("Id_Classe")%>",
						dataType: "html",
						data: { calendario: "<%=idcal%>", summary: "<%=Avviso%>", description:"Accedi ad Umanet per i dettagli", end: "<%=Scadenza%>" }
					}) /* .ajax */
					.done(function( ans ) {

					//alert("../../../../googleapi/addevento.php?calendario=<%=idcal%>&summary=<%=Avviso%>&description=<%=Azione%>&end=<%=Scadenza%>");

						//alert(ans);
						window.location.href = "<%=Request.ServerVariables("HTTP_REFERER")%>";

					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});

 </script>


	</body>
	</html>
