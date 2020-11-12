<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Piano di recupero</title>

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
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>

    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

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
	 <!--<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>-->



   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%
 ' 	lettura dei parametri passati alla pagina
 

CodiceAllievo=request.querystring("CodiceAllievo")
cartella=request.querystring("cartella")

		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
     
	 
	</div>
<%
q="select Cognome, Nome from Allievi where CodiceAllievo='"&CodiceAllievo&"'"
set rsStud=ConnessioneDB.execute(q)

%>
	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Argomenti da recuperare per <%=rsStud("Cognome")%>&nbsp;<%=rsStud("Nome")%> </h1>

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
							<a href="#">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Libro</a>
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

				      <div class="box-content">

   <%

   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
 '*******
   QuerySQL="SELECT * FROM MODULI_CLASSE  " &_
 " WHERE Id_Classe='" & session("id_classe")  &"' and Visibile=1 order by Posizione asc"
  Set rsTabellaModuli = ConnessioneDB.Execute(QuerySQL)
  'response.write("* "&QuerySQL&"<br>")
  intestazione=""
  do while not rsTabellaModuli.eof 
	Modulo=rsTabellaModuli("ID_Mod")
	TitoloCapitolo=rsTabellaModuli("Titolo")
		QuerySQL="SELECT * FROM MODULI_PARAGRAFI " &_
	" WHERE ID_Mod='" & Modulo  &"' order by Posizione"
		Set rsTabellaParagrafi = ConnessioneDB.Execute(QuerySQL)
	 ' response.write("** "&QuerySQL&"<br>")
	 	contpar=0
		do while not rsTabellaParagrafi.eof 
		CodiceTest=rsTabellaParagrafi("ID_Paragrafo")
		Paragrafo=rsTabellaParagrafi("Paragrafo")
		'response.write("- Risorsa:<b> "&rsTabellaParagrafi("URL_O")&"</b><br>")
         %>
		<%		
		query="Select TOP(1) Id_Prefrase from Frasi WHERE Frasi.Id_Arg='" & CodiceTest&"' "
	'	 response.write("** "&query&"<br>")
		Set rsTabellaSvolte = ConnessioneDB.Execute(query)
		if not rsTabellaSvolte.eof then
			QuerySQL="SELECT * FROM ParagrafiSottoparagrafi2 " &_
			" WHERE ID_Paragrafo='" & CodiceTest  &"' order by Posizione"
			Set rsTabellaSottoParagrafi = ConnessioneDB.Execute(QuerySQL)
			' response.write("*** "&QuerySQL&"<br>")
			 sottopar=0
			  intestazione=""
			  cont=0
			do while not  rsTabellaSottoParagrafi.eof 
			
				if not rsTabellaSottoParagrafi.eof then
				sottopar=1
				'response.write("<br>"&intestazione)
				'response.write("<br>"& rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")&"/"&rsTabellaSottoParagrafi("Titolo"))
					if StrComp(intestazione, rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")&"/"&rsTabellaSottoParagrafi("Titolo")) = 0 then
					' Response.Write("Le due stringhe sono uguali")
					else 
						'i=0
						  if cont=0 then 
						    intestazione=rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")&"/"&rsTabellaSottoParagrafi("Titolo")
							response.write("<b><h4>"&rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")&"</b></h4>")
							cont=1
						  end if
						'response.write("--- Sottoparagrafo:<b> "&rsTabellaSottoParagrafi("Titolo")&"</b><br>")
						response.write("<b><a href='"&rsTabellaSottoParagrafi("URL")&"' target=blank><i class='icon-cloud'></i> "&rsTabellaSottoParagrafi("Titolo")&"</a></b>")
							%>
					<%end if %>  	

					<%
					CodiceSottopar=rsTabellaSottoParagrafi("ID_Sottoparagrafo")
					QuerySQL="SELECT * " &_
					"FROM preFrasi WHERE preFrasi.Id_Paragrafo='" & CodiceTest & "' and preFrasi.Id_Sottoparagrafo='" & CodiceSottopar & "' and ID_Prefrase not in (Select Id_Prefrase from Frasi WHERE Frasi.Id_Arg='" & CodiceTest & "' and Id_Stud='"&CodiceAllievo&"') and Id_Mod='" &Modulo  & "'"&_
					" order by Posizione;"

				end if
				compiti=1
				Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					%>
					<% if rsTabella.eof and rsTabella.bof then%>
						<div class="alert alert-success">
							<b><%=response.write("Non ci sono compiti assegnati oppure li hai gia' svolti")%></b>
						</div>
						<%end if
					i=0%>
	            <% do while (not rsTabella.eof) and compiti=1

					query="Select Id_Prefrase from Frasi WHERE Frasi.Id_Prefrase='" & rsTabella("ID_Prefrase")&"'"
					Set rsTabellaSvolte1 = ConnessioneDB.Execute(query)
					if not rsTabellaSvolte1.eof then

					 %>
						<% if (i=0) then %>
							<ul style="list-type: square;margin:0px;">
						<%end if %>
							<li style="list-type:square;margin:0px;">
						 <% if rsTabella("img")=1  then
								image1="  <i class='icon-picture' title='richiede immagine'></i>"
							else
								image1=""
							end if %>
							<% if rsTabella("Estesa")=true  then
								image="  <i class='glyphicon-edit' title='testo esteso'></i>"
							else
								image=""
							end if
							image=image&image1
							%>
							<% if ((instr(rsTabella("Quesito"),"tp://")<>0) or (instr(rsTabella("Quesito"),"tps://")<>0)) then%>
							<a title="Apri la risorsa esterna per rispondere alle seguenti domande" href="<%=rsTabella("Quesito")%>" target="_blank"> <i class="icon-cloud"></i> Apri risorsa esterna per rispondere alle seguenti domande </a>
							<%else
							%>
							<a title="Scade il <%=rsTabella("Scadenza")%>" target="_blank" href="2inserisci_frase.asp?estesa=<%=rsTabella.fields("Estesa")%>&by_UECDL=<%=by_UECDL%>&Tipo=0&Quesito=<%=rsTabella.fields("Quesito")%>&Cartella=<%=Cartella%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&prefrase=1&ID_Prefrase=<%=rsTabella("ID_Prefrase")%>&Scadenza=<%=rsTabella("Scadenza")%>&Img=<%=rsTabella("Img")%>&cFile=<%=rsTabella(("Files"))%><% if CodiceSottopar<>"" then%>&CodiceSottopar=<%=rsTabella("Id_Sottoparagrafo")%>&Sottoparagrafo=<%=Sottoparagrafo%><%end if%>"><%=rsTabella("Posizione")%>.<%=image%><%=Server.HTMLEncode(rsTabella("Quesito"))%> </a>
						<% end if %>
							</li>
					</ul>

					<%end if%> 

					 
               	<%   	i=i+1
						rsTabella.movenext
					loop

				
					rsTabellaSottoParagrafi.movenext
					loop


					if sottopar=0 then   ' NON HO SOTTOPARAGRAFI
					'*****
					 
									QuerySQL="SELECT * " &_
								"FROM preFrasi WHERE preFrasi.Id_Paragrafo='" & CodiceTest & "' and ID_Prefrase not in (Select Id_Prefrase from Frasi WHERE Frasi.Id_Arg='" & CodiceTest & "' and Id_Stud='"&CodiceAllievo&"') and Id_Mod='" &Modulo  & "' "&_
								" order by Posizione;"
								
									if StrComp(intestazione, rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")) = 0 then
									' Response.Write("Le due stringhe sono uguali")
									else 
										if contpar=0 then

											intestazione=rsTabellaModuli("Titolo")&"/"&rsTabellaParagrafi("Paragrafo")
											'response.write(QuerySQL&"<br>")
											response.write("<b><br><h4>"&rsTabellaModuli("Titolo")&"</b></h4>")
											contpar=1
										end if
										response.write("<b><h5><a href='"&rsTabellaParagrafi("URL_O")&"' target=blank><i class='icon-cloud'></i> "&rsTabellaParagrafi("Paragrafo")&"</a></b></h5>")
											%>
									<%end if
									Set rsTabella = ConnessioneDB.Execute(QuerySQL)
								%>
								<% if rsTabella.eof and rsTabella.bof then%>
									<div class="alert alert-success">
										<b><%=response.write("Non ci sono compiti assegnati oppure li hai gia' svolti")%></b>
									</div>
									<%end if
								i=0%>
							<% do while (not rsTabella.eof)  

								query="Select Id_Prefrase from Frasi WHERE Frasi.Id_Prefrase='" & rsTabella("ID_Prefrase")&"'"
								'response.write(query&"<br>")
								Set rsTabellaSvolte1 = ConnessioneDB.Execute(query)
								if not rsTabellaSvolte1.eof then

								%>
									<% if (i=0) then %>
										<ul style="list-type: square;margin:0px;">
									<%end if %>
										<li style="list-type:square;margin:0px;">
								  <% if rsTabella("img")=1  then
										image1="  <i class='icon-picture' title='richiede immagine'></i>"
									else
										image1=""
									end if %>
									<% if rsTabella("Estesa")=true  then
										image="  <i class='glyphicon-edit' title='testo esteso'></i>"
									else
										image=""
									end if
										image=image&image1
									%>
										<% if ((instr(rsTabella("Quesito"),"tp://")<>0) or (instr(rsTabella("Quesito"),"tps://")<>0)) then%>
										<a title="Apri la risorsa esterna per rispondere alle seguenti domande" href="<%=rsTabella("Quesito")%>" target="_blank"> <i class="icon-cloud"></i> Apri risorsa esterna per rispondere alle seguenti domande </a>
										<%else
										%>
										<a target="_blank" title="Scade il <%=rsTabella("Scadenza")%>" href="2inserisci_frase.asp?estesa=<%=rsTabella.fields("Estesa")%>&by_UECDL=<%=by_UECDL%>&Tipo=0&Quesito=<%=rsTabella.fields("Quesito")%>&Cartella=<%=Cartella%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&prefrase=1&ID_Prefrase=<%=rsTabella("ID_Prefrase")%>&Scadenza=<%=rsTabella("Scadenza")%>&Img=<%=rsTabella("Img")%>&cFile=<%=rsTabella(("Files"))%><% if CodiceSottopar<>"" then%>&CodiceSottopar=<%=rsTabella("Id_Sottoparagrafo")%>&Sottoparagrafo=<%=Sottoparagrafo%><%end if%>"><%=rsTabella("Posizione")%>.<%=image%><%=Server.HTMLEncode(rsTabella("Quesito"))%> </a>
									<% end if %>
										</li>
								</ul>

								<%end if%> 

								
							<%   	i=i+1
									rsTabella.movenext
								loop

					end if
								 
					'*****

        end if ' 	if not rsTabellaSvolte.eof then
                rsTabellaParagrafi.movenext
                loop 

               rsTabellaModuli.movenext
             loop

else
	response.write("Non puoi visualizzare i dati degli altri studenti")
end if
                
                %>
               <br>
               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6>

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

        <!-- #include file = "../include/colora_pagina.asp" -->

	<% if Session("Scaduto") = true then
	Session("Scaduto") = ""
	Session.Contents.Remove(Session("Scaduto"))
	%>

  <script type="text/javascript" src="../js/refresh_session.js"></script>
	<script>

	$( document ).ready(function() {

	var t = setTimeout(function(){

	alert("Il compito non è aperto in questo momento, chiedi spiegazioni al prof.");

	clearTimeout(t);

	},200);

	});

	</script>

	<% end if %>

	</body>

 </html>
