<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Crea Metafora</title>

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
  Cartella=Request.QueryString("Cartella")
  TitoloCapitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")

  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")

		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo

Select case CodiceTest
   Case Cartella&"_U_2_3"
   Case Cartella&"_U_2_5" 
	if CodiceSottopar<>"" then
		QuerySQL="SELECT * " &_
"FROM preNavigazione WHERE Id_Paragrafo='" & CodiceTest & "' and Id_Sottoparagrafo='" &CodiceSottopar  & "'and Id_Mod='" &Modulo  & "' order by Posizione;"
		else
		 QuerySQL="SELECT * " &_
"FROM preNavigazione WHERE Id_Paragrafo='" & CodiceTest & "' and Id_Mod='" &Modulo  & "' order by Posizione;"
		end if

   Case Cartella&"_U_2_8"	

end Select

 'QuerySQL="SELECT Verifica from Paragrafi where ID_Paragrafo='"&CodiceTest&"'"
 'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'Verifica=rsTabella(0)
  
		

		'response.write("<br>"&QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 if rsTabella.eof and rsTabella.bof then
 compiti=0
 else
 compiti=1
 end if

 'response.write("compiti="&compiti)
 %>
	 
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Crea Metafora </h1>

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
						<li>
							<a href="#"><%=response.write(TitoloCapitolo)%></a>
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
				        <h3> <i class="icon-reorder"></i>  <%=response.write(Paragrafo)%>
						<%if Verifica=1 then%>
						
							 <a title="Ripassa" href="3visualizza_modello_verifiche.asp?Classe=<%=Classe%>&Cartella=<%=Cartella%>&TitoloModulo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>"><i class="glyphicon-zoom_in"></i> </a>
 								<a title="Esegui test" href="3esegui_verifica_paragrafo.asp?Cartella=<%=Cartella%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>"><i class="icon-pencil"></i> </a>

						<%end if%>

                        
                         </h3>
			          </div>
				      <div class="box-content">

   <%' se la query che preleva i compiti non restituisce risultati
  ' response.write("compiti="&compiti)
    ' if rsTabella.eof and rsTabella.bof then
	if compiti=0 then
	%>
   <div class="alert alert-error">
   
             <b><%=response.write("Non ci sono compiti assegnati")%></b>
   </div>

<%else

   Select Case CodiceTest%>
   <% Case Cartella&"_U_2_3" 'Topolino%>
    <% Case Cartella&"_U_2_5" 'Navigazione%>
<%
	 QuerySQL="SELECT * " &_
	"FROM preNavigazione WHERE  Id_Paragrafo='" & CodiceTest & "' and ID_Premetafora not in (Select Id_Premetafora from M_Navigazione WHERE  Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "' "&_
" order by Posizione;"

%>

	 <% Case Cartella&"_U_2_8" 'Navigazione 
	' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
  end Select

 
	'response.write("<br>???"&QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		 %>

		<div class="row-fluid">
		 <div class="span12">
		   <div class="box">
              <div class="box-content">
                   <% if rsTabella.eof and rsTabella.bof then%>
                     <div class="alert alert-success">
                     	<b><%=response.write("Hai gia' svolto tutti i compiti assegnati")%></b>
                     </div>
					<%end if
                    end if
                i=0%>

	            <% do while (not rsTabella.eof) and compiti=1%>
                   <% if (i=0) then %>
 				     <ul style="list-style-type: none;">
		           <%end if %>
                    <li style="list-type:square">
                   <%image=""%>
                     <% if ((instr(rsTabella("Quesito"),"tp://")<>0) or (instr(rsTabella("Quesito"),"tps://")<>0)) then%>
					 <a title="Apri la risorsa esterna per rispondere alle seguenti domande" href="<%=rsTabella("Quesito")%>" target="_blank"> <i class="icon-cloud"></i> Apri risorsa esterna per rispondere alle seguenti domande </a>
					 <%else
					 %>
				 		<a title="Scade il <%=rsTabella("Scadenza")%>" href="inserisci_metafore.asp?Tipo=0&Quesito=<%=rsTabella.fields("Quesito")%>&id_classe=<%=Session("id_classe")%>&Cartella=<%=Cartella%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&premetafora=1&ID_Premetafora=<%=rsTabella("ID_Premetafora")%>&Scadenza=<%=rsTabella("Scadenza")%>&Img=0&cFile=0"><%=rsTabella("Posizione")%>.<%=image%><%=Server.HTMLEncode(rsTabella("Quesito"))%> </a>
				
				       <% end if %>


                    </li>
				<%
					i=i+1
					rsTabella.movenext
				loop%>
               </ul>
               </div>
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
