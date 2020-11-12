<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Risultati verifica</title>

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

 

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

<%


%>

</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo
Titolo=request.querystring("Titolo")
cartella=request.querystring("cartella")
shorturl=request.querystring("shorturl")
paragrafo=request.querystring("paragrafo")
cod=request.querystring("cod")
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/"
urlCorrezione=url&shorturl 
urlCorrezione=Replace(urlCorrezione,"\","/")	
' aggiungere controllo che solo lo stud può vedere i suoi risultati
if (strcomp(cod,session("CodiceAllievo"))<>0) and (strcomp(session("admin"),"true")<>0) then
'response.redirect("https://elexpo.net")
end if

	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Risultati verifica </h1>

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
							<a href="#"><%=response.write(Titolo)%></a>
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
                         </h3>
			          </div>
				      <div class="box-content">
							 
								
								
<%
 



Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
objXMLDoc.async = False    
objXMLDoc.load urlCorrezione
Set domande = objXMLDoc.documentElement.selectNodes("Domanda")   
Set sentiment = objXMLDoc.documentElement.selectNodes("Sentiment") 
  'Set nodeListClient = xmlDom.selectNodes("/result/items/client")
Dim domanda    
i=1
For Each domanda In domande
	set testi = domanda.selectNodes("Testo")
	set modelli= domanda.selectNodes("Modello")
	set corrispondenze = domanda.selectNodes("Corrispondenza")
	
	For Each testo In testi
	  response.write("<b>"&i&") "&testo.text&"</b>")
	   
	Next

	For Each modello In modelli
	  if (strcomp(right(modello.text,3),"(0)") = 0) and (modello.text<>"") then
	     response.write("<br><font color=red>"&left(modello.text, len(modello.text)-3) &"</font>")
	  else
	  	 response.write("<br><font color=green>"&left(modello.text, len(modello.text)-3)&"</font>")
	  end if
	Next
	For Each corrispondenza In corrispondenze
	  response.write("<br><b>Sentiment risposta("&i&") : "&corrispondenza.text&"%</b><br><br>")
	   
	Next
	i=i+1
Next   
For Each x In sentiment
	  response.write("<br><br><b>Sentiment totale : "&x.text&"%</b><br><br>")
Next
%>
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->
 
   
 

	</body>

 </html>
