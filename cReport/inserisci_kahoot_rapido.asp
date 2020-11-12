<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script type="text/javascript" src="../../js/google.js"></script><title>Inserisci punteggi</title>

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


<!-- Easy pie -->
    <link rel="stylesheet" href="../../css/plugins/easy-pie-chart/jquery.easy-pie-chart.css">
	<script src="../../js/plugins/easy-pie-chart/jquery.easy-pie-chart.min.js"></script>

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
 


	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserisci punteggi quiz </h1>

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
				        <h3> <i class="icon-reorder"></i>  <%=response.write(request("txtTitolo"))%>
                         </h3>
			          </div>
				      <div class="box-content">
							<div class="row-fluid">
								<div class="span12">
									<div class="box">
								 


<%

			Titolo=Request("txtTitolo")  
			Data=Request("txtData")     
			Titolo = Replace(Titolo, Chr(34), "'")
			Titolo=  Replace(Titolo,"'",Chr(96))
			TipoVoto="S"			
			Scrutini=1
			Classifica=1
			QuerySQL="INSERT INTO [2ESERCITAZIONI_SINGOLI] (Descrizione,Data,Id_Classe,Scrutini,Classifica,TipoVoto) SELECT '" & Titolo  & "','" & Data & "','" & id_classe & "'," & Scrutini & "," & Classifica & ",'" & TipoVoto & "';"
	
		  ' response.write(QuerySQL)
		   ConnessioneDB.Execute QuerySQL 
		'   response.write("Esercitazione inserita correttamente")
		   
		   'prelevo il codice dell'esercitazione appena inserita
		   QuerySQL="SELECT MAX([ID_Esercitazione]) "&_
		   " FROM [2ESERCITAZIONI_SINGOLI];" 
			Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
			ID_ESER=rsTabella1(0) 
		'	ID_ESER=0

	strText = Request.Form("MyTextArea")
    arrLines = Split(strText, vbCrLf)
    i=0
    j=0
    cont=0
    response.write("<br><h3>Punteggi assegnati:"&ubound(arrLines)&"</h3>")
	consegnato=""
        for a=0 to ubound(arrLines)
                riga=arrLines(a)
				colonne=split(riga,"	")
				codiceallievo=trim(colonne(0))
				pc=trim(colonne(1))
				QuerySQL="select Cognome,Nome from Allievi where  CodiceAllievo='"&codiceallievo &"'"
				set rsTabellaNew= ConnessioneDB.Execute(QuerySQL)
				if not rsTabellaNew.eof then
					QuerySQL="INSERT INTO [2CREDITI] (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & codiceallievo & "','" & pc & "';"
					ConnessioneDB.Execute(QuerySQL)
					response.write("<br>"&rsTabellaNew("Cognome") & " " &left(rsTabellaNew("Nome"),1)&". Pt."&pc&"") 
					consegnato=consegnato&"'"&codiceallievo&"'"&","
				else
					response.write("<br>ATTENZIONE : "&codiceallievo & " non corrisponde a nessun studente della classe...") 
				end if

        next 
		consegnato=left(consegnato,len(consegnato)-1) ' tolgo ,
		QuerySQL="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") order by Cognome;"
		set rsTabellaNew= ConnessioneDB.Execute(QuerySQL)
		response.write("<h3>Assenti :</h3>")
		do while not rsTabellaNew.eof
			response.write("<br><font color='red'> "&rsTabellaNew("Cognome") &" " & left(rsTabellaNew("Nome"),1)&"."&"</font> ")
				pc=0 ' per ogni assente metto 0 punti
				QuerySQL="INSERT INTO [2CREDITI] (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & ID_ESER & "','" & rsTabellaNew("CodiceAllievo") & "','" & pc & "';"
				ConnessioneDB.Execute(QuerySQL)

			rsTabellaNew.movenext
		loop
 

%>	
 										 

										 
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
 
   
 

	</body>

 </html>
