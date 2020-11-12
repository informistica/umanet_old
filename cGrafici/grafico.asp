<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script type="text/javascript" src="../../js/google.js"></script><title>Grafici andamenti</title>

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

  <script src="https://cdn.jsdelivr.net/npm/chart.js@2.9.3/dist/Chart.min.js"></script>
   <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
 

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

 <style>
    .chart-container {
      position: relative;
      height: 100vh;
      width: 100%;
      padding:0;
      margin:0;
    }
    
    .container-fluid-graph{
      margin:0;
      padding:0;
    }
  </style>


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

' QuerySQL="SELECT Verifica from Paragrafi where ID_Paragrafo='"&CodiceTest&"'"
 'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'Verifica=rsTabella(0) 


	%>


      <%
      primachiamata=request.querystring("primachiamata")
      if primachiamata<>"" then
    classe=request.querystring("classe")
    session("classe")=classe
    else
     classe=session("classe")
    end if
    classe=session("classe")
    anno="as_1920"
   ' urlreport=protocollo&dominio&homesito& "/grafici/"&anno&"/report&"& classe &".asp"
    urlreport="../../grafici/"&anno&"/report&"& classe &".json"  'json risolve problema caratteri speciali
    urlreport=Replace(urlreport,"\","/")
    
    ' per il debug
    'urlpage=homeserverlocal&Request.ServerVariables("PATH_INFO")&"?classe="&classe
    'per la versione http
     urlpage=protocollo&dominio&Request.ServerVariables("PATH_INFO")&"?"&Request.ServerVariables("QUERY_STRING")
  '  response.write(urlreport)
   ' response.write("<br>"&urlpage)
  
  %>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Feedback consapevolezza </h1>

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
							<a href="#">Andamenti</a>
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
				        <h3> <i class="icon-reorder"></i> Andamenti in classifica <small> (Seleziona il periodo)</small>
                         </h3>
			          </div>
				      <div class="box-content">
							<div class="row-fluid">
								<div class="span12">
									<div class="box">
										<div class="box-content">

                                         <%
   
    
     ' Response.Write("<br>QUERY_STRING"&Request.ServerVariables("QUERY_STRING"))
      'Response.Write("<br>PATH_INFO"&Request.ServerVariables("PATH_INFO"))
     
      'Response.Write("<br>"&url)   
%>
    <div class="container-fluid">
    
    <div id="periodi">
      <form name="formPeriodi"  method="get">
        <select name="periodoInizio" id="periodoInizio">
        </select>
        <select name="periodoFine" id="periodoFine">
        </select>
        <input type="button" class="btn" id="btnAggiorna" value="Invia" onclick="aggiorna();">
      </form>
    </div><hr>
      <div id="perInput">
      Inserisci il numero dello studente, oppure più numeri studente separati da , <br>
      <input type="text" name="studenti" id="studenti" placeholder="n oppure n1,n2,n3,...">
      <button id="invia">Invia</button>
      <div id="tutti">
        <br>
        <button id="mostraTutti">Mostra tutti gli studenti</button>
      </div>
    </div>

    <div id="perIlGrafico" class="chart-container">
      <canvas id="grafico"></canvas>
    </div>
  </div>
  <script src="primaVersione/secondarie.js"></script>
  <script>
  
   function aggiorna(){
            let PI=document.getElementById("periodoInizio").value;
            let PF=document.getElementById("periodoFine").value;
            let urlfetch ="<%=urlpage%>"+"&periodoInizio="+PI+"&periodoFine="+PF;
            document.formPeriodi.action = urlfetch;
			document.formPeriodi.submit();
      }
    
     
    $("#tutti").hide();
    let risultati, intestazione;
    let file;
    //Prendo il file con la memoria
    fetch("<%=urlreport%>")
      .then(d => d.json())
      .then(d => {
        file = d;
        finito()
      })
      .catch(e => console.error(e));

    function finito(alunni = []) {
      risultati = file.risultati;
      intestazione = file.intestazione;
      let periodoInizio = prendiParametro("periodoInizio");
      let periodoFine = prendiParametro("periodoFine");
      let classe = prendiParametro("classe");
      let date = prendiDate(intestazione);
      let inizio = trova(periodoInizio, date);
      let fine = trova(periodoFine, date);
      inserisciDate(date, inizio, fine);
      date = prendiDateN(intestazione, inizio,fine);
      let voti = prendiVoti(risultati,inizio,fine);
      let massimo = calcolaMassimo(voti);
      let nomi = prendiNomi(risultati);
      let datasets = creaDatasets(nomi, voti, alunni);
      resetCanvas(); //Serve per evitare che i grafici si sovrappongano
      creaGrafico(date, datasets, massimo);
      if (alunni.length != 0)
        $("#tutti").show();
      else $("#tutti").hide();
    }

    $("#invia").click(() => {
      let s = $("#studenti").val().split(",");
      //Tolgo tutti i valori che potrebbero dare problemi
      let filtrato = s.filter(value => value.match(/^\d+$/)); //Prendo solo cifre
      filtrato = filtrato.filter(value => value < risultati.length) //Tolgo i numeri non corrispondenti a nessun alunno
      filtrato = filtrato.filter(value => value.length != 0); //Tolgo gli elementi vuoti(nel caso non si inserisca il numero tra una virgola e un'altra)
      if (filtrato.length == 0)
        alert("Nessun valore inserito è valido");
      else if (filtrato.length != s.length)
        alert("I valori non validi sono stati scartati");
      filtrato.sort((a, b) => a - b); //Sorting in ordine crescente
      if (filtrato.length != 0)
        finito(filtrato)
      $("#studenti").val("");
    });
    $("#mostraTutti").click(() => {
      finito();
    })

    function prendiParametro(nomeParametro) {
      let risultato = null,
        momentaneo = [];
      let parametri = location.search.substr(1).split("&");
      for (let i = 0; i < parametri.length; i++) {
        momentaneo = parametri[i].split("=");
        if (momentaneo[0] === nomeParametro)
          risultato = decodeURIComponent(momentaneo[1]);
      }
      return risultato;
    }

    function inserisciDate(d, i, f) {
      $("#periodoInizio").empty();
      $("#periodoFine").empty();
      for (let i = 0; i < d.length; i++) {
        let option = document.createElement("option");
        option.value = d[i];
        option.innerText = d[i];
        $("#periodoInizio").append(option);
        option = document.createElement("option");
        option.value = d[i];
        option.innerText = d[i];
        $("#periodoFine").append(option);
      }

      $("#periodoInizio").children().eq(i).attr("selected", "selected");
      $("#periodoFine").children().eq(f).attr("selected", "selected");
    }

    function trova(e, v) {
      for (let i = 0; i < v.length; i++)
        if (v[i] == e)
          return i;
      return -1;
    }
  </script>

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

        <!-- #include file = "../include/colora_pagina.asp" -->
 
   
 

	</body>
  <script>
   //window.onload = aggiorna;
  </script>

 </html>
