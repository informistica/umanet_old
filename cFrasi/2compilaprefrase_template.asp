<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script type="text/javascript" src="../../js/google.js"></script><title>Crea Frase</title>

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
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo

' QuerySQL="SELECT Verifica from Paragrafi where ID_Paragrafo='"&CodiceTest&"'"
 'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'Verifica=rsTabella(0) 


	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Crea Frase </h1>

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
                         </h3>
			          </div>
				      <div class="box-content">
							<div class="row-fluid">
								<div class="span12">
									<div class="box">
										<div class="box-content">
										ciao	
 										<ul class="pagestats style-3">
											<li>
                                                <div class="spark">
													<div title="% di Frasi svolte" class="chart" data-percent="20" data-color="#368ee0" data-trackcolor="#d5e7f7">
												20 %
                                                    </div>
												</div>
												<div class="bottom">
                                                
                                                  
													<span class="name">10 su 20</span>
                                                    <span class="name">PF.20 </span>
                                                      </span>
												</div>
											</li>
                                            
										</ul>

										<ul class="pagestats style-3">
											<li>
												<div class="spark">
													<div class="chart easyPieChart" data-percent="73" data-color="#368ee0" data-trackcolor="#d5e7f7" style="width: 80px; height: 80px; line-height: 80px;">43%<canvas width="80" height="80"></canvas></div>
												</div>
												<div class="bottom">
													<span class="name">CPU Usage</span>
												</div>
											</li>
										</ul>

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

 </html>
