<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   <meta charset="utf-8">
   <title>Narrazione della Navigazione UWWW</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	 
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	 
	<!-- jQuery UI -->
    
     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	 

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

 <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" >

	<div id="navigation">
     
        <% 
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Narr@zione </h1> 
                    
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
							<a href="#more-login.html">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-files.html">Libro U-WWW</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Metafore</a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                 
                 
                 
                 
              
				 
   <%
   
    CodiceMetafora = Request.QueryString("CodiceMetafora")
  Cartella = Request.QueryString("Cartella")
  Paragrafo = Request.QueryString("Paragrafo")
  QuerySQL="Select * from M_Navigazione where CodiceMetafora="& cint(CodiceMetafora) &";"
  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
  Modulo=rsTabella("Id_Mod")
  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&CodiceMetafora&".txt"
  
  url=Replace(url,"\","/")
  Set objTextFile = objFSO.OpenTextFile(url, ForReading)
  sReadAll = objTextFile.ReadAll
 'sReadAll=url
  objTextFile.Close
  
 
Autista = ucase(rsTabella("Autista"))
Destinazione = ucase(rsTabella("Destinazione"))
Carburante = ucase(rsTabella("Carburante"))
Luogo = ucase(rsTabella("Luogo"))
Strada = ucase(rsTabella("Strada"))
Strada_KO = ucase(rsTabella("Strada_KO"))
Strada_OK = ucase(rsTabella("Strada_OK"))
Cespugli = ucase(rsTabella("Cespugli"))
Lupo = ucase(rsTabella("Lupo"))
Cestino = ucase(rsTabella("Cestino"))
Distanza = ucase(rsTabella("Distanza"))
Sintesi=sReadAll
set rsTabella=nothing
   %>              
                 
                 
                 
                 
              	<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>UMANET EXPLORER </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     <center>
                       <font size="+0">
Il contesto in cui avviene la navigazione riguarda  <%=Luogo%> <br /><br />
Entrati nella mente troviamo il protagonista di questa esplorazione : <%=Autista%> seduto sulla sedia che rappresenta la sua mente. <br />Davanti a sé ha uno schermo in cui può osservare l'andamento della navigazione verso  <%=Destinazione%>. 
<br /> <br>

<iframe width="420" height="315" src="https://www.youtube.com/embed/6MLijXA3W8o?rel=0" frameborder="0" allowfullscreen></iframe><br />
<br />
 
Una prima connessione possibile è quella della coerenza, <br />cioè una situazione  in cui <%=Autista%> sceglie il comportamento <%=Strada_OK%>  soddisfando le aspettative di <%=Luogo%>. <br /><br />
<iframe width="420" height="315" src="https://www.youtube.com/embed/NLGU5lRtfuw?rel=0" frameborder="0" allowfullscreen></iframe><br /><br />
Ma se <%=Luogo%> manifesta <%=Cespugli%> questo è un segnale di pericolo che, <br /> se non adeguatamente gestito, preannuncia una crisi. <br /><br />
<iframe width="420" height="315" src="https://www.youtube.com/embed/4KCBcTbg8wY?rel=0" frameborder="0" allowfullscreen></iframe><br /><br />
Se il segnale <%=Cespugli%> non viene gestito oppure <%=Autista%> sceglie la strada <%=Strada_KO%>  <br />allora  <%=Autista%> andrà incontro a  <%=Lupo%>.<br /><br />
<iframe width="420" height="315" src="https://www.youtube.com/embed/czw1CI1MYac?rel=0" frameborder="0" allowfullscreen></iframe><br />
<br />A questo punto per superare la crisi sarà necessario abbandonare <%=Cestino%>  <br />in modo da riportare la situazione in positivo.<br /><br />
<iframe width="420" height="315" src="https://www.youtube.com/embed/UwcC2ytz3Gg?rel=0" frameborder="0" allowfullscreen></iframe><br />
<br />
<b>Morale della Metafora </b><br /><br />
<% response.write(Sintesi)%>
</font>
                      </center>
               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> 
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

