<%@ Language=VBScript %>

<% if session("DB") <> "1" then
	Response.Redirect "../../home.asp"
	end if
	
%>	
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
   
   <title>Gestione Schermo Nero Simulator</title>   
	
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
<meta charset="utf-8">
    
    


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
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />
       
       
      
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta oppure hai cercato di leggere i dati degli altri studenti!")
location.href="../../../../"
//location.href=window.history.back();
 }
    </script>
    <script type="text/javascript" src="../js/selezionatutti.js"></script>
    
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
     
  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") or ((session("CodiceAllievo") <> Request.QueryString("cod")) and (session("admin") = false))then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
	
     <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  <% end if %>

	<script src="../js/jquery.tablesorter.pager.js"></script>
	
	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
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
						<h3> <i class="icon-comments"></i> Schermo Nero Simulator </h3> 
                    
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
							<a href="#">Admin</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Classifica</a>
                           
						</li>
                         
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
                
				<% 
				sessione = Request.QueryString("sessione")
				QuerySQLTitolo = "SELECT Titolo FROM Sessioni_SNS WHERE Id_Sessione = '"&sessione&"';"
				set rsNomeSessione = ConnessioneDB.Execute(QuerySQLTitolo)
				%>
				
				<div class="row-fluid">
				  <div class="span12">
			        
			        <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  SESSIONE: <%=UCase(rsNomeSessione(0))%></h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">
							
							<form method="post" action="cSns/inseriscisessione.asp">
								<center>
								
								
								<% response.write("<table class='table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped '>")
response.write("<thead><th><b>P.</b></th><th><b>Nome</b></th><th><b><center>Best</center></b></th><th><center><b>Last</b></center></th><th><center><b>Totale (Tentativi)</b></center></th></thead>")

QuerySQL = "SELECT CodiceAllievo, MAX(Risultato) as Risultato, SUM(Risultato) as Totale FROM Risultati_SNS WHERE Sessione = '"&sessione&"' GROUP BY CodiceAllievo order by Totale desc, CodiceAllievo asc" 
'response.write(QuerySQL)
set rsTabella = ConnessioneDB.Execute(QuerySQL)
i = 1
do while not rsTabella.EOF
	
	response.write("<tr>")
	
	QuerySQLNome = "SELECT Cognome, Nome FROM Allievi WHERE CodiceAllievo = '"&rsTabella("CodiceAllievo")&"';"
	set rsNome = ConnessioneDB.Execute(QuerySQLNome)
	
	nome = rsNome("Cognome")&" "&left(rsNome("Nome"),1)&"."
	
	QuerySQLData = "SELECT Data FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' and Risultato = '"&rsTabella("Risultato")&"';"
	set rsData = ConnessioneDB.Execute(QuerySQLData)
	datamigliore = rsData(0)
	
	QuerySQLUltimo = "SELECT Data, Risultato FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' AND Data = (SELECT MAX(Data) FROM Risultati_SNS WHERE CodiceAllievo = '"&rsTabella("CodiceAllievo")&"'  and Sessione = '"&sessione&"');"
	set rsUltimo = ConnessioneDB.Execute(QuerySQLUltimo)
	
	dataultimo = rsUltimo("Data")	

	migliore = formattatempo(rsTabella("Risultato"))
	ultimo = formattatempo(rsUltimo("Risultato"))
	
	
	 QuerySQLTotNum = "SELECT count(*) FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"';"
	 set rsTotNum = ConnessioneDB.Execute(QuerySQLTotNum)
	 ntentativi = rsTotNum(0)
	
	' QuerySQLTot = "SELECT Risultato FROM Risultati_SNS WHERE Sessione = '"&sessione&"' and CodiceAllievo = '"&rsTabella("CodiceAllievo")&"';"
	' set rsTot = ConnessioneDB.Execute(QuerySQLTot)
	
	' totale = 0
	
	' do while not rsTot.EOF
	
		' totale = totale + rsTot("Risultato")
	
	' rsTot.movenext
	' loop
	
	response.write("<td>"&i&".</td><td>"&nome&"</td><td><center>"&migliore&" <span style='font-size:11px'>("&left(datamigliore, 19)&")</span></center></td><td><center>"&ultimo&" <span style='font-size:11px'>("&left(dataultimo, 19)&")</center></span></td><td><center>"&formattatempo(rsTabella("Totale"))&" ("&ntentativi&")</center></td>")
	
	response.write("</tr>")
	
	i = i+1
rsTabella.movenext
loop


response.write("</table>")


Function formattatempo(tempoiniz)

	min = 0
	sec = 0
	ore = 0
	
	tempo = tempoiniz
	
	if tempo >= 60 then
		ore = fix(tempoiniz/3600)
		tempoiniz = tempoiniz - 3600*ore
		min = fix(tempoiniz/60)
		tempoiniz = tempoiniz - 60*min
		sec = tempoiniz Mod 60
	else
		sec = tempo
	end if
	
	stampa = ""
	
	if tempo>=3600 then
		stampa = stampa & ore & "<span style='font-size:11px'>h </span> "
	end if
	
	if min < 10 and tempo >= 3600 then
		stampa = stampa & "0"
	end if
	
	if tempo >= 60 then
		stampa = stampa & min & "<span style='font-size:11px'>m </span>"
	end if
	
	if sec < 10 and tempo >=60 then
		stampa = stampa & "0"
	end if
	
	stampa = stampa &  sec & "<span style='font-size:11px'>s</span>"

	formattatempo = stampa
	
End Function

%>		
								</center>
							</form>
							
							
							</div>
                      
                      </div>
			        </div>
				    
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
		
	</body>

 </html>