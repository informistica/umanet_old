<%@ Language=VBScript %>

<% if session("DB") <> "1" then
	Response.Redirect "../../home.asp"
	end if
	
%>	

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
							<a href="#">Gestione Schermo Nero Simulator</a>
                           
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
				        <h3> <i class="icon-reorder"></i>  NUOVA SESSIONE</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">
							
							<form method="post" action="cSns/inseriscisessione.asp">
								<center>
								<table><tr>
								<td>Nome Sessione:</td><td><input name="nomesessione" id="nomesessione" type="text" class="input-xxlarge"></td></tr>
								<tr><td>Privata (0/1):</td><td><input name="tiposessione" id="tiposessione" type="text" class="input-xxlarge"></td></tr>
								<tr><td>Chiave (se privata):&nbsp;</td><td><input name="passwordsessione" id="passwordsessione" type="text" class="input-xxlarge"></td></tr></table><br>
								<button id="inviosess" type="button" class="btn btn-primary" onClick="controlloinvio()">Invia</button>
								</center>
							</form>
							
							
							</div>
                      
                      </div>
			        </div>
							
							
							
					
					<div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> CHIUDI O MODIFICA SESSIONI</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">					
							
							<% QuerySQL1 = "SELECT * FROM Sessioni_SNS WHERE Aperta = '1'"
							set rsSessioniAperte = ConnessioneDB.Execute(QuerySQL1) %>
							
							<table class="table table-hover table-nomargin">
								<tr>
									<th><b>ID</b></th><th><b>Titolo<b></th><th><b>Data</b></th><th><b>Azione</b></th>
								</tr>
								
								<% do while not rsSessioniAperte.EOF %>
								
								<tr>
								<td><%=rsSessioniAperte("Id_Sessione")%></td>
								<td><a href="classificasns.asp?sessione=<%=rsSessioniAperte("Id_Sessione")%>"><%=rsSessioniAperte("Titolo")%></a></td>
								<td><%=rsSessioniAperte("Data")%></td>
								<td><a href="#modal-1" onClick="modifica(<%=rsSessioniAperte("Id_Sessione")%>, '<%=rsSessioniAperte("Titolo")%>', '<%=rsSessioniAperte("Privata")%>', '<%=rsSessioniAperte("Chiave")%>')" data-toggle="modal"><i style="text-decoration:none" class="icon-pencil" title="Modifica sessione"></i></a>&nbsp;&nbsp;
								<a style="text-decoration:none" href="cSns/chiudisessione.asp?id=<%=rsSessioniAperte("Id_Sessione")%>"> <i class="icon-eye-close" title="Chiudi sessione"></i></a></td>
								</tr>
								
								<% rsSessioniAperte.movenext
								loop %>
								
							</table>
							
							
							</div>
                      
                      </div>
			        </div>
							
						 
							
							
							<div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  RIAPRI O ELIMINA SESSIONI</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">
						
							<center>Attenzione: scegliendo elimina  verranno eliminati  tutti i  risultati della sessione scelta.</center><br>
							
							<% QuerySQL2 = "SELECT * FROM Sessioni_SNS WHERE Aperta = '0'"
							set rsSessioniChiuse = ConnessioneDB.Execute(QuerySQL2) %>
							
							<table class="table table-hover table-nomargin">
								<tr>
									<th><b>ID</b></th><th><b>Titolo</b></th><th><b>Data</b></th><th><b>Azione</b></th>
								</tr>
								
								<% do while not rsSessioniChiuse.EOF %>
								
								<tr>
								<td><%=rsSessioniChiuse("Id_Sessione")%></td>
								<td><a href="classificasns.asp?sessione=<%=rsSessioniChiuse("Id_Sessione")%>"><%=rsSessioniChiuse("Titolo")%></a></td>
								<td><%=rsSessioniChiuse("Data")%></td>
								<td><a style="text-decoration:none" href="cSns/riaprisessione.asp?id=<%=rsSessioniChiuse("Id_Sessione")%>"><i class="icon-eye-open" title="Riapri sessione"></i> </a>&nbsp;&nbsp;
								<a style="text-decoration:none" href="cSns/rimuovisessione.asp?id=<%=rsSessioniChiuse("Id_Sessione")%>"><i class="icon-remove" title="Elimina sessione"></i></a></td>
								</tr>
								
								<% rsSessioniChiuse.movenext
								loop %>
								
							</table>
							
						</div>
                      
                      </div>
			        </div>
			        
			        
			        
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
		
		<form id="mod" action="cSns/modificasessione.asp" method="post">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Modifica Sessione</h3>
				</div>
				<div class="modal-body">
					<b>Titolo</b><br><input name="titolomodifica" id="titolomodifica" type="text" class="input-xxlarge" style="width: 97%"><br>
					<b>Privata (0/1)</b><br><input name="tipomodifica" id="tipomodifica" type="text" class="input-xxlarge" style="width: 97%">
					<br>
					<b>Chiave (se privata)</b><br><input name="chiavemodifica" id="chiavemodifica" type="text" class="input-xxlarge" style="width: 97%"><br><br>
				</div>
				<div class="modal-footer">
					<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
					<button type="button" id="inviamodifica" class="btn btn-primary" onClick="controllamodifica()">Invia</button>
				</div>
			</div>
		</form>
		
		
		
		<script>
		
		function modifica(id, titolo, tipo, chiave){
			document.getElementById("titolomodifica").value=titolo;
			document.getElementById("tipomodifica").value=tipo;
			document.getElementById("chiavemodifica").value=chiave;
			document.getElementById("mod").action="cSns/modificasessione.asp?id="+id;
		}
		
		function controllamodifica(){
			var titolo = document.getElementById("titolomodifica").value.trim();
			var tipo = document.getElementById("tipomodifica").value;
			
			if(titolo == ""){
				alert("Il nome della sessione è obbligatorio");
			}else if(tipo != "0" && tipo != "1"){
				alert("Il tipo può essere 0 oppure 1");
			}else{
				document.getElementById("inviamodifica").type="submit";
			}
			
		}
		
		function controlloinvio(){
			
			var testo = document.getElementById("nomesessione").value;
			var tipo = document.getElementById("tiposessione").value;
			
			if(testo.trim()==""){
				alert("Il nome della sessione è obbligatorio");
			}else if(tipo != "0" && tipo != "1"){
				alert("Il tipo può essere 0 oppure 1");
			}else{
				document.getElementById("inviosess").type="submit";
			}
			
		}
		
		</script>
		
	</body>

 </html>

