<%@ Language=VBScript %>

<% if session("DB") <> "1" then
	Response.Redirect "../../home.asp"
	end if
	
%>	
<% Session.CodePage = 65001 %>

<!doctype html>
<html>
<head>
   
   <title>Approva Frasi RimorchiApp</title>   
	
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
						<h3> <i class="icon-comments"></i> Approva Frasi RimorchiApp </h3> 
                    
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
							<a href="#">Approva Frasi RimorchiApp</a>
                           
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
				        <h3> <i class="icon-reorder"></i>  CONTROLLA FRASI</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">
							
							<center>
							
							<% QuerySQL = "SELECT DISTINCT Categoria FROM RimorchiApp WHERE 1=1 order by Categoria asc"
								set rsTabellaCategorie = ConnessioneDB.Execute(QuerySQL)
								
							do while not rsTabellaCategorie.EOF
							
								QuerySQLC = "SELECT count(*) FROM RimorchiApp WHERE Categoria = '"&rsTabellaCategorie("Categoria")&"' and Approvata = '0';"
								set rsNumero = ConnessioneDB.Execute(QuerySQLC)
								
							%>
							
							<a href="approvarimorchiapp.asp?mostra=<%=rsTabellaCategorie("Categoria")%>"><button class="btn"><b><%=rsTabellaCategorie("Categoria")%></b> (<%=rsNumero(0)%>)</button></a>&nbsp;
							
							<%
							
							rsTabellaCategorie.MoveNext
							Loop
							
							%>	 
							
							<% QuerySQLC = "SELECT count(*) FROM RimorchiApp WHERE 1=1 and Approvata = '0'"
								set rsNumero = ConnessioneDB.Execute(QuerySQLC)
								%>
							
							<a href="approvarimorchiapp.asp"><button class="btn"><b>Tutte</b> (<%=rsNumero(0)%>)</button></a>
							
							</center>
							<br>
							
							<table class="table table-hover table-nomargin">
								<tr>
									<th><b>Testo</b></th><th><b>Categoria<b></th><th><b>Azione</b></th>
								</tr>
								
								<%
								
								mostra = Request.QueryString("mostra")
								
								if mostra <> "" then
									QuerySQL = "SELECT * FROM RimorchiApp WHERE Categoria = '"&mostra&"' and Approvata = '0' order by Categoria asc, Testo asc;"
								else
									QuerySQL = "SELECT * FROM RimorchiApp WHERE 1=1 and Approvata = '0' order by Categoria asc, Testo asc"
								end if
								
								
								Set rsTabella = ConnessioneDB.Execute(QuerySQL)
								
								do while not rsTabella.EOF
								
									'QuerySQLR = "SELECT REPLACE('"&rsTabella("Testo")&"','""','') FROM RimorchiApp WHERE 1=1;"
									'response.write(QuerySQLR)
									'ConnessioneDB.Execute(QuerySQLR)
									
									categ = split(rsTabella("Categoria"), " ")
									
									QuerySQLAut = "SELECT Cognome, Nome FROM Allievi WHERE CodiceAllievo = '"&rsTabella("Autore")&"';"
									'response.write(QuerySQLAut)
									set rsAut = ConnessioneDB.Execute(QuerySQLAut)
									
									autore = trim(rsAut("Cognome"))&" "&trim(left(rsAut("Nome"), 1))&"."
									
								%>
								
								<tr>
									<td><%=Replace(rsTabella("Testo"), Chr(34), "")%></td><td><%=categ(1)%></td><td><a href="#modal-1" onclick="modifica(<%=rsTabella("Id_Frase")%>, '<%=Replace(rsTabella("Testo"), Chr(34), "")%>', '<%=rsTabella("Categoria")%>', '<%=autore%>')" data-toggle="modal"><i style="text-decoration:none" class="icon-pencil"></i></a>&nbsp;&nbsp;<a style="text-decoration:none" href="cRimorchiApp/rimuovifrase.asp?id=<%=rsTabella("Id_Frase")%>&prov=approva"><i class="icon-remove"></i></a></td>
								</tr>
								
								<%
								
								rsTabella.MoveNext
								loop
								
									
								%>	
								
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
		
		<form id="mod" action="cRimorchiApp/approvafrase.asp" method="post">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Modifica Frase</h3>
				</div>
				<div class="modal-body">
					<b>Categoria</b><br>
					<select id="categoriamodifica" name="categoriamodifica" class="input-xxlarge">
						<option value="Frasi Impossibili">Frasi Impossibili</option>
						<option value="Frasi Pessime">Frasi Pessime</option>
						<option value="Frasi Scontate">Frasi Scontate</option>
						<option value="Frasi Carine">Frasi Carine</option>
					</select><br>
					<b>Autore</b><br><input name="autoremodifica" id="autoremodifica" disabled type="text" class="input-xxlarge" style="width: 97%">
					<br><br>
					<b>Frase</b><br><input name="frasemodifica" id="frasemodifica" type="text" class="input-xxlarge" style="width: 97%">
					<br><br>
				</div>
				<div class="modal-footer">
					<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
					<button type="button" id="inviamodifica" class="btn btn-primary" onclick="controllamodifica()">Invia e Approva Frase</button>
				</div>
			</div>
		</form>
		
		
		
		<script>
		
		function modifica(id, frase, categoria, autore){
			document.getElementById("categoriamodifica").value=categoria;
			document.getElementById("frasemodifica").value=frase;
			document.getElementById("autoremodifica").value=autore;
			document.getElementById("mod").action="cRimorchiApp/approvafrase.asp?id="+id;
		}
		
		function controllamodifica(){
			var categoria = document.getElementById("categoriamodifica").value.trim();
			var frase = document.getElementById("frasemodifica").value.trim();
						
			if(categoria == "" || frase == ""){
				alert("Devi compilare tutti i campi!");
				document.getElementById("inviamodifica").type="button";
			}else if(frase.length > 150){
				alert("La lunghezza della frase supera il massimo dei caratteri!");
				document.getElementById("inviamodifica").type="button";
			}else{
				document.getElementById("inviamodifica").type="submit";
			}
			
		}
		
		</script>
		
	</body>

 </html>