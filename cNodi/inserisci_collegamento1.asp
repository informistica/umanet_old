<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Collega nodi</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">
    
    


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery UI -->
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.draggable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
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
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
    </script>
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
	
	
    
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

<script language="javascript" type="text/javascript" >
		function myLink(id) {
		var testo = prompt("Inserisci testo del collegamento", "");
		if (testo != null) {
		alert(testo+" "+id);
    }
	return testo;
}
	</script>
  


   
</head>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>'>
  <% end if %>


	<div id="navigation">
     
   <%
   Stato=Request.QueryString("Stato")
  CodiceTest=Request.QueryString("CodiceTest") 
  if CodiceTest<>"" then 
     Session("CodiceTest")=CodiceTest
  end if 
		  Capitolo=Request.QueryString("Capitolo")
		  Paragrafo=Request.QueryString("Paragrafo")
		  Modulo=Request.QueryString("Modulo")
		  Nome=Request.QueryString("Nome")
		  Cognome=Request.QueryString("Cognome")
		  Cartella=Request.QueryString("Cartella")
		  Corso=Request.QueryString("Corso") ' serve per distinguere quando è stato scelto il corso da visualizzare
		  Id_n1=Request.QueryString("Id_n1")   'id del nodo di partenza del link (href che punto all'ancora nel documento)
		  Id_n2=Request.QueryString("Id_n2")  ' 'id del nodo di arrivo del link   (ancora puntata dall'href)
		  L1=Request.QueryString("L1") ' livello del primo nodo da cui parte il link (chi, cosa, dove, ecc...)
		  L2=Request.QueryString("L2")' livello del secondo nodo a cui arriva il link (chi, cosa, dove, ecc...)
		  T2=Request.QueryString("T2") ' testo nel livello di arrivo da visualizzare sull'arco che collega i nodi
		  T2 = Replace(T2, "'",Chr(96))
		 
		  T2 = Replace(T2,chr(133),"a"&Chr(96))
		  T2 = Replace(T2,chr(236),"i"&Chr(96))
		  T2 = Replace(T2,chr(237),"i"&Chr(96))
		  T2 = Replace(T2,chr(242),"o"&Chr(96))
		  T2 = Replace(T2,chr(243),"o"&Chr(96))
		  T2 = Replace(T2,chr(151),"u"&Chr(96))
		  T2 = Replace(T2,chr(250),"u"&Chr(96))
		 T2 = Replace(T2,chr(138),"e"&Chr(96))
		 T2 = Replace(T2,chr(130),"e"&Chr(96))		  
   %>
	
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
						<h1> <i class="glyphicon-link"></i> Collega nodi</h1> 
                    
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
							<a href="#">Verifica</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							 <a href="#">Collega nodi</a> 
                             
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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%> : <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 						<%
						
						   QuerySQL="INSERT INTO Link (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) VALUES (" & clng(Id_n1) & "," &L1 & ", " & clng(Id_n2) & "," & L2 & ",'" & Session("CodiceAllievo")& "','" &T2& "');"						 
						   response.write(QuerySql&"<br>")
						   ConnessioneDB.Execute QuerySQL 	
						   
						   	QuerySQL = "UPDATE Nodi SET NLink = NLink + 1 WHERE CodiceNodo = '"&clng(Id_n1)&"';"
							ConnessioneDB.Execute(QuerySQL)
							QuerySQL = "UPDATE Nodi SET NLink = NLink + 1 WHERE CodiceNodo = '"&clng(Id_n2)&"';"
							ConnessioneDB.Execute(QuerySQL)
						   
							
							QuerySQL = "SELECT * FROM Nodi WHERE CodiceNodo = '"&clng(Id_n1)&"'"
							Set rsTabella = ConnessioneDB.Execute(QuerySQL)
							proprietario = rsTabella("Id_Stud")
							voto = rsTabella("Voto")
							
							if proprietario = Session("CodiceAllievo") and voto < 3 then
								QuerySQL = "UPDATE Nodi SET Voto = "&(voto+1)&" WHERE CodiceNodo = '"&clng(Id_n1)&"';"
								ConnessioneDB.Execute(QuerySQL)
							end if
							
						
						%>	
						<span class="alert-success">
<h5>Inserimento effettuato correttamente, continua collegare...</h5>						
  					</span>						
	 
				<%
				Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo

				%> 
				 
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

