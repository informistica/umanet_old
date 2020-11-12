<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica Login</title>   
   
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
	
     <script src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>
     
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
   
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

 


   
</head>
<%cla=Request.QueryString("cla")
 
  
    CodiceAllievo = Request.form("txtCodiceAllievo") 
	NewCodiceAllievo = Request.form("txtNewCodiceAllievo")
    PwdAllievo=Request.form("txtPwdAllievo")
	
   ' NewPwdAllievo=Request.form("txtNewPwd")
	'NewPwdAllievo1=Request.form("txtNewPwd1")
	
	 OldPwdAllievo=Request("PwdAllievoSHA256Old")
	 NewPwdAllievo= Request("PwdAllievoSHA256")
	 NewPwdAllievo1=Request("PwdAllievoSHA256")
	
	

	
	%>
<body>

	<div id="navigation">
     
        <% 
		
 	
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
						<h1> <i class="icon-comments"></i> Modifica Login </h1> 
                    
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
							<a href="#more-files.html">Account</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Login</a>
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
				        <h3> <i class="icon-reorder"></i> Modifica Login </h3>
			          </div>
				      <div class="box-content">
                      
 <%
' response.write("a<br>"& txtMessaggio1)
'  response.write("<br>CodiceAllievo="& CodiceAllievo)
' 
'	response.write("<br>NewCodiceAllievo="& NewCodiceAllievo)
'	response.write("<br>NewPwdAllievo="& NewPwdAllievo)
'	response.write("<br>NewPwdAllievo1="& NewPwdAllievo1)
 %>
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
            
            <% 
			 CodiceAllievo1 = Replace(CodiceAllievo, "'", "''")  
			QuerySQL="SELECT Allievi.Cognome,Allievi.Nome,Allievi.Password FROM Allievi WHERE Allievi.CodiceAllievo='" & CodiceAllievo1 & "' and Allievi.PasswordSHA256='"& OldPwdAllievo&"';"
  'Assegna alla variabile il risultato della query prodotta utilizzando il metodo Execute(stringa della query) dell'oggetto connessione
'  response.write(QuerySQL)
  
  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
  
  
  
   If (rsTabella.EOF) Then 
 ' Session("Loggato") = False %>
    <div class="alert alert-error">
  Codice studente inesistente o password in uso errata !!! <br><br>
   <h6 align="center"><a href="#" onClick="javascript:history.back();"> Indietro </a></h6> 
  </div> 
   
  
       <!-- richiama la pagina per reinserire le chiavi d'accesso al data base, è necessario ripassare all'indietro alcuni parametri -->
      <!-- <a href="richiama_test.asp?CodiceCorso=&CodiceTest=<%=CodiceTest%>&TitoloTest=<%=Titolo_Test%>"> <h4></a></h4> -->
 
 
 <%  		 
				  

  Else 'se il codice è corrispondente verifico la corrispondenza della password inserita con quella contenuta nel data base associata a quel codice 
  
if (NewPwdAllievo<>NewPwdAllievo1) or (NewPwdAllievo="") or (NewPwdAllievo1="")  then
 '  Session("Loggato") = False%>
     
              <div class="alert alert-error">
Password e conferma password sono diverse !!! <br><br>
  <h6 align="center"><a href="#" onClick="javascript:history.back();"> Indietro </a></h6> 
  </div> 

 <% 		 
			 			 
else 
 
        
		'Cognome=rsTabella.Fields("Cognome") 
	   ' Nome=rsTabella.Fields("Nome") 
	   ' Response.Cookies("Dati")("Cognome") = rsTabella.Fields("Cognome") 
       ' Response.Cookies("Dati")("Nome") = rsTabella.Fields("Nome") 
	   ' pwdcri= md5(NewPwdAllievo)
		
		QuerySQL ="UPDATE Allievi SET   Allievi.PasswordSHA256 = '" &NewPwdAllievo& "'" &_
		" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo1 & "'"
		
		'response.write(QuerySQL)
	    ConnessioneDB.Execute(QuerySQL)
	
	   ' il codiceallievo se lo cambio devo aggiornare tutte le domande,frasi,nodi,metafore, perchè altrimenti lo studente sparisce dalla classifica
	   
	  '  QuerySQL ="UPDATE Domande  SET Domande.Id_Stud ='" &NewCodiceAllievo& "' WHERE Domande.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Frasi  SET Frasi.Id_Stud ='" &NewCodiceAllievo& "' WHERE Frasi.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Nodi  SET Nodi.Id_Stud ='" &NewCodiceAllievo& "' WHERE Nodi.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE M_Navigazione  SET M_Navigazione.Id_Stud ='" &NewCodiceAllievo& "' WHERE M_Navigazione.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE M_Topolino  SET M_Topolino.Id_Stud ='" &NewCodiceAllievo& "' WHERE M_Topolino.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE 	M_Desideri  SET M_Desideri.Id_Stud ='" &NewCodiceAllievo& "' WHERE 	M_Desideri.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		
'	
'		QuerySQL ="UPDATE Link  SET Link.Id_Stud ='" &NewCodiceAllievo& "' WHERE Link.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE LinkTopolino  SET LinkTopolino.Id_Stud ='" &NewCodiceAllievo& "' WHERE LinkTopolino.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE LinkNavigazione  SET LinkNavigazione.Id_Stud ='" &NewCodiceAllievo& "' WHERE LinkNavigazione.Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Gruppi_composizione  SET Id_Stud ='" &NewCodiceAllievo& "' WHERE Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE 2CREDITI  SET Id_Stud ='" &NewCodiceAllievo& "' WHERE Id_Stud= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Risultati  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Risultati1  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Visualizzazioni  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		QuerySQL ="UPDATE Visualizzazioni1  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)	
'		QuerySQL ="UPDATE 4PERIODI_CLASSIFICA  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)	
'		
'		 QuerySQL ="UPDATE AVVISI  SET CodiceAllievo2='" &NewCodiceAllievo& "' WHERE CodiceAllievo2= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'		 QuerySQL ="UPDATE AVVISI  SET CodiceAllievo='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		'ConnessioneDB.Execute(QuerySQL)
'		
'		
'		QuerySQL ="UPDATE FORUM_MESSAGES  SET CodiceAllievo ='" &NewCodiceAllievo& "' WHERE CodiceAllievo= '"&CodiceAllievo1 & "';"
'		ConnessioneDB.Execute(QuerySQL)
'	
		
	  	%>
<b>
      <br>
   <div class="alert alert-success">
              Password modificata correttamente
           <!--   <h6><a href="../../home.asp"> Torna all'Home Page ed effettua il Login con i nuovi dati... </a></h6> 
      -->
           </div>         
      <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
      <% 'session.Abandon()%>
 <%    end if 		  
    
 End if 
  %>
  
                      
              
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

