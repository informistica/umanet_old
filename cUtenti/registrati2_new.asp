<%@ Language=VBScript %>
<!doctype html>
<html><head>
   
   <title>Registrati</title>   
   
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
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  
   
</head>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim stato
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  stato=Request.QueryString("stato")
%>
 
<!-- #include file = "adovbs.inc" -->
<%
' VENGONO SOSTITUITI GLI APICI (') CON DUE APICI ('')
' PER EVITARE IL PROBLEMA "SQL INJECTION"

nome=Replace(Request("nome"), "'","") 
nome=ucase(left(nome,1))&lcase(right(nome,len(nome)-1))
cognome=Replace(Request("cognome"), "'","")
cognome=ucase(left(cognome,1))&lcase(right(cognome,len(cognome)-1))
username = Replace(Request("username"), "'","")
password = Replace(Request("password"), "'","")
password_conferma = Replace(Request("password_conferma"), "'","")
passwordsha256=Request("passwordsha256")

mipiace = Replace(Request("mipiace"), "'","")
nonmipiace=Request("nonmipiace")
descriviti=Request("descriviti")
classe=Request("classe")
sezione=Request("sezione")
id_classe=Request("id_classe")
email=Request("email")
tag=Request("tag")
if tag="" then 
tag="1920A"
end if
probabilita=30
 
   
'in_quiz=1
' CONTROLLA INNANZITUTTO SE TUTTI I CAMPI SONO STATI COMPILATI
' CORRETTAMENTE

'IF username <> "" and password <> "" and Instr(email, "@") > 0 and Instr(email, ".") > 0 then

 

 
		' PERCORSO DEL DATABASE
		   StringaConnessione = Request.Cookies("Dati")("StrConn")
		 
			Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
		 
			
			%>   
                <!-- #include file = "../var_globali.inc" --> 
			   <!-- #include file = "../stringhe_connessione/stringa_connessione_no_session.inc" -->
               
			<%  
		
		
		Set RecSet = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"'"
		RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
		
		' CONTROLLA SE L'USERNAME INSERITO E' GIA' STATO USATO
		
		IF Not RecSet.Eof Then
		
			' USERNAME GIA' USATO
			' IMPOSTA LA VARIABILE "USATO" SU TRUE
			' (IN MODO DA POTER FAR DOPO UN CONTROLLO IF...)
			
			usato = True
			
			Else
			' ALTRIMENTI ... USERNAME NON USATO
			' IMPOSTA LA VARIABILE "USATO" SU FALSE
			
			usato = False
		End IF
		
		' Chiude la connessione al DB
		
		RecSet.Close
		Set RecSet = Nothing
		
		' FA LA CONDIZIONE PER VERIFICARE SE L'USERNAME
		' IMMESSO E' GIA' STATO USATO...%>
        
        
        
        
        
        
<body class='login'>
	<div id="navigation">
     
        
          	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
     
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Registrazione </h1> 
                    
					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->	 
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
			 
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      
				      <div class="box-content">
                      
 
 							<%	IF usato = True then
		
		' USERNAME GIA' USATO.
		%>
		<hr>
		<div class="alert alert-error">
        Codice allievo inserito gi&agrave; in uso! 
        </div>
		<hr>
		<%
		Else
		
		 
				' NICK NON USATO...
				' PROCEDE ALLA SUA REGISTRAZIONE...
					QuerySQL="Select * from Setting where Id_Classe='" & id_classe & "';" 
					Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
					' se raggiungo il limite ricomncio
					in_quiz=cint(rsTabella("In_Quiz"))
					max_in_quiz=cint(rsTabella("Max_In_Quiz"))
					if (in_quiz=max_in_quiz+1) then
					   in_quiz=1
					end if   
				
				
				Set RecSet = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM Allievi Order By CodiceAllievo Desc"
				RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
				
				RecSet.Addnew
				
				RecSet("CodiceAllievo") = username
				RecSet("Password") = password
				RecSet("PasswordSHA256")=passwordsha256
				RecSet("Cognome") = cognome
				RecSet("Nome") = nome

				RecSet("Classe") = classe
				'RecSet("Sezione") = sezione
				RecSet("Anno")="2015-2016"
				RecSet("Id_Classe") = id_classe
				RecSet("In_Quiz") = in_quiz
				
				RecSet("Mipiace") = mipiace
				RecSet("Nonmipiace") = nonmipiace
				RecSet("Descriviti") = descriviti
				RecSet("Stile") = "blue"
				RecSet("Email") = email
				RecSet("Tags") = tag
				RecSet("Probabilita") = probabilita
				
				
				RecSet.Update
				
				' CHIUDE LA CONNESSIONE AL DB
				RecSet.Close
				Set RecSet = Nothing
				  ' per gestire in_quiz
				 
						 '  dim objFSO,objCreatedFile
							'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
							'	Dim sRead, sReadLine, sReadAll, objTextFile
								'Set objFSO = CreateObject("Scripting.FileSystemObject")
				'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_reg.txt"
				'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'				objCreatedFile.WriteLine(in_quiz)
				'				objCreatedFile.Close
				
				QuerySQL ="UPDATE Setting SET In_Quiz = " & cint(in_quiz)+1 & "  WHERE Id_Classe ='" &id_classe &"';"
					 ConnessioneDB.Execute(QuerySQL)
				
				'if (cint(classe)=6) or (cint(classe)=7) then
						Set RecSet = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"' and Password='"&password&"';"
						RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
						id=RecSet("CodiceAllievo")
						RecSet.Close
						Set RecSet = Nothing
						
				
					QuerySQL="SELECT * FROM anni_scolastici where Attivo=1"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					id_as=rsTabella("ID_AS")
					session("id_as")=id_as '  
					nome_as=rsTabella("Nome")
				
				 QuerySQL="INSERT INTO stud_as_classe (Id_Stud,Id_As,Id_Classe) SELECT '" & username & "'," &  session("id_as") & ",'" & id_classe & "';"
				 
				' response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL 
				 
						
						'trasferisco in un file include usato anche da cClasse/promuoviti.asp
			
   
   %>
   
   <!-- #include file = "../include/inizializzaDB.asp" -->  
				 
				 <div class="alert alert-success">
                 
       				 Registrazione avvenuta con successo ! 
                     <br>
                     <% if Session("DB")=1 then %>
                    <a href="form_login3.asp?id_classe=<%=id_classe%>&cartella=<%=classe%>"> Ora puoi effettuare il Login</a>
                    <% else %>
                     <a href="form_login2.asp?id_classe=<%=id_classe%>&cartella=<%=classe%>"> Ora puoi effettuare il Login</a>
                    <% end if%>
       			 </div>
		<%		'Response.Redirect "form_login.asp?id_classe="&id_classe&"&divid="&divid
				END IF
				%>	 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
     
		              <div class="box-content"> 
                     
                      
                      
               <h6 align="center"><a  onClick="javascript:history.back();"> Indietro </a></h6> 
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
        
      
			 
	</body>

 </html>

