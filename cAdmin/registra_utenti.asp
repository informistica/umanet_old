<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Registra </title>   
   
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
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	 <!--<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>-->
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
 ' 	lettura dei parametri passati alla pagina
 function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"a"&Chr(96))
  sAns = Replace(sAns,chr(225),"a"&Chr(96))
  sAns = Replace(sAns,chr(232),"e"&Chr(96))
  sAns = Replace(sAns,chr(233),"e"&Chr(96))
  sAns = Replace(sAns,chr(236),"i"&Chr(96))
  sAns = Replace(sAns,chr(237),"i"&Chr(96))
  sAns = Replace(sAns,chr(242),"o"&Chr(96))
  sAns = Replace(sAns,chr(243),"o"&Chr(96))
  sAns = Replace(sAns,chr(249),"u"&Chr(96))
  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
  sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
  sAns=  Replace(sAns,"&","e") 
  sAns=  Replace(sAns,"/","-") 
  sAns=  Replace(sAns,"\","-") 
  sAns=  Replace(sAns,"?",".") 
  sAns=  Replace(sAns,"*","x") 
  sAns=  Replace(sAns,"<","_")
  sAns=  Replace(sAns,">","_") 
  
ReplaceCar = sAns
end function
'on error resume next
Function password_mista()
  ' Creo la variabile "caratteri" contenente tutti i
  ' numeri da 0 a 9 e tutte le lettere dalla A alla Z
  caratteri = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
  Randomize()
  Do Until len(password) = 4 
    ' Genero un valore casuale compreso tra 1 e 37
    ' dove 1 corrisponde al numero 0 e 37 alla lettera Z 
    carattere = Int((37 * Rnd) + 1)
    ' Aggiorno la variabile "password" usando Mid per individuare
    ' all'interno della stringa "caratteri" il numero o la lettera
    ' che corrisponde al numero memorizzato nella variabile "carattere"
    password = password & Mid(caratteri,carattere,1) 
  Loop 
  password_mista = password
End Function

  txtDomande = Request.Form("MyTextArea")
  id_classe=Request.querystring("id_classe")
  classe=Request.querystring("classe")
  
		
		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo
		
		 	%>	          
	</div>  
    
	<div class="container-fluid" id="content">
       
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Registra Utenti </h1> 
                    
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
							<a href="#"> </a>
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
				        <h3> <i class="icon-reorder"></i>   
                        	 
                        
                         </h3>
			          </div>
				      <div class="box-content">
                                       
 
  
	 				 
		<div class="row-fluid">
		 <div class="span12">
		   <div class="box">     
              <div class="box-content">       
                   
                     <div class="alert alert-success">
                     	<b><%=response.write("Hai gia' svolto tutti i compiti assegnati")%></b>
                     </div>    			
				       
                       <div class="alert alert-error">
                                 <b><%=response.write("Non ci sono compiti assegnati")%></b>
                       </div>
                   
                   
                   
                  <%   
        'strText = MyTextArea.Value
		strText = txtDomande
        arrLines = Split(strText, vbCrLf)
    k=1
	For Each strLine in arrLines
	 cognome=left(strLine,instr(strLine,"-")-1)
	 cognome=ucase(left(cognome,1))&lcase(right(cognome,len(cognome)-1))
	 cognome=ReplaceCar(trim(cognome)) 
	 nome=right(strLine,len(strLine)-instr(strLine,"-"))
	 nome=ucase(left(nome,1))&lcase(right(nome,len(nome)-1))
	 nome=ReplaceCar(trim(nome)) 
	 username=left(trim(Cognome),3)&"."&left(trim(Nome),3)
	 id=username
     password=lcase(password_mista)
		 
		' response.write  username &" " &password  &"<br>"
		 
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
				
				 CodiceAllievo = username
				 Password = password
				 Cognome = cognome
				 Nome = nome
				 Classe = classe
				 Anno="2015-2016"
				 Id_Classe = id_classe
				 In_Quiz = in_quiz
				 Stile = "darkblue"
		
				mipiace=""
				nonmipiace=""
				descriviti=""
				email=""
				QuerySQL="  INSERT INTO Allievi (CodiceAllievo,Nome,Cognome,Password,Classe,Anno,In_Quiz,Id_Classe,Stile)  SELECT '" & CodiceAllievo & "','"&Nome&"','"&Cognome& "','"&Password& "','"&Classe&  "','"&Anno& "',"&In_Quiz& ",'"&Id_Classe&  "','"&Stile&"';"
				
				response.write(QuerySQL&"<br>")
				ConnessioneDB.execute(QuerySQL)
				
				
				
				 
				
				 
				
			 
				
			 
				  ' per gestire in_quiz
	
				QuerySQL ="UPDATE Setting SET In_Quiz = " & cint(in_quiz)+1 & "  WHERE Id_Classe ='" &id_classe &"';"
				' response.write(QuerySQL)
				ConnessioneDB.Execute(QuerySQL)
				
			 
				session("id_as")=2 ' poi farò query per persacer anno attivo		
				 QuerySQL="INSERT INTO stud_as_classe (Id_Stud,Id_As,Id_Classe) SELECT '" & username & "'," &  session("id_as") & ",'" & id_classe & "';"
				 
				 'response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL 
				 
						
						'trasferisco in un file include usato anche da cClasse/promuoviti.asp
			
   
   %>
   <!-- #include file = "../include/inizializzaDB.asp" -->  
		 
	<%	 
		 

       k=k+1
	Next
	 
             %>      
                   
                   
                
               </div> 	
               <br>
               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> 
                            
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

