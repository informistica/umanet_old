<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Bilancia Quiz</title>   
   
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
<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
		
      'on error resume next
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
						<h1> <i class="icon-comments"></i> Bilanciamento del Quiz </h1> 
                    
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
				        <h3> <i class="icon-reorder"></i>Test bilanciato </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     
                      <%                      'Lettura dei dati memorizzati nei cookie. 
   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  vf = Request.QueryString("vf") 
  rm = Request.QueryString("rm")
   
  
  
Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
   DataTest = gira_data()
   
   Stato=Request.QueryString("Stato")
   Modulo=Request.QueryString("Modulo")
   CodiceTest=Request.QueryString("CodiceTest") ' se svolgo tutto il modulo contiene l'id del modulo
   'Definizione query SQL per contare il numero di domande del test.
   NUMTEST=Request.QueryString("NUMTEST")

if strcomp(vf,"1")=0 then
	 stringQuery=" VF=1 and "
     else  
	   if strcomp(rm,"1")=0  then 
			stringQuery=" Multiple=1 and "
		else
		   stringQuery="Multiple=0 and VF=0 and "
		 end if
	end if
	
if (Stato=0) then 
 'Definzione codice SQl della query per ricercare le domande del paragrafo 
    if CodiceSottopar<>"" then
	
			 QuerySQL="SELECT count(*) " &_
             "FROM DomandeQuiz WHERE " &_
            stringQuery& " Id_Arg='" & CodiceTest & "'  and Id_Sottoparagrafo='" & CodiceSottopar & "' ;"
	else
   QuerySQL="SELECT count(*) " &_
             "FROM DomandeQuiz Where " &_
             stringQuery&" Id_Arg='" & CodiceTest & "' ;"
    end if
  
  
  
    'Assegna alla variabile il risultato della query prodotta utilizzando il metodo Execute(stringa della query) dell'oggetto connessione
else 
'Definzione codice SQl della query per ricercare le domande del modulo
'QuerySQL="SELECT count(*) " &_
'             "FROM Domande " &_
'             "WHERE Domande.Id_Mod='" & Modulo & "' ;"
			 
			  QuerySQL="SELECT count(*) " &_
             "FROM DomandeQuiz WHERE " &_
             stringQuery&" Id_Mod='" & Modulo & "' ;"
end if   
response.write("<br>"&QuerySQL)
   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    NumDom=rsTabella(0).value 'Assegno a NumDom numero delle domande
	' se devo correggere vero falso
	if  strcomp(vf,"1")=0 then
	 stringQuery=" VF=1 and "
     else  
	     if strcomp(rm,"1")=0 then 
			stringQuery=" Multiple=1 and "
		else
		   stringQuery=" Multiple=0 and VF=0 and "
		 end if
	end if
if (Stato=0) then 
     if CodiceSottopar<>"" then	
			 QuerySQL="SELECT *" &_
             "FROM DomandeQuiz WHERE " &_
            stringQuery& " Id_Arg='" & CodiceTest & "'  and Id_Sottoparagrafo='" & CodiceSottopar & "' order by CodiceDomanda asc;"
	else
			 QuerySQL="SELECT *" &_
             "FROM DomandeQuiz Where " &_
             stringQuery&" Id_Arg='" & CodiceTest & "' order by CodiceDomanda asc;"
    end if
   
   
   'Definizione query SQL per la lettura delle risposte esatta nel test scelto.
   
else
   QuerySQL="SELECT * " &_
             "FROM DomandeQuiz WHERE " &_
             stringQuery&" Id_Mod='" & Modulo & "' order by CodiceDomanda asc;"
end if  
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    response.write(QuerySQL)
  'Calcolo del numero di risposte esatte.  
  
  i=1
 
  Do While not(rsTabella.EOF) 
	  
     
    ' if (Escludi=0) and (rsTabella.fields("In_Quiz")<> clng(NUMTEST)) then ' se non era inclusa la includo 
        
     
		QuerySQL ="UPDATE Domande SET In_Quiz = " & clng(Request.Form("In_Quiz_" & i & "")) & ", Segnalata = " & clng(Request.Form("segnalata" & i & "")) & " WHERE CodiceDomanda =" &rsTabella(0) &";"	
		response.write(QuerySQL & "<br>")
		ConnessioneDB.Execute QuerySQL 
		 
     i=i+1
     rsTabella.MoveNext 			' passa alla prossima domanda
   Loop 
   
 %>
   		
		
			 
		 
			 <span class="alert-danger"> <h4>TEST CORRETTO</h4></span>
		    <h5> <span class="alert-error"> ...</span></h5>
			<h5> <span class="alert-info">  ... </span></h5>
			<h5> <span class="alert-success"> ...</span></h5>
              
			 
          
                      
                      
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

