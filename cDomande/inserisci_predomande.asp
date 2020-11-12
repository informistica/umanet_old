<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci predomande</title>   
   
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

	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
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

  


   
</head>

<%

  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>' data-layout-topbar="fixed">
  <% end if %>


	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
 <%
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
 
TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Paragrafo=Request.QueryString("Paragrafo")
  CodiceTest = Request.QueryString("CodiceTest") 
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome") 
  Cartella=Request.QueryString("Cartella")
  'Scadenza=Request.Form("txtScadenza")
  Scadenza=Request.Form("date3")
  Num = Request.Form("TxtNum") ' numero di domande che si vogliono inserire
  
   by_UECDL=Request.QueryString("by_UECDL")
   Segnalibro=Request.QueryString("Segnalibro")
   BoxApro=Request.QueryString("BoxApro")
  
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserisci domande</h1> 
                    
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
							<a href="#">Crea compito</a>
                            
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
				        <h3> <i class="icon-reorder"></i><%=Capitolo%> : <%=TitoloParagrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				  <% if num<>"" then %>
<form method="POST" name="dati" class="form-vertical" action="inserisci_predomande1.asp?Segnalibro=<%=Segnalibro%>&BoxApro=<%=BoxApro%>&by_UECDL=<%=by_UECDL%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Scadenza=<%=Scadenza%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->

		
		 
		 
			<p>
	 
				 
			
		 <% for k=1 to Num%>
		  <p><input class="input-xxlarge" type="text" name="txtDomanda<%=k%>" value=""  maxlength="150" tabindex="<%=k%>">
		  <b> 
			N.<%=k%> </b>  <b>Img</b>&nbsp;<input type="checkbox" title="Selezionare se è previsto il caricamento di immagini" name="chkImmagine<%=k%>" value="si">&nbsp;<b>File</b>&nbsp;<input type="checkbox" title="Selezionare se è previsto il caricamento di file" name="chkFile<%=k%>" value="si"></p> 
		  <p>
			<%next %>
		  <p><input type="submit" value="Invia" name="B1" class="btn"></p> <!--Definisce i due bottoni del form -->
		</form> <!-- Chiude l'interfaccia -->
		</div>
		</div>
		</div>
<% else%>
  
    <span class="alert-info"><b>
    <% 
       response.write("Quante domande vuoi inserire ?") %></b><br></span>
       <form class="form-vertical" method="POST" form action="inserisci_predomande.asp?Segnalibro=<%=Segnalibro%>&BoxApro=<%=BoxApro%>&by_UECDL=<%=by_UECDL%>&Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceSottopar=<%=CodiceSottopar%>" >
        <p align="center" class="titolo">
        <input class="input-mini" type="text" name="txtNum" size="1"> <br><br>
      
      
        <p align="center"><input type="submit" value="Invia" name="B1" class="btn"></p>
        </form>
        
        <br><hr>
       <form method="POST" class="form-vertical" name="dati" action="inserisci_predomande1.asp?Segnalibro=<%=Segnalibro%>&BoxApro=<%=BoxApro%>&by_UECDL=<%=by_UECDL%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Scadenza=<%=Scadenza%>&CodiceSottopar=<%=CodiceSottopar%>" >
        
       <span class="alert-info"> 
        <b>Oppure incolla l'elenco (una su ogni riga)</b>  
        </span>
        
        
        
        <textarea name="MyTextArea" rows=8 cols=70 class="input-block-level" placeholder="aggiungi di seguito alla domanda il $ se &egrave; previsto il caricamento di immagini, aggiungi il # se &egrave; previsto il caricamento di file " ></textarea> 
        
        
        <p align="center"><b>
		<% response.write("Scadenza ?") %><br> </b>
        <!--<input type="text" name="txtScadenza" size="10" value="gg/mm/aaaa">-->
        <i class="icon-calendar"></i>&nbsp;<b>Data:</b> 
        
        <input type="text" name="date3" id="datepicker" class="input-medium datepick" /></p>
        </p>
      
        <p align="center"><input type="submit" value="Invia" name="B1" class="btn"></p>
        </form>
         
        </p>
         
    
    <% end if%>
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

