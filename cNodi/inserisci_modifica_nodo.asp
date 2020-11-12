<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica Nodo </title>   
   
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
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>'>
	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
 <% response.Buffer=true
   on error resume next
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
   
   Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  CodiceNodo=Request.QueryString("CodiceNodo")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("Cartella")
  ID=CodiceNodo 
  Chi=Request.QueryString("Chi")
  Cosa=Request.QueryString("Cosa")
  Dove=Request.QueryString("Dove")
  Quando=Request.QueryString("Quando")
  Come=Request.QueryString("Come")
  Perche=Request.QueryString("Perche")
  Quindi=Request.QueryString("Quindi")
  MO=Request.QueryString("MO")
   VAL=Request.QueryString("VAL")
  if MO<>"" then 
 	Modulo=MO
  end if  
 
  'Nome=Request.QueryString("Nome")
  'Cognome=Request.QueryString("Cognome")
   



  

Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

 ''
url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" &  Cartella &"/"&MO&"_Nodi/"&MO&"_"&Paragrafo&"_"&ID&".txt"

url=Replace(url,"\","/")



'response.write(url1)
'response.write(url)


' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)

' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll
'response.write(url)
'response.write("<br>"&url)

'objTextFile.Close

 
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Modifica nodo</h1> 
                    
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
							 <a href="#">Modifica nodo</a>
                          
                             
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
				        <h3> <i class="icon-reorder"></i>   "<%=Capitolo%> : <%=Paragrafo%>"</h3>
			          </div>
				      <div class="box-content">
                      
                      
                      
                      <form method="POST"  class='form-vertical' action="inserisci_modifica_nodo1.asp?Cartella=<%=Cartella%>&CodiceNodo=<%=CodiceNodo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&MO=<%=MO%>" >  
                      
                      
  
  	<div class="control-group">
		<label for="textfield" class="control-label"><b>Chi</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtChi"  value="<%=Chi%>" class="input-xxlarge">
	    	</div>
    </div>
  
  <div class="control-group">
		<label for="textfield" class="control-label"><b>Cosa</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR1Cosa"  value="<%=Cosa%>" class="input-xxlarge">
	    	</div>
    </div>
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Dove</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR2Dove" value="<%=Dove%>" class="input-xxlarge">
	    	</div>
    </div>
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Quando</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR3Quando"  value="<%=Quando%>" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Come</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR4Come"  value="<%=Come%>" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Perch&egrave;</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR5Perche"  value="<%=Perche%>" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Quindi</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtREQuindi"  value="<%=Quindi%>" class="input-xxlarge">
	    	</div>
    </div>
    
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Sintesi</b></label>
	     	<div class="controls">
		     <p><textarea class="input-block-level" rows="6" name="S1" ><%=Response.write(sReadAll)%></textarea></p>
 
	    	</div>
    </div>
    
	 <div class="form-actions">
			<button type="submit" class="btn btn-primary" name="B1">Invia</button>
	</div>
   
 
</form> 
                      
                      
                      
 
 		     
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

