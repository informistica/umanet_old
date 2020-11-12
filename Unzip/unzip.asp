<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Unzip</title>   
   
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
       
   
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<%

homesito=Request.QueryString("homesito")
homeserver=Request.QueryString("homeserver")
id_classe=Request.QueryString("id_classe")
Materia=Request.QueryString("Materia")
Cartella=Request.QueryString("Cartella")
IDPARENT=Request.QueryString("IDPARENT")
ID=Request.QueryString("ID")
CodiceAllievo=Request.QueryString("CodiceAllievo")
CodiceAllievo = Replace(CodiceAllievo," ","%20") ' per evitare errore su url che non ammette spazi 
Social=Request.QueryString("Social")
scegli=Request.QueryString("scegli")
bacheca=Request.QueryString("bacheca")
RCount=Request.QueryString("RCount")

           
                 
%>


<body class='theme-<%=session("stile")%>'>
	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Decompressione archivio</h1> 
                    
					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->	 
                    </div>
				</div>
                
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  Decompressione in corso ... </h3>
			          </div>
				      <div class="box-content">
    
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     <%
					 
					 ' response.write("<br>"& homeserver  ) 
'					    response.write("<br>"&homesito  )
'						response.write("<br>" & Materia  )  
'						response.write("<br>"& Cartella  )
'    					 response.write("<br>"& Social  )  
'						 response.write("<br>"& IDPARENT  )  
'						 response.write("<br>"& CodiceAllievo  )  
'						  response.write("<br>"& Session("Nomefilezip")  )  
'						  response.write("<br>"& Session("NomefilezipSint")  ) 
'						   response.write("<br>"& Session("NomefilezipOrig") )
						 
	 
	    
		 
		 
	 %>            
                      <form method="get" action="unzip.php" target="_self" name="dati">
				 
				<input type="hidden" name="nome" value="<%=Session("Nomefilezip")%>">
                	<input type="hidden" name="nomeSint" value="<%=Session("NomefilezipSint")%>">
                    	<input type="hidden" name="nomeOrig" value="<%=Session("NomefilezipOrig")%>">					
                <br /><input type="hidden" name="CodiceAllievo" value="<%=CodiceAllievo%>">
                <br /> <input type="hidden" name="Cartella" value="<%=Cartella%>">
                <br /> <input type="hidden" name="IDPARENT" value="<%=IDPARENT%>">
                <br /> <input type="hidden" name="Materia" value="<%=Materia%>">
                <br /> <input type="hidden" name="Social" value="<%=Social%>">
                <br /> <input type="hidden" name="homesito" value="<%=homesito%>">
                <br /> <input type="hidden" name="homeserver" value="<%=homeserver%>">
                 
                  <br /> <input type="hidden" name="scegli" value="<%=scegli%>">
                  <br /> <input type="hidden" name="id_classe" value="<%=id_classe%>">
                  <br /> <input type="hidden" name="bacheca" value="<%=bacheca%>">
                  <br /> <input type="hidden" name="ID" value="<%=ID%>">
                  <br /> <input type="hidden" name="RCount" value="<%=RCount%>">
                 
<INPUT class="btn" TYPE="Submit"  VALUE = "Attendere decompressione in corso ..." name="ApplyMessage" id="ApplyMessage"><br><small>Se non vieni reindirizzato controlla che il browser non blocchi l'apertura di popup</small>
</form>
              
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
     <script type="text/javascript">
	
		 
$(window).load(function () {
	   
	   $('#ApplyMessage').click();
	  
	 
	    event.stopPropagation();
	    
	});
	
</script>


 </html>

