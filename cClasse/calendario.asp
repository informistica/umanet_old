<%@ Language=VBScript %>
<html>
<head>
   
   <title>Calendario</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<meta charset="utf-8">
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	
<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- Fullcalendar -->
	<link rel="stylesheet" href="../../css/plugins/fullcalendar/fullcalendar.css">
	<link rel="stylesheet" href="../../css/plugins/fullcalendar/fullcalendar.print.css" media="print">
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
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- FullCalendar -->
	<script src="../../js/plugins/fullcalendar/fullcalendar.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>

	<!--[if lte IE 9]>
		<script src="js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="img/apple-touch-icon-precomposed.png" />

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  <style>
.loader {
display: block;
position: fixed;
left: 0px;
top: 0px;
width: 100%;
height: 100%;
z-index: 9999;
background: #fafafa url(../image/page-loader.gif) no-repeat center center;
text-align: center;
color: #999;
}
</style>
   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" onload="caricaeventi()">  
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
						<h1> <i class="icon-calendar"></i> Calendario </h1> 
                    
					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->	 
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				 
				 
                 
<%  QuerySQL="SELECT * " &_
"FROM Classi WHERE ID_Classe='" & Session("Id_Classe") & "';" 
'response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)  %>
          
                                
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      
				      <div class="box-content nopadding">

                     <% idcal = rsTabella("Url_calendar")
						cartella = rsTabella("Classe")%>
					 
					<div class="calendar"></div>
					  <div class="loader"></div>


                    <% 
					
					QuerySQL = "SELECT * FROM CAT_CAT WHERE Descrizione = 'Compiti' AND Id_Classe = '"&Session("Id_Classe")&"'"
					
					'response.write(QuerySQL)
					
					set rsTabella = ConnessioneDB.Execute(QuerySQL)
					
					IdCatUrl = rsTabella("Id_Categoria")
					
					%>
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
        
		
		<script>
		
		$( document ).ready(function() {
		//una volta caricato il DOM esegui questo codice
		//console.log( "DOM ready!" );
		$(window).load(function() {
		//una volta caricata l'intera pagina (immagini e i frames..) esegui questo codice
		//console.log( "Page ready!" );
		//faccio scomparire l'immagine di caricamento
		//$(".loader").fadeOut("slow");
		});
		});
		
		function caricaeventi(){
					$.ajax({
						method: "POST",
						url: "../../../../googleapi/geteventi.php?id=<%=idcal%>&DB=<%=Session("DB")%>",
						dataType: "html",
						data: { id: "<%=idcal%>" }
					}) /* .ajax */
					.done(function( ans ) {
						
						
						testo = ans.substring(0,ans.length-5);
						//alert(testo);
						
						if($(".calendar").length > 0)
						{
							var date = new Date();
							var d = date.getDate();
							var m = date.getMonth();
							var y = date.getFullYear();
							var url;
							
							
							eventi = testo.split("|");
							
							for(var i=0; i<eventi.length; i++){
							var evento = eventi[i].split(",");
							dataev = new Date(evento[1]);
							idPost=evento[2];
							//url="https://www.umanetexpo.net/expo2015Server/UECDL/script/cSocial/default0.asp?categoria=Compiti&id_categoria=<%=IdCatUrl%>&id_classe=<%=Session("Id_Classe")%>&cartella=<%=cartella%>&scegli=2&ID="+idPost;
							url="https://www.umanetexpo.net/expo2015Server/UECDL/script/cSocial/ShowMessage.asp?categoria=Compiti&id_categoria=<%=IdCatUrl%>&id_classe=<%=Session("Id_Classe")%>&cartella=<%=cartella%>&scegli=2&Zip=0&RCount=1&visibile=1&privato=0&ID="+idPost+"&TParent="+idPost;
							
							 //console.table(evento);
							$('.calendar').fullCalendar('addEventSource', [
								{
									title: evento[0],
									/* start: new Date(dataev.getFullYear(), dataev.getMonth(), dataev.getDate(), dataev.getHours(), dataev.getMinutes()), */ 
									start: new Date(dataev.getFullYear(), dataev.getMonth(), dataev.getDate(), 23, 59),
									allDay: false,
									url: url
								}
							]);
							}
							
							
						}
						
						$(".loader").fadeOut("slow");
						
					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					$(".loader").fadeOut("slow");					
					alert(textStatus+": "+errorThrown);
					});
					
					
					}
					 </script>

			 
	</body>

 </html>

