<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Promuoviti</title>   
   
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
  id_classe=Request.QueryString("id_classe") 
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  Classe=Request.QueryString("Classe")
  avanti=Request.QueryString("avanti")
 
  
   
		
		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		
		if Classe="" then
  QuerySQL="SELECT Classe FROM Classi where ID_Classe='"&id_classe&"'"
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  Classe=rsTabella("Classe")
  end if
		 
		 QuerySQL ="UPDATE Allievi SET Id_Classe = '" & id_classe & "',  Classe = '" & Classe & "' WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	
	 ConnessioneDB.Execute(QuerySQL)
	'response.write("<br>"&QuerySQL)
		
	if avanti=1 then	
		QuerySQL="SELECT * FROM anni_scolastici where Attivo=1"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					id_as=rsTabella("ID_AS")
					session("id_as")=id_as '  
	 
	 
		QuerySQL="INSERT INTO stud_as_classe (Id_Stud,Id_As,Id_Classe) SELECT '" & CodiceAllievo & "'," &  session("id_as") & ",'" & id_classe & "';"
				 
				' response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL 
	
		'elimino associazioni
		QuerySQL = "DELETE FROM AssociazioniAllievi WHERE CodiceAllievo = '"&CodiceAllievo&"' OR UtenteAssociato = '"&CodiceAllievo&"';"
		ConnessioneDB.Execute(QuerySQL)
		
		if Session("Admin") = false then
		
			QuerySQL = "SELECT * FROM Allievi WHERE CodiceAllievo = '"&CodiceAllievo&"';"
			'response.write QuerySQL
			set rsImg = ConnessioneDB.Execute(QuerySQL)
			
			urlimg = rsImg("Url_img")
			'response.write urlimg
		
					if urlimg <> "" then
					
						'sposto file immagine nella nuova classe
						Dim FileObject
						Set FileObject=CreateObject("Scripting.FileSystemObject")
						
						urlold = "C:/inetpub/umanetroot/expo2015Server/UECDL/Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&session("id_classe_img")&"/Profili/thumb/"&urlimg
						urlnew = "C:/inetpub/umanetroot/expo2015Server/UECDL/Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&classe&"/Profili/thumb/"&urlimg
						
						'response.write urlold & "<br>" & urlnew
						
						FileObject.MoveFile urlold, urlnew
						
						urlold = "C:/inetpub/umanetroot/expo2015Server/UECDL/Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&session("id_classe_img")&"/Profili/img/"&urlimg
						urlnew = "C:/inetpub/umanetroot/expo2015Server/UECDL/Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&classe&"/Profili/img/"&urlimg
						
						FileObject.MoveFile urlold, urlnew

						Set FileObject=Nothing
					
					end if
		
	     end if
		
		'response.write("ok")
		
		
		id=CodiceAllievo ' importate perchè l'include usa id 
		 
 	%>
    
     <!-- #include file = "../include/inizializzaDB.asp" -->  	          
	
	<% end if%>
	</div>  
    
	<div class="container-fluid" id="content">
       
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Aggiornamento </h1> 
                    
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
				      <div class="box-title">
				        <h3>  Promuoviti
                        
                         </h3>
			          </div>
				      <div class="box-content">
                                       
   <%' se la query che preleva i compiti non restituisce risultati
     %>
     
   <div class="alert alert-success">
             <b><%=response.write("Aggiornamento eseguito")%></b><br>
              <b><%=response.write("Effettua il logout e entra nella nuova classe ")%></b>
             
             
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

