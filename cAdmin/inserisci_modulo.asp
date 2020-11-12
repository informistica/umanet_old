<!-- richiama_test.asp -->
<%@ Language=VBScript %>

<%  if session("admin") = false then
		response.redirect("../../../../index.html") 
		end if
		%>

<%
 
  'Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag
  'Dim ConnessioneDB , rsTabella,QuerySQL

   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
   ' ********************ATTENZIONE per EXPO crea problemi se inserisco moduli
Session("DB2")=1   
%> 
 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 <%
  Id_Classe=Request.QueryString("Id_Classe")
  divid=request.QueryString("divid")
  
  classe=Request.QueryString("classe")
   cartella=Request.QueryString("cartella")
  
  posizione= Request("posizione")
   umanet= Request("umanet")
 ' response.Write("jhjhk="&posizione)
  Titolo = Request.Form("TxtTitolo")
  Num = Request.Form("TxtNum") ' numero di paragrafi che si vogliono inserire
  ID_Mod=Request.Form("txtID_Mod")
%>
<html>
<head>
   

        
        <!-- #include file = "../var_globali.inc" --> 
  		
		<!-- #include file = "../service/controllo_sessione.asp" -->
       
        <!-- #include file = "../service/formatta_data_LO.asp" -->
       
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
    
    <!-- Datepicker new-->
	<link rel="stylesheet" href="../../css/plugins/datepicker/datepicker.css">
    
    


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
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
-->
  
<!-- Datepicker --> 

<!-- <script src="../js/plugins/datepicker/bootstrap-datepicker.it.js"></script> -->
  
 <title>Inserisci Modulo</title>

</head>
<body class='theme-<%=session("stile")%>'>
        
	<div id="navigation">
		  <!-- #include file = "../include/navigation.asp" -->
        
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
 
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					 
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Configurazione moduli didattici <%=classe%> </h1> 
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
							<a href="#more-files.html">Admin</a>
					 <i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-files.html">Moduli Didattici</a>
					 
						</li>
						 
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				
				<div class="row-fluid">
					 <div class="box">
							<div class="box-title">
								<h3>
									<i class="icon-reorder"></i>
									Inserisci nuovo modulo
								</h3>
								<div class="actions">
									<a href="#" class="btn btn-mini content-refresh"><i class="icon-refresh"></i></a>
									<a href="#" class="btn btn-mini content-remove"><i class="icon-remove"></i></a>
									<a href="#" class="btn btn-mini content-slideUp"><i class="icon-angle-down"></i></a>
								</div>
							</div>
							<div class="box-content">
							
							
							
							
						
<% if num<>"" then %>
<form method="POST" form action="inserisci_modulo1.asp?umanet=<%=umanet%>&Id_Classe=<%=Id_Classe%>&ID_Mod=<%=ID_Mod%>&Titolo=<%=Titolo%>&classe=<%=classe%>&cartella=<%=cartella%>&Num=<%=Num%>&divid=<%=divid%>&posizione=<%=posizione%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
		 
		 <br>
		   <FIELDSET style="width:auto"><LEGEND><b>
           Inserisci i titoli dei paragrafi del modulo :<font color=#FF0000 size="4"> <%Response.write (Titolo & " " & ID_Mod) %> 
           </font></b></LEGEND>
		    ID <input type="text" name="txtCopertina" value="<%=ID_Mod%>_<%=0%>" size="10" maxlength="10" ><br>
            Risorsa del Modulo (pagine)<input type="text" name="txtPg0" value="" size="3" maxlength="5" >
            <input type="text" name="txtPg1" value="" size="3" maxlength="5" >
            <input type="text" name="txtPg2" value="" size="3" maxlength="5" ><br>
			Risorsa del Modulo (url)<input type="text" name="txtURLMod" class="input-text-xxlarge" >
           <br>
			
			
			
			<p>
			  
		 <% for k=1 to Num%>
		 <hr>
		 <b><%=k%>) Paragrafo </b>   <br>
		   <input type="text" name="txtId<%=k%>" value="<%=ID_Mod%>_<%=k%>" size="10" maxlength="10" > ID<br>
         
		  
		   <input type="text" name="txtDomanda<%=k%>">  Titolo<br>
		   <input type="text" name="txtUrl<%=k%>" >Url risorsa<br>
		  
		  
		  	
							<div class="accordion accordion-widget" id="accordionContenitore">
								<div class="accordion-group">
									<div class="accordion-heading">
										<a id="cap<%=k%>" class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordionContenitore" href="#c<%=k%>">
											 <%=k%>) Sottoparagrafi
										</a>
									</div>
									<div id="c<%=k%>" class="accordion-body collapse">
										<div class="accordion-inner">
											 
										  <b>Sotto paragrafi  </b> 
										  <p><textarea class="input-block-level" rows="5" name="txtSottoparagrafi<%=k%>" placeholder="Inserisci i sottoparagrafi uno sotto l'altro"></textarea></p>
										  <br>
										   <b>Url dei sotto paragrafi  </b> 
										  <p><textarea class="input-block-level" rows="5" name="txtUrlSottoparagrafi<%=k%>" placeholder="Inserisci gli url delle risorse uno sotto l'altro"></textarea></p>
										  <br>
										  Pagine
										  <%for j=1 to 18%>
										  <input type="text" name="txtPg<%=k%>_<%=j%>" value="" size="2" maxlength="5" >
										  <%next%><br>
										  </p> 
										</div>
									</div>
								</div>
							</div>

		  
		  
		 
		  
		  
			<%next %>
		  <p><input type="submit" value="Invia" name="B1"><input type="reset" value="Reimposta" name="B2"></p> <!--Definisce i due bottoni del form -->
		</form> <!-- Chiude l'interfaccia -->
		 </fieldset>
		 

<% else%>

 
   <form method="POST"  class='form-horizontal form-striped' form action="inserisci_modulo.asp?umanet=<%=umanet%>&Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&cartella=<%=cartella%>&divid=<%=divid%>&posizione=<%=posizione%>" >

   <div class="control-group">                                        
		<label class="control-label"><b>ID del modulo</b></label>
		
   <%
    if umanet<>"" then
	    QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&cartella&"';"
	  else
	  QuerySQL="SELECT max(posizione) FROM MODULI_NOT_UMANET where Cartella='"&cartella&"';"
	 end if
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1
	'InStrRev([inizio,]stringa1,stringa2[,compara])
	ConnessioneDB.close
	Set ConnessioneDB=nothing
	
	%>     
	<div class="controls"><input type="text" name="txtID_Mod" class="input-large" value="<%=cartella%>_<%=posizione%>" > </div>
	</div>
	

	<div class="control-group">                                        
		<label class="control-label"><b>Titolo del modulo</b></label>
		<div class="controls"><input type="text" name="txtTitolo" value="" class="input-large"> </div>
	</div>
	
	<div class="control-group">                                        
		<label class="control-label"><b>Quanti paragrafi vuoi inserire in questo modulo?</b></label>
		<div class="controls"><input type="text" name="txtNum" value="" class="input-large"> </div>
	</div>
	<br>
    <p><input class="btn btn-primary" type="submit" value="Invia" name="B1"></p>
    </form>
	<hr>
	<div class="box-title">
		<h3>
		<i class="icon-reorder"></i>
			Inserisci nuovo modulo rapido
		</h3>							 
	</div>
	<div class="box-content">
	 <form method="POST" class="form-vertical" name="dati" action="inserisci_modulo1_rapido.asp?umanet=<%=umanet%>&Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&cartella=<%=cartella%>&divid=<%=divid%>&posizione=<%=posizione%>" >
        <div class="control-group">  
			<label class="control-label"><b>ID del modulo</b></label>
			<div class="controls"><input type="text" name="txtID_Mod" value="" class="input-large"> </div>
		</div>
		<div class="control-group">                                        
			<label class="control-label"><b>Titolo del modulo</b></label></div>
			<div class="controls"><input type="text" name="txtTitolo" value="" class="input-large"> </div>
		</div>	
        <span class="alert-info"> 
        	<b>Incolla url e paragrafo (uno su ogni riga separati da uno spazio bianco)</b>  
        </span>
		<div class="control-group"> 
           <textarea name="MyTextArea" rows=8 cols=70 class="input-block-level" placeholder="https://www.w3schools.com/css/css_table.asp Tables    vedi test4.asp" ></textarea> 
		</div>
        <p style ="text-align:center"><input class="btn btn-primary" type="submit" value="Invia" name="B2"  rel="tooltip" title="inserisci in blocco"></p>
     </form>
	</div>
	
<hr>
	<div class="box-title">
		<h3>
		<i class="icon-reorder"></i>
			Inserisci nuovo modulo rapido con sottoparagrafi
		</h3>							 
	</div>
	<div class="box-content">
	 <form method="POST" class="form-vertical" name="dati" action="inserisci_modulo1_rapido1.asp?sottoparagrafi=1&umanet=<%=umanet%>&Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&cartella=<%=cartella%>&divid=<%=divid%>&posizione=<%=posizione%>" >
        <div class="control-group">  
			<label class="control-label"><b>ID del modulo</b></label>
			<div class="controls"><input type="text" name="txtID_Mod" value="" class="input-large"> </div>
		</div>
		<div class="control-group">                                        
			<label class="control-label"><b>Titolo del modulo</b></label></div>
			<div class="controls"><input type="text" name="txtTitolo" value="" class="input-large"> </div>
		</div>	
        <span class="alert-info"> 
        	<b>Incolla &&&nomeparagrafo nomesottoparagrafo url (uno su ogni riga, attenzione al prefisso &&& nel nomeparagrafo)</b>  
        </span>
		<div class="control-group"> 
           <textarea name="MyTextArea" rows=8 cols=70 class="input-block-level" placeholder="vedi test4_modrapid.asp" ></textarea> 
		</div>
        <p style ="text-align:center"><input class="btn btn-primary" type="submit" value="Invia" name="B2"  rel="tooltip" title="inserisci in blocco"></p>
     </form>
	</div>


	<!-- </div> -->
    
    <fieldset><legend>Trasferisci modulo</legend> 
    <!-- Per utilizzare un modulo gi� esistente in altra classe -->
    <div style="overflow:scroll; height:300px;">
    <iframe src="seleziona_origine.asp?umanet=<%=umanet%>" name="postmessage" id="postmessage" width="100%" height="100%" frameborder="0" SCROLLING="yes" border="0" class="iframe"></iframe>
    </div>
  </fieldset>  
  
   <fieldset><legend>Condividi lavoro sul modulo</legend> 
    <!-- Per utilizzare un modulo gi� esistente in altra classe -->
    <div style="overflow:scroll; height:300px;">
    <iframe src="seleziona_origine.asp?umanet=<%=umanet%>&condividi=1" name="postmessage" id="postmessage" width="100%" height="100%" frameborder="0" SCROLLING="yes" border="0" class="iframe"></iframe>
    </div>
  </fieldset>  
  
 <% end if%>   
   
	 
</div>

               <!-- #include file = "../include/colora_pagina.asp" -->

</div>
</div>
</body>
</html>