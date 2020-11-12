<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Sviluppa metafora</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<meta charset="UTF-8">
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	 
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	 
	<!-- jQuery UI -->
    
     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	 

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>

	<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>
 
  <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione e' scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="../../css/jquery-ui.css" />

  


   
</head>

<%
 ' Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" >

  <% end if %>
	<div id="navigation">
     
        <% 
		
 
		 
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set objFSO = CreateObject("Scripting.FileSystemObject")
    			
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo") ' TOPOLINO ED OBIETTIVI
 ' Modulo=Request.QueryString("Modulo")
 ' Cartella=Request.QueryString("Cartella")
  Num = cint(Request.QueryString("Num"))
  Num=Num+1
   daNavigazione=Request.QueryString("daNavigazione")
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  ThreadParent=Request.QueryString("ThreadParent")
  CodiceAllievo=Request.QueryString("CodiceAllievo")
   DATA=Request.QueryString("DATA")
  Collegata=CodiceMetafora
  daSviluppa=1 'serve per la funzione inserisci_metafore che dovrà aggiungere il collegamento del pf 
   
 
  'Sviluppa=Request.QueryString("Sviluppa") ' è settato se sono chiamata da sviluppa metafora devo inserire e linkare con l codice della chiamante
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
						<h1> <i class="icon-comments"></i> Sviluppa metafora </h1> 
                    
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
							<a href="#more-files.html">Libro U-WWW</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Metafore</a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
          <% QuerySQL="Select * from M_Navigazione where CodiceMetafora=" & cint(CodiceMetafora)& ";"
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	 ' response.write(QuerySQL)
	  Cartella=rsTabella.fields("Cartella")
	  Modulo=rsTabella.fields("Id_Mod")
	  Codice_Test=rsTabella.fields("Id_Arg")
	  Autore=rsTabella("Id_Stud")
	  ' response.write("//"&rsTabella("Autista"))
	   SELECT CASE Request.Form("rdSviluppa")
     CASE "1"
       Li=1
      CASE "2"
       Li=2
	  CASE "3"
       Li=3
     CASE "4"
       Li=4
	 CASE "5"
       Li=5
	 CASE "6"
       Li=6
	 CASE "7"
       Li=7
	 CASE "8"
       Li=8
	 CASE "9"
       Li=9
	 CASE "10"
       Li=10
     CASE ELSE
     	Li="Destinazione"
     END SELECT  
	  
%>
       
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%>:&nbsp;<%=Paragrafo%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		 			  <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora</h3>
							</div>
							<div class="box-content" id="storia">
								 
                    		</div>
							<div class="box-content">
								<form name="dati" onSubmit="return validateForm();"  action="inserisci_metafora_patente1.asp?CodiceMetafora=<%=CodiceMetafora%>&prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&daNavigazione=<%=daNavigazione%>&daSviluppa=1&Li=<%=Li%>" method="POST" class="form-vertical">
								 <input type="hidden" value="<%=Cartella%>" id="cartella">
								  <input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
								   <input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
								   <input type="hidden" value="<%=ThreadParent%>" id="ThreadParent">
								   <input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
								      <input type="hidden" value="<%=Modulo%>" id="Modulo">
									   <input type="hidden" value="<%=Paragrafo%>" id="Paragrafo">
									 
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Autista</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="SOGGETTO protagonista" class="input-xxlarge"  name="txtAutista" id="txtAutista" maxlength="148"  value="<%=rsTabella("Autista")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Destinazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtDestinazione" id="txtDestinazione" maxlength="148" value="<%=rsTabella.fields(Li)%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Carburante</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtCarburante" id="txtCarburante" maxlength="148" value="<%=rsTabella.fields("Carburante")& "(?)"%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Luogo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtLuogo" id="txtLuogo" maxlength="148" value="<%=rsTabella.fields("Luogo")& "(?)"%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Comportamento" class="input-xxlarge"   name="txtStrada" id="txtStrada" maxlength="148" value="<%=rsTabella.fields("Strada")& "(?)"%>" >
										</div>
									</div>
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="76" >
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Comportamento adeguato" class="input-xxlarge"  name="txtStrada_ok" id="txtStrada_OK" maxlength="148" value="?" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" checked="true">  
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Comportamento inadeguato" class="input-xxlarge"   name="txtStrada_ko" id="txtStrada_KO" maxlength="148" value="<%=rsTabella.fields("Strada_KO")& "(?)"%>">
										</div>
									</div>
                                     
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cespugli</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Feedback di pericolo" class="input-xxlarge"   name="txtCespugli" id="txtCespugli" maxlength="148" value="<%= rsTabella.fields("Cespugli")& "(?)" %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Lupo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Conseguenze negative" class="input-xxlarge"  name="txtLupo" id="txtLupo" maxlength="148" value="<%= rsTabella.fields("Lupo")& "(?)" %>" >
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cestino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Attaccamenti, eccessi da abbandonare" class="input-xxlarge"  name="txtCestino" maxlength="148" id="txtCestino" value="<%= rsTabella.fields("Cestino")& "(?)" %>">
										</div>
									</div>
                                    
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" placeholder="Num. da 1 a 5" class="input-small"  name="txtDistanza" id="txtDistanza" value="<%=rsTabella.fields("Distanza")%>">
										</div>
									</div>
                                    
                                    
                                    
									 <div class="control-group">
										<label for="textarea" class="control-label">Data</label>
										<div class="controls">
											<input type="text"  oninput="aggiornaStoria();" onfocus="aggiornaStoria();" name="txtDATA" value="<%=DATA%>" class="input-small">
										</div>
									</div>
								 
                                     
                                     
                                     
                                   	<%
	
	 'Prelovo la spiegazione della metafora topolino, che qua andrà estesa 
	     
				url=Server.MapPath(homesito)& "/Db"&Session("DB")& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&CodiceMetafora&".txt" 'per il server on line
				 url=Replace(url,"\","/")
	 		   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
				'sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
   
	%> 
  								<div class="form-actions">
										
										<button type="button" class="btn" onClick="copia_testo();" name="b1">Inizia spiegazione</button>	
                                         
									</div>
 									 <div class="control-group">
										<label for="textarea" class="control-label"><b>Sintesi</b></label>
										<div class="controls">
											<textarea maxlength="910" name="S1" id="textarea" rows="5" class="input-block-level"><%=Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
                                    

									<% if (ucase(session("CodiceAllievo"))=ucase(Autore))  or session("admin")=true then   %>
                                    <div class="form-actions">
										<button type="button" onclick="inserisci_metafore(1)" class="btn btn-primary" name="b1" id="b1">Invia</button>&nbsp;&nbsp;&nbsp;&nbsp;
									</div>
									<%end if%>
                                   
                                    
                                 </form>
                     
                     
                     
                      
                      
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
		
		
         
 <script language="javascript">



// JavaScript Document
function inserisci_metafore(tipo){
var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  ThreadParent=document.getElementById("ThreadParent").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;		  
	 
			txtAutista=document.getElementById("txtAutista").value;
			txtDestinazione=document.getElementById("txtDestinazione").value;
			txtCarburante=document.getElementById("txtCarburante").value;
			txtLuogo=document.getElementById("txtLuogo").value;
			txtStrada=document.getElementById("txtStrada").value;
			txtStrada_OK=document.getElementById("txtStrada_OK").value;
			txtStrada_KO=document.getElementById("txtStrada_KO").value;
			txtCespugli=document.getElementById("txtCespugli").value;
			txtLupo=document.getElementById("txtLupo").value;
			txtCestino=document.getElementById("txtCestino").value;
			txtDistanza=document.getElementById("txtDistanza").value;
			textarea=document.getElementById("textarea").value; 
			
			dati2="&CodiceMetafora="+CodiceMetafora+"&ThreadParent="+ThreadParent+"&txtAutista="+txtAutista+"&txtDestinazione="+txtDestinazione+"&txtCarburante="+txtCarburante+"&txtLuogo="+txtLuogo+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtCespugli="+txtCespugli+"&txtCestino="+txtCestino+"&txtLupo="+txtLupo+"&txtDistanza="+txtDistanza+"&S1="+textarea;
			 
	  
	  
	
	dati="cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&Codice_Test="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo; 
    var url = "7_sviluppa_metafora_ajax.asp?"+dati+dati2;			   
	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
	  if (xhttp.readyState == 4 && xhttp.status == 200) {
						var testo = xhttp.responseText;	
						document.getElementById("b1").disabled=true;
						testoJSON=JSON.parse(testo);
						stato=testoJSON["stato"];
						alert(stato);
						CodiceMetafora=testoJSON["id"];
						if (CodiceMetafora != 0)
							 window.location.href = "sintesi_metafore.asp?cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&CodiceTest="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&CodiceMetafora="+ThreadParent;
		 
			
	  }
	};
	xhttp.open("GET", url, true);
	xhttp.send();	
		
}
	

function validateForm() {
//alert(document.getElementById("txtAutista").value);
  var autista =  document.getElementById("txtAutista").value;
  var destinazione =  document.getElementById("txtDestinazione").value;
  var carburante =  document.getElementById("txtCarburante").value;
  var luogo =  document.getElementById("txtLuogo").value;
  var strada =  document.getElementById("txtStrada").value;
  var stradaok =  document.getElementById("txtStrada_ok").value;
  var stradako =  document.getElementById("txtStrada_ko").value;
  var cespugli =  document.getElementById("txtCespugli").value;
  var lupo =  document.getElementById("txtLupo").value;
  var cestino =  document.getElementById("txtCestino").value;
  var distanza =  document.getElementById("txtDistanza").value;
  
  if (autista == "") {
    alert("Non hai inserito il soggetto");
    return false;
  }
  else if (destinazione == ""){
    alert("Non hai inserito l'obiettivo");
    return false;
  }
  else if (carburante == ""){
    alert("Non hai inserito la motivazione");
    return false;
  }
  else if (luogo == ""){
    alert("Non hai inserito il contesto");
    return false;
  }
  else if (strada == ""){
    alert("Non hai inserito il comportamento");
    return false;
  }
  else if (stradaok == ""){
    alert("Non hai inserito il comportamento ok");
    return false;
  }
  else if (stradako == ""){
    alert("Non hai inserito il comportamento ko");
    return false;
  }
  else if (cespugli == ""){
    alert("Non hai inserito i segnali di pericolo");
    return false;
  }
  else if (lupo == ""){
    alert("Non hai inserito le conseguenze negative");
    return false;
  }
  else if (cestino == ""){
    alert("Non hai inserito il ciò che va abbandonato");
    return false;
  }
  else if ((distanza == "") || isNaN(distanza)) {
    alert("Non hai inserito il numero che indica la distanza dall'obiettivo");
    return false;
  }
}

function copia_testo(){
	var testo;
	//alert(document.dati.txtAutista.value);
	testo=document.dati.txtAutista.value.toUpperCase() + " "+ document.dati.txtDestinazione.value.toUpperCase() + " "+ document.dati.txtCarburante.value.toUpperCase() + " "+document.dati.txtLuogo.value.toUpperCase() + " "+document.dati.txtStrada.value.toUpperCase() + " "+document.dati.txtStrada_ok.value.toUpperCase() + " "+document.dati.txtStrada_ko.value.toUpperCase() + " "+document.dati.txtCespugli.value.toUpperCase() + " "+ " "+document.dati.txtLupo.value.toUpperCase() + " "+ " "+document.dati.txtCestino.value.toUpperCase() + " ";
	document.dati.S1.value=testo;
	 } 
	 

 function aggiornaStoria() {
        var narrazione = "";
		
			//Metafora navigazione
		if (document.dati.txtAutista.value != "")
            Autista = document.getElementById("txtAutista").value.toUpperCase(); //riferimento tramite Id
        else
            Autista = "SOGGETTO";
        if (document.dati.txtDestinazione.value != "")
            Destinazione = document.dati.txtDestinazione.value.toUpperCase(); //riferimento tramite gerarchia DOM
        else
            Destinazione = "OBIETTIVO";
        if (document.dati.txtCarburante.value != "")
            Carburante = document.dati.txtCarburante.value.toUpperCase();
        else
            Carburante = "MOTIVAZIONE";
        if (document.dati.txtLuogo.value != "")
            Luogo = document.dati.txtLuogo.value.toUpperCase();
        else
            Luogo = "SITUAZIONE";
        if (document.dati.txtStrada.value != "")
            Strada = document.dati.txtStrada.value.toUpperCase();
        else
            Strada = "COMPORTAMENTO";
        if (document.dati.txtStrada_OK.value != "")
            Strada_OK = document.dati.txtStrada_OK.value.toUpperCase();
        else
            Strada_OK = "COMPORTAMENTO ADEGUATO";
        if (document.dati.txtStrada_KO.value != "")
            Strada_KO = document.dati.txtStrada_KO.value.toUpperCase();
        else
            Strada_KO = "COMPORTAMENTO INADEGUATO";
        if (document.dati.txtCespugli.value != "")
            Cespugli = document.dati.txtCespugli.value.toUpperCase();
        else
            Cespugli = "FEEDBACK";
        if (document.dati.txtLupo.value != "")
            Lupo = document.dati.txtLupo.value.toUpperCase();
        else
            Lupo = "CONSEGUENZE NEGATIVE";
        if (document.dati.txtCestino.value != "")
            Cestino = document.dati.txtCestino.value.toUpperCase();
        else
            Cestino = "ATTACCAMENTI";

        plurale = Autista.search(/ e /i); //se è presente e oppure E è >0
        plurale1 = Autista.search(","); //faccio mettere ; per indicare il prurale
        if ((plurale == -1) && (plurale1 == -1)) {
            volere = "vuoi";
            raggiungere = "raggiungerai";
            avere = "hai";
            scegliere = "scegli";
            avvicinarsi = "ti avvicina";
            avvicinarsi1 = "avvicinarti";
            avvicinarsi2 = "avvicinarsi";
            allontanarsi = "ti allontana";
            allontanarsi1 = "ti sei allontanato troppo hai";
            scontrarsi = "e ti sei scontrato";
            continuare = "continua";
            fare = "ci sei quasi fai";
            dovere = "devi";
            ti_vi = "ti";
        }
        else {
            volere = "volete";
            raggiungere = "raggiungerete";
            avere = "avete";
            scegliere = "scegliete";
            avvicinarsi = "vi avvicina";
            avvicinarsi2 = "avvicinarsi";
            avvicinarsi1 = "avvicinarvi";
            allontanarsi = "vi allontana";
            allontanarsi1 = "vi siete allontanati troppo avete";
            scontrarsi = "e vi siete scontrati";
            continuare = "continuate";
            fare = "ci siete quasi fate";
            dovere = "dovete";
            ti_vi = "vi";
        }
        narrazione = narrazione + "\n " + Autista + " " + volere + "  raggiungere " + Destinazione + " ?";
        narrazione = narrazione + "NO! <br>\n\n  Mancando " + Carburante.replace("voglia", "") + " per raggiungere " + Destinazione + " , " + Autista + " nel contesto " + Luogo + " non " + raggiungere + " l'obiettivo ! ";
        narrazione = narrazione + "<br>\n\n " + Autista + " " + volere + " raggiungere " + Destinazione + " ?";
        narrazione += "<br>\n\n   " + Autista + "  quale  " + Strada + " " + scegliere + " ?  ";
        narrazione += "<br>\n\nATTENZIONE  " + Autista + "  la scelta  " + Strada_KO + " " + allontanarsi + " da  " + Destinazione;
        narrazione += "<br>\n\n " + Cespugli + " " + ti_vi + " segnalano il pericolo ! ";
        narrazione += "<br>\n\n :-(  " + Autista + "  " + allontanarsi1 + " scelto la strada chiusa  " + Strada_KO + " " + scontrarsi + " con " + Lupo + " \n ";
        narrazione += "<br>\n\n  " + Autista + "  per risolvere la situazione " + dovere + "  abbandonare  " + Cestino + " cosi' da " + avvicinarsi1 + " a " + Destinazione + "  ";
        narrazione += "<br>\n\n  " + Autista + "  quale  " + Strada + " " + scegliere + " ?  ";
        narrazione += "<br>\n\n  " + Autista + "  la scelta  " + Strada_OK + " " + avvicinarsi + " a  " + Destinazione + "  " + continuare + " così !  ";
        narrazione += "<br>\n\n Coraggio " + fare + " l'ultimo passo ! '";
        narrazione += "<br>\n :-) COMPLIMENTI  " + Autista + " " + avere + " raggiunto " + Destinazione + "!!!";
			 
        document.getElementById("storia").innerHTML = narrazione;
    }


</script>
			 
	</body>

 </html>

