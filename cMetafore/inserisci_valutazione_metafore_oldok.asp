<%@ Language=VBScript %>

<!doctype html>
<html>
<head>
   
     <% 
	 Cartella=Request.QueryString("Cartella")
	  Codice_Test=Request.QueryString("CodiceTest")
	  
	  
	   If Request.QueryString("damodifica")<>"" Then
		 response.write("<script>alert('Modifica della metafora effettuata correttamente'); </script>")

	  end if

	 Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
 <title>Topolino</title>
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 <title>Navigazione</title>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Client/Server</title>
							<%End Select%>
  
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">
    
       <!-- Le styles -->
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap.css" rel="stylesheet">
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap-responsive.css" rel="stylesheet">
    
    <link href="../../../guida/css/pageguide.css" rel="stylesheet">
   <link rel="stylesheet" href="../../css/style.css">
    
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
	
    
   
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!");

location.href="classifica.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"

//location.href=window.history.back();
 }
 </script>
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
immagini = new Array(2); 
immagini[0]="../img/clienteosteno.jpg";
immagini[1]="../img/clienteostesi.jpg";
function scambia_n(cont) {
	if (cont = 1) { alert (" Attenzione per simulare il terremoto devi inserire un evento paradossale che abbia significati in contrasto, la risposta del client deve deludere l'aspettativa del server!");
	document["rappresentazione"].src = immagini[dati.txtTipoEvento.value];
	}else{
	document["rappresentazione"].src = immagini[dati.txtTipoEvento.value];
	}
	
}

 </script>
<script type="text/javascript" src="../js/selezionatutticheckbox.js"></script>
<script type="text/javascript" src="../js/deselezionatutticheckbox.js"></script>
 


   
</head>
<% 
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    
    <body  class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" >

  <% end if %>

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  
'  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
      'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBCopiatestonline")

	'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=" & Server.MapPath("../database/Copiaditestonline.mdb")

%>
     

<%

' homesito="/expo2015Server/UECDL"   
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceAllievo=Request.QueryString("cod")
  if CodiceAllievo="" then
    CodiceAllievo=Request.QueryString("CodiceAllievo")
  end if
  cla=Request.QueryString("cla")
 
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  if Paragrafo="" then
   Paragrafo=TitoloParagrafo
  end if
  DaInserimento=Request.QueryString("DaInserimento") ' vale 1 se sono chiamata dopo inserisci_metafora1, anzichè da studente_domande, in tal caso devo fare la query per reecuperare i parametri.
  
  
  Modulo=Request.QueryString("Modulo")
  if Modulo="" then
  Modulo=session("Modulo")
  end if
  MO=Request.QueryString("MO")
  VAL=Request.QueryString("VAL")
  URL=Request.QueryString("URL")
  DATA=cdate(Request.QueryString("DATA"))
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  ID=CodiceMetafora 
  
  Segnalata= Request.QueryString("Segnalata")

 'response.write(Cartella&"_U_3_3")
' response.write("<br>Allora"&Codice_Test)
  
 
 Select Case Codice_Test
	Case Cartella&"_U_2_3" 
	QuerySQL="Select * from M_Topolino where CodiceMetafora=" & cint(CodiceMetafora)&";"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	DATA=rsTabella("Data")
	response.write(QuerySQL)
	
	 if  Request.QueryString("Topolino")<>"" then
	  Topolino=Request.QueryString("Topolino")
	  Formaggio=Request.QueryString("Formaggio")
	  Fame=Request.QueryString("Fame")
	  Labirinto=Request.QueryString("Labirinto")
	  Strada=Request.QueryString("Strada")
	  Strada_OK=Request.QueryString("Strada_OK")
	  Strada_KO=Request.QueryString("Strada_KO")
	  Testata=Request.QueryString("Testata")
	  Distanza=Request.QueryString("Distanza")
	  
	 else
	  Topolino=rsTabella("Topolino")
	  Formaggio=rsTabella("Formaggio")
	  Fame=rsTabella("Fame")
	  Labirinto=rsTabella("Labirinto")
	  Strada=rsTabella("Strada")
	  Strada_OK=rsTabella("Strada_OK")
	  Strada_KO=rsTabella("Strada_KO")
	  Testata=rsTabella("Testata")
	  Distanza=rsTabella("Distanza")
	  
	 end if 
	 Pi=rsTabella("Pi") ' codice della metafora precedente
	 Pf=rsTabella("Pf") ' ' codice della metafora seguente
     rsTabella.close
 
  Case Cartella&"_U_2_5"
     
	 QuerySQL="Select * from M_Navigazione where CodiceMetafora=" & cint(CodiceMetafora)&";"
	 Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	 DATA=rsTabella("Data")
   if  Request.QueryString("Autista")<>"" then
	  Autista=Request.QueryString("Autista")
	  Destinazione=Request.QueryString("Destinazione")
	  Carburante=Request.QueryString("Carburante")
	  Luogo=Request.QueryString("Luogo")
	  Strada=Request.QueryString("Strada")
	  Strada_OK=Request.QueryString("Strada_OK")
	  Strada_KO=Request.QueryString("Strada_KO")
	  Lupo=Request.QueryString("Lupo")
	  Cestino=Request.QueryString("Cestino")
	  Cespugli=Request.QueryString("Cespugli")
	  Distanza=Request.QueryString("Distanza")
	 ' Paragrafo=Request.QueryString("TitoloParagrafo")
 	else
	  Autista=rsTabella("Autista")
	  Destinazione=rsTabella("Destinazione")
	  Carburante=rsTabella("Carburante")
	  Luogo=rsTabella("Luogo")
	  Strada=rsTabella("Strada")
	  Strada_OK=rsTabella("Strada_OK")
	  Strada_KO=rsTabella("Strada_KO")
	  Lupo=rsTabella("Lupo")
	  Cestino=rsTabella("Cestino")
	  Cespugli=rsTabella("Cespugli")
	  Distanza=rsTabella("Distanza") 
	end if
	 Pi=rsTabella("Pi") ' codice della metafora precedente
	 Pf=rsTabella("Pf") ' ' codice della metafora seguente
	rsTabella.close
	
Case Cartella&"_U_2_8" ' dbdesideri
	 '  response.write("cioo")
	   
	    QuerySQL="Select * from M_Desideri where CodiceMetafora=" & cint(CodiceMetafora)&";"
	 Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	 DATA=rsTabella("Data")
   if  Request.QueryString("SoggettoC")<>"" then
	  SoggettoC=Request.QueryString("SoggettoC")
	  DomandaC=Request.QueryString("DomandaC")
	  MotivazioneC=Request.QueryString("MotivazioneC")
	  DesiderioC=Request.QueryString("DesiderioC")
	  BisognoC=Request.QueryString("BisognoC")
	  SoggettoS=Request.QueryString("SoggettoS")
	  RispostaS=Request.QueryString("RispostaS")
	  MotivazioneS=Request.QueryString("MotivazioneS")
	  DesiderioS=Request.QueryString("DesiderioS")
	  BisognoS=Request.QueryString("BisognoS")
	  TipoEvento=cint(Request.QueryString("TipoEvento"))
	 
	  TolleranzaC=Request.QueryString("TolleranzaC")
	 
 	else
	  SoggettoC=rsTabella("SoggettoC")
	  DomandaC=rsTabella("DomandaC")
	  MotivazioneC=rsTabella("MotivazioneC")
	  DesiderioC=rsTabella("DesiderioC")
	  BisognoC=rsTabella("BisognoC")
	  SoggettoS=rsTabella("SoggettoS")
	  RispostaS=rsTabella("RispostaS")
	  MotivazioneS=rsTabella("MotivazioneS")
	  DesiderioS=rsTabella("DesiderioS")
	  BisognoS=rsTabella("BisognoS")
	  TipoEvento=cint(rsTabella("TipoEvento"))
	  TolleranzaC=rsTabella("TolleranzaC")
	 

	end if
	 Pi=rsTabella("Pi") ' codice della metafora precedente
	 Pf=rsTabella("Pf") ' ' codice della metafora seguente
	rsTabella.close
	   
	   
	   
End Select

 
if MO<>"" then 
 Modulo=MO
end if  
QuerySQLApp=QuerySQL ' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close

  
if (1=1) OR (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' 
'else è alla fine

Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

  
   
    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url=Replace(url,"\","/")
 
 

 
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url5="C:\Inetpub\umanetroot\expo2015\logMetaf.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url5, True)
'				objCreatedFile.WriteLine("Pf="&Pf)
'				objCreatedFile.Close
'



' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)

' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll
'sReadAll=url
response.write(url)
objTextFile.Close

 

   
%>
	<div id="navigation">
      
		
 
	
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Valuta o modifica metafora </h1> 
                    
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
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%>:&nbsp;<%=Paragrafo%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 		<% 'response.write("pi="&Pi)
 'response.write("<br>"&Codice_Test)	
'response.write("DBQ=" & Server.MapPath("../database/Copiaditestonline.mdb"))
 
 %>						 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
               
						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora N.(<%=CodiceMetafora%>)</h3>
							</div>
							<div class="box-content">
								<form name="dati" action="inserisci_modifica_metafora1.asp?davalutazione=1&VALORE=<%=VAL%>&Cartella=<%=Cartella%>&cla=<%=cla%>&cod=<%=CodiceAllievo%>&CodiceAllievo=<%=CodiceAllievo%>&CodiceMetafora=<%=CodiceMetafora%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&MO=<%=MO%>" method="POST" class="form-vertical">
							
                                 
                                
                                <% Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
                              
                                <fieldset id="Parametri">
                                
                                  <div class="control-group">
										<label for="textfield" class="control-label"><b>Topolino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" placeholder="Soggetto protagonista" class="input-xxlarge"  name="txtTopolino"  value="<%=Topolino%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Formaggio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtR1Formaggio"  value="<%=Formaggio%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Fame</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtR2Fame"  value="<%=Fame%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Labirinto</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtR3Labirinto"  value="<%=Labirinto%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" placeholder="Obiettivo" class="input-xxlarge"  name="txtR4Strada"  value="<%=Strada%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6" checked="true">  
											<input type="text" placeholder="Strategia vincente" class="input-xxlarge"  name="txtR5Strada_OK"  value="<%=Strada_OK%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" >
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge"  name="txtREStrada_KO"  value="<%=Strada_KO%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Testata</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Conseguenze della strategia perdente" class="input-xxlarge"  name="txtRETestata"  value="<%=Testata%>">
										</div>
									</div>
                                       </fieldset>
                                       
                                     <span id="idDistanza">
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Num. da 1 a 5" class="input-small"  name="txtREDistanza"  value="<%=Distanza%>">
										</div>
									</div>
                                 </span>
                              	
                                    
                                    <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea name="S1" id="textarea" rows="5" class="input-block-level"><%=Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
                                    
								
								
                                 <div class="accordion" id="accordion3">
									<div class="accordion-group">      
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseMail"><center>
												
                                                <i class="icon-edit" title="Sviluppa"></i>
                                                </center>
											</a>
										</div>
										<!--<div id="collapseMail" class="accordion-body collapse">-->
											<div id="collapseMail" class="accordion-body">
                                            <div class="accordion-inner">
 <center>
 <a title="Esegui simulazione interattiva" href="6_simula_metafora_topolino.asp?CodiceMetafora=<%=CodiceMetafora%>">Simula</a> 
 <br> <br>
<a title="Sviluppa narrazione multimediale"  onClick="SviluppaMetaforaTopolino();">Sviluppa</a> <br><br>
  <%if session("admin")=true then%>
  <br> Solo admin<br>
<a title="Interpreta nella  Metafora della Navigazione" href="inserisci_metafore.asp?CodiceMetafora=<%=CodiceMetafora%>&Modulo=<%=Modulo%>&CodiceTest=<%=Cartella%>_U_2_5&Capitolo=Interfaccia UWWW&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&daTopolino=1" >Invia a Patente</a> <br> <br>
<%end if%>
<span id="btnSxDx">
<input type="button" class="btn" name="indietro" value="<< Indietro " onClick="Precedente()" title="Zoom indietro">
<input type="button" class="btn" name="avanti" value="Avanti >> " onClick="Successiva()" title="Zoom avanti"> 
</span>
<input type="hidden"  name="Pi" value="<%=Pi%>">
<input type="hidden"   name="Pf" value="<%=Pf%>">
										   </center>
                     						 </div>                       
										</div>
                                     </div>  
                                     <%if (session("Admin")=true) then %>
                                     
                                      <div class="control-group">
										<label for="textarea" class="control-label">Data</label>
										<div class="controls">
											<input type="text" name="txtDATA" value="<%=DATA%>" class="input-small">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                     
                             
								<% else 
                                   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
                                     <%if session("admin")=true then%>
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input disabled="disabled" type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
									<%end if%>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT disabled="disabled"  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                <% end if %>
                                 
                                <%end if %>
 
                                    
                                    <div class="form-actions">
										<button type="submit" class="btn btn-primary" name="b1">Aggiorna</button>
								 
									</div>
								
                                <%  Case Cartella&"_U_2_5" ' metafora  METAFORA NAVIGAZIONE%>
                                
                                
                                <fieldset id="Parametri">
                                
                                <div class="control-group">
										<label for="textfield" class="control-label"><b>Autista</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" placeholder="Soggetto protagonista" class="input-xxlarge"  name="txtAutista<%=i%>"  value="<%= Autista %>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Destinazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtR1Destinazione<%=i%>" value="<%=Destinazione%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Carburante</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtR1Carburante<%=i%>" value="<%=Carburante%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Luogo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtR1Luogo<%=i%>" value="<%=Luogo%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" placeholder="Obiettivo" class="input-xxlarge"   name="txtR1Strada<%=i%>" value="<%=Strada%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6" checked="true">  
											<input type="text" placeholder="Strategia vincente" class="input-xxlarge"   name="txtR1Strada_OK<%=i%>" value="<%=Strada_OK%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" >
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge"  name="txtR1Strada_KO<%=i%>" value="<%=Strada_KO%>" >
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cespugli</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge"   name="txtR1Cespugli<%=i%>" value="<%= Cespugli %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Lupo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" placeholder="Strategia perdente" class="input-xxlarge"  name="txtR1Lupo<%=i%>" value="<%= Lupo %>" >
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cestino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" placeholder="Conseguenze della strategia perdente" class="input-xxlarge"  name="txtR1Cestino<%=i%>" value="<%= Cestino %>">
										</div>
									</div>
                                    
                                        </fieldset>
                                       
                                     <span id="idDistanza">
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" placeholder="Num. da 1 a 5" class="input-small"  name="txtREDistanza"  value="<%=Distanza%>">
										</div>
									</div>
                                 </span>  
                                    <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea name="S1" id="textarea" rows="5" class="input-block-level"><%=Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
                                    
									
								
                                 <div class="accordion" id="accordion4">
									<div class="accordion-group">      
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion4" href="#collapseMail1"><center>
												
                                                <i class="icon-edit" title="Sviluppa"></i>
                                                </center>
											</a>
										</div>
										<div id="collapseMail1" class="accordion-body">
											<div class="accordion-inner">
 <center>
 <a title="Esegui simulazione interattiva" href="6_simula_metafora_navigazione.asp?CodiceMetafora=<%=CodiceMetafora%>">Simula</a> 
 <br> <br>
 <a title="Esegui narrazione multimediale" href="6_narrazione_metafora_navigazione.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Paragrafo=<%=Paragrafo%>">Narrazione</a>
  <br> <br>
 <a title="Sviluppa narrazione multimediale"  onClick="SviluppaMetaforaPatente();" >Sviluppa</a> <br> <br>
<span id="btnSxDx">
<input type="button" class="btn" name="indietro" value="<< Indietro " onClick="Precedente()" title="Zoom indietro">
<input type="button" class="btn" name="avanti" value="Avanti >> " onClick="Successiva()" title="Zoom avanti"> 
</span>
<input type="hidden"  name="Pi" value="<%=Pi%>">
<input type="hidden"   name="Pf" value="<%=Pf%>">
									 </center>	   
                     						 </div>                       
										</div>
                                     </div>  
                                     <%if (session("Admin")=true) then %>
                                     
                                      <div class="control-group">
										<label for="textarea" class="control-label">Data</label>
										<div class="controls">
											<input type="text" name="txtDATA" value="<%=DATA%>" class="input-small">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                     
                             
								<% else 
                                   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
                                
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input disabled="disabled" type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT disabled="disabled"  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                <% end if %>
                                 
                                <%end if %>
 
                                    
                                    <div class="form-actions">
										<button type="submit" class="btn btn-primary" name="b1">Aggiorna</button>
								 
									</div>
                                
                                
                                
                                <% Case Cartella&"_U_2_8" ' metafora  METAFORA db desideri%>
                                
                                
                                
                                
                                 <fieldset id="idClient">
                               
                                <div class="control-group">
										<label for="textfield" class="control-label"><b>Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" placeholder="Soggetto che manifesta un aspettativa" class="input-xxlarge"   name="txtSoggettoC<%=i%>"  value="<%= SoggettoC %>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Domanda</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" placeholder="Aspettativa" class="input-xxlarge"  name="txtDomandaC<%=i%>" value="<%=DomandaC%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"   name="txtMotivazioneC<%=i%>" value="<%=MotivazioneC%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"   name="txtDesiderioC<%=i%>" value="<%=DesiderioC%>">
										</div>
									</div>
                               
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge"  name="txtBisognoC<%=i%>" value="<%=BisognoC%>">
										</div>
									</div>
                                    </fieldset>
                                     <div class="control-group" id="idTolleranza">
										<label for="textfield" class="control-label"><b>Tolleranza del Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" placeholder="Indice della tensione che può sopportare" class="input-mini"  name="txtTolleranzaC<%=i%>" value="<%=TolleranzaC %>" >
										</div>
									</div>
                                    
                                    
                                     <hr>
                                     <div class="line"></div>
                                     <fieldset id="idServer">
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Server</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6"> 
											<input type="text" placeholder="Soggetto che risponde alla richiesta" class="input-xxlarge"   name="txtSoggettoS<%=i%>" value="<%=SoggettoS%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Risposta</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" checked="true">  
											<input type="text" placeholder="Risposta alla richiesta" class="input-xxlarge"  name="txtRispostaS<%=i%>" value="<%=RispostaS%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Ragioni che sostengono la Risposta" class="input-xxlarge"  name="txtMotivazioneS<%=i%>" value="<%=MotivazioneS%>" >
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" placeholder="Desiderio che sostiene la Motivazione" class="input-xxlarge"   name="txtDesiderioS<%=i%>" value="<%=DesiderioS %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge"  name="txtBisognoS<%=i%>" value="<%=BisognoS%>" >
										</div>
									</div>
                                    <%'="?="&TipoEvento%>
                               </fieldset>     
                                         <div class="control-group">
                                         <span id="tipoEvento">
										<label for="textfield" class="control-label"><b>Tipo di evento</b></label>
										<div class="controls">
                                             <select name="txtTipoEvento" style="width:auto" onChange="scambia_n(txtTipoEvento.value);"> 
                                               <%if TipoEvento=0 then %>
                                                   <option   value="1">Coerente
                                                 <option selected   value="0">Paradossale
                                                 <%else%>
                                                  <option selected  value="1">Coerente
                                                 <option   value="0">Paradossale
                                                
                                                 <%end if%>
                                             </select> <br><br>
   
                                                <% if TipoEvento=0 then%>
                                                      <img src="../../img/clienteosteno.jpg" name="rappresentazione" width="500px" height="300px">
                                                  <%else %>
                                                   <img src="../../img/clienteostesi.jpg" name="rappresentazione" width="500px" height="300px">                                                    
                                                 <%end if%>
                                                </center><br>
										</div>
                                        </span>
									</div>
                                    
                                    <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea name="S1" id="textarea" rows="5" class="input-block-level"><%=Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
                                    
									
								
                                
									<div class="accordion-group">      
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion5" href="#collapseMail5"><center>
												
                                                <i class="icon-edit" title="Sviluppa"></i>
                                                </center>
											</a>
										</div>
										<div id="collapseMail5" class="accordion-body">
											<div class="accordion-inner">
 <center>
 <a title="Esegui simulazione interattiva" href="6_simula_metafora_desideri.asp?CodiceMetafora=<%=CodiceMetafora%>">Simula</a>
 <br> <br>
  <!--
   <a title="Esegui narrazione multimediale" href="6_narrazione_metafora_desideri.asp?CodiceMetafora=<%'=CodiceMetafora%>&Cartella=<%'=Cartella%>&Paragrafo=<%'=Paragrafo%>">Narrazione</a>
   -->
  <br> <br>
 <a title="Sviluppa narrazione multimediale"  onClick="SviluppaMetaforaDesideri();" >Sviluppa</a> <br> <br>
<span id="btnSxDx">
<input type="button" class="btn" name="indietro" value="<< Indietro " onClick="Precedente()" title="Zoom indietro">
<input type="button" class="btn" name="avanti" value="Avanti >> " onClick="Successiva()" title="Zoom avanti"> 
</span>
<input type="hidden"  name="Pi" value="<%=Pi%>">
<input type="hidden"   name="Pf" value="<%=Pf%>"><br>
<img src="../../img/printer.jpg" alt="Stampa questa metafora" onClick="stampa_navigazione();">

									 </center>	   
                     						 </div>                       
										</div>
                                     </div>  
                                     <%if (session("Admin")=true) then %>
                                     
                                      <div class="control-group">
										<label for="textarea" class="control-label">Data</label>
										<div class="controls">
											<input type="text" name="txtDATA" value="<%=DATA%>" class="input-small">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                     
                             
								<% else 
                                   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
                                
                                    <div class="control-group">
										<label for="textarea" class="control-label">Valutazione</label>
										<div class="controls">
											<input disabled="disabled" type="text" name="txtVAl" value="<%=VAL%>" class="input-mini">
                                    
										</div>
									</div>
                                        <div class="control-group">
										<label for="textarea" class="control-label">Segnalata</label>
										<div class="controls">
											  
                                              <% if (Segnalata=1)  then  %>
                                         
											 <INPUT disabled="disabled"  TYPE="RADIO" name="cb1" checked="true" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"  value="0">No  	          
                                            <% else %>
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1" value="1">Si  
                                             <INPUT disabled="disabled" TYPE="RADIO" name="cb1"   checked="true" value="0">No  
                                           
										<% end if %>
                      					 
										</div>
									</div>
                                <% end if %>
                                 
                                <%end if %>
 
                                    
                                    <div class="form-actions">
										<button type="submit" class="btn btn-primary" name="b1">Aggiorna</button>
								 
									</div>
                                
                                
                                
                                
                                
                                <% end select %>
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
			</div>
            
     <% else
response.write(ucase(CodiceAllievo) & "=" & ucase(Session("CodiceAllievo")))
%> 
 <script language="javascript" type="text/javascript"> 
window.alert("Non puoi visualizzare i dati degli altri studenti!");
location.href="studente_domande.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"
</script>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if %>
    <script language="javascript" type="text/javascript"> 
 function Successiva() {
	 //window.alert(dati.Pf.value);
//location.href="../home.asp"
//location.href=window.history.back();
if (dati.Pf.value==0)
	{
	   alert("Non ci sono Metafore figlio");
	   return 0;
	}
 else
  
	{
		
		
	    document.dati.action = "inserisci_valutazione_metafore.asp?VAL=<%=VAL%>&CodiceAllievo=<%=CodiceAllievo%>&CodiceMetafora=<%=Pf%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>";
		document.dati.submit();	
    }


 }
 
  function Precedente() {
 // window.alert(dati.Pi.value);
  if (dati.Pi.value==0)
	{
	   
	   alert("Non ci sono Metafore genitore");
	   return 0;
	}
 else
  
	{
	    document.dati.action = "inserisci_valutazione_metafore.asp?VAL=<%=VAL%>&CodiceAllievo=<%=CodiceAllievo%>&CodiceMetafora=<%=Pi%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>";
		document.dati.submit();	
    }
//location.href="../home.asp"
//location.href=window.history.back();
 }
 
  function SviluppaMetaforaPatente() {
	     if (dati.Pf.value!=0)
	{
	   
	   alert("Puoi sviluppare solo le metafore che non hanno ancora delle metafore figlio");
	   return 0;
	}
 else{
	    document.dati.action = "6_sviluppa_metafora_patente.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daNavigazione=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
	 }
  }
   function SviluppaMetaforaTopolino() {
	   
	   if (dati.Pf.value!=0)
	{
	   
	   alert("Puoi sviluppare solo le metafore che non hanno ancora delle metafore figlio");
	   return 0;
	}
 else{
	    document.dati.action = "6_sviluppa_metafora_topolino.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daTopolino=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
	}
    
 }
 
 function SviluppaMetaforaDesideri() {
	    if (dati.Pf.value!=0)
	{
	   
	   alert("Puoi sviluppare solo le metafore che non hanno ancora delle metafore figlio");
	   return 0;
	}
 else{
	    document.dati.action = "6_sviluppa_metafora_desideri.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daDesideri=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
 }
 }
 
 
 function stampa_navigazione() {
    document.dati.action = "7_stampa_scheda_metafora.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&daNavigazione=1";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
}

function stampa_topolino() {
    document.dati.action = "7_stampa_scheda_metafora.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&daTopolino=1";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
	}	
		
 </script>   
 
 
   
    <script src="../../../guida/docs/lib/bootstrap/js/bootstrap-dropdown.js"></script>
    <script src="../../../guida/docs/lib/google-code-prettify/prettify.js"></script>

    <script src="../../../guida/js/jquery.pageguide.js"></script>
    <script language="javascript">
      /**
       * Helper Functions
       */

      // View source of current page in a new window
      function viewsource(e){
        window.open("view-source:" + window.location, 'jquery.pageguide.source');
      }

      // Smooth scroll to anchor
      function scrollTo(e) {
        e.preventDefault();

        var anchor = e.currentTarget.hash.slice(1);
            $t = $('a[name=' + anchor + ']');

        if (!$t.size()) return;

        var dvh = $(window).height(),
            dvtop = $(window).scrollTop(),
            eltop = $t.offset().top,
            mgn = {top: 100, bottom: 100};

        var scrollTo = eltop - mgn.top;

        $('html,body').animate({
          scrollTop: scrollTo
        }, {
          duration: 500
        });
      }

      // Example guides
	  
	  </script>
   
 
     
                                <% Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
<script language="javascript" type="text/javascript" src="../jsguide/topolino.js"> </script> 
								<% Case Cartella&"_U_2_5" 'Navigazione%>
<script language="javascript" type="text/javascript" src="../jsguide/navigazione.js"> </script>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
<script language="javascript" type="text/javascript" src="../jsguide/clientserver.js"> 
</script>                              	   
							<%End Select%>
     
      

    
 
		</div> <!--fine main-->
        </div>
        
         

			 
	</body>

 </html>

