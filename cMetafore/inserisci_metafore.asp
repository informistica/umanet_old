<%@ Language=VBScript %>



<!doctype html>
<html>
<head>
   
   <title>Inserisci metafora</title>   
   
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
	
	<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>
	
    
   
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  <script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!");

location.href="studente_domande.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"

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
	if (cont == 1) { alert (" Attenzione per simulare il terremoto devi inserire un evento paradossale che abbia significati in contrasto, la risposta del client deve deludere l'aspettativa del server!");
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
  cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  if Paragrafo="" then
   Paragrafo=TitoloParagrafo
  end if
  DaInserimento=Request.QueryString("DaInserimento") ' vale 1 se sono chiamata dopo inserisci_metafora1, anzichè da studente_domande, in tal caso devo fare la query per reecuperare i dati.
  
  'response.write(Codice_Test)
  Modulo=Request.QueryString("Modulo")
  if Modulo="" then
  Modulo=Session("Modulo")
  end if
  MO=Request.QueryString("MO")
  VAL=Request.QueryString("VAL")
  URL=Request.QueryString("URL")
  DATA=cdate(Request.QueryString("DATA"))
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  ID=CodiceMetafora 
  Cartella=Request.QueryString("Cartella")
  Segnalata= Request.QueryString("Segnalata")

  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else 
		Select Case Codice_Test
			Case Cartella&"_U_2_3" 
			Case Cartella&"_U_2_5"
			Case Cartella&"_U_2_8"
		End Select
  %>
    
    <body  class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" onload="cleartextarea();" >

  <% end if %>

<%
  ' dichiarazione delle variabili per contenere i dati (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  
  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
      'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBCopiatestonline")
    
	
	
    
	ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../database/Copiaditestonline.mdb")

 homesito="/expo2015/UECDL"   
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceAllievo=Request.QueryString("cod")
  if CodiceAllievo="" then
    CodiceAllievo=Request.QueryString("CodiceAllievo")
  end if
   if CodiceAllievo="" then
    CodiceAllievo=Session("CodiceAllievo")
  end if
  
  

   ID_Premetafora=Request.QueryString("ID_Premetafora")

 Select Case Codice_Test
	Case Cartella&"_U_2_3" 
	 Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa 
  Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
   if Modulo="" then
  Modulo=Session("Modulo")
  end if
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito")
  prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella")

  Num = Request.QueryString("Num")
  Num=Num+1
  
 
  Case Cartella&"_U_2_5"
     
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo") ' TOPOLINO ED OBIETTIVI
  ParagrafoNuovo="Navigazione nella Rete della Vita" ' VALORE CABLATO
  Modulo=Request.QueryString("Modulo")
  if Modulo="" then
  Modulo=Session("Modulo")
  end if
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  'Quesito=Request.QueryString("Quesito")
  'prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella")
  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = cint(Request.QueryString("Num"))
  Num=Num+1
  daTopolino=Request.QueryString("daTopolino") ' vale 1 se sono stata chiamata da spiegazione_metafora_topolino tramite invia a -->
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  if daTopolino<>"" then
    
   ' eseguo la query con codiemetafora per ricreare tutti i dati sopra che non ho tramite QueryString
    QuerySQL="Select * from M_Topolino where CodiceMetafora=" & cint(CodiceMetafora)& ";"
	'response.write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	
	'Privato=rsTabella.fields("Privato") 
	'rsTabella.close
     Cartella=rsTabella.fields("Cartella")
	 'Codice_Test=rsTabella.fields("Id_Arg")
	 Modulo=rsTabella.fields("Id_Mod")
	   
  end if
	
Case Cartella&"_U_2_8" ' dbdesideri
	 '  response.write("cioo")
	  Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
 TipoEvento=Request.QueryString("TipoEvento")
  if strcomp(TipoEvento&"","")=0 then
	     TipoEvento=1
	  end if
 ' TipoEvento=1
  'Request.Cookies("Dati")("CodiceTest")= Codice_Test
  
  Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  if Modulo="" then
  Modulo=Session("Modulo")
  end if
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito")
  prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella")

  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = Request.QueryString("Num")
  Num=Num+1
	 

 
	   
	   
End Select

 
if MO<>"" then 
 Modulo=MO
end if  
QuerySQLApp=QuerySQL ' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close

  
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' 
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
'response.write(sReadAll)
objTextFile.Close

 

   
%>
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
						<h1> <i class="icon-comments"></i> Inserisci metafora </h1> 
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%>:&nbsp;<% 
						if daTopolino<>"" then 
							response.write(ParagrafoNuovo) ' "Navigazione nella Rete della Vita"
						else
							response.write(Paragrafo)
						
						end if
						%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 		<% 'response.write("pi="&Pi)
 'response.write("<br>"&Codice_Test)	
'response.write("DBQ=" & Server.MapPath("../database/Copiaditestonline.mdb"))
 
 %>						 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                  <%
				  %>
                  <div class="alert-success" id="messaggio" style="display:none";>
                  Inserimento avvenuto!<br>  
				  1) Sei vuoi sviluppare la metafora aprila dal Quaderno UWWWW la metafora</a> <br>
				  2) Se ne vuoi inserie un'altra scrivila qui sotto...
                  </div>
				   
						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i>  Metafora <%=CodiceMetafora%></h3>
							</div>
							 <div class="box-content" id="storia">
								 
                    		</div>
							<div class="box-content">
								
                                
								
							
                                
                                <% Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
                                <fieldset>
<form name="dati" class="form-vertical" action="inserisci_metafora_topolino1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>" method="POST"> 
									<input type="hidden" value="<%=Cartella%>" id="cartella">
								  <input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
								   <input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
								   <input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
								      <input type="hidden" value="<%=Modulo%>" id="Modulo">
									   <input type="hidden" value="<%=Paragrafo%>" id="Paragrafo"> 
									   <input type="hidden" value="<%=ID_Premetafora%>" id="ID_Premetafora"> 
                              
                                
                                  <div class="control-group">
										<label for="textfield" class="control-label"><b>Topolino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Soggetto protagonista" class="input-xxlarge"  name="txtTopolino" id="txtTopolino" maxlength="148"  value="<%=Topolino%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Formaggio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Obiettivo da raggiungere" class="input-xxlarge"  name="txtFormaggio"  id="txtFormaggio" maxlength="148"  value="<%=Formaggio%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Fame</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Motivazione che spinge verso l'obiettivo" class="input-xxlarge"  name="txtFame" id="txtFame" maxlength="148"  value="<%=Fame%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Labirinto</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Contesto in cui si svolge l'azione" class="input-xxlarge"  name="txtLabirinto" id="txtLabirinto" maxlength="148"  value="<%=Labirinto%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Obiettivo" class="input-xxlarge"  name="txtStrada" id="txtStrada" maxlength="148"  value="<%=Strada%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6" checked="true">  
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Strategia vincente" class="input-xxlarge"  name="txtStrada_OK" id="txtStrada_OK"  maxlength="148"  value="<%=Strada_OK%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" >
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Strategia perdente" class="input-xxlarge"  name="txtStrada_KO" id="txtStrada_KO"  maxlength="148" value="<%=Strada_KO%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Testata</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Conseguenze della strategia perdente" class="input-xxlarge"  name="txtTestata" id="txtTestata"  maxlength="148"  value="<%=Testata%>">
										</div>
									</div>
                                       </fieldset>
                                       
                                     <span id="idDistanza">
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" oninput="aggiornaStoria(1);" onfocus="aggiornaStoria(1);" placeholder="Num. da 1 a 5" class="input-small"  name="txtDistanza" id="txtDistanza"  value="<%=Distanza%>">
										</div>
									</div>
                                 </span>
								 <div class="form-actions">
										<button type="button" class="btn" onClick="copia_testo(1);" name="b1">Inizia spiegazione</button>	
									</div>
                                  <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea maxlength="910" name="S1" id="textarea" rows="5" class="input-block-level"></textarea> 
										</div>
									</div>
                                    
									
								
                            
                                      
                                    
                                    <div class="form-actions">
										<button type="button" onClick="inserisci_metafore(0);" class="btn btn-primary" name="b1">Invia</button>
								 
									</div>
								
                                <%  Case Cartella&"_U_2_5" ' metafora  METAFORA NAVIGAZIONE%>
                       <fieldset>

                                
<form name="dati" method="POST" class="form-vertical"  action="inserisci_metafora_patente1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=ParagrafoNuovo%>&Modulo=<%=Modulo%>" >  
								<input type="hidden" value="<%=Cartella%>" id="cartella">
								  <input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
								   <input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
								   <input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
								      <input type="hidden" value="<%=Modulo%>" id="Modulo">
									   <input type="hidden" value="<%=Paragrafo%>" id="Paragrafo">  
									   <input type="hidden" value="<%=ID_Premetafora%>" id="ID_Premetafora">                              
<% 

'response.write("ID_Premetafora="&ID_Premetafora)
if daTopolino=1 then  
 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <%  
   '' 
   ' eseguo la query con codiemetafora per ricreare tutti i dati sopra che non ho tramite QueryString
    QuerySQL="Select * from M_Topolino where CodiceMetafora=" & cint(CodiceMetafora)& ";"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'response.write(QuerySQL)
       Autista=rsTabella.fields("Topolino") 
	  Lupo=rsTabella.fields("Testata") 
      Destinazione=rsTabella.fields("Formaggio") 
	  Carburante=rsTabella.fields("Fame")
	  Luogo=rsTabella.fields("Labirinto")
	  Strada=rsTabella.fields("Strada")
	  StradaKO=rsTabella.fields("Strada_KO")
	  StradaOK=rsTabella.fields("Strada_OK") 
	  rsTabella.close	   
	  
	  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&CodiceMetafora&".txt" 'per il server on line
				 url=Replace(url,"\","/")
				' response.write(url)
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
	  
  else   
 	  Lupo=""
      Destinazione=""
	  Carburante=""
	  Luogo=""
	  Strada=""
	  StradaKO=""
	  StradaOK=""
	  sReadAll=""
end if
Cespugli=""
Cestino=""


 %>                              
                                <div class="control-group">
										<label for="textfield" class="control-label"><b>Autista</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="SOGGETTO protagonista" class="input-xxlarge"  name="txtAutista"  id="txtAutista"  maxlength="148"  value="<%= Autista %>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Destinazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="OBIETTIVO da raggiungere" class="input-xxlarge"  name="txtDestinazione" id="txtDestinazione"  maxlength="148" value="<%=Destinazione%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Carburante</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="MOTIVAZIONE che spinge verso l'obiettivo" class="input-xxlarge"  name="txtCarburante" id="txtCarburante"  maxlength="148" value="<%=Carburante%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Luogo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="SITUAZIONE in cui si svolge l'azione" class="input-xxlarge"  name="txtLuogo" id="txtLuogo"  maxlength="148" value="<%=Luogo%>">
										</div>
									</div>
                                    
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="COMPORTAMENTO" class="input-xxlarge"   name="txtStrada" id="txtStrada"  maxlength="148" value="<%=Strada%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_OK</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6" checked="true">  
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="COMPORTAMENTO ADEGUATO" class="input-xxlarge"   name="txtStrada_OK" id="txtStrada_OK"  maxlength="148" value="<%=Strada_OK%>">
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Strada_KO</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" >
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="COMPORTAMENTO INADEGUATO" class="input-xxlarge"  name="txtStrada_KO" id="txtStrada_KO"  maxlength="148" value="<%=Strada_KO%>" >
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cespugli</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="FEEDBACK di pericolo" class="input-xxlarge"   name="txtCespugli" id="txtCespugli"  maxlength="148" value="<%= Cespugli %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Lupo</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="CONSEGUENZE NEGATIVE" class="input-xxlarge"  name="txtLupo" id="txtLupo"  maxlength="148" value="<%= Lupo %>" >
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Cestino</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="ATTACCAMENTI, eccessi da abbandonare" class="input-xxlarge"  name="txtCestino"  maxlength="148" id="txtCestino" value="<%= Cestino %>">
										</div>
									</div>
                                    
                                        </fieldset>
                                       
                                     <span id="idDistanza">
                                         <div class="control-group">
										<label for="textfield" class="control-label"><b>Distanza</b></label>
										<div class="controls">
                                         
											<input type="text" oninput="aggiornaStoria(2);"  onfocus="aggiornaStoria(2)" placeholder="Num.da 1 a 5" class="input-small" id="txtDistanza"  name="txtDistanza"    value="<%=Distanza%>">
										</div>
									</div>
                                 </span>  
									<div class="form-actions">
										<button type="button" class="btn" onClick="copia_testo(2);" name="b1">Inizia spiegazione</button>	
									</div>
                                    <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea maxlength="910" name="S1" id="textarea" rows="5" class="input-block-level">
                                             <%=sReadAll%>
                                            </textarea> 
										</div>
									</div>
                                    
									
								
                               
                                    
                                    <div class="form-actions">
										<button type="button" onClick="inserisci_metafore(1);" class="btn btn-primary" name="b1">Invia</button>
									</div>
                                
                                
                                
                                <% Case Cartella&"_U_2_8" ' metafora  METAFORA db desideri%>
    
     
                                
   <form class="form-vertical" method="POST" name="document"  action="inserisci_metafora_dbdesideri1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>" >   
							<input type="hidden" value="<%=Cartella%>" id="cartella">
								  <input type="hidden" value="<%=CodiceAllievo%>" id="CodiceAllievo">
								   <input type="hidden" value="<%=CodiceMetafora%>" id="CodiceMetafora">
								   <input type="hidden" value="<%=Codice_Test%>" id="Codice_Test">
								      <input type="hidden" value="<%=Modulo%>" id="Modulo">
									   <input type="hidden" value="<%=Paragrafo%>" id="Paragrafo"> 
									 <input type="hidden" value="<%=ID_Premetafora%>" id="ID_Premetafora"> 
                                
                                
                                 <fieldset id="idClient">
                               
                                <div class="control-group">
										<label for="textfield" class="control-label"><b>Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="1"> 
                                            <input type="text" oninput="aggiornaStoria(3);" placeholder="Soggetto che manifesta un aspettativa" class="input-xxlarge"   name="txtSoggettoC"  maxlength="148" id="txtSoggettoC" value="<%= SoggettoC %>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Domanda</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="2"> 
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Aspettativa" class="input-xxlarge"  name="txtDomandaC"  id="txtDomandaC"  maxlength="148" value="<%=DomandaC%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="3"> 
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"   name="txtMotivazioneC"  maxlength="148"  id="txtMotivazioneC" value="<%=MotivazioneC%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="4"> 
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Desiderio che sostiene l'Aspettativa" class="input-xxlarge"   name="txtDesiderioC"  maxlength="148"  id="txtDesiderioC" value="<%=DesiderioC%>">
										</div>
									</div>
                               
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="5"> 
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge"  name="txtBisognoC"  maxlength="148" id="txtBisognoC" value="<%=BisognoC%>">
										</div>
									</div>
                                    </fieldset>
                                     <div class="control-group" id="idTolleranza">
										<label for="textfield" class="control-label"><b>Tolleranza del Client</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="10" >
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Indice della tensione che può sopportare" class="input-mini"  name="txtTolleranzaC"  maxlength="148" id="txtTolleranzaC" value="<%=TolleranzaC %>" >
										</div>
									</div>
                                    
                                    
                                     <hr>
                                     <div class="line"></div>
                                     <fieldset id="idServer">
                                    	<div class="control-group">
										<label for="textfield" class="control-label"><b>Server</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="6"> 
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Soggetto che risponde alla richiesta" class="input-xxlarge"   name="txtSoggettoS"  maxlength="148" id="txtSoggettoS" value="<%=SoggettoS%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Risposta</b></label>
										<div class="controls">
                                         <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="7" checked="true">  
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Risposta alla richiesta" class="input-xxlarge"  name="txtRispostaS" id="txtRispostaS"  maxlength="148" value="<%=RispostaS%>" >
										</div>
									</div>
                                     	<div class="control-group">
										<label for="textfield" class="control-label"><b>Motivazione</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Ragioni che sostengono la Risposta" class="input-xxlarge"  name="txtMotivazioneS"  maxlength="148" id="txtMotivazioneS" value="<%=MotivazioneS%>" >
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Desiderio</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="8" >
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Desiderio che sostiene la Motivazione" class="input-xxlarge"   name="txtDesiderioS"  maxlength="148" id="txtDesiderioS" value="<%=DesiderioS %>">
										</div>
									</div>
                                    <div class="control-group">
										<label for="textfield" class="control-label"><b>Bisogno</b></label>
										<div class="controls">
                                          <input TYPE="RADIO" class="radio"  name="rdSviluppa"  title="Seleziona per sviluppare" value="9" >
											<input type="text" oninput="aggiornaStoria(3);" placeholder="Bisogno che sostiene il Desiderio" class="input-xxlarge"  name="txtBisognoS"  maxlength="148" id="txtBisognoS" value="<%=BisognoS%>" >
										</div>
									</div>
                                  
                               </fieldset> 
                                         <div class="control-group">
                                         <span id="tipoEvento">
										<label for="textfield" class="control-label" disabled="true"><b>Tipo di evento</b></label>
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
    										<center>
                                                <% if TipoEvento=0 then%>
                                                      <img src="../../img/clienteosteno.jpg" name="rappresentazione" width="500px" height="300px">
                                                  <%else %>
                                                   <img src="../../img/clienteostesi.jpg" name="rappresentazione" width="500px" height="300px">                                                    
                                                 <%end if%>
                                                </center><br>
										</div>
									</div>
                                    </span>
                                    <div class="control-group" id="Boxtext">
										<label for="textarea" class="control-label"><b>Spiegazione</b></label>
										<div class="controls">
											<textarea maxlength="910" name="S1" id="textarea" rows="5" class="input-block-level"><%=Response.write(sReadAll)%> </textarea> 
										</div>
									</div>
                                    

                                    <div class="form-actions">
									<button type="button" onClick="inserisci_metafore(2);" class="btn btn-primary" name="b1">Invia</button>
								 
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
'response.write(ucase(CodiceAllievo) & "=" & ucase(Session("CodiceAllievo")))
%> 
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
	    document.dati.action = "inserisci_valutazione_metafore.asp?VAL=<%=VAL%>&CodiceAllievo=<%=CodiceAllievo%>&CodiceMetafora=<%=Pf%>&Cartella=<%=Cartella%>&Num=<%=Num%>&Modulo=<%=Modulo%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>";
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
	    document.dati.action = "6_sviluppa_metafora_patente.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daNavigazione=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
 }
   function SviluppaMetaforaTopolino() {
	    document.dati.action = "6_sviluppa_metafora_topolino.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daTopolino=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
    
 }
 
 function SviluppaMetaforaDesideri() {
	    document.dati.action = "6_sviluppa_metafora_desideri.asp?CodiceMetafora=<%=CodiceMetafora%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&daDesideri=1&Num=1&DATA=<%=DATA%>";
		document.dati.submit();	
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

<script language="javascript" type="text/javascript"> 

function copia_testo(tipo){
	var testo;
 switch (tipo) {
		case 1:  //Topolino
		testo=document.dati.txtTopolino.value.toUpperCase() + " "+ document.dati.txtFormaggio.value.toUpperCase() + " "+ document.dati.txtFame.value.toUpperCase() + " "+document.dati.txtLabirinto.value.toUpperCase() + " "+document.dati.txtStrada.value.toUpperCase() + " "+document.dati.txtStrada_OK.value.toUpperCase() + " "+document.dati.txtStrada_KO.value.toUpperCase() + " "+document.dati.txtTestata.value.toUpperCase() ;
		break;
		case 2: //Navigazione
		testo=document.dati.txtAutista.value.toUpperCase() + " "+ document.dati.txtDestinazione.value.toUpperCase() + " "+ document.dati.txtCarburante.value.toUpperCase() + " "+document.dati.txtLuogo.value.toUpperCase() + " "+document.dati.txtStrada.value.toUpperCase() + " "+document.dati.txtStrada_OK.value.toUpperCase() + " "+document.dati.txtStrada_KO.value.toUpperCase() + " "+document.dati.txtCespugli.value.toUpperCase() + " "+ " "+document.dati.txtLupo.value.toUpperCase() + " "+ " "+document.dati.txtCestino.value.toUpperCase() + " ";
		break;
 
 }
	document.dati.S1.value=testo;
	 } 
	 
    function aggiornaStoria(tipo) {
        var narrazione = "";
		switch (tipo) {
		case 1:
			//Metafora topolino
		if (document.dati.txtTopolino.value != "")
			Topolino = document.dati.txtTopolino.value.toUpperCase();
		else
			Topolino = "SOGGETTO";
		if (document.dati.txtFormaggio.value != "")
			Formaggio = document.dati.txtFormaggio.value.toUpperCase();
		else
			Formaggio = "OBIETTIVO";
		if (document.dati.txtFame.value != "")
			Fame = document.dati.txtFame.value.toUpperCase();
		else
			Fame = "MOTIVAZIONE";

		if (document.dati.txtLabirinto.value != "")
			Labirinto = document.dati.txtLabirinto.value.toUpperCase();
		else
			Labirinto = "CONTESTO";
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
		if (document.dati.txtTestata.value != "")
			Testata = document.dati.txtTestata.value.toUpperCase();
		else
			Testata = "CONSEGUENZE NEGATIVE";

		plurale = Topolino.search(/ e /i); //se è presente e oppure E è >0
		plurale1 = Topolino.search(","); //faccio mettere , per indicare il prurale
		var narrazione = "";
		if ((plurale == -1) && (plurale1 == -1)) {
			volere = "vuoi";
			raggiungere = "raggiungerai";
			avere = "hai";
			scegliere = "scegli";
			avvicinarsi = "ti avvicina";
			allontanarsi = "ti allontana";
			allontanarsi1 = "ti sei allontanato troppo hai";
			scontrarsi = "e ti sei scontrato";
			continuare = "continua";
			fare = "ci sei quasi fai";
		}
		else {
			volere = "volete";
			raggiungere = "raggiungerete";
			avere = "avete";
			scegliere = "scegliete";
			avvicinarsi = "vi avvicina";
			allontanarsi = "vi allontana";
			allontanarsi1 = "vi siete allontanati troppo avete";
			scontrarsi = "e vi siete scontrati";
			continuare = "continuate";
			fare = "ci siete quasi fate";
		}
		contSi = 0;
		contNo = 0;
		Motivato = 0;

		narrazione += "\n\n " + Topolino + " " + volere + " raggiungere " + Formaggio + " ?";

		//document.getElementById("storia").innerHTML=dati.Storia.value;


		narrazione += " NO! <br>\n\n Mancando " + Fame + " per raggiungere " + Formaggio + " , " + Topolino + " nel contesto " + Labirinto + " non " + raggiungere + " l'obiettivo ! ";
		narrazione += "<br>\n\n " + Topolino + " " + volere + " raggiungere " + Formaggio + " ?" + "Si!";
		narrazione += "<br>\n\n   " + Topolino + "  quale  " + Strada + " " + scegliere + " ?  ";

		narrazione += "<br>\n\nATTENZIONE  " + Topolino + "  la scelta  " + Strada_KO + " " + allontanarsi + " da  " + Formaggio;
		narrazione += "<br>\n\n :-(  " + Topolino + " " + allontanarsi1 + " scelto la strada chiusa  " + Strada_KO + " " + scontrarsi + " con " + Testata + " \n ";
		narrazione += "<br>\n\n  " + Topolino + "  quale  " + Strada + " " + scegliere + " ?  ";
		narrazione += "<br>\n\n  " + Topolino + "  la scelta  " + Strada_OK + " " + avvicinarsi + " a  " + Formaggio + "  " + continuare + " così !  ";
		narrazione += "<br>\n\n Coraggio " + fare + " l'ultimo passo ! '";
		narrazione += "<br>\n :-) COMPLIMENTI  " + Topolino + " " + avere + " raggiunto " + Formaggio + "!!!";
			break;
		case 2:
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
			break;
		case 3:
			//Metafora client/server
			break;
		
		}



       

        document.getElementById("storia").innerHTML = narrazione;
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
	 
function cleartextarea(){
	 document.getElementById("textarea").value="";
  }	  
	  
	  </script>
	  
	  
	  
	  
	   <script language="javascript" type="text/javascript"> 

function inserisci_metafore(tipo){
var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo,errore,ID_Premetafora;
		  errore=0;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;	
		  ID_Premetafora=document.getElementById("ID_Premetafora").value; 
	switch(tipo) {
	  case 0:	  	
		  txtTopolino=document.getElementById("txtTopolino").value;
		  txtFormaggio=document.getElementById("txtFormaggio").value;
		  txtFame=document.getElementById("txtFame").value;
		  txtLabirinto=document.getElementById("txtLabirinto").value;
		  txtStrada=document.getElementById("txtStrada").value;
		  txtStrada_OK=document.getElementById("txtStrada_OK").value;
		  txtStrada_KO=document.getElementById("txtStrada_KO").value;
		  txtTestata=document.getElementById("txtTestata").value;
		  txtDistanza=document.getElementById("txtDistanza").value;
		  textarea=document.getElementById("textarea").value;
		  if (txtTopolino=="" || txtFormaggio=="" || txtFame=="" || txtLabirinto=="" || txtStrada=="" || txtStrada_OK=="" || txtStrada_KO=="" || txtTestata==""  || txtDistanza=="" || textarea==""){
		   errore=1;
		   alert("Compila tutti i campi");
		  }
		   if (isNaN(txtDistanza)) {
			   errore=1;
			   alert("La distanza deve essere un numero");
		  }
		  dati2="&ID_Premetafora="+ID_Premetafora+"&txtTopolino="+txtTopolino+"&txtFormaggio="+txtFormaggio+"&txtFame="+txtFame+"&txtLabirinto="+txtLabirinto+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtTestata="+txtTestata+"&txtDistanza="+txtDistanza+"&S1="+textarea;		
		 // da riabilitare per inviare con post per il problema del limiti di lunghezza delle textarea
		 // dati2="&txtTopolino="+txtTopolino+"&txtFormaggio="+txtFormaggio+"&txtFame="+txtFame+"&txtLabirinto="+txtLabirinto+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtTestata="+txtTestata+"&txtDistanza="+txtDistanza;		 
		 break;
	  case 1:
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
			 if (txtAutista=="" || txtDestinazione=="" || txtCarburante=="" || txtLuogo=="" || txtStrada=="" || txtStrada_OK=="" || txtStrada_KO=="" || txtCespugli=="" || txtLupo=="" || txtCestino==""  || txtDistanza=="" || textarea==""){
			   errore=1;
			   alert("Compila tutti i campi");
		  }
		   if (isNaN(txtDistanza)) {
			   errore=1;
			   alert("La distanza deve essere un numero");
		  }
			dati2="&ID_Premetafora="+ID_Premetafora+"&txtAutista="+txtAutista+"&txtDestinazione="+txtDestinazione+"&txtCarburante="+txtCarburante+"&txtLuogo="+txtLuogo+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtCespugli="+txtCespugli+"&txtCestino="+txtCestino+"&txtLupo="+txtLupo+"&txtDistanza="+txtDistanza+"&S1="+textarea;
			//dati2="&txtAutista="+txtAutista+"&txtDestinazione="+txtDestinazione+"&txtCarburante="+txtCarburante+"&txtLuogo="+txtLuogo+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtCespugli="+txtCespugli+"&txtCestino="+txtCestino+"&txtLupo="+txtLupo+"&txtDistanza="+txtDistanza;
			break;
	  case 2:
			txtSoggettoC=document.getElementById("txtSoggettoC").value;
			txtDomandaC=document.getElementById("txtDomandaC").value;
			txtMotivazioneC=document.getElementById("txtMotivazioneC").value;
			txtDesiderioC=document.getElementById("txtDesiderioC").value;
			txtBisognoC=document.getElementById("txtBisognoC").value;
			txtSoggettoS=document.getElementById("txtSoggettoS").value;
			txtRispostaS=document.getElementById("txtRispostaS").value;
			txtMotivazioneS=document.getElementById("txtMotivazioneS").value;
			txtDesiderioS=document.getElementById("txtDesiderioS").value;
			txtBisognoS=document.getElementById("txtBisognoS").value;
			txtTipoEvento=document.getElementById("txtTipoEvento").value;
			txtTolleranzaC=document.getElementById("txtTolleranzaC").value;
			textarea=document.getElementById("textarea").value;
			
			 if (txtSoggettoC=="" || txtDomandaC=="" || txtMotivazioneC=="" || txtDesiderioC=="" || txtBisognoC=="" || txtSoggettoS=="" || txtRispostaS=="" || txtMotivazioneS=="" || txtDesiderioS=="" || txtBisognoS==""  || txtTolleranzaC=="" || textarea==""){
			   errore=1;
			   alert("Compila tutti i campi");
		  }
		   if (isNaN(txtTolleranzaC)) {
			   errore=1;
			   alert("La distanza deve essere un numero");
		  }
			
			
			dati2="&ID_Premetafora="+ID_Premetafora+"&txtSoggettoC="+txtSoggettoC+"&txtDomandaC="+txtDomandaC+"&txtMotivazioneC="+txtMotivazioneC+"&txtDesiderioC="+txtDesiderioC+"&txtBisognoC="+txtBisognoC+"&txtSoggettoS="+txtSoggettoS+"&txtRispostaS="+txtRispostaS+"&txtMotivazioneS="+txtMotivazioneS+"&txtDesiderioS="+txtDesiderioS+"&txtBisognoS="+txtBisognoS+"&txtTipoEvento="+txtTipoEvento+"&S1="+textarea+"&txtTolleranzaC="+txtTolleranzaC;		 
			//dati2="&txtSoggettoC="+txtSoggettoC+"&txtDomandaC="+txtDomandaC+"&txtMotivazioneC="+txtMotivazioneC+"&txtDesiderioC="+txtDesiderioC+"&txtBisognoC="+txtBisognoC+"&txtSoggettoS="+txtSoggettoS+"&txtRispostaS="+txtRispostaS+"&txtMotivazioneS="+txtMotivazioneS+"&txtDesiderioS="+txtDesiderioS+"&txtBisognoS="+txtBisognoS+"&txtTipoEvento="+txtTipoEvento+"&txtTolleranzaC="+txtTolleranzaC;		 
		
			break;
	} 
	
	if (errore==0){
		dati="cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&Codice_Test="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo; 
		var url = "7_inserisci_metafora_ajax.asp?"+dati+dati2;			   
		var xhttp = new XMLHttpRequest();
		var testojson,stato;
		xhttp.onreadystatechange = function() {
		  if (xhttp.readyState == 4 && xhttp.status == 200) {
							var testo = xhttp.responseText;	
							testoJSON=JSON.parse(testo);
							stato=testoJSON["stato"];
							alert(stato);
							CodiceMetafora=testoJSON["id"];
							if (CodiceMetafora != 0)
								 window.location.href = "sintesi_metafore.asp?cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&CodiceTest="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&CodiceMetafora="+CodiceMetafora;
			 
				
		  }
		};

	    // FUNZIONA MA VA REPLICATO IN TUTTE LE PAGINE cMetafore 
		//var url = "7_inserisci_metafora_ajax.asp?"+dati+dati2;	
		//testo=encodeURIComponent(textarea);
	 	//params="S1="+testo;
		//xhttp.open('POST', url) 
		//xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')
		//xhttp.send(params);



		xhttp.open("GET", url, true);
		xhttp.send();	
	}	
}
	
	
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

