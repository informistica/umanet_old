<%@ Language=VBScript %>
<!doctype html>
<%
 Dim Num_Quiz,rand,Quiz,orderby
Function domandaplus()
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 Cartella=rsTabella(13)
	 Modulo=rsTabella(10)
	 'Paragrafo=rsTabella(15)
	 Paragrafo=rsTabella.fields("Titolo")
	' response.write("PARAGRAFO="&Paragrafo)
	 Id=rsTabella(0)
	 
	 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	response.write(sReadAll)
	objTextFile.Close
End Function 

Function imgdomanda()
     Cartella=rsTabella(13)
	 Modulo=rsTabella(10)
	 Paragrafo=rsTabella.fields("Titolo")
	 Id=rsTabella(0)
	 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".jpg"
     url=Replace(url,"\","/")
	 %>
	 <img src="<%=url%>">
	 <%
	 
end function
NUMTEST=request.form("NUMTEST")
%>

<%Sub setInQuizOrderBy()
' genera un numero casuale per scegliere quale quiz e quale ordinamento per le domande   
             Num_Quiz=rsTabella(0) 
			if Num_Quiz=-1 then
			   Quiz=-1
			   randomize()
			    do 
					rand=rnd()
				loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=7)
				orderby=left((rand*5),1)
			   
			else
			 
				'response.write("NUM_QUIZ="&Num_Quiz)
				randomize()
				do 
					rand=rnd()
				loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=Num_Quiz)
				Quiz=left((rand*5),1)
				' Response.write("QUIZ="&Quiz)
				 do 
					rand=rnd()
				loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=7)
				orderby=left((rand*5),1)
			end if
end sub %>
<% Response.Buffer=True %>
<html>
<head>
   
   <title>Esegui test</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 

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
	<!-- jQuery UI -->
	
     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	<script src="../../js/demonstration.min.js"></script>
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
     
    
    <script language="javascript" type="text/javascript"> 
function showText() {window.alert("La pagina richiesta non è al momento abilitata")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
 <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")

location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<% Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body class='theme-<%=session("stile")%>'>
  <% end if  %>
  
   


	<div id="navigation">
     
        <% 
		 
 ' per generare un ordinamento casuale delle domande in base ad uno dei seguenti campi
 Dim order(8)
 
 
order(0)="" ' non lo uso 
order(1)="CodiceDomanda" 
order(2)="Quesito" 
order(3)="Risposta1"
order(4)="Risposta2"
order(5)="Risposta3"  
order(6)="Risposta4" 
order(7)="Data" 
 
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
						<h1> <i class="icon-comments"></i> Esegui test </h1> 
                     <%' if session("DB")=1 then%>
                   <!--     <a title="Condividi link alla pagina" href="#" onClick="javascript:PopUpWindow(600,400);return false;"><i class="glyphicon-share_alt"> </i> <small>Condividi</small> </a>  -->
                      <%' end if%>
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
							<a href="#more-files.html">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Verifica</a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                 
<%  Stato=Request.QueryString("Stato") '=0 se svolto test del paragrafo 1 se svolgo quello del modulo
   Stato0=Request.QueryString("Stato0")
   Modulo=Request.QueryString("Modulo") 
   Capitolo=Request.QueryString("TitoloCapitolo")
   Paragrafo=Request.QueryString("TitoloParagrafo")
   'Raccolta dei dati digitati dall'utente e salvati nel cookie
   TitoloTest=Request.Cookies("Dati")("TitoloTest")
'   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceAllievo=Session("CodiceAllievo")
  ' if (CodiceAllievo="") then
  '    response.Redirect("../home.asp")
  ' end if
    Tutti=request.querystring("Tutti")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   CodiceTest = Request.QueryString("CodiceTest") ' se svolgo tutto il modulo (stato=1) contiene l'Id del modulo e non del paragrafo
   Verifica=clng(Request.QueryString("Verifica")) ' se sono stato chiamato da verifica il test il valore vale 1 , serve per segnalre le domande da correggere
    filecsv=Request.form("txtCSV")
	
	
	 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"'"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'Privato=rsTabella.fields("Privato") 
	TestAbilitato=rsTabella.fields("TestAbilitato")
	rsTabella.close

 Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  

 if  (Session("Admin")=True) or (TestAbilitato=1) then  ' else è alla fine
  %>
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%><% if strcomp(CodiceTest,"1_0")<>0 then%> : <%=Paragrafo%><%end if%> 
                          <% if  Sottoparagrafo<>"" then %>
                          /&nbsp;<%=Sottoparagrafo%>
                         <% end if%>
                        
                        </h3>
			          </div>
				      <div class="box-content">
                      
<%
if (Stato=0) then 

   if CodiceSottopar<>"" then
	     QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0 and Domande.VF=0;"
	   else
       QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Multiple=0  and Domande.VF=0;"
	end if
 else 
			' il In_QUiz=-2 quando inserisco un test ma non vogli che sia visibile, probabilmente non serve usando TestAbilitati=0    
          QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and Domande.VF=0;"
    
  end if	
  
  %>
  
  <%		
	    ' Const ForReading = 1, ForWriting = 2, ForAppending = 8
'		 Dim sRead, sReadLine, sReadAll, objTextFile
'		 Set objFSO = CreateObject("Scripting.FileSystemObject")  
'		   	url="C:\Inetpub\umanetroot\anno_2013-2014\logEsegui.txt"
'						Set objCreatedFile = objFSO.CreateTextFile(url, True)
'						objCreatedFile.WriteLine(QuerySQL & "orderby="&orderby)
'						objCreatedFile.Close 
			
		'	response.write(QuerySQL&"<br>")
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
 
If rsTabella.BOF=True And rsTabella.EOF=True Then %>
   <div class="alert alert-error">
                    Test non ancora disponibile!
                     </div>
   				 <% rsTabella.close()%>
<% else %>
<div class="alert-success">
      <% if Verifica=1 Then %>
		<p align="center"><b>ESEGUI VERIFICA DEL TEST </b></p>
		<%else%>
		<p align="center"><b>ESEGUI TEST</b></p>
		<%end if%>
		<p align="center"><b><%Response.write (TitoloTest) %></b></p> <!-- stampa il titolo del test -->
		
		<%  ' non serve visto che dopo la queruy ne faccio una latra che sovrascrive rsTbella
		'if (Stato=0) then 
'		 'Definzione codice SQl della query per ricercare le domande del paragrafo 
'		 ' mi serve anche il titolo del paragrafo per ricostruire il nome del file che contiene la domanda plus
'		   
'		   ' codice per la generazione di quiz random : seleziono il numero di quiz presenti 
'		   
		  ' if CodiceSottopar<>"" then
'	       QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
'		   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
'		   " WHERE Domande.Id_Arg='" & CodiceTest & "'  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0  and Domande.VF=0;"
'		  
'		   else
'     QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
'		   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
'		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Multiple=0  and Domande.VF=0;"
'        
'			end if
'		   
		   
		
			 
		'	 Set objFSO = CreateObject("Scripting.FileSystemObject")  
'			url="C:\Inetpub\umanetroot\anno_2013-2014\logEsegui319.txt"
'						Set objCreatedFile = objFSO.CreateTextFile(url, True)
'						objCreatedFile.WriteLine(QuerySQL & "orderby="&orderby)
'						objCreatedFile.Close 
						
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			call setInQuizOrderBy()
			'orderby=1
			'Quiz=1
			
			 if NUMTEST<>"" then
		   Quiz=NUMTEST
		   end if	
		   ' Response.write("QUIZ="&Quiz)
		 %>
		<p align="center"><b><%Response.write ("N."&Quiz) %></b>
  
		
		 <%
		if (Stato=0) then   
		  if CodiceSottopar<>"" then
	    
		    QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "'  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) order by Domande." & order(orderby)& " asc;"
		  
		   else
    QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) order by Domande." & order(orderby)& " asc;"
        
			end if

		   ' utilizzo il numero casuale per accedere al vettore che contiene le possibilità di ordinamento, potrò farlo anche per asc e desc	   
		else 
		'Definzione codice SQl della query per sapere quanti quiz ci sono
		 QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and Domande.VF=0;"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			'call setInQuizOrderBy()
			
 	 
		QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & " or Domande.In_Quiz=-1) order by Domande." & order(orderby)& " asc;"
		
		
		end if   
		'da passare a calcola risultato per ricreare la stessa query
		ordina=order(orderby)
		
%>
  </div> <!--aalert-succes -->
<%
			 '  response.write(QuerySQL&"<br>")			
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		 'response.write(QuerySQL)
		'Creazione di una pagina HTML dinamica con i test. 
		'Le domande sono individuate da un nome del tipo NAME=i, dove i e' il numero
		'della domanda. Il test e' indipendente dal numero di domande memorizzato.
		'Dopo la compilazione del test, la pagina richiama calcola_risultato.asp
		'che effettua il calcolo del risultato raggiunto.  
		If rsTabella.BOF=True And rsTabella.EOF=True Then %>
           <div class="alert alert-error">
                            Test non ancora disponibile!
                             </div>
                         <% rsTabella.close()%>
        <% else %>    
	 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		              <div class="box-content"> 
                      
                       <FORM name="formQuiz"  class='form-vertical' METHOD="POST" ACTION="calcola_risultato.asp?ordina=<%=ordina%>&Verifica=<%=Verifica%>&Stato=<%=Stato%>&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&CodiceAllievo=<%=CodiceAllievo%>&Quiz=<%=Quiz%>&orderby=<%=orderby%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>">
		  <%i=1 'inizializza la variabile i (contatore delle domande)
		 
		  dim objFSO,objCreatedFile
		  Const ForReading = 1, ForWriting = 2, ForAppending = 8
		  Set objFSO = CreateObject("Scripting.FileSystemObject")
		  
		 'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DEL NODO

'Create the FSO.
filecsvD=filecsv&"_D.csv"
filecsvR=filecsv&"_R.csv"
filecsvRE=filecsv&"_RE.csv"

urlAppD=Server.MapPath(homesito)&"/app/"&filecsvD
urlAppR=Server.MapPath(homesito)&"/app/"&filecsvR
urlAppRE=Server.MapPath(homesito)&"/app/"&filecsvRE
urlAppD=Replace(urlAppD,"\","/")
urlAppR=Replace(urlAppR,"\","/")
urlAppRE=Replace(urlAppRE,"\","/")

Set objCreatedFileD = objFSO.CreateTextFile(urlAppD, True)
Set objCreatedFileR = objFSO.CreateTextFile(urlAppR, True)
Set objCreatedFileRE = objFSO.CreateTextFile(urlAppRE, True)
response.write("Creo i file<br>")
response.write("ulr ="&urlAppR &"<br>")
response.write("ulr ="&urlAppD &"<br>")
response.write("ulr ="&urlAppR &"<br>")
response.write("ulr ="&urlAppRE &"<br>")

 
		  Do until rsTabella.EOF   ' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
		  
		 '  Set objFSO = CreateObject("Scripting.FileSystemObject")  
'			url="C:\Inetpub\umanetroot\anno_2013-2014\logEsegui395.txt"
'						Set objCreatedFile = objFSO.CreateTextFile(rsTabella("CodiceDomanda"), True)
'						objCreatedFile.WriteLine(QuerySQL & "orderby="&orderby)
'						objCreatedFile.Close 
		  
		  url=rsTabella.Fields("URL_Teoria")
    'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
'da sistemare url come questo sotto
' url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Replace(url,"\","/")
'sReadAll=url
'response.write(sReadAll)
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll
objTextFile.Close









		   if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
			<% ' se sono in modalità verifica aggiungo un bottone per la segnalazione della domanda
			   if Verifica=1 then %>  
			   <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore"> 
                  <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="right" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a></span>
			   
			  <%end if %>
               <B>
              <% if session("Admin")=true then%>
              
               
                <%=rsTabella.Fields("Quesito")%>
			  &nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &"."%>">.</a></B>
			   (<%=rsTabella.Fields("Cognome") & left(rsTabella.Fields("Nome"),1) &". RE="& rsTabella("RispostaEsatta")%>)
              <%else%>
              
              <%=rsTabella.Fields("Quesito")%>
			  &nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &"."%>">.</a></B>
			   (<%=rsTabella.Fields("Cognome") & left(rsTabella.Fields("Nome"),1) &"."%>)
		    
               <%end if%>
           
              
			  <%' verifico se devo inserire l'immagine come domanda o il testo plus
			      if rsTabella("Img")=1 then 
				     imgdomanda()
				  else%>    
                 <textarea rows="6" name="S1" value="ciao" cols="116"><%=Response.write(domandaplus())%> </textarea><br>
              <%end if %>
              
		  
		  <%else
		  ' aggiungo alla domanda la possibilità di sapere di chi è tramite il titolo dell'href
		  %>
          
        						  <div class="control-group">
									<% if session("Admin")=true then %>
                                    	 
                                         <h5> <%'=i & ") "%> <%=rsTabella.Fields("Quesito")%>&nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &". RE="& rsTabella("RispostaEsatta")%>">. </a></h5>
                                         <%else%>
                                          <h5> <%'=i & ") "%> <%=rsTabella.Fields("Quesito")%>&nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &"."%>">. </a></h5>
                                         <%end if%>
                                         <%
										 ' scrivo sul file
										 objCreatedFileD.WriteLine(replace(rsTabella.Fields("Quesito"),"'",Chr(96))&" - ")
										 %>  
										<div class="controls">
                                              <% if Verifica=1 then %>
                                            
			   <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore">   <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="right" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a></span>
			   </b>
			 
			 <br><br>
			  <%end if %>
	
		   <%end if %>
		   
			  <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="1">
			  <%=rsTabella.Fields("Risposta1")%><BR>
			  <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="2">
			  <%=rsTabella.Fields("Risposta2")%><BR> 
			  <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="3">
			  <%=rsTabella.Fields("Risposta3")%><BR> 
			  <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="4">
			  <%=rsTabella.Fields("Risposta4")%> <BR>
              
              <%
			   stringa=rsTabella("Risposta1")&" , " & rsTabella("Risposta2")&" , " & rsTabella("Risposta3")&" , " &rsTabella("Risposta4")&" - "
			   stringa=  Replace(stringa,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi 
   			   stringa=  Replace(stringa,Chr(39),Chr(96)) 
' 
			   objCreatedFileR.WriteLine(stringa)
			   objCreatedFileRE.WriteLine(rsTabella("RispostaEsatta")&" - ")
			  %>
			 	</div>
			</div>
            <hr>
			 
		   <% i = i+ 1 
			   rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande%>
		   <% Loop %>
		  <P>
			   <b>Inserisci codice di sessione:</b>
               <input type="text" class="input-mini" value="0" name="txtSessione"><br>
								<button type="button" onClick="invia_test();" class="btn btn-primary">Invia le <%=i-1%> risposte del test</button>
                                </P>
		   </FORM>
		<% End If %>
		<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
		   ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		  
          objCreatedFileD.Close
		  objCreatedFileR.Close
		  objCreatedFileRE.Close
                     
                      
%>
   <%END IF%>



               
                      </div>         
			        </div>
			      </div>
			    </div>
	
<% end if%>         
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         
 <script language="javascript" type="text/javascript">

function invia_test() {
	if (document.formQuiz.txtSessione.value=="0") 
	  if (confirm("Non hai inserito il codice per tracciare il quiz, inviare comunque?")) {  
		document.formQuiz.submit();	
	 }
	  
	 if (document.formQuiz.txtSessione.value!="0") 
	     document.formQuiz.submit();	
}

 
function PopUpWindow(w,h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
 
window.open('../cSocial/share.asp','share.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);

}
</script>
			 
	</body>

 </html>

