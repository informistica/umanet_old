<%@ Language=VBScript %>
<%
 Dim Num_Quiz,rand,Quiz,orderby
 on error resume next
%>
<!doctype html>
<%Function domandaplus()
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
End Function %>

<% Sub inserisci_immagini()
	
	if rsTabella("Img")=1 then
			QuerySQL="Select * from Domande_Img where Id_Domanda="& rsTabella("CodiceDomanda")&";"
			url= "../Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Domande/Img" ' vuole il percorso relativo della cartella
			url=Replace(url,"\","/")   
			Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)   
			%><div class="immagini" style="overflow:auto;"><%   
			do while not rsTabella2.eof
				   	
				   urlimg=url&"/"& rsTabella2("Url") ' aggiungo al percorso il nome del file
				   %>
				   <p align="center">
                   
				   <img src="<%=urlimg%>" border="1"> <br>
				  <%  'response.write(urlimg)
				     response.write(rsTabella2("Nome"))%><br>
				   </p>
				 <% rsTabella2.movenext
		   loop%>
           </div>
	
	<%end if
end sub%>

<% Response.Buffer=True
NUMTEST=request.querystring("NUMTEST")

 %>

<html>
<head>
   
   <title>Esegui Quiz con domande a risposta multipla</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
		 <meta charset="utf-8">

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
<%if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

  <% end if %> 
 


	<div id="navigation">
     
        <% 
		
 Dim order(9)
order(0)="" ' non lo uso 
order(1)="CodiceDomanda" 
order(2)="Quesito" 
order(3)="Risposta1"
order(4)="Risposta2"
order(5)="Risposta3"  
order(6)="Risposta4" 
order(7)="Data" 

Sub setInQuizOrderBy()
' genera un numero casuale per scegliere quale quiz e quale ordinamento per le domande  
	if CodiceSottopar<>"" then
	       QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
		   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "'  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0  and Domande.VF=0;"
		  
		   else
			if stato=1 then
			
			QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
					   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
					   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and Domande.VF=0;"  
			  else
				QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
					   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
					   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Multiple=0  and Domande.VF=0;"    
				end if
			end if
			
			Set rsTabella=ConnessioneDB.execute(QuerySQL)
			 if not isnull(rsTabella(0)) then
			   Num_Quiz=rsTabella(0)
			 else
			  Num_Quiz=0
			 end if  
             if Num_Quiz>0 then
				'response.write("NUM_QUIZ="&Num_Quiz)
				'response.write("stronzo dentro")
				randomize()
				do 
					rand=rnd()
				loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=Num_Quiz)
				Quiz=left((rand*5),1)
				'response.write("Quiz="&Quiz)
				
				' Response.write("QUIZ="&Quiz)
				 do 
					rand=rnd()
				loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=7)
				orderby=left((rand*5),1)
				'response.write("orderby="&orderby)
			end if
end sub %>
<%'response.write("Instring="& InStr("1234","0"))
'RispoD="4"
'RispoE="123"
'esa=0
'for k=1 to len(RispoD)
'    if ( InStr(RispoE,mid(RispoD,k,1))<>0) then
'	   esa=esa+1
'	end if   
'   'response.write( mid(RispoD,k,1) & "<br>" )
'next 
'response.write("esa="&esa)

  
    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
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
						<h1> <i class="icon-comments"></i> Quiz con risposte multiple </h1> 
                    
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
				 
                 
 <%
    Capitolo=Request.QueryString("Capitolo")
	if Capitolo="" then
	 Capitolo=Request.QueryString("TitoloCapitolo")
	end if
   Paragrafo=Request.QueryString("Paragrafo")
   if Paragrafo="" then
	 Paragrafo=Request.QueryString("TitoloParagrafo")
	end if
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
  Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if

 %>                
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h4> <i class="icon-reorder"></i>  <%Response.write (Capitolo)%>:  <%Response.write (Paragrafo)%>
   <% if Sottoparagrafo<>"" then
  Response.write ("/"&Sottoparagrafo)
                   end if%>
                        
                        </h4>
			          </div>
				     
                      
  <%Stato=Request.QueryString("Stato") '=0 se svolto test del paragrafo 1 se svolgo quello del modulo
   Stato0=Request.QueryString("Stato0")
   Modulo=Request.QueryString("Modulo") 
    Tutti=Request.QueryString("Tutti") 
	Cartella=Request.QueryString("Cartella") 
   'Raccolta dei dati digitati dall'utente e salvati nel cookie

'   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceAllievo=Session("CodiceAllievo")
   if (CodiceAllievo="") then
      response.Redirect("../home.asp")
   end if
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   CodiceTest = Request.QueryString("CodiceTest") ' se svolgo tutto il modulo (stato=1) contiene l'Id del modulo e non del paragrafo
   Verifica=clng(Request.QueryString("Verifica")) ' se sono stato chiamato da verifica il test il valore vale 1 , serve per segnalre le domande da correggere
    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"'"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	TestAbilitato=rsTabella.fields("TestAbilitato")
	rsTabella.close
if  (Session("Admin")=True) or (TestAbilitato=1) then ' else è alla fine%> 

 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				   
                    <%' Controllo subito se esiste almeno una domanda per quel modulo altrimenti salto tutto , chiudo end if prima dell'ultimo end if alla fine 
 if Stato=0 then ' verifico esistenza domande del paragrafo altrimenti del capitolo
 
  if CodiceSottopar<>"" then
	 
		   QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=1  and Lingua='"&Lingua&"';"
   
	   else
 QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and Domande.Multiple=1  and Lingua='"&Lingua&"';"
	end if
 
 
  
		   ' <>0 per escludere le domande degli stud che non faranno parte di quiz e <>-2 per escludere le domande inserite dall'admin (In_Quiz=-1) che non devono essere ancora visibili (che quindi metterò =-2
 
 else
 
 QuerySQL="SELECT * "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=1 and  Domande.Segnalata=0  and Lingua='"&Lingua&"';"
  
  CodiceTest=Modulo ' serve per inserire il codice del test (cioè del modulo) anzichè 1_0 nella tabella risultati1
end if	   
  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
   If rsTabella.BOF=True And rsTabella.EOF=True Then 
     rsTabella.close()  %>
     

                     
                      <div class="alert alert-error">
                   Test non ancora disponibile!
                     </div>
       <% Else %>
 
          

<%if (Stato=0) then 
	 'Definzione codice SQl della query per ricercare le domande del paragrafo 
	 ' mi serve anche il titolo del paragrafo per ricostruire il nome del file che contiene la domanda plus
	   
	   ' codice per la generazione di quiz random : seleziono il numero di quiz presenti 
	   
	   if CodiceSottopar<>"" then
	 
		     QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
	   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
	   " WHERE Domande.Multiple=1 and Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0   and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "';"
   
	   else
  QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
	   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
	   " WHERE Domande.Multiple=1 and  Domande.Segnalata=0  and Domande.Id_Arg='" & CodiceTest & "'  and Lingua='"&Lingua&"';"
	end if
	   
	   
	
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		'response.Write(QuerySQL) 
		Num_Quiz=rsTabella(0) 
		'response.write("NUM_QUIZ="&Num_Quiz)
		'genero un numero casuale compreso tra 1 e Num_Quiz per selezionare le domande di quel quiz
		
		if Num_Quiz=-1 then
		
			' se ci sono solo le domande dell'admin  (In_Quiz=-1) non eseguo la generazione rnd per il NumQuiz
			' che si bloccherebbe perchè il numero generato non sarà mai minore di -1 , ma la faccio qua solo
			' per scegliere il tipo di ordinamento
			Quiz=-1
			'orderby=1
			randomize()
			do 
					rand=rnd()
			loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=7)
			orderby=left((rand*5),1)
			'response.write("Order by="&orderby)
			
		else
 
				call  setInQuizOrderBy()

		end if 
		
	   ' Response.write("QUIZ="&Quiz)

	 
	   
	    if CodiceSottopar<>"" then
	  
	    QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
	   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
	   " WHERE Domande.Multiple=1 and Domande.Id_Arg='" & CodiceTest & "'  and  Domande.Segnalata=0 and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"'  order by Domande." & order(orderby)& " asc;"
	   
   
	   else
  QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
	   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
	   " WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Arg='" & CodiceTest & "' AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"'  order by Domande." & order(orderby)& " asc;"
	   
	end if
	   
	   
	   
	   %>

<%else ' if stato=0
'Definzione codice SQl della query per sapere quanti quiz ci sono
 QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
   " WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Mod='" & Modulo & "' or Domande.In_Quiz=-1  and Lingua='"&Lingua&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Num_Quiz=rsTabella(0) 
	'response.write("NUM_QUIZ="&Num_Quiz)
	'response.write(QuerySQL)
	if Num_Quiz=-1 then
	' se ci sono solo le domande dell'admin  (In_Quiz=-1) non eseguo la generazione rnd per il NumQuiz
		' che si bloccherebbe perchè il numero generato non sarà mai minore di -1 , ma la faccio qua solo
		' per scegliere il tipo di ordinamento
		Quiz=-1
		'orderby=1
		randomize()
		do 
				rand=rnd()
		loop until (clng(left((rand*5),1))>0) and (clng(left((rand*5),1))<=7)
		orderby=left((rand*5),1)
		'response.write("Order by="&orderby)
	else 
		call setInQuizOrderBy()
		'Quiz=1
	end if 
	
	 if NUMTEST<>"" then
				  if strcomp(NUMTEST,"-1")=0 then
					stringaQuery="and Domande.In_Quiz like '%' "
					' response.write(stringaQuery)
				  else
					stringaQuery="and (Domande.In_Quiz="&Quiz &" or Domande.In_Quiz=-1)  "
				   Quiz=NUMTEST
				  end if 
			 else
				stringaQuery="and (Domande.In_Quiz="&Quiz &" or Domande.In_Quiz=-1)  "
			 end if  
	 %>
<div align="center"><font size="4" color=#FF0000><b><%'Response.write ("N."&Quiz) %></b></font></div> <!-- stampa il titolo del test -->

 <%
  
	

QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
   " WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Mod='" & Modulo & "' "& stringaQuery &"  and Lingua='"&Lingua&"'  order by Domande." & order(orderby)& " asc;"


end if %> 
<span align="center">
<% if Verifica=1 Then %>
  <div class="alert alert-success">ESEGUI CONTROLLO QUALITA` DEL TEST   (vere 0,1,2,3,4) 
 
  
<%else%>
   <div class="alert alert-success"> ESEGUI TEST A RISPOSTA MULTIPLA  (vere 0,1,2,3,4) 
<%end if%>  
<%
		if NUMTEST<>"" then
			Quiz=NUMTEST
			 if strcomp(NUMTEST,"-1")=0 then 
			    Response.write ("<p align='center'><b>Tutti i quiz</b>")
			 else
				 %>
				<p align="center"><b><%Response.write ("N."&Quiz) %></b> <!-- stampa il titolo del test -->			
			 <%
			end if
		  end if%>
		  </div>
</span>
		  <%

Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'response.write(QuerySQL)
'Creazione di una pagina HTML dinamica con i test. 
'Le domande sono individuate da un nome del tipo NAME=i, dove i e' il numero
'della domanda. Il test e' indipendente dal numero di domande memorizzato.
'Dopo la compilazione del test, la pagina richiama calcola_risultato.asp
'che effettua il calcolo del risultato raggiunto.      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then %>
  <div class="alert alert-error">Test non ancora disponibile!</div>
  <p><h5><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
</H5>
<% Else %>

 


  <FORM  name="formQuiz" class="form-vertical form-bordered" METHOD="POST" ACTION="5_calcola_risultato_multiple_1.asp?Lingua=<%=Lingua%>&Verifica=<%=Verifica%>&Stato=<%=Stato%>&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&CodiceAllievo=<%=CodiceAllievo%>&Quiz=<%=Quiz%>&orderby=<%=orderby%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>&NUMTEST=<%=NUMTEST%>">
  <%
    dim objFSO,objCreatedFile
		  Const ForReading = 1, ForWriting = 2, ForAppending = 8
		  Set objFSO = CreateObject("Scripting.FileSystemObject")
		 
  i=1 'inizializza la variabile i (contatore delle domande)
  Do until rsTabella.EOF  ' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
  
  	  url=rsTabella.Fields("URL_Teoria")
    'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
url=Replace(url,"\","/")
'sReadAll=url
'response.write(sReadAll)
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll

objTextFile.Close
  
  %> <div class="control-group">
   <%if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
     <% if session("admin")=true then%>
			<a  href="../cDomande/inserisci_valutazione.asp?Lingua=<%=Lingua%>&traduzione=1&Multiple=<%=rsTabella("Multiple")%>&ORA=<%=left(rsTabella("Ora"),5)%>&DATA=<%=rsTabella("Data")%>&Tipodomanda=<%=rsTabella("Tipo")%>&Cartella=<%=rsTabella("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabella("CodiceDomanda")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=rsTabella("Titolo")%>&Quesito=<%=rsTabella("Quesito")%>&R1=<%=rsTabella("Risposta1")%> &R2=<%=rsTabella("Risposta2")%>&R3=<%=rsTabella("Risposta3")%>&R4=<%=rsTabella("Risposta4")%>&RE=<%=rsTabella("RispostaEsatta")%>&MO=<%=rsTabella("ID_Mod")%>&VAL=<%=rsTabella("Voto")%>&VF=<%=rsTabella("VF")%>&URL=<%=rsTabella("URL_Teoria")%>&INQUIZ=<%=rsTabella("In_Quiz")%>&VALINQUIZ=<%=rsTabella("In_QuizStud")%>&Segnalata=<%=rsTabella("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">                                                     
			   <b>(<%=rsTabella("CodiceDomanda")%>)</b></a>
			   <%end if%>
	
	<% ' se sono in modalità verifica aggiungo un bottone per la segnalazione della domanda
	   if Verifica=1 then %>  
	  	 <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore">   <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="top" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a></span> 
	  <%end if %>
      <label class="control-label"><B> <%'=i & ") "%><%=rsTabella.Fields("Quesito")%>
	  &nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &"."%>">.</a></B></label>
	  
   
	 <textarea class="input-block-level" rows="<%=round((len(domandaplus()))/15)%>" name="S1" value="ciao" ><%=Response.write(domandaplus())%> </textarea><br>
  <% 
  
     else
  ' aggiungo alla domanda la possibilità di sapere di chi è tramite il titolo dell'href
  %>
     <label class="control-label"><B> 
	  <% if session("Admin")=true then%>
         <%'=i & ") "%> <%=rsTabella.Fields("Quesito")%>&nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) & ". RE=" & rsTabella("RispostaEsatta") & "."%>">.</a></B>
	  <% else%>
	  <%'=i & ") "%> <%=rsTabella.Fields("Quesito")%>&nbsp;<a href="#" title="<%=rtrim(rsTabella.Fields("Cognome")) &" "& left(rsTabella.Fields("Nome"),1) &"."%>">.</a></B>
      <%end if%>
      </label>
	   <% if session("admin")=true then%>
											  <a  href="../cDomande/inserisci_valutazione.asp?Lingua=<%=Lingua%>&traduzione=1&Multiple=<%=rsTabella("Multiple")%>&ORA=<%=left(rsTabella("Ora"),5)%>&DATA=<%=rsTabella("Data")%>&Tipodomanda=<%=rsTabella("Tipo")%>&Cartella=<%=rsTabella("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabella("CodiceDomanda")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=rsTabella("Titolo")%>&Quesito=<%=rsTabella("Quesito")%>&R1=<%=rsTabella("Risposta1")%> &R2=<%=rsTabella("Risposta2")%>&R3=<%=rsTabella("Risposta3")%>&R4=<%=rsTabella("Risposta4")%>&RE=<%=rsTabella("RispostaEsatta")%>&MO=<%=rsTabella("ID_Mod")%>&VAL=<%=rsTabella("Voto")%>&VF=<%=rsTabella("VF")%>&URL=<%=rsTabella("URL_Teoria")%>&INQUIZ=<%=rsTabella("In_Quiz")%>&VALINQUIZ=<%=rsTabella("In_QuizStud")%>&Segnalata=<%=rsTabella("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">                                                     
			   <b>(<%=rsTabella("CodiceDomanda")%>)</b></a>
			   <%end if%>
	  <% if Verifica=1 then %>
	   <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore">  <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="top" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a></span>
	   </b>
	  
	 
	  <%end if %>
	  </LEGEND>
   
   <%end if 
   inserisci_immagini() ' funzione per inserire eventuali immagini alla domanda
   %>

      <label class='checkbox'>
      <INPUT TYPE="checkbox" NAME="C1_<%=i%>" VALUE="1">
      <%=rsTabella.Fields("Risposta1")%><BR>
      </label>
      <label class='checkbox'>
      <INPUT TYPE="checkbox" NAME="C2_<%=i%>" VALUE="2">
      <%=rsTabella.Fields("Risposta2")%><BR> 
      </label>
      <label class='checkbox'>
      <INPUT TYPE="checkbox" NAME="C3_<%=i%>"  VALUE="3">
      <%=rsTabella.Fields("Risposta3")%><BR> 
      </label>
      <label class='checkbox'>
      <INPUT TYPE="checkbox" NAME="C4_<%=i%>"  VALUE="4">
      <%=rsTabella.Fields("Risposta4")%> <BR>
      </label>
      <label class='checkbox'>
       <INPUT TYPE="checkbox" NAME="C5_<%=i%>"  VALUE="5">
       
      Nessuna <BR>
      </label>
 

</div>
   <% i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande%>
   <% Loop %>
    
 
    <div class="form-actions">
	<P>
			   <b>Inserisci codice di sessione:</b>
               <input type="text" class="input-mini" value="0" name="txtSessione"><br>
								<button type="button" onClick="invia_test();" class="btn btn-primary">Invia le <%=i-1%> risposte del test</button>
                                </P>									 
									
     
  </div>
   </FORM>
<% End If %>
<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
   ConnessioneDB.Close : Set ConnessioneDB = Nothing %>
    
    
                     
                     
                     
                      
                      
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
    <%end if ' chiudo l'if preliminare che controlla l'esistenza di almeno una domanda %>		

<% else 
 %>
 <BODY onLoad="showText();"> 
 <%    
	
  ' Response.Redirect "../home.asp"
      end if %>

<script language="javascript" type="text/javascript">

function invia_test() {
	if (document.formQuiz.txtSessione.value=="0") 
	  if (confirm("Non hai inserito il codice per tracciare il quiz, inviare comunque?")) {  
		document.formQuiz.submit();	
	 }
	  
	 if (document.formQuiz.txtSessione.value!="0") 
	     document.formQuiz.submit();	
}

</script>

 </html>

