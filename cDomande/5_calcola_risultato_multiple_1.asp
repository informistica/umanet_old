<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Risultato TEST </title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
		<meta charset="utf-8">

<!-- Bootstrap -->
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

	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
      
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>'>
	<div id="navigation">
     
        <% 
		'on error resume next
    Response.Buffer=True 
   Dim  Quiz
   
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,Segnalata
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
    Dim order(9)
order(0)="" ' non lo uso 
order(1)="CodiceDomanda" 
order(2)="Quesito" 
order(3)="Risposta1"
order(4)="Risposta2"
order(5)="Risposta3"  
order(6)="Risposta4" 
order(7)="Data" 
order(8)="Risposta5" 
	
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <!-- #include file = "tabella_corrispondenze.inc" -->

        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Risultato test a risposta multipla</h1> 
                    
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
							<a href="#">Risultato</a>
                             
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                  <%
	   CodiceTest = Request.Cookies("Dati")("CodiceTest")
	    Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
  ' CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
  Verifica=Request.QueryString("Verifica")
  if Verifica="" then 
     Verifica=0
  end if
  
   Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if
  
  
Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
   DataTest = gira_data()
   
     SessioneQuiz=Request.Form("txtSessione")
	 Quiz=clng(Request.QueryString("Quiz"))
   Stato=Request.QueryString("Stato")
   Stato0=Request.QueryString("Stato0")
   Tutti=Request.QueryString("Tutti")
   Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   Modulo=Request.QueryString("Modulo")
   CodiceTest=Request.QueryString("CodiceTest") ' se svolgo tutto il modulo contiene l'id del modulo
   CodiceAllievo=Request.QueryString("CodiceAllievo") 
   'parametro generato random da esegui test per scegliere il quiz da eseguire di cui ora calcolo il risultato
 
   orderby=clng(Request.QueryString("orderby"))
   'Definizione query SQL per contare il numero di domande del test.
 NUMTEST=request.querystring("NUMTEST")

Function calcola_risposta(codice_risposta)
    ' scorre tutta la tabella, appena trovo il codice restituisco il numero associato 
	' esempio codice_risposta=1010 -> calcola_risposta=13
	for j=0 to 16
	 ' response.write("<br>"& v1(j) & "=?" & codice_risposta)
	    if v1(j) = codice_risposta then  
		    calcola_risposta=v2(j)
			'response.write("OK calcola_risposta="&calcola_risposta)
		end if
	next 
End Function

 
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
 

if (Stato=0) then 
 'Definzione codice SQl della query per ricercare le domande del paragrafo 
  ' QuerySQL="SELECT count(*) " &_
'             "FROM Domande " &_
'             "WHERE Domande.Multiple=1 and Domande.Id_Arg='" & CodiceTest & "' AND Domande.In_Quiz="&Quiz&" or Domande.In_Quiz=-1 ;"
 if CodiceSottopar<>"" then
			 QuerySQL="SELECT count(*) " &_
             "FROM Domande " &_
             "WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Arg='" & CodiceTest & "'  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "'  AND (Domande.In_Quiz="&Quiz&"  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"' ;"
	   else
 
   QuerySQL="SELECT count(*) " &_
             "FROM Domande " &_
             "WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Arg='" & CodiceTest & "' AND (Domande.In_Quiz="&Quiz&"  or Domande.In_Quiz=-1)   and Lingua='"&Lingua&"' ;"
	end if

 
   
'    Assegna alla variabile il risultato della query prodotta utilizzando il metodo Execute(stringa della query) dell'oggetto connessione
else 
'Definzione codice SQl della query per ricercare le domande del modulo
'QuerySQL="SELECT count(*) " &_
'             "FROM Domande " &_
'             "WHERE Domande.Multiple=1 and Domande.Id_Mod='" & Modulo & "' AND Domande.In_Quiz="&Quiz&"or Domande.In_Quiz=-1 ;"
QuerySQL="SELECT count(*) " &_
             "FROM Domande " &_
             "WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Mod='" & Modulo & "' "& stringaQuery&"  and Lingua='"&Lingua&"' ;"
end if   

    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    NumDom=rsTabella(0).value 'Assegno a NumDom numero delle domande
if (Stato=0) then 
      if CodiceSottopar<>"" then
			
			 QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Cognome,Nome " &_
             "FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud   " &_
           "  WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Arg='" & CodiceTest & "'   and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "'   AND  (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"'   order by Domande." & order(orderby)& " asc;"
	   else
 
   QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Cognome,Nome " &_
             "FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud   " &_
           "  WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Arg='" & CodiceTest & "' AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"'   order by Domande." & order(orderby)& " asc;"
	end if
 
 

 
else

			
 '  QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud  " &_
'             "FROM Domande " &_
'             " WHERE Domande.Multiple=1 and Domande.Id_Mod='" & Modulo & "' AND Domande.In_Quiz=" &Quiz & " or Domande.In_Quiz=-1 order by Domande." & order(orderby)& " asc;"
  QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Cognome,Nome  " &_
             "FROM  Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud   " &_
             " WHERE Domande.Multiple=1 and  Domande.Segnalata=0 and Domande.Id_Mod='" & Modulo & "' "& stringaQuery&"  and Lingua='"&Lingua&"'  order by Domande." & order(orderby)& " asc;"

end if  
   
   
		   
   
   
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    'response.write(QuerySQL)
  'Calcolo del numero di risposte esatte.  
  i=1
  inbianco=0
  RisposteKO = 0   			  'contatore delle risposte esatte
  RisposteOK = 0   			  'contatore delle risposte errate
  ReDim RispDate(NumDom+1)    'dimensionamento dell'array dinamico che tiene traccia delle risposte date
  ReDim RispEsatte(NumDom+1)  'dimensionamento dell'array dinamico che tiene traccia delle risposte esatte
  ReDim RispDate1(NumDom+1)    'dimensionamento dell'array dinamico che tiene traccia delle risposte date
  ReDim RispEsatte1(NumDom+1) 
  ReDim Errori(NumDom+1) 	  'dimensionamento dell'array dinamico che tiene traccia degli errori
  ReDim RispDateStr(NumDom+1)    'dimensionamento dell'array dinamico che tiene traccia delle risposte date in formato stringa
  ReDim RispEsatteStr(NumDom+1)  'dimensionamento dell'array dinamico che tiene traccia delle risposte esatte in formato stringa
 
  Stringone="" ' conterrò la concatenazione delle stringhe di  tutte le risposte esatte, mi interesserà la sua lunghezza per il calcolo del 
  ' risultato
  Do While not(rsTabella.EOF) ' per ogni risposta confronta la risposta esatta con quella data dall'utente
    if rsTabella.Fields("RispostaEsatta")=0 then
	 RispostaEsatta=5
	else
	 RispostaEsatta=rsTabella.Fields("RispostaEsatta") 'legge dal risultato e memorizza in una variabile d'appoggio la risposta esatta
    end if
	 ' creo lo stringone
	 Stringone = Stringone & Cstr(RispostaEsatta)
	 'legge il valore associato all'oggetto avente per nome il numero contenuto nella variabile i, in base al valore ricava la risposta data  
  
	' per leggere i checkbox delle risposte
 
	SELECT CASE Request.Form("C1_" & i & "")
     CASE "1"
       Risposta1=1
     CASE ELSE
     	Risposta1=0
     END SELECT  
	 SELECT CASE Request.Form("C2_" & i & "")
     CASE "2"
       Risposta2=1
     CASE ELSE
        Risposta2=0
     END SELECT  
	 SELECT CASE Request.Form("C3_" & i & "")
     CASE "3"
       Risposta3=1
     CASE ELSE
     	Risposta3=0
     END SELECT  
	 SELECT CASE Request.Form("C4_" & i & "")
     CASE "4"
       Risposta4=1
     CASE ELSE
     	Risposta4=0
     END SELECT  
	   SELECT CASE Request.Form("C5_" & i & "")
     CASE "5"
       Risposta5=1
     CASE ELSE
     	Risposta5=0
     END SELECT  
	 
	 ' per leggere il checkbox della segnalazione
	  SELECT CASE Request.Form("Check" & i & "")
     CASE "1"
       Segnalata=1
     CASE ELSE
     	Segnalata=0
     END SELECT  
	 
     dim a
     a=1
	 
	 
	'response.write(a&"<br>ciao")
				
	 ' creo una stringa binaria con le 4 risposte 
	 CodiceRisposta= Cstr(Risposta1) & Cstr(Risposta2) & Cstr(Risposta3) & Cstr(Risposta4) & Cstr(Risposta5)   
    'response.write("<br>352")    
	RispostaData=clng(calcola_risposta(CodiceRisposta))
	   
	 '********************************
	 ' per rendere più raffinata la correzione ed evitare di segnare tuta rossa una domanda che è in parte giusta
	 ' non devo convertire in numero e confrontare l'uguaglianza dei numeri (RispData=RIspEsatta) 
	 ' ma analizzare i singoli caratteri formanti le due stringhe (quella della risdata e quella della rispesatta
	 'per fare ciò  
	 
	' response.write("<br>RispostaData="&RispostaData)
'	  IF (RispostaData=0) THEN
'		 
'	 response.write("=0")
'	 end if
' 
		 
				
	'response.write("<br>373;"&i&"<br>RispostaData="&RispostaData) 
	
     RispDate(i) = RispostaData ' memorizza il valore della risposta data (i) tipo numero124
	 
	 RispDateStr(i)=CodiceRisposta ' ' memorizza il valore della risposta data (i) tipo stringa "124" 
   
     
	 IF (RispostaData=0) THEN
		RispDate1(i)= "IN BIANCO" 
		inbianco=inbianco+1
     ELSE
	  
	  ' la metto in commento perchè nel caso di risposta singola conteneva il testo della risposta, ma ora 
	  ' devo usare un altro modo perchè ci spossono essere più risposte
      ' RispDate1(i) =  rsTabella.Fields(RispostaData).value ' 
       'Response.Write(rsTabella.Fields(1+RispostaData).value)
     END IF
      IF (Segnalata=1) THEN
	   %> 
	  	<!-- #include file = "inserisci_segnalazione_include.asp" -->
	            
	<% 
		    RispDate1(i)= "SEGNALATA"
			 
	 END IF
	 
	 
     RispEsatte(i) = RispostaEsatta ' memorizza nel vettore risposte esatte il valore della risposta esatta (i)
	 RispEsatteStr(i)=Cstr(RispostaEsatta)
     'RispEsatte1(i) = rsTabella.Fields(1+RispostaEsatta).value ' *****si blocca qua non gli piace l'assegnazione
    ' RispEsatte1(i)="Ciao"
	'********************************
	'qua devo modificare il tipo di confronto, devo analizzare le stringhe componenti le risposte
	 
	 'IF (RispostaEsatta=RispostaData) THEN  ' se sono uguali incrementa il numero delle risposte ok e pone a 0 l'elemento i del vettore errori 
'           RisposteOK = RisposteOK +1
'           Errori(i)=0 				'0 = domanda i esatta
'     ELSE       					'1 = domanda i errata  
'           
'		   Errori(i)=1				'se sono diversi incrementa il numero delle risposte ko e pone a 1 l'elemento i del vettore errori
'           RisposteKO = RisposteKO +1 
'		   ' qua devo aggiungere le modifiche per distinguere la gravità della risposta errata, tutte errate, 2 su 3 , ecc...
'		     
'     END IF
    ' analizzo carattere per carattere della rispostadata per vedere se esiste nella stringa della risposta esatta 
    errata=false
	' response.write("<br>RispostaData="&RispostaData)
	for k=1 to len(cstr(RispostaData))
	  ''response.write("411=---- <br><br>")
		if (InStr(cstr(RispostaEsatta),mid(cstr(RispostaData),k,1))<>0) then
		   
		   RisposteOK=RisposteOK+1
		   if errata=false then
			Errori(i)=0 
		   end if
		else
		   RisposteKO=RisposteKO+1
		   Errori(i)=1
		   errata=true	
		end if   
		'response.write("errata=" & errata & " ---- <br><br>")
		 
    next 
	
 
	 if errata=true then
		Errori(i)=1
	 end if
	 'caso ad hoc per evitare che se do una sola risposta ed è giusta, ma ce ne sono altre in risposteesatte, mi dia Errori(i)=0
	 ' come nel caso della domanda 12 che se indovino la prima rispo me di tutto verde
      if (len(cstr(RispostaData))<len(cstr(RispostaEsatta))) or (len(cstr(RispostaData))>len(cstr(RispostaEsatta)))  then
	    Errori(i)=1
	  end if

     i = i + 1						' incrementa i
     rsTabella.MoveNext 			' passa alla prossima domanda
   Loop 
   
   'Calcolo della percentuale di domande corrette.
   '******************************** 
   'qua devo modificare il calcolo del risultato 
   ' RisposteOK terrà conto dell'analisi dei singoli caratteri delle stringhe e il denominatore dipendere dalla lunghezza dello stringone delle 
   ' risposteesatte di tutto il quiz
    NumRispTotale=len(Stringone) 
                
				 %>
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  Quiz su "<%=Capitolo%> : <%=Paragrafo%>"</h3>
			          </div>
				      <div class="box-content">
                      

   
 
   <!-- stampa la tabella per offire l'opportunità di visualizzare le correzioni -->
  
  <div class="box-content nopadding">
								<table class="table  table-nomargin ">
									<thead>
										<tr>
											<th class='hidden-480'>Domanda</th>
											<th>Quesito</th>
											<th class='hidden-350'>Risposta data</th>
											<th>Risposta esatta</th>
											<th class='hidden-350'>Approfondisci</th>
										</tr>
									</thead>
			
                                    <tbody>
                                    
                                 <%	rsTabella.Movefirst ' torna all'inizio delle domande
		   	    i=1
				ok=0
				ko=0
				
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
 
				
				
				Do While Not rsTabella.EOF %>
				<% numRDate=len(cstr(RispDate(i))) ' lunghezza della risposta, cioè quante sono es. 124 -> len=3
				   numREsatte=len(cstr(RispEsatte(i)))
				   contDate=1
				   ContEsatte=1
				   ' mi serve per vedere quante righe devo predisporre nella tabella per quella domanda
				   if numRDate>numREsatte then
				      maxLen=numRDate
				   else
				      maxLen=numREsatte
				   end if
				   
				   if Errori(i)=1 then %>  <!-- se la risposta è errata usa il colore rosso -->
					  <tr><td  class='hidden-480' valign=top rowspan=<%=maxLen%>><b><center><%=i%></center></b></td>
					  <td rowspan=<%=maxLen%>> <font color="red"><%=rsTabella.Fields("Quesito")%></font></td>
					  <!-- Devo inserire il controllo per stabilire se usare il rosso o il verde
					   se RispDate esiste in RispEsatte uso verde altrimenti rosso : questo vale per la prima riga-->
					  <%' faccio il controllo bidirezionale  
					    
						if strcomp(RispDate1(i),"IN BIANCO")=0 then%>
						 <td class='hidden-350'> <font color="red">
									<%=RispDate1(i)%>
									
									</font>
							 </td>
						
						<%else
							if InStr(RispEsatte(i),left(RispDate(i),1)) then %>
							  <td class='hidden-350'> <font color="green">
										<%=rsTabella.Fields(1+clng(left(RispDate(i),1)))%>
										
								  </font> 
							  </td>
						  <%else %>
							 <td class='hidden-350'> <font color="red">
									<%=rsTabella.Fields(1+clng(left(RispDate(i),1)))%>
									
									</font>
							 </td>
						  <%end if%>
                     <% end if%>
                      
					   <% if InStr(RispDate(i),left(RispEsatte(i),1)) then %>
					     <td> <font color="green">
						 		<%=rsTabella.Fields(1+clng(left(RispEsatte(i),1)))%>
                                <%'ok=ok+1%>
                         </font> </td>
					  <%else %>
					     <td> <font color="red">
						 		<%=rsTabella.Fields(1+clng(left(RispEsatte(i),1)))%>
                                 <%ko=ko+1%>
                              </font> </td>
					  <%end if%>
                      <%
					    ' url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
						url=rsTabella.Fields("URL_Teoria")
    'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
url=Replace(url,"\","/")
'url=url3
'response.write(url)
' GESTION ERRORI 


' COMMENTO PER CORREGGERE 
' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)

' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll
'sReadAll=url
'response.write(sReadAll)
objTextFile.Close
					  %>
		 
					  <td class='hidden-350' valign=top rowspan=<%=maxLen%>>  <center>
                       <% if session("Admin")=true then%>
              
               <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="left" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a>
					 
              <%else%>
              
              <a data-original-title="Spiegazione" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="left" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a>
					 
			    <%end if%>
                    
                    </td>
					  </tr>
					  <%j=2
						contDate=contDate+1 ' cioè 2
						contEsatte=contEsatte+1
						'analizzo la risposta e per ogni numero che la forma visualizzo il campo corrispondente 
						while (j<=maxLen)%>
						 <tr>
						 <%if contDate<=numRdate then %>
						    <%if InStr(RispEsatte(i),mid(RispDate(i),j,1)) then %>
						  	    <td class='hidden-350'><font color="green"><%=rsTabella.Fields(1+clng(mid(RispDate(i),j,1)))%></font></td>
                                   <%ok=ok+1%>
							<%else %>
								<td class='hidden-350'><font color="red"><%=rsTabella.Fields(1+clng(mid(RispDate(i),j,1)))%></font></td> 
                                   <%ko=ko+1%>  
							<%end if %>
						 
						 <%else%>
						   <td class='hidden-350'>&nbsp; </td>
						 <%end if	
						  if contEsatte<=numREsatte then %>
						     <%if InStr(RispDate(i),mid(RispEsatte(i),j,1)) then %>
						  	    <td class='hidden-350'><font color="green"><%=rsTabella.Fields(1+clng(mid(RispEsatte(i),j,1)))%></font></td>
                                <%ok=ok+1%>
							<%else %>
								<td class='hidden-350'><font color="red"><%=rsTabella.Fields(1+clng(mid(RispEsatte(i),j,1)))%></font></td>   
                                  <%ko=ko+1%>    
							<%end if %>
						 <%else%>
						   <td class='hidden-350'>&nbsp; </td>
						 <%end if	
						j=j+1
						contDate=contDate+1  
						contEsatte=contEsatte+1
						%>
						 </tr>
						<%wend%>
						</font>		   
					 
                  
			    <% else %>      <!-- se la risposta è correta usa il colore verde -->
	
									  <tr><td  class='hidden-480' valign=top rowspan=<%=maxLen%>><b><center><%=i%></center></b></td>
					  <td rowspan=<%=maxLen%>> <font color="green"><%=rsTabella.Fields("Quesito")%></font></td>
					  <td class='hidden-350'> <font color="green"> <%=rsTabella.Fields(1+clng(left(RispDate(i),1)))%></font> </td>
					  <td> <font color="green"> <%=rsTabella.Fields(1+clng(left(RispEsatte(i),1)))%></font></td>
					  <td valign=top  class='hidden-350' rowspan=<%=maxLen%>>
                       
                      <%ok=ok+1%><center>
						  <span data-original-title="Spiegazione"  class="btn" rel="popover" data-trigger="hover" title="" data-placement="left" data-content="<%=sReadAll%>">
						  <i class="icon-question-sign"></i></center></span>
					 
					  </td>
					  </tr>
					  <%j=2
						contDate=contDate+1 ' cioè 2
						contEsatte=contEsatte+1
						'analizzo la risposta e per ogni numero che la forma visualizzo il campo corrispondente 
						while (j<=maxLen)%>
						 <tr>
						 <%if contDate<=numRdate then %>
						  <td class='hidden-350'><font color="green"><%=rsTabella.Fields(1+clng(mid(RispDate(i),j,1)))%></font></td>     
                           <%ok=ok+1%>
						 <%else%>
						   <td>&nbsp; </td>
						 <%end if	
						  if contEsatte<=numREsatte then %>
                            
							  <td><font color="green"><%=rsTabella.Fields(1+clng(mid(RispEsatte(i),j,1)))%></font></td>
						 <%else%>
						   <td>&nbsp; </td>
						 <%end if	
						j=j+1
						contDate=contDate+1  
						contEsatte=contEsatte+1
						%>
						 </tr>
						<%wend%>
						</font>		   


				<%End if%>
								 
				</tr>
				<% rsTabella.movenext
				i=i+1
				Loop %>	
                </tbody>
			</table> 
</div>
  
                                    
                                    
                  <% 
             'Risultato = (RisposteOK/(i-1))*100
	'response.write("<br>RispoOk="&RisposteOK)
	'response.write("<br>Len stringone="&NumRispTotale)
	'Risultato = (RisposteOK-(RisposteKO/2))/(NumRispTotale)*100
    Risultato = (NumRispTotale - ko)/NumRispTotale*100
	if Risultato<0 then
	Risultato=0
	end if
	'Risultato_relativo = (RisposteOK/(i-inbianco-1))*100
   Risultato_relativo=0
    
    DataTest=date()
   'Esecuzione della query per inserire il risultato del test nella tabella Risulati
   if (clng(Verifica)<>1) then ' inserisco i risultati solo se non sono in modalità verifica
   
	   if (Stato=0) then 
		   QuerySQL="  INSERT INTO Risultati (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Lingua) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0)   & "'," &Quiz & "," &SessioneQuiz  & ",2,'"&Lingua&"';"
	   else 
			QuerySQL="  INSERT INTO Risultati1 (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Lingua) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0)   & "'," &Quiz & "," &SessioneQuiz  & ",2,'"&Lingua&"';"
	   end if 
	 '  response.write(QuerySQL)
	   ConnessioneDB.Execute QuerySQL 
   end if
   
   
   
  
  if (Round(Risultato,0)*10/100)<6 then
   Response.Write("<H4><span class='alert-error'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*10/100 &"</span></H4>")
   else
   Response.Write("<H4><span class='alert-success'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*10/100 &"</span></H4>")
   end if
 
   
   Response.Write("<H5>  Hai risposto  <span class='alert-success'> correttamente " & ok-1 &"</span> volte, e  <span class='alert-error'>sbagliato "& ko & "</span>  volte<BR></h5>") 

  %>
  
   <br>
             
                                    
                                      <a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&Tutti=<%=Tutti%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>" target="_blank"><input type="button" class="btn" value="Continua verifica e lascia aperta questa pagina"></i>  </a>      <br><br>
       <a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&Tutti=<%=Tutti%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>" target="_Top"><input type="button" class="btn-primary" value="Continua verifica e chiudi questa pagina"></i>  </a>         
                      
                      
									
  
  
   	<p>     
   		
		
				 
				 
                   
                   
 <!-- Non chiude ?
		  			  <div class="box-content"> 
                      
            			   <h6 align="center"><a onClick="javascript:window.close();"> Chiudi </a></h6> 
                      </div>  
                    -->         
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

