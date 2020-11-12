<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Risultati Test</title>   
   
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

	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  
	<div id="navigation">
     
        <% 
 
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,Segnalata
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Dim order(8)
	order(0)="" ' non lo uso 
	order(1)="CodiceDomanda" 
	order(2)="Quesito" 
	order(3)="Risposta1"
	order(4)="Risposta2"
	order(5)="Risposta3"  
	order(6)="Risposta4" 
	order(7)="Data" 	
	on error resume next
	Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")	
 
  CodiceTest=Request.QueryString("CodiceTest") ' se svolgo tutto il modulo contiene l'id del modulo
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Risultato test </h1> 
                    
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
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				       <h3> <i class="icon-reorder"></i>  <%=Request.QueryString("Capitolo")%>
					   <% if strcomp(CodiceTest,"1_0")<>0 then%>
                        : <%=Request.QueryString("Paragrafo")%>
						<%end if%> </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			   <div class="box-content"> 
                     
                       <% CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
  Verifica=Request.QueryString("Verifica")
Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
   DataTest = gira_data()
    SessioneQuiz=Request.Form("txtSessione")
	 Quiz=clng(Request.QueryString("Quiz"))
   Stato=Request.QueryString("Stato")
    Stato0=Request.QueryString("Stato0")
	Tutti=Request.QueryString("Tutti")
	Cartella=Request.QueryString("Cartella")
   Modulo=Request.QueryString("Modulo")
    Capitolo=Request.QueryString("Capitolo")
  
   CodiceAllievo=Request.QueryString("CodiceAllievo") 
    CodiceTest=Request.QueryString("CodiceTest")
   'parametro generato random da esegui test per scegliere il quiz da eseguire di cui ora calcolo il risultato
  
   orderby=clng(Request.QueryString("orderby"))
   
    Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  ordina= Request.QueryString("ordina") 
  
    Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if
  
  numtest=request.querystring("NUMTEST")
  
  
  
  
   
  
  
  
  
   
   'Definizione query SQL per contare il numero di domande del test.
   
   
   
'   QuerySQL="SELECT count(*)" &_
'             "FROM Domande INNER JOIN " &_
'             "(Test INNER JOIN ComposizioneTest ON " &_
'             "Test.CodiceTest = ComposizioneTest.CodiceTest) " &_
'             "ON Domande.CodiceDomanda = ComposizioneTest.CodiceDomanda " &_
'             "WHERE Test.CodiceTest='" & CodiceTest & "';"
	
	
	
	
	
		
	
	
	
	
	
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
		  if CodiceSottopar<>"" then
	    
		    QuerySQL="SELECT count(*) " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "'  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) and Lingua='"&Lingua&"' ;"
		  
		   else
     QuerySQL="SELECT count(*) " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) and Lingua='"&Lingua&"' ;"
        
			end if

		   ' utilizzo il numero casuale per accedere al vettore che contiene le possibilità di ordinamento, potrò farlo anche per asc e desc	   
		else 
		'Definzione codice SQl della query per sapere quanti quiz ci sono
		  QuerySQL="SELECT count(*) " &_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and  Domande.Segnalata=0 and Domande.Multiple=0  and Domande.VF=0 and Lingua='"&Lingua&"' ;"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			'call setInQuizOrderBy()
			
			 %>
			<p align="center"><b><%Response.write ("N."&Quiz) %></b> <!-- stampa il titolo del test -->
		
		 <%
		  
			 
		  QuerySQL="SELECT count(*) " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and  Domande.Segnalata=0 and Domande.VF=0 "& stringaQuery&" and Lingua='"&Lingua&"' ;"
		
		
		end if   
'response.write(QuerySQL)
   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    NumDom=rsTabella(0).value 'Assegno a NumDom numero delle domande
	'response.write("NumDom="&NumDom)
	
 'Dim objFSO,objCreatedFile
'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
'	Dim sRead, sReadLine, sReadAll, objTextFile
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	 url="C:\Inetpub\umanetroot\anno_2012-2013\logCalcolaRisultati.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(i & " " & NumDom & "<br>" & QuerySQL)
'				objCreatedFile.Close

	if (Stato=0) then   
		  if CodiceSottopar<>"" then
	    
		    QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0  and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) and Lingua='"&Lingua&"'  order by Domande." & ordina& " asc;"
		  
		   else
    QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and  Domande.Multiple=0  and Domande.VF=0 AND (Domande.In_Quiz=" &Quiz & "  or Domande.In_Quiz=-1) and Lingua='"&Lingua&"' order by Domande." & ordina& " asc;"
        
			end if

		   ' utilizzo il numero casuale per accedere al vettore che contiene le possibilità di ordinamento, potrò farlo anche per asc e desc	   
		else 
		
		
	
		
		'Definzione codice SQl della query per sapere quanti quiz ci sono
		 QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and  Domande.Segnalata=0 and Domande.Multiple=0  and Domande.VF=0 and Lingua='"&Lingua&"' ;"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			'call setInQuizOrderBy()
			
			 
		QuerySQL="SELECT Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE Domande.Id_Mod='" & Modulo & "' and  Domande.Segnalata=0 and Domande.Multiple=0  and Domande.VF=0 and Lingua='"&Lingua&"' " &stringaQuery & " order by Domande." & ordina& " asc;"
		
		
		end if   
		

'
'if (Stato=0) then 
'
'   if CodiceSottopar<>"" then
'	     QuerySQL="SELECT * "&_
'		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
'		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0 and Domande.VF=0;"
'	   else
'       QuerySQL="SELECT * "&_
'		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
'		   " WHERE Domande.Id_Arg='" & CodiceTest & "' and Domande.Multiple=0  and Domande.VF=0;"
'	end if
' else 
'			' il In_QUiz=-2 quando inserisco un test ma non vogli che sia visibile, probabilmente non serve usando TestAbilitati=0    
'          QuerySQL="SELECT * "&_
'		  " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
'		   " WHERE Domande.Id_Mod='" & Modulo & "' and Domande.Multiple=0  and Domande.VF=0;"
'    
'  end if	
	
 

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'	 url="C:\Inetpub\umanetroot\anno_2012-2013\logCalcolaRisultati1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(i & " " & NumDom & "<br>" & QuerySQL)
'				objCreatedFile.Close

 


   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 '   response.write(QuerySQL)
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
  'if rsTabella.EOF then response.write("caz...")
  Do While not(rsTabella.EOF) ' per ogni risposta confronta la risposta esatta con quella data dall'utente
   ' response.write("Ci sono")
	 RispostaEsatta=rsTabella.Fields("RispostaEsatta") 'legge dal risultato e memorizza in una variabile d'appoggio la risposta esatta
     'legge il valore associato all'oggetto avente per nome il numero contenuto nella variabile i, in base al valore ricava la risposta data  
     SELECT CASE Request.Form("" & i & "")
     CASE "1"
       RispostaData=1
     CASE "2"
       RispostaData=2
     CASE "3"
       RispostaData=3
     CASE "4"
       RispostaData=4
     CASE ELSE
     	RispostaData=0
     END SELECT  
	
	 ' per leggere il checkbox
	  SELECT CASE Request.Form("Check" & i & "")
     CASE "1"
       Segnalata=1
     CASE ELSE
     	Segnalata=0
     END SELECT  
	 
     dim a
     a=1
     
	 
	
	 
	 
	 
     RispDate(i) = RispostaData     ' memorizza nel vettore risposte date il valore della risposta data (i)
    ' Response.write(rsTabella.Fields.Count)
    ' Response.write(rsTabella.Fields(a).value)
    ' Response.write(rsTabella.Fields(a+1).value)
     'Response.write(rsTabella.Fields(1).value)
     'Response.write(rsTabella.Fields(2).value)
     'Response.write(rsTabella.Fields(3).value)
     
	 IF (RispostaData=0) THEN
		RispDate1(i)= "IN BIANCO" 
		inbianco=inbianco+1
     ELSE
       RispDate1(i) =  rsTabella.Fields(1+RispostaData).value
       'Response.Write(rsTabella.Fields(1+RispostaData).value)
     END IF
      IF (Segnalata=1) THEN	  
			
	  
	  %> 
	  	<!-- #include file = "inserisci_segnalazione_include.asp" -->
	            
	<% 
			RispDate1(i)= "SEGNALATA"
			 
	END IF
	 
	 
     RispEsatte(i) = RispostaEsatta ' memorizza nel vettore risposte esatte il valore della risposta esatta (i)
    
	
			'	dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'			   url="C:\Inetpub\umanetroot\anno_2012-2013_2\log.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(1+RispostaEsatta)
'				objCreatedFile.Close
	
	
	
	
    RispEsatte1(i) = rsTabella.Fields(1+RispostaEsatta).value ' *****si blocca qua non gli piace l'assegnazione
    'RispEsatte1(i)=1
	'response.write("1+RispostaEsatta"& 1+RispostaEsatta)
	'response.write("<br>R" & RispostaEsatta & "   "& RispostaData )
	 IF (RispostaEsatta=RispostaData) THEN  ' se sono uguali incrementa il numero delle risposte ok e pone a 0 l'elemento i del vettore errori 
           RisposteOK = RisposteOK +1
           Errori(i)=0 				'0 = domanda i esatta
     ELSE       					'1 = domanda i errata  
           Errori(i)=1				'se sono diversi incrementa il numero delle risposte ko e pone a 1 l'elemento i del vettore errori 
           RisposteKO = RisposteKO +1   
     END IF
     i = i + 1						' incrementa i
     rsTabella.MoveNext 			' passa alla prossima domanda
   Loop 
   
   'Calcolo della percentuale di domande corrette. 
   ' response.write(RisposteOK & "  " & (i-1))
	Risultato = (RisposteOK/(i-1))*100
    Risultato_relativo = (RisposteOK/(i-inbianco-1))*100
   
   '  response.write("Risultato="&Risultato)
    DataTest=date()
   'Esecuzione della query per inserire il risultato del test nella tabella Risulati
  ' if (Verifica<>1) then ' inserisco i risultati solo se non sono in modalità verifica
   
	   if (Stato=0) then 
		   QuerySQL="  INSERT INTO Risultati (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Lingua) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0)   & "'," &Quiz & "," &SessioneQuiz  & ",1,'"&Lingua&"';"
	   else 
			QuerySQL="  INSERT INTO Risultati1 (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Lingua) SELECT '" & CodiceAllievo & "','" & Modulo & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0)    & "'," &Quiz & "," &SessioneQuiz  & ",1,'"&Lingua&"';"
	   end if 
	  ' response.write(QuerySQL)
	   ConnessioneDB.Execute QuerySQL 
  ' end if
   'Stampa del risultato all'utente
   
   if (Round(Risultato,0)*8/100)<6 then
   Response.Write("<H4><span class='alert-error'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*8/100 &"</span></H4>")
   else
   Response.Write("<H4><span class='alert-success'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*8/100 &"</span></H4>")
   end if
   
   %>
	 
   <%
   Response.Write("<H5>Su un totale di " & NumDom & " domande ci sono <span class='alert-success'>" & RisposteOK & " risposte corrette</span> e <span class='alert-error'>" & RisposteKO & " risposte errate</span></h5> <BR>")
   %>
	 
   <%  
   
    if (Round(Risultato,0)*8/100)<6 then
   Response.Write("<H4><span class='alert-error'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato_relativo,0) & "% - Voto = " &  Round(Risultato_relativo,0)*8/100 &"</span></H4>")
   else
   Response.Write("<H4><span class='alert-success'>Risultato assoluto del test N."&Quiz&": " & Round(Risultato_relativo,0) & "% - Voto = " &  Round(Risultato_relativo,0)*8/100 &"</span></H4>")
   end if
   
    %>
	 
   <%
   Response.Write("<H5>Su un totale di " & NumDom-inbianco & " domande risposte ci sono <span class='alert-success'> " & RisposteOK & " risposte corrette</span> e <span class='alert-error'> " & NumDom-inbianco-RisposteOK & " risposte errate </span></h5><BR>")

  %>
   <!-- stampa la tabella per offire l'opportunità di visualizzare le correzioni -->
   	</font>
                  
                  
         <table class="table table-hover table-nomargin table-colored-header"> 
			<tr>
				<th><b>Domanda</b> </th>
		 
				<th><b>Quesito</b> </th>
				<th><b>Risposta Data</b> </th>
				<th><b>Risposta Esatta</b> </th>
				<th><b>Correzione</b> </th>
			</tr>
			</font>
			<%	rsTabella.Movefirst ' torna all'inizio delle domande
		   	    i=1
				Do While Not rsTabella.EOF %>
				<tr>
				<% if Errori(i)=1 then %>  <!-- se la risposta è errata usa il colore rosso -->
				
				  <td valign=top><b><%=i%></b></td>
			 
				  <td valign=top><b><%=rsTabella.Fields("Quesito")%></b></td>
				  <td valign=top> <font color="red"><%=RispDate1(i)%></font> </td>
				  <td valign=top> <font color="red"><%=RispEsatte1(i)%></font> </td>				
				   

			    <% else %>      <!-- se la risposta è correta usa il colore verde -->
	
				  <td valign=top><b><%=i%></b> </td>
			 
				  <td valign=top><b><%=rsTabella.Fields("Quesito")%></b></td>
	   			  <td valign=top><font color="green"><%=RispDate1(i)%> </font></td>
				  <td valign=top><font color="green"><%=RispEsatte1(i)%></font> </td>				
				  

				<%End if%>
                <%
				url=rsTabella.Fields("URL_Teoria")
				url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
				url=Replace(url,"\","/")
				 
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
				<% rsTabella.movenext
				i=i+1
				Loop %>	
			</table>     
              
                      </div>         
			        </div>
			      </div>
			    </div>
	
                      
                      
                        <a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&Tutti=<%=Tutti%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>" target="_blank"><input type="button" class="btn" value="Continua verifica e lascia aperta questa pagina"></i>  </a>      <br><br>
       <a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&Tutti=<%=Tutti%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>" target="_Top"><input type="button" class="btn-primary" value="Continua verifica e chiudi questa pagina"></i>  </a>         
                      
                      
                      
                      
                      
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

