<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Risultato TEST V/F</title>   
   
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
		
    Response.Buffer=True 
	on error resume next
   Dim  Quiz
   
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
						<h1> <i class="icon-comments"></i> Risultato quiz Vero/Falso</h1> 
                    
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
	  
  ' CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
  Verifica=Request.QueryString("Verifica")
Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
   DataTest = gira_data()
   
    SessioneQuiz=Request.Form("txtSessione")
	 Quiz=clng(Request.QueryString("Quiz"))
   Stato=Request.QueryString("Stato")
    Tutti=Request.QueryString("Tutti")
    Stato0=Request.QueryString("Stato0")
	Cartella=Request.QueryString("Cartella")
   Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   Modulo=Request.QueryString("Modulo")
   CodiceTest=Request.QueryString("CodiceTest") ' se svolgo tutto il modulo contiene l'id del modulo
   CodiceAllievo=Request.QueryString("CodiceAllievo") 
   'parametro generato random da esegui test per scegliere il quiz da eseguire di cui ora calcolo il risultato
  
   orderby=clng(Request.QueryString("orderby"))
    Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
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

function trovaRisposta(RispostaData)
if RispostaData=0 then
    trovaRisposta="Falso"
else
   if RispostaData=1 then
     trovaRisposta="Vero"
  else
     if RispostaData=3 then
     trovaRisposta="Bianca"
	 end if
  end if
end if
end function


 'response.write("aaaaNUMTESt="&NUMTEST&"cmp="&strcmp(NUMTEST,"-1"))
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
             " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0   and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0 and Domande.VF=1 AND (Domande.In_Quiz="&Quiz&"  or Domande.In_Quiz=-1) and Lingua='"&Lingua&"' ;"
	   else
 
       QuerySQL="SELECT count(*) " &_
              " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
             " WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0  and Domande.Multiple=0 and Domande.VF=1 AND (Domande.In_Quiz="&Quiz&"  or Domande.In_Quiz=-1)  and Lingua='"&Lingua&"' ;"
	end if


  
			
    'Assegna alla variabile il risultato della query prodotta utilizzando il metodo Execute(stringa della query) dell'oggetto connessione
else 
'Definzione codice SQl della query per ricercare le domande del modulo
QuerySQL="SELECT count(*) " &_
               " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
			 "WHERE Domande.Id_Mod='" & Modulo & "' and  Domande.Segnalata=0 and Domande.Multiple=0 and Domande.VF=1 " &stringaQuery & "  and Lingua='"&Lingua&"' ;"
			 
		 
		
			 
end if   

 ' response.write(QuerySQL)
   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
    NumDom=rsTabella(0).value 'Assegno a NumDom numero delle domande
	
	
 'Dim objFSO,objCreatedFile
'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
'	Dim sRead, sReadLine, sReadAll, objTextFile
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	 url="C:\Inetpub\umanetroot\anno_2012-2013\logCalcolaRisultati.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(i & " " & NumDom & "<br>" & QuerySQL)
'				objCreatedFile.Close

	
	Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
 
	
if (Stato=0) then 
     if CodiceSottopar<>"" then
	  
		     QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Domande.Multiple " &_
               " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
           "  WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and Domande.Id_Sottoparagrafo='" & CodiceSottopar & "' and Domande.Multiple=0 and Domande.VF=1  and Lingua='"&Lingua&"'  AND (Domande.In_Quiz=" &Quiz & " or Domande.In_Quiz=-1) order by Domande." & order(orderby)& " asc;"
		  
		   else
     QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Domande.Multiple,Cognome,Nome " &_
               " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
           "  WHERE Domande.Id_Arg='" & CodiceTest & "' and  Domande.Segnalata=0 and Domande.Multiple=0 and Domande.VF=1  and Lingua='"&Lingua&"' AND (Domande.In_Quiz=" &Quiz & " or Domande.In_Quiz=-1) order by Domande." & order(orderby)& " asc;"
        
		
			end if
		   
 
else
   QuerySQL="SELECT CodiceDomanda,Quesito,Risposta1,Risposta2,Risposta3,Risposta4,RispostaEsatta,URL_Teoria,Id_Stud,Domande.Multiple,Cognome,Nome  " &_
             " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
             " WHERE Domande.Id_Mod='" & Modulo & "' and  Domande.Segnalata=0 and Domande.Multiple=0 and Domande.VF=1  and Lingua='"&Lingua&"' " &stringaQuery & " order by Domande." & order(orderby)& " asc;"

end if  

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'	 url="C:\Inetpub\umanetroot\anno_2012-2013\logCalcolaRisultati1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(i & " " & NumDom & "<br>" & QuerySQL)
'				objCreatedFile.Close

 

'response.write(QuerySQL)
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
  
  Do While not(rsTabella.EOF) ' per ogni risposta confronta la risposta esatta con quella data dall'utente
     RispostaEsatta=rsTabella.Fields("RispostaEsatta") 'legge dal risultato e memorizza in una variabile d'appoggio la risposta esatta
     'legge il valore associato all'oggetto avente per nome il numero contenuto nella variabile i, in base al valore ricava la risposta data  
     SELECT CASE Request.Form("" & i & "")
     CASE "1"
       RispostaData=1
     CASE "0"
       RispostaData=0
    ' CASE "3"
'       RispostaData=3
'     CASE "4"
'       RispostaData=4
     CASE ELSE
     	RispostaData=3
     END SELECT  
	'response.write(RispostaData &"-")
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
     
	 IF (RispostaData=3) THEN
		RispDate1(i)=3 
		inbianco=inbianco+1
     ELSE
	     IF RispostaData=0 then
             RispDate1(i)=0
       'Response.Write(rsTabella.Fields(1+RispostaData).value)
	     ELSE
		     RispDate1(i)=1
		 END IF
     END IF
	 
	' response.write("<br>"&inbianco&"-"&RispDate1(i))
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
	
	
	
	
   ' RispEsatte1(i) = rsTabella.Fields(1+RispostaEsatta).value ' *****si blocca qua non gli piace l'assegnazione
     RispEsatte1(i) = RispostaEsatta
	'RispEsatte1(i)=1
	'response.write("1+RispostaEsatta"& 1+RispostaEsatta)
	'response.write("<br>" & QuerySQL)
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
    Risultato = (RisposteOK/(i-1))*100
	if (i-inbianco-1)<>0 then
    Risultato_relativo = (RisposteOK/(i-inbianco-1))*100
	else
	 Risultato_relativo=0
	end if
	
   
    
    DataTest=date()
   'Esecuzione della query per inserire il risultato del test nella tabella Risulati
  ' if (Verifica<>1) then ' inserisco i risultati solo se non sono in modalità verifica
   
	   if (Stato=0) then 
		   QuerySQL="  INSERT INTO Risultati (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Lingua) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0)  & "'," &Quiz & "," &SessioneQuiz  & ",0,'"&Lingua&"';"
	   else 
			QuerySQL="  INSERT INTO Risultati1 (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo, Lingua) SELECT '" & CodiceAllievo & "','" & Modulo & "', '" & DataTest & "', '" & FormatDateTime(now, 4) & "','" & Round(Risultato,0) & "'," &Quiz& "," &SessioneQuiz  & ",0,'"&Lingua&"';"
	   end if 
	   'response.write(QuerySQL)
	   ConnessioneDB.Execute QuerySQL 
  ' end if%>
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  Quiz su "<%=Capitolo%> : <%=Paragrafo%>"</h3>
			          </div>
				      <div class="box-content">
                      
 

<%  
if (Round(Risultato,0)*8/100)<6 then%>
<div class="alert-error">
<%else%>
<div class="alert-success">
<%end if%>
<%

 Response.Write("<h4>Risultato assoluto del test N."&Quiz&": " & Round(Risultato,0) & "% - Voto = " &  Round(Risultato,0)*8/100 &"</H4>")
   %>
	 
	 
   <%
   Response.Write("<H5>Su un totale di " & NumDom & " domande ci sono " & RisposteOK & " risposte corrette e " & RisposteKO & " risposte errate <BR></h5>")
   %>
</div>	   
 <hr>
 
   <%
   if (Round(Risultato_relativo,0)*8/100)<6 then%>
<div class="alert-error">
<%else%>
<div class="alert-success">
<%end if%>  
   <% Response.Write("<h4>Risultato relativo del test N."&Quiz&": " & Round(Risultato_relativo,0) & "% - Voto = " &  Round(Risultato_relativo,0)*8/100 &"</H3>")
    %>
 
   <%
   Response.Write("<H5>Su un totale di " & NumDom-inbianco & " domande risposte ci sono " & RisposteOK & " risposte corrette e " & NumDom-inbianco-RisposteOK & " risposte errate <BR></h5>")

  %>
  </div>
  <hr>
   <!-- stampa la tabella per offire l'opportunità di visualizzare le correzioni -->
  
  <div class="box-content nopadding">
								<table class="table table-hover table-nomargin table-striped">
									<thead>
										<tr>
											<th class='hidden-480'>Domanda</th>
											<th>Quesito</th>
											<th class='hidden-350'>Risposta data</th>
											<th>Risposta esatta</th>
											<th class='hidden-480'>Approfondisci</th>
										</tr>
									</thead>
			
                                    <tbody>
                                    
                                    
                                    <%	rsTabella.Movefirst ' torna all'inizio delle domande
		   	    i=1
				Do While Not rsTabella.EOF %>
				<tr>
				<% if Errori(i)=1 then %>  <!-- se la risposta è errata usa il colore rosso -->
				
				  <td  class='hidden-480' valign=top><b><%=i%></b></td>
				 
				  <td valign=top><b><%=rsTabella.Fields("Quesito")%></b></td>
				  <td class='hidden-350' valign=top> <font color="red"><%=trovaRisposta(RispDate1(i))%> </td>
				  <td valign=top> <font color="red"><%=trovaRisposta(RispEsatte1(i))%> </td>				
				 

			    <% else %>      <!-- se la risposta è correta usa il colore verde -->
	
				  <td class='hidden-480' valign=top><b><%=i%></b> </td>
				 
				  <td valign=top><b><%=rsTabella.Fields("Quesito")%></b></td>
	   			  <td valign=top class='hidden-350'><font color="green"><%=trovaRisposta(RispDate1(i))%> </td>
				  <td valign=top><font color="green"><%=trovaRisposta(RispEsatte1(i))%> </td>				
				

				<%End if%>
                <%
				url=rsTabella.Fields("URL_Teoria")
    'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
url=Replace(url,"\","/")
'sReadAll=url
'response.write(sReadAll)
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll

objTextFile.Close
				%>
							
                             <td valign=top  class='hidden-350'>
                       
                       
                       
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
				</tbody>
		</table>
</div>
  
                                    
                                    
                                    
                                    
   
    <!--   <a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&Tutti=<%=Tutti%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>" target="_top"><input type="button" class="btn-primary" value="Continua verifica"></i>  </a>    -->                  
	 
  
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

