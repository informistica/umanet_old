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
NUMTEST=request.querystring("NUMTEST")
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
   <meta charset="utf-8">
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
	 
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
    
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Stati Generali della Legalità - 24/10/2017 </h1> 
						<h3>Foglio di Correzione</h3>
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
     QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"'"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'Privato=rsTabella.fields("Privato") 
	TestAbilitato=rsTabella.fields("TestAbilitato")
	rsTabella.close

 Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  Lingua = Request.QueryString("Lingua")
  if Lingua="" then 
    Lingua="it"
  end if	

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
 


NumeroDomande = 20 ' usare numero PARI!!!
	Tempo = 30
	
	QuerySQL = "SELECT TOP("&NumeroDomande&") * FROM Leg_Domande WHERE VF = 0"
	'QuerySQL = QuerySQL & "order by "&order(orderby)&";"  
	'response.write(QuerySQL)
  
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
		<p align="center"><b>RISPOSTE ESATTE DEL TEST</b></p>
		<%end if%>
		<p align="center"><b><%Response.write (TitoloTest) %></b></p> <!-- stampa il titolo del test -->
		
		<%  
						
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			call setInQuizOrderBy()
			'orderby=1
			'Quiz=1
			
			 if NUMTEST<>"" then
		   Quiz=NUMTEST
		   end if	
		   ' Response.write("QUIZ="&Quiz)
		 %>
		 
		
		 <%
		 QuerySQL="SELECT top(20) Domande.*, Paragrafi.Titolo,Allievi.Cognome,Allievi.Nome " &_
		   " FROM Allievi INNER JOIN (Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Allievi.CodiceAllievo = Domande.Id_Stud " &_
		   " WHERE(Id_Arg = 'Expo_9_3' OR Id_Arg = 'Expo_9_5' OR Id_Arg = 'Expo_9_6') AND (Segnalata = 0) and VF=0 AND (Multiple = 0)"
		
		
		
		
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
                      
                       <FORM name="formQuiz"  class='form-vertical' METHOD="POST" ACTION="calcola_risultato.asp?Lingua=<%=Lingua%>&ordina=<%=ordina%>&Verifica=<%=Verifica%>&Stato=<%=Stato%>&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&CodiceAllievo=<%=CodiceAllievo%>&Quiz=<%=Quiz%>&orderby=<%=orderby%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>">
		  <%i=1 'inizializza la variabile i (contatore delle domande)
		 
		  dim objFSO,objCreatedFile
		  Const ForReading = 1, ForWriting = 2, ForAppending = 8
		  Set objFSO = CreateObject("Scripting.FileSystemObject")
		
		    dim domande(5)
		 
		  Do until rsTabella.EOF   ' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
		  
		 '  Set objFSO = CreateObject("Scripting.FileSystemObject")  
'			url="C:\Inetpub\umanetroot\anno_2013-2014\logEsegui395.txt"
'						Set objCreatedFile = objFSO.CreateTextFile(rsTabella("CodiceDomanda"), True)
'						objCreatedFile.WriteLine(QuerySQL & "orderby="&orderby)
'						objCreatedFile.Close 
		  
		  url=rsTabella.Fields("URL_Teoria")
    'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Server.MapPath(homesito)& "/Db"&Session("DB")& right(url, len(url)-2)
url=Replace(url,"\","/")
'sReadAll=url
'response.write(sReadAll)
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll

objTextFile.Close

		   if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
			<% ' se sono in modalità verifica aggiungo un bottone per la segnalazione della domanda
			   if Verifica=1 then %>  
			     <a  href="../cDomande/inserisci_valutazione.asp?Lingua=<%=Lingua%>&traduzione=1&Multiple=<%=rsTabella("Multiple")%>&ORA=<%=left(rsTabella("Ora"),5)%>&DATA=<%=rsTabella("Data")%>&Tipodomanda=<%=rsTabella("Tipo")%>&Cartella=<%=rsTabella("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabella("CodiceDomanda")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=rsTabella("Titolo")%>&Quesito=<%=rsTabella("Quesito")%>&R1=<%=rsTabella("Risposta1")%> &R2=<%=rsTabella("Risposta2")%>&R3=<%=rsTabella("Risposta3")%>&R4=<%=rsTabella("Risposta4")%>&RE=<%=rsTabella("RispostaEsatta")%>&MO=<%=rsTabella("ID_Mod")%>&VAL=<%=rsTabella("Voto")%>&VF=<%=rsTabella("VF")%>&URL=<%=rsTabella("URL_Teoria")%>&INQUIZ=<%=rsTabella("In_Quiz")%>&VALINQUIZ=<%=rsTabella("In_QuizStud")%>&Segnalata=<%=rsTabella("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">                                                          
			   <b>(<%=rsTabella("CodiceDomanda")%>)</b></a>
			   <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore"> 
                  <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="top" data-content="<%=sReadAll%>">
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
                                         
										<div class="controls">
                                              <% if Verifica=1 then %>
                                             <a  href="../cDomande/inserisci_valutazione.asp?Lingua=<%=Lingua%>&traduzione=1&Multiple=<%=rsTabella("Multiple")%>&ORA=<%=left(rsTabella("Ora"),5)%>&DATA=<%=rsTabella("Data")%>&Tipodomanda=<%=rsTabella("Tipo")%>&Cartella=<%=rsTabella("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabella("CodiceDomanda")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=rsTabella("Titolo")%>&Quesito=<%=rsTabella("Quesito")%>&R1=<%=rsTabella("Risposta1")%> &R2=<%=rsTabella("Risposta2")%>&R3=<%=rsTabella("Risposta3")%>&R4=<%=rsTabella("Risposta4")%>&RE=<%=rsTabella("RispostaEsatta")%>&MO=<%=rsTabella("ID_Mod")%>&VAL=<%=rsTabella("Voto")%>&VF=<%=rsTabella("VF")%>&URL=<%=rsTabella("URL_Teoria")%>&INQUIZ=<%=rsTabella("In_Quiz")%>&VALINQUIZ=<%=rsTabella("In_QuizStud")%>&Segnalata=<%=rsTabella("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">                                                          
			   <b>(<%=rsTabella("CodiceDomanda")%>)</b></a>
			   <INPUT TYPE="checkbox" NAME="Check<%=i%>" VALUE="1"  title="Notifica un errore all'autore">   <a data-original-title="Spiegazione (<%=rsTabella("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>)" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="top" data-content="<%=sReadAll%>">
						<center>  <i class="icon-question-sign"></i></center></a></span>
			   </b>
			 
			 <br><br>
			  <%end if %>
	
		   <%end if %>
		   
		 <% domande(1)=rsTabella.Fields("Risposta1")
		 domande(2)=rsTabella.Fields("Risposta2")
		 domande(3)=rsTabella.Fields("Risposta3")
		 domande(4)=rsTabella.Fields("Risposta4")
		 
		 %>
			  <%=clng(rsTabella.Fields("RispostaEsatta"))&") "%>
			  <%=domande(clng(rsTabella.Fields("RispostaEsatta")))%>  <BR>
			 	</div>
			</div>
            <hr>
			 
		   <% i = i+ 1 
			   rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande%>
		   <% Loop %>
		  <P>
			  
		   </FORM>
		<% End If %>
		<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
		   ConnessioneDB.Close : Set ConnessioneDB = Nothing %>
		  
                     
                      

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

