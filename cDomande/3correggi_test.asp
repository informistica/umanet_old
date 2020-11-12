<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Bilancia Quiz</title>   
   
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
       
       
     
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
 
</head>
<%Function domandaplus()
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 Cartella=rsTabella(13)
	 Modulo=rsTabella(10)
	 Paragrafo=rsTabella.fields("Titolo")
	 Id=rsTabella(0)
 
	 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	response.write(sReadAll)
	objTextFile.Close
End Function %>
<% Response.Buffer=True 
on error resume next
%>


<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">  

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
						<h1> <i class="icon-comments"></i> Bilancia Quiz </h1> 
                    
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
							<a href="#">....</a>
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
				        <h3> <i class="icon-reorder"></i>  .... </h3>
			          </div>
				      <div class="box-content">
   <% 
 
    StringaConnessione= Request.Cookies("Dati")("StrConn")   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
   Stato=Request.QueryString("Stato") '=0 se svolto test del paragrafo 1 se svolgo quello del modulo
   Modulo=Request.QueryString("Modulo") 
   'Raccolta dei dati digitati dall'utente e salvati nel cookie
   TitoloTest=Request.Cookies("Dati")("TitoloTest")
'   CodiceTest = Request.Cookies("Dati")("CodiceTest")
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   CodiceTest = Request.QueryString("CodiceTest") ' se svolgo tutto il modulo (stato=1) contiene l'Id del modulo e non del paragrafo
    NUMTEST=Request.form("txtNUMTEST")
	if NUMTEST="" then
	   NUMTEST=Request.querystring("NUMTEST")
	end if
  'leggo quanti test sono presenti ed offro la possibilità di scegliere quale modificare
Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  vf=Request.QueryString("vf")  ' correggo tipo vero falso
  rm=Request.QueryString("rm")  ' correggo tipo vero falso

if NUMTEST="" then%>
<form method="POST" class="form-horizontal" action="3correggi_test.asp?Stato=<%=Stato%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>&vf=<%=vf%>&rm=<%=rm%>">

<%
	QuerySQL="SELECT MAX([In_Quiz]) AS [NUMQUIZ] FROM DOMANDE WHERE Id_Arg='" & CodiceTest & "'"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	NUMQUIZ=rsTabella(0) 
	response.Write("<h3>Quale test vuoi modificare ?</h3>")
	%>
	
	<select name="txtNUMTEST">
	 <% for i=1 to NUMQUIZ %>
	   <option value="<%=i%>"><%=i%> Quiz </option>
	<%next%>
	<br> <input type="submit" value="Invia" name="B1"> </p> 
</form>                   
 
 
 <%else
  

%>
 
 
<p align="center"><b><font  color=#FF0000>ESCLUDI(X) ed INCLUDI(V) DOMANDE DAL TEST N.(<%=NUMTEST%>)</font></b></p>
<p align="center"><b><%Response.write (TitoloTest) %></b></font></p> <!-- stampa il titolo del test -->

<%  
if strcomp(vf,"1")=0 then
	 stringQuery=" VF=1 and "
     else  
	   if strcomp(rm,"1")=0  then 
			stringQuery=" Multiple=1 and "
		else
		   stringQuery="Multiple=0 and VF=0 and "
		 end if
	end if
	
	'stringQuery="Domande.In_Quiz="&NUMTEST&" and "&stringQuery
if (Stato=0) then 
     if CodiceSottopar<>"" then	
			 QuerySQL="SELECT * " &_
             "FROM DomandeQuiz WHERE " &_
            stringQuery& " Id_Arg='" & CodiceTest & "'  and Id_Sottoparagrafo='" & CodiceSottopar & "' order by Domande.CodiceDomanda asc;"
	else
			 QuerySQL="SELECT * " &_
             "FROM DomandeQuiz Where " &_
             stringQuery&" Id_Arg='" & CodiceTest & "' order by CodiceDomanda asc;"
    end if
   
   
   'Definizione query SQL per la lettura delle risposte esatta nel test scelto.
   
else
   QuerySQL="SELECT *" &_
             "FROM DomandeQuiz WHERE " &_
             stringQuery&" Id_Mod='" & Modulo & "' order by CodiceDomanda asc;"
end if  
    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   response.write(QuerySQL)

%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then %>
  <H4>Test non ancora disponibile!<h4>
  <p><h5><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
</H5>
<% Else %>
  <FORM class="form-horizontal" METHOD="POST" ACTION="3correggi_test1.asp?Stato=<%=Stato%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&NUMTEST=<%=NUMTEST%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>&vf=<%=vf%>&rm=<%=rm%>">
 <div class="control-group">
		 
  <%i=1 
   numDomInQuiz=0 'inizializza la variabile i (contatore delle domande)
  Do until rsTabella.EOF  ' esegue un ciclo e ad ogni iterazione crea un quiz (con 4 valori possibili) avente per nome il numero contenuto nella variabile i 
   if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
    <B> <%=i & ") "%><%=rsTabella.Fields("Quesito")%></B> 
	 <textarea rows="6" name="S1" value="ciao" cols="116"><%=Response.write(domandaplus())%> </textarea><br>
  <%else
     ' se la domanda è In_Quiz=NUMTEST allora la scrivo in verde , per ogni domanda offro una V per abilitarle e una X per disabilitarla
       if (rsTabella.Fields("In_Quiz")=clng(NUMTEST) ) then
	   numDomInQuiz=numDomInQuiz+1
	   %>
     <font color="#00CC66" ><b>V</b><INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="0"><B><%=rsTabella.Fields("In_Quiz")&")"%>  <%=rsTabella.Fields("Quesito")%></B></font><font color="#00CC66"><b> X</b></font> <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="5"> <br>
       <%else%>
	      <font color="#00CC66"><b>V</b></font><font color="#FF0000"><INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="0"><B> <%=rsTabella.Fields("In_Quiz")&")"%> <%=rsTabella.Fields("Quesito")%></B></font><font color="#FF0000"><b> X</b></font> <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="5"> <br>
       
	   <% end if%>
   <%end if %>
   <% if strcomp(vf,"1")=0 then%>
   <%else%>
      <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="1">
      <%=rsTabella.Fields("Risposta1")%><BR>
      <INPUT TYPE="RADIO" NAME="<%=i%>" VALUE="2">
      <%=rsTabella.Fields("Risposta2")%><BR> 
      <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="3">
      <%=rsTabella.Fields("Risposta3")%><BR> 
      <INPUT TYPE="RADIO" NAME="<%=i%>"  VALUE="4">
      <%=rsTabella.Fields("Risposta4")%> <BR>
      <%end if%>
    </FIELDSET>
   <% i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande%>
   <% Loop %>
   <P>
      <INPUT TYPE="SUBMIT" NAME="submit" VALUE="Invia le <%=numDomInQuiz%> risposte del test sul totale di <%=i%>"> <!-- crea il bottone per inviare le riposte alla pagina che calcola il risultato -->
   </P>
   
 
   </div>
   </FORM>
<% End If %>
<% rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
   ConnessioneDB.Close : Set ConnessioneDB = Nothing 
end if ' chiudi l'if iniziale quello della scelta preliminare del test%>
  
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     
                      
                      
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
         

			 
	</body>

 </html>

