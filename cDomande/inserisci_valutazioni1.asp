<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci valutazioni</title>   
   
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
	 <!-- Notify -->
	<link rel="stylesheet" href="../../css/plugins/gritter/jquery.gritter.css">
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
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.draggable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
    <!-- Notify -->
	<script src="../../js/plugins/gritter/jquery.gritter.min.js"></script>


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
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
    </script>
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
    
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  <script type="text/javascript">
	
		 
$(window).ready(function () {	   
	
	   $('#msg').click();
	   
	  // event.stopPropagation();
	    
	});
	
</script>


   
</head>

<%
  Response.Buffer = true
 ' On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>'>
  <% end if %>


	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
 <%
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
  Dim Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Dim objFSO,objCreatedFile
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim sRead, sReadLine, sReadAll, objTextFile
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   Cartella=Request.QueryString("Cartella")
   BoxApro=Request.QueryString("BoxApro")
   esporta=Request.QueryString("esporta")
    tutto=Request.QueryString("tutto")
	 Modulo=Request.QueryString("Modulo")


   
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-question-sign"></i>Valutazioni</h1> 
                    
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
							<a href="../cClasse/home_app.asp?id_classe=<%=session("id_classe")%>">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Compiti</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							 <a href="#">Valutazioni</a> 
                             
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
				        <h3> <i class="icon-reorder"></i></h3>
			          </div>
				      <div class="box-content">
                      
 
 	<% 
	    if esporta<>"" and session("admin")=True then
Set objFSO = CreateObject("Scripting.FileSystemObject")
		if tutto<>"" then
			'************SONO QUI 
			url=Server.MapPath(homesito & "/script/cDomande/esportate_xls")&"/"&classe&"_"&Modulo&".xls"
		else

			url=Server.MapPath(homesito & "/script/cDomande/esportate_xls")&"/"&classe&"_"&CodiceTest&".xls"
 
		end if
			url=Replace(url,"\","/")
			if objFSO.FileExists(url) then 
		    	objFSO.DeleteFile url
			end if
			response.write("Creo il file:"&url&"<br>")
			Set objCreatedFile = objFSO.CreateTextFile(url, True)
			
		    riga="<table><thead><tr><th><b>Domanda</b></th><th><b>R1</b></th><th><center><b>R2</b></center></th><th><center><b>R3</b></center></th><th><center><b>R4</b></center></th><th><center><b>SEC</b></center></th><th><center><b>RE</b></center></th><th><center><b>Spiegazione</b></center></th></tr></thead><tbody>"
		    objCreatedFile.WriteLine(riga)

			  response.write(riga&"<br>")
   end if
	 
	 
	 NumRec=clng(Request.Form("TxtNUMREC"))
   
  ' response.write(numrec)
  for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
   Domanda = Request.Form("txtDomanda"&k)
   ID=Request.Form("txtCodiceDomanda"&k)
   R11 = Request.Form("txtR1"&k)
   R1=Replace(R11,"'","''")
   R22 = Request.Form("txtR2"&k)
   R2=Replace(R22,"'","''")
   R33 = Request.Form("txtR3"&k)
   R3=Replace(R33,"'","''")
   R44 = Request.Form("txtR4"&k)
   R4 = Replace(R44,"'","''")
  
  
   'Spiegazione=Request.Form("S1")
   'TestoDomandaPlus=Request.Form("TestoDomandaPlus")
     
	 
'	 Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
' 	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logInQuiz.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(Request.Form("txtINQUIZ"&k))
'				objCreatedFile.Close
'	 
'	 
	 
	  
   RE = clng(Request.Form("txtRE"&k))
   VAL=clng(Request.Form("txtVAL"&k))
   INQUIZ=clng(Request.Form("txtINQUIZ"&k))
   DATA=cdate(Request.Form("txtDataDomanda"&k))
    Segnalata=Request.Form("txtSegnalata"&k)
	Esportata=Request.Form("txtEsportata"&k)
   if Segnalata="" then
     Segnalata=0
   end if
   if Esportata="" then
     Esportata=0
   end if
   ' per la spiegazione della domanda 
   ' url=Server.MapPath(homesito)&"/"& Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	'url3=Replace(url,"\","/")
	'url=url3

  ' per il testo della domanda plus
    ' url4=Server.MapPath(homesito)& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
    ' url4=Replace(url4,"\","/")
	 
if session("Admin")=true and esporta="" then  
      QuerySQL ="UPDATE Domande SET Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE &  "', Voto = '" & VAL & "', In_Quiz = " & INQUIZ &", Data= '" & DATA & "', Segnalata= '" & Segnalata & "' WHERE CodiceDomanda =" &ID&";"
	
	 ConnessioneDB.Execute(QuerySQL)
	 response.write("<br>"&QuerySQL)
end if
 url=Request.Form("url"&k)    
Spiegazione=Request.Form("txtSpiegazione"&k)
if clng(Segnalata)=1 then
 
' Aggiorno la spiegazione
	' se è segnalata aggiorno file di spiegazione
	objFSO.DeleteFile url
	'  response.Write("<br>Cancello : " &url)
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	'' Write a line with a newline character.
	objCreatedFile.WriteLine(Spiegazione)
 '   response.Write("<br>Creo : " &Spiegazione)
	''Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
end if 

  if esporta<>"" and session("admin")=True and esportata<>0 then
  secondi=20
   riga=" <tr><td>"&Domanda&"</td><td>"&R1&"</td><td>"&R2&"</td><td>"&R3&"</td><td>"&R4&"</td><td>"&secondi&"</td><td>"&RE&"</td><td>"&Spiegazione&"</td></tr>"
 objCreatedFile.WriteLine(riga)
response.write(riga&"<br>")

 end if



'response.write(QuerySQL) %> <br> <%
'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

'Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
 '	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logStud.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
'Create the FSO.
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(Cartella)
'response.write(url)
'objFSO.DeleteFile url
'Set objCreatedFile = objFSO.CreateTextFile(url, True)
' Write a line with a newline character.
'objCreatedFile.WriteLine(Spiegazione)
'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
'objCreatedFile.Close
' per aggiornare la domanda plus
'if Tipodomanda=1 then
'	objFSO.DeleteFile url4
'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
'	' Write a line with a newline character.
'	objCreatedFile.WriteLine(TestoDomandaPlus)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
'	objCreatedFile.Close
'end if 
next 

 riga=" </tbody></table>"
 response.write(riga&"<br>")
 objCreatedFile.WriteLine(riga)
objCreatedFile.Close
On Error Resume Next
If Err.Number = 0 Then
%>
<span class="alert-success">
<%
Response.Write "Modifica avvenuta! " 
'response.Redirect "../cClasse/home_app.asp?id_classe="&Session("Id_Classe")

 
Else%>
<span class="alert-error">
<%

Response.Write Err.Description 
Err.Number = 0
End If%>
</span>

  
 
 



 



								 
	 
	 
				 
				 
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      <p>
										 
									</p>
									<p> <span class="invisible"><a id="msg" href="#modal-1" role="button" class="btn notify" data-notify-title="Modifica effettuata!" data-notify-message="Torna al Libro... ">Stai per essere reindirizzato</a></span>
									<!--	<a href="#modal-1" role="button" class="btn notify" data-notify-time="1000" data-notify-title="Success!" data-notify-message="The user has been successfully edited.">Timed fade notification (1second)</a>
										 <a href="#modal-1" role="button" class="btn notify" data-notify-title="WARNING!" data-notify-message="Please refresh the cache!" data-notify-sticky="true">Sticky notification</a> 
                      	<a id="msg2" href="#modal-4" role="button" class="btn" data-toggle="modal">Alert</a>
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


<% 
'Response.AddHeader "REFRESH","1;URL=../cClasse/home_app.asp?id_classe="&session("Id_Classe")&"&dividApro="&BoxApro

 
%>
 </html>

