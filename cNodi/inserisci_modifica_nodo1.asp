<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica nodo</title>   
   
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
    
    


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
 
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
   
    <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
     
	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
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
 tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 tNod=request.querystring("tNod")
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Modifica nodo</h1> 
                    
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
							 <a href="#">Modifica nodo</a> 
                             
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
   <% Dim Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,voto
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   davalutazione=Request.QueryString("davalutazione")
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceNodo = Request.QueryString("CodiceNodo")
    CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   votobase=Request.Form("txtVAL")
   
    DATA = cdate(Request.Form("txtDATA"))   %>           
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%> : <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 <%  CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
  ' CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceAllievo = Request.querystring("cod")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
  Segnalata=Request.Form("txtSegnalata"&k)
   if Segnalata="" then
     Segnalata=0
   end if  
    
  ID=Request.QueryString("CodiceNodo")
  Chi = Request.Form("txtChi")
 ' Chi = Request.QueryString("Chi")
  Chi = Replace(Chi, Chr(34), "'")
  Chi=  Replace(Chi,"'",chr(96))
  
   Cosa = Request.Form("txtR1Cosa")
   'Cosa = Request.QueryString("Cosa")
   Cosa = Replace(Cosa, Chr(34), "'")
   Cosa=  Replace(Cosa,"'",chr(96))


  Dove = Request.Form("txtR1Dove")
  '  Dove = Request.QueryString("Dove")
   Dove = Replace(Dove, Chr(34), "'")
   Dove=  Replace(Dove,"'",chr(96))

   Quando = Request.Form("txtR1Quando")
   'Quando = Request.QueryString("Quando")
   Quando = Replace(Quando, Chr(34), "'")
   Quando=  Replace(Quando,"'",chr(96))
 
   Come = Request.Form("txtR1Come")
 'Come = Request.QueryString("Come")
   Come = Replace(Come, Chr(34), "'")
   Come=  Replace(Come,"'",chr(96))

   Perche = Request.Form("txtR1Perche")
   'Perche = Request.QueryString("Perche")
   Perche = Replace(Perche, Chr(34), "'")
   Perche=  Replace(Perche,"'",chr(96))
   
   Quindi = Request.Form("txtR1Quindi")
   'Quindi = Request.QueryString("Quindi")
   Quindi = Replace(Quindi, Chr(34), "'")
   Quindi=  Replace(Quindi,"'",chr(96))
   
   Sintesi=Request.Form("S1")
   Sintesi= Replace(Sintesi, Chr(34), chr(96))
   Sintesi= Replace(Sintesi, Chr(130), chr(138))  ' é in è
   Sintesi=  Replace(Sintesi,"'",chr(96))


  Spiegazione=Request.Form("S1")
 ' response.write(Sintesi)
 
   errore=0
   ' response.write(errore)
   if ((len(Chi)=0) or (len(Cosa)=0) or (len(Dove)=0) or (len(Quando)=0) or (len(Come)=0) or (len(Perche)=0) or(len(Quindi)=0) or(len(Sintesi)=0) ) then
      errore=2
	  
	  
   end if 
  ' response.write(len(Chi) & " - " & len(Cosa) & " - " &len(Dove) & " - " &len(Quando) & " - " &len(Come) & "-  " &len(Perche) & " - " &len(Quindi) & "-  " &len(Sintesi) & "  -")
   'response.write("Dove="&Dove)
  if (errore=0) then 
     

	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"  'per server on-line
	url=Replace(url,"\","/")
	 
	 
	  ' devo vedere se il setting è tale da richiedere voto=1 come default oppure no  
		QuerySQL1="Select * from Setting where Id_Classe='"& Session("Id_Classe")&"';"
		Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
		Valutato=rsTabella.fields("Valutato") 
		DVAbilitato=rsTabella.fields("DVAbilitato")
		rsTabella.close
		
		QuerySQL1="Select Id_Stud from Nodi where CodiceNodo="&CodiceNodo&";"
		Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
		CodiceAllievo=rsTabella.fields(0) 
		 'response.Write(QuerySQL1)
		rsTabella.close
	
	
			
			sintesi2=""
			
			if Valutato=1 then
						 if len(trim(sintesi))<40 then
							 
								voto=0
								Segnalata=1
								sintesi2="         SPIEGAZIONE TROPPO CORTA!"
							 
						 else 
						       if DVAbilitato=1 then 
									randomize()
									rand=rnd()
									if len(trim(sintesi))>400 then
										 if (clng(left(rand,1)) mod 2)= 0 then ' se il numero casuale è pari (testa o croce)
											Voto=2
											sintesi2=  "             HAI OTTENUTO 1 BONUS!"
										 else
											Voto=1
											sintesi2= "             POTEVI OTTENERE 1 BONUS! MA NON SEI STATO FORTUNATO"					
										end if
										
									else
										Voto=1
										
									end if
									Segnalata=0
							  else
							       Voto=1
								   Segnalata=0
							  end if  ' if DVAbilitato=1 then 
						 end if	 '  if len(trim(sintesi))<40 then
				
				
				else					
						Voto=votobase
				end if ' if Valutato=1 then
			
	 
	 
	 
	 
	 
 
	'response.write(url)
		if (session("Admin")=True)  then 
		  
		  ' Per aggiornare anche il voto
		 ' QuerySQL ="UPDATE Nodi SET Chi = '" & Chi & "', Cosa= '" & Cosa & "',Dove= '" & Dove & "',Quando= '" & Quando & "', Come= '" & Come & "', Perche= '" & Perche & "', Quindi = '" & Quindi & "', Voto = '" & voto & "',Data='" & DATA &"',Segnalata='" & Segnalata &"'  WHERE CodiceNodo =" &ID&";"
		'  
		' senza aggiornare il voto 
			QuerySQL ="UPDATE Nodi SET Chi = '" & Chi & "', Cosa= '" & Cosa & "',Dove= '" & Dove & "',Quando= '" & Quando & "', Come= '" & Come & "', Perche= '" & Perche & "', Quindi = '" & Quindi & "',Data='" & DATA &"',Segnalata='" & Segnalata &"',Voto=" & request.form("txtVAL") & "  WHERE CodiceNodo =" &ID&";"
		 'response.write("1"&QuerySQL)
			
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		   QuerySQL ="UPDATE Nodi SET Chi = '" & Chi & "', Cosa= '" & Cosa & "',Dove= '" & Dove & "',Quando= '" & Quando & "', Come= '" & Come & "', Perche= '" & Perche & "', Quindi = '" & Quindi & "',Segnalata='" & Segnalata &"'  WHERE CodiceNodo =" &ID&";"
		  'response.write("2"&QuerySQL)
		   end if 
		   
		end if

'	 response.write(ucase(session("CodiceAllievo")) & "="& ucase(CodiceAllievo))
'		response.write("<br>?"&session("Admin")) 
		 ConnessioneDB.Execute(QuerySQL)
	 
	
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DEL NODO
	'response.write(url)
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Create the FSO.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'CANCELLA LA VECCHIA VERSIONE DEL FILE11
	objFSO.DeleteFile url
	'On Error Resume Next
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
    objCreatedFile.WriteLine(Sintesi)
	objCreatedFile.WriteLine(Sintesi2)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
	
	'On Error Resume Next
	
	response.redirect "../cClasse/quaderno.asp?stile="&session("stile")&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&DataClaq2="&Session("DataClaq2")&"&DataClaq="& Session("DataClaq")&"&tCap="&tCap&"&tSot="& tSot&"&tNod="& tSot
	
	If Err.Number = 0 Then
	%><span class="alert-success">
	<%Response.Write "Modifica avvenuta! "
	
	response.redirect "../cClasse/quaderno.asp?stile="&session("stile")&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&DataClaq2="&Session("DataClaq2")&"&DataClaq="& Session("DataClaq")&"&tCap="&tCap&"&tSot="&tSot&"&tNod="&tNod
	Else%>
    <span class="alert-error">
	<%Response.Write Err.Description 
	'response.write(errore)
	Err.Number = 0
	End If%>
    </span>
	<%

else%>
    <span class="alert-error">
    <%'response.write(len(Chi)&" " &len(Cosa)&" " &len(Dove)&" " &len(Quando)&" " &len(Come)&" " &len(Perche)&" " &len(Quindi)&" " &len(Sintesi)&" ") %>
    <%response.write(errore& " Controlla che non ci siano campi lasciati vuoti")%>
 </span>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%

end if 





   %>
	</font>   
	<% if davalutazione<>"" then ' se sono stata chimata da inserisci_valutazione devo tornare alla pagina studente.asp altrimenti no %> 
		  <h5><a href="../cClasse/quaderno.asp?DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>&cla=<%=cla%>&CodiceAllievo=<%=CodiceAllievo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua 
		a valutare o modificare i nodi...</a></h5>
	      <% else %>
		
		
		
      <h5><a href="../cClasse/studente_quiz.asp?Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&testnodo=1">Continua a modificare i nodi...</a></h5>
	
	<%end if %>
	
	<p>&nbsp;</p>
	 
			
	<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
 <h5 ><a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>"> Torna a scelta verifica </a></h5> 	
 									 
	 
	 
				 
				 
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

