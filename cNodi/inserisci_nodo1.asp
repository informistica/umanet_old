<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci nodo</title>   
   
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
       
       
       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
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
  


   
</head>
<% Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    %>
 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
  <% Response.Buffer = true
  prenodo=Request.QueryString("prenodo") ' serve per capire il chiamante e quindi sapere se alla fine devo redirectare ad home_ver o home_app
   ID_Prenodo=Request.QueryString("ID_Prenodo") 
   
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
      
     <% if ID_Prenodo<>"" then
	  querySQL="Select * from Nodi where Id_Stud='" & Session("CodiceAllievo") & "' and (Id_Prenodo="&clng(ID_Prenodo)&");" 
	  response.write(querySQL)
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	          If not(rsTabella.BOF=True And rsTabella.EOF=True) Then%>	
	 
                      <BODY onLoad="showText3();"> 
                                Stai per essere reindirizzato all'home page ... </BODY>           
                <% end if %> 
	  <%end if%>		
	    
  <body class='theme-<%=session("stile")%>'>
       
  <% end if %>
  <%
  
   Dim Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Sintesi
   Dim Chi,Cosa,Dove,Quando,Come,Perche,Quindi
   
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   CodiceTest = Request.QueryString("CodiceTest")
   Modifica=Request.QueryString("Modifica")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   
   by_UECDL=Request.QueryString("by_UECDL")
  %>

	<div id="navigation">
	
		 
        
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
       
         
	</div>
    
  
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserimento nodo</h1> 
                    
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
							 <a href="#">Inserisci nodo</a>                            
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%>: <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	<% 
		
	
		 
	   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
	   DataTest = Request.Cookies("Dati")("DataTest")
	   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
	   CodiceCap=Request.Cookies("Dati")("CodiceCap")
	  Num=Request.QueryString("Num")
	Capitolo=Request.QueryString("Capitolo")
	 CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
	
	Paragrafo=Request.QueryString("Paragrafo")
	Modulo=Request.QueryString("Modulo")
	DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
	 '  if tipo="0" then
		   Chi = Request.Form("txtChi")
		   Chi = Replace(Chi,  Chr(34),Chr(96))
		   Chi = Replace(Chi,  "'",Chr(96))
		    
		  
	  
	   Cosa = Request.Form("txtR1Cosa")
	   Cosa = Replace(Cosa, Chr(34),Chr(96))
	   Cosa = Replace(Cosa, "'",Chr(96))
	    
	   
	
	
	   Dove = Request.Form("txtR2Dove")
	   Dove = Replace(Dove, Chr(34),Chr(96))
	    Dove = Replace(Dove, "'",Chr(96))
	   
	   
	
	   Quando = Request.Form("txtR3Quando")
	   Quando = Replace(Quando, Chr(34),Chr(96))
	   Quando = Replace(Quando, "'",Chr(96))
	 
	  
	   Come = Request.Form("txtR4Come")
	   Come = Replace(Come, Chr(34),Chr(96))
	   Come = Replace(Come, "'",Chr(96))
	    
	
	   Perche = Request.Form("txtR5Perche")
	   Perche=  Replace(Perche,Chr(34),Chr(96))
	   Perche=  Replace(Perche,"'",Chr(96))
	  
	   
	   Quindi = Request.Form("txtREQuindi")
	   Quindi = Replace(Quindi,Chr(34),Chr(96))
	   Quindi = Replace(Quindi,"'",Chr(96))
	 
	  
	   
	   Sintesi=Request.Form("S1")
	   Sintesi= Replace(Sintesi, Chr(34),Chr(96))
	   Sintesi= Replace(Sintesi, "'",Chr(96))
	   Sintesi= Replace(Sintesi, Chr(130), chr(138))
	  
	
	   
	   
	   
	   if ( (len(Chi)=0) or (len(Cosa)=0) or (len(Dove)=0) or (len(Quando)=0) or (len(Come)=0) or (len(Perche)=0) or(len(Quindi)=0) ) then
	 '  Response.Redirect("inserisci_test.asp?Cartella=Cartella&Num=0&Cognome=Cognome&Nome=Nome&CodiceTest=CodiceTest&Capitolo=Capitolo&Paragrafo=Paragrafo&Modulo=Modulo") 
	   ' Response.Redirect("inserisci_test.asp") 
	   errore=2
	  
	   end if
	   
	 if (errore=0) then
	   
	   ' devo vedere se il setting è tale da richiedere voto=1 come default oppure no  
	'	QuerySQL1="Select * from Setting"
'		Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
'		Valutato=rsTabella.fields("Valutato") 
'		DVAbilitato=rsTabella.fields("DVAbilitato")
'		rsTabella.close
'	
'	
'	
	
	
	'	if Valutato=1 then
'			 if len(sintesi)>300 then
'				voto=2
'			 else
'				 if len(sintesi)<40 then
'					voto=0
'				 else
'					 Voto=1 ' valore di default
'				 end if
'			 end if 
'		else
'			Voto=0
'		end if
			
		'	Sintesi2=""
'			
'			if Valutato=1 then
'						 if len(trim(sintesi))<40 then
'							 
'								voto=0
'								Segnalata=1
'								Sintesi2="         SPIEGAZIONE TROPPO CORTA!"
'							 
'						 else 
'						       if DVAbilitato=1 then 
'									randomize()
'									rand=rnd()
'									if len(trim(sintesi))>400 then
'										 if (clng(left(rand,1)) mod 2)= 0 then ' se il numero casuale è pari (testa o croce)
'											Voto=2
'											Sintesi2=  "             HAI OTTENUTO 1 BONUS!"
'										 else
'											Voto=1
'											Sintesi2= "             POTEVI OTTENERE 1 BONUS! MA NON SEI STATO FORTUNATO"					
'										end if
'										
'									else
'										Voto=1
'										
'									end if
'									Segnalata=0
'							  else
'							       Voto=1
'								   Segnalata=0
'							  end if  ' if DVAbilitato=1 then 
'						 end if	 '  if len(trim(sintesi))<40 then
'				
'				
'				else					
'						Voto=0
'				end if ' if Valutato=1 then
'			
'	
	
	
	QuerySQL1="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
	Valutato=rsTabella.fields("Valutato") 
	DVAbilitato=rsTabella.fields("DVAbilitato")
	
	rsTabella.close
	if Valutato=1 then
		  if len(trim(sintesi))<40 then
				    if Img<>"" then ' se carico l'immagine va bene anche senza commento
						voto=1
						Segnalata=0
					 else
					    voto=0
						Segnalata=1
						Sintesi2="         TROPPO CORTA!"				    
					 end if
					 
		  else
		      voto=1
			  segnalata=0
		  end if
			
				
			if DVAbilitato=1 then 
				randomize()
				rand=rnd()
				randomize()
				rand2=rnd()
				
				if len(trim(sintesi))>400 then
				     if ((clng(left(rand*100,1)) mod 2)= 0) and ((clng(left(rand2*100,1)) mod 2)= 0)then ' se il numero casuale è pari (testa o croce)
					 	Voto=2
						Sintesi2= "             HAI OTTENUTO 1 BONUS!"
					 else
					    Voto=1
						Sintesi2= ""					
					end if
				else
				     if len(trim(sintesi))<40 then
				     
					    voto=0
					 else
					     voto=1		    
					 end if
					
				end if
				Segnalata=0
			end if	
	      
					 
	 		   
			 
	else
			Voto=0
	end if
	
				
			
			
		
		if strcomp(preNodo,1)=0 then
		   ID_Prenodo=clng(ID_Prenodo)
		else
		   ID_Prenodo=0
		end if
		
		
	  QuerySQL="INSERT INTO Nodi (Chi, Cosa, Dove,Quando,Come,Perche,Quindi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,ID_Prenodo,In_Quiz,ID_Sottoparagrafo,NLink) SELECT '" & Chi & "','" & Cosa & "', '" & Dove & "','" & Quando & "','" & Come & "','" & Perche & "','" & Quindi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) & "'," & ID_Prenodo &"," & Session("In_Quiz") &",'" & CodiceSottoPar & "',0;" 
	 ' response.write(QuerySQL&"<br>")
	   ConnessioneDB.Execute QuerySQL 
	  
	'	prelava ID dell'ultimo record inserito
	
		QuerySQL = "SELECT CodiceNodo,Cartella FROM Nodi WHERE CodiceNodo=(Select Max(CodiceNodo) FROM Nodi);" 
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		ID=rsTabella(0)
		CARTA=rsTabella(1)
		
		url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" 'per il server on line
	   'url=Server.MapPath("/ECDL/")& "/" & CARTA &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" ' per localhost
	   
		'response.write(url)
	   ' url1=  "../" &CARTA & "/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	   
	'    
		  
	
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DEL NODO
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
	 
	url3=Replace(url,"\","/")
	url=url3
	'response.write(url3)
	'response.write("<br>URL:"&url)
	
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
	objCreatedFile.WriteLine(Sintesi)
	objCreatedFile.WriteLine(Sintesi2) ' messaggio su eventuale bonus
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
	'response.write(url)
	if Tipo="1" then 'CREAZIONE FILE DI TESTO PER INSERIRE LA DOMANDA
	
		url4=Replace(url4,"\","/")
		 
		Set objCreatedFile = objFSO.CreateTextFile(url4, True)
		' Write a line with a newline character.
		objCreatedFile.WriteLine(Domanda)
		'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
		objCreatedFile.Close
	end if 
	'response.write("<br>" & url)
	
	'On Error Resume Next
	%>
    <div class="alert-success">
	<%If Err.Number = 0 Then
	
	Response.Write "Inserimento avvenuto! "
	Else %>
	<div class="alert-error">
	<% Response.Write Err.Description 
	Err.Number = 0
	End If
  %>
 
	   
		 
			
		  <h5><a href="inserisci_nodo.asp?Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua ...</a></h5>
		 
				
				<%
				 if (prenodo<>"") then 'se sono stato chiamato da compilaprenodo devo ritornare ad home_app
				  ' se sono stato chiamato da compilaprefrase di home_app_uecdl devo tornare li
						if by_UECDL<>"" then %>
					   <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
		<h5><a href="../cClasse/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Apprendimento... </a></h5> 
					<%else%>
						<!-- REDIRECT INTELLIGENTE  -->
		<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna al Libro </a></h5> 
		 
          <h5><a href="1compilaprenodo.asp?Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&CodiceTest=<%=CodiceTest%>&Modulo=<%=Modulo%>&prefrase=<%=prefrase%>"> Torna alla pagina dei Nodi... </a></h5> 
        
					<%end if	
					
				 else%>
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
	 <h5 ><a href="../cClasse/scegli_azione_test.asp?id_classe=<%=Session("Id_Classe")%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>"> Torna a scelta verifica </a></h5> 	
    
    
                
				<%end if 
	else%>
    <div class="alert-error">
	<%  if (errore=1) then
		 response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
	  end if 
	  if (errore=2) then
		response.write("Controlla che non ci siano campi lasciati vuoti")
	  end if %>
		<a href="#" onClick="history.go(-1);return false;">Indietro</a>
	  <%
	end if 			

  
%>
	</div> 
				 
				 
                   
                   
 
		  			  
                             
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

