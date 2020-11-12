<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Spiegazione nodi</title>   
   
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

<% 'per il copia incolla

 
  Response.Buffer = true
  'On Error Resume Next  
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
 
 daQuaderno = Request.QueryString("daQuaderno")
	
	if daQuaderno <> 1 then daQuaderno = 0 end if
 
  QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	Nodi=rsTabellaCI("Nodi")
	
if daQuaderno = 1 then
	Nodi = 0
	end if
	
	rsTabellaCI.close
' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close

 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  Codice_Test=Request.QueryString("CodiceTest") 
  Codice_Allievo=Request.QueryString("CodiceAllievo") 
   Dim objFSO, objTextFile
  Dim liv(8) ' serve per indicizzare il chi,cosa,....
  liv(1)="Chi"
  liv(2)="Cosa"
  liv(3)="Dove"
  liv(4)="Quando"
  liv(5)="Come"
  liv(6)="Perch&egrave;"
  liv(7)="Quindi"
  
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Spiegazione nodi</h1> 
                    
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
							<a href="#">Apprendimento</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							 <a href="#">Nodi</a> 
                             
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%> : <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
	<%			 
				 
                   costQuerySQL1="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Paragrafi.Posizione, Nodi.Cartella" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Cartella "

costQuerySQL2="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Paragrafi.Posizione,Nodi.Id_Stud,Nodi.Cartella" &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Chi,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.Data,Nodi.In_Quiz,Paragrafi.Posizione,Nodi.Id_Stud,Nodi.Cartella "

'if (clng(Stato)=0) or (clng(Stato0)=0) then  
 if clng(Stato)=0 then
 'Definzione codice SQl della query per ricercare i nodi del paragrafo 
  
   if (Session("Admin")=True) or (Privato=0) or (Nodi=1) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato  
		QuerySQL=costQuerySQL1 &_
		" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' and Nodi.Chi<>'?'  " &_   
		" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
	else
	    QuerySQL=costQuerySQL2 &_
		" HAVING Paragrafi.ID_Paragrafo='" & Codice_Test & "' and Nodi.Chi<>'?' and Nodi.Id_Stud='"& Session("CodiceAllievo") &_   
		"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
	end if

else 		 

    if (Session("Admin")=True) or (Privato=0) or (Nodi=1) then  'se vero visualizzo tutte i nodi del paragfrafo altrimenti solo quelle dello       studente loggato
		QuerySQL= costQuerySQL1 &_
		" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Chi<>'?' " &_ 
		" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
    else
	    QuerySQL=costQuerySQL2 &_
		" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Chi<>'?'  and Nodi.Id_Stud='"& Session("CodiceAllievo") &_   
		"' ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;"
	
	end if

end if    
'response.write(querySQL)

'la parte che segue sostituisce questa sopra nella definizioane della query da esguire






 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 cartella=rsTabella("Cartella")
      
%>
<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
 
%><center>
  <H4>Nodi della rete non ancora disponibili!</h4></center>
  
<% Else
  
	  i=1 'inizializza la variabile i (contatore delle domande)
	  Do until rsTabella.EOF
	  'response.Write(rsTabella(12))
		if (strcomp(rsTabella(12),"12/12/2112")<>0) then  'apro l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
					 
				 
					ID=rsTabella(3)
					url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
					 ' NB c'è una / nell'url locale
				
					' url=Server.MapPath("/ECDL") & "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
					   url1= "../" & Cartella & "/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
				
				url=Replace(url,"\","/")
				 
				'response.write(url)
				' Open file for reading.
				Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
			'	sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
				%>
				<%' devo controllare se ID nodo esiste nella tabella dei link in tal caso leggo la L1 ed in quella posizione invece dell'ancora metto href
										  '0		   1		 2			3		4			5          6
				QuerySql="Select Link.ID_Link, Link.Id_n1, Link.L1, Link.Id_n2, Link.L2, Link.Id_Stud,Link.Testo2 FROM Link WHERE Id_n1="&ID&";"
				 
			
				Set rsLink = ConnessioneDB.Execute(QuerySQL)
				If rsLink.BOF=True And rsLink.EOF=True Then  ' se il nodo non compare nella tabella link allora metto tutte ancore
				%>
			
					  <table class="table table-hover table-nomargin table-bordered table-condensed" >
							<thead><tr>
							   <th width="10%"><b>Nodo n</b>.<%=rsTabella.fields("CodiceNodo")%></td>
							  <th width="60%"><%=rsTabella.fields("Titolo")%></td>
							  <th width="20%"><%=rsTabella.fields("Cognome")%></td>
							</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
                            </thead>
                            <tbody>
							<tr><td><b><a name="<%=ID%>_1">Chi</a></b></td><td colspan=3><p align="center"><b><%=rsTabella.fields("Chi")%></b></th></tr>
							<tr><td><b><a name="<%=ID%>_2">Cosa</a></b></td><td colspan=2><p align="center"><%=rsTabella.fields("Cosa")%></td></tr>
							<tr><td><b><a name="<%=ID%>_3">Dove</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Dove")%></td></tr>
							<tr><td><b><a name="<%=ID%>_4">Quando</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Quando")%></td></tr>
							<tr><td><b><a name="<%=ID%>_5">Come</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Come")%></td></tr>
							<tr><td><b><a name="<%=ID%>_6">Perch&egrave;</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Perche")%></td></tr>
							<tr><td><b><a name="<%=ID%>_7">Quindi</a></b></td><td colspan=3><p align="center"><%=rsTabella.fields("Quindi")%></td></tr>
							<tr>
							<td colspan=3>
							<p align="center">
							 <textarea rows="<%=1+round((len(sReadAll))/70)%>" name="TestoDomandaPlus" class="input-block-level"><% 
							 Response.write(sReadAll)%> </textarea><br>
							</td>
							</tr>
                            </tbody>
				</table>
				<br>
				<%else ' devo mettere href nel livello indicato %> 
					
					
					<table class="table table-hover table-nomargin table-bordered" >
							<thead><tr>
							  <th width="10%"><b>Nodo n</b>.<%=rsTabella.fields("CodiceNodo")%></td>
							  <th width="60%"><%=rsTabella.fields("Titolo")%></td>
							  <th width="20%"><%=rsTabella.fields("Cognome")%></td>
							  <th width="10%">Link to</td>
							  </tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
							</thead>
							<%' per ogni livello di ogni nodo vedo i link che ha ad altri nodi, e metto una stellina per ognuno
							  ' per ogni livello controllo il rsLink, se trovo che il livello è coinvolto in un link metto href, la prima volta metto il <td> le altre aggiungo allo stesso <td>
							   for i=1 to 7 %>
                               <tbody>
							   <tr>
							   <td><b><a name="<%=ID%>_<%=i%>" title="<%=ID%>_<%=i%>"><%=liv(i)%></a></b></td><td colspan=2><p align="center"><%=rsTabella(4+i)%> </td>
								<td> 			
								<%	 rsLink.Movefirst()
									 Do until rsLink.EOF
											L1=rsLink("L1")
											Id_n1=rsLink("Id_n1")
											Id_n2=rsLink("Id_n2")
											L2=rsLink("L2")
											T2=rsLink("Testo2")
										   if i=L1 then%>													
												 <a href="#<%=Id_n2&"_"&L2%>" title="<%=Id_n2&"_"&L2&":"&T2%>"> <i class="glyphicon-link"></i></a> 
										   <%end if  
										  rsLink.Movenext()
										Loop%>
								</td></tr>
							  <% next
								
							 %>
							 
							<tr>
							<td colspan=4>
							<p align="center">
							 <textarea rows="<%=1+round((len(sReadAll))/100)%>" name="TestoDomandaPlus" class="input-block-level" ><% 
							 Response.write(sReadAll)%> </textarea><br>
							</td>
							</tr>
                            </tbody>
				</table>
				<br>	
				
				<%end if %>
			<%
			
       i = i+ 1 
	   	end if  'chiudo l'if che serve per saltare il nodo se è uno di quelli inseriti alla registrazione con data 12/12/2112 per il quale non esiste la spiegazione
	
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente i nodi
    Loop 
 End If 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 %>
   
   <br>
     <a target="_blank" href="../cMap/spiegazione_mappa.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>&cod=<%=Session("CodiceAllievo")%>&daQuaderno=<%=daQuaderno%>&idclasse=<%=Session("Id_Classe")%>">Apri rete di nodi interattiva</a> <br>
	 <br>
	  <a href="inserisci_collegamento.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci un collegamento nella rete</a>
 
	
  <!-- 
  <a href="../cClasse/scegli_azione_app.asp?Cartella=<%=Cartella%>&Stato=1&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>">	Indietro </a>
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

