<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Spiegazione Quiz</title>   
   
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
   
</head>

<%Function domandaplus()
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 Cartella=rsTabella.fields("Cartella")
	 Modulo=rsTabella.fields("ID_Mod")
	 'Paragrafo=rsTabella(15)
	 Paragrafo=rsTabella.fields("Titolo")
	' response.write("PARAGRAFO="&Paragrafo)
	 Id=rsTabella.fields("CodiceDomanda")
	'homesito="/anno_2010-2011_ITC/ECDL"
	 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	response.write(url)
	objTextFile.Close
End Function %>
 
 
 <%Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    

  'Dichiara la variabili per contenere i dati digitati dall'utente (codice allievo, password, codice corso
     'Dichiara le variabili per interagire con il data base (connessione, stringa per contenere la query, stringa per contenere i risultati della query

  
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
'homesito="/anno_2010-2011_ITC/ECDL"
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("Cartella")
  
  
  
  
  Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
  		'response.write("Stato"&stato)				
%>

 <%'response.write("Stati :  " & stato & " " & stato0) 
 if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY class='theme-<%=session("stile")%>' onLoad="showText2();"> </BODY>
  <% else %>
     <% if (CIAbilitato=0) then ' disabilito copia incolla%>
        <body class='theme-<%=session("stile")%>'  oncontextmenu="return false" ondragstart="return false" onselectstart="return false">  
        <%else%>
      
      <body class='theme-<%=session("stile")%>'>

        <%end if%>
  <% end if %>
	<div id="navigation">
     
        <% 
		
   'per il copia incolla
  ' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where  Id_Classe='" & session("Id_Classe") &"';"
'response.Write(QuerySQL)
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabella.fields("CIAbilitato") 
	Privato=rsTabella.fields("Privato") 
	VotoAttivo=rsTabella.fields("VotoAttivo") 
	rsTabella.close
	  
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		 
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Spiegazione Quiz </h1> 
                    
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
							<a href="#more-blank.html">Approfondimento</a>
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
				        <h3> <i class="icon-reorder"></i>  	<%=Capitolo & ":"&Paragrafo%> </h3>
			          </div>
				      <div class="box-content">
                      


                      
                      
 
 	<div class="row-fluid">
					<div class="span12">
						<div class="box">
							<div class="box-title">
								<h3>
									<i class="icon-table"></i>
								Elenco domande guidate <a title="Consulta le domande libere"  href="spiegazione_test_1.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>"><i class="icon-unlock"></i></a>
                                
                                 
								</h3>
							</div>
							<div class="box-content nopadding">
                            
                                                  <%   
  
 
 
if (clng(Stato)=0) or (clng(Stato0)=0) then 
' 'Definzione codice SQl della query per ricercare le domande del paragrafo 
	
  if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le domande del PARAGRAFO altrimenti solo quelle dello       studente loggato  
  
	QuerySQL="SELECT * from PREDOMANDE1 where ID_Paragrafo='" & Codice_Test & "' order by Id_Predomanda;"
	else
	 
	end if 

else 

  if (Session("Admin")=True) or (Privato=0) then  'se vero visualizzo tutte le domande del MODULO altrimenti solo quelle dello       studente loggato  
   	
 QuerySQL="SELECT * from PREDOMANDE1 where ID_Mod='" & Modulo & "' order by Id_Predomanda;"  
 else
  
 end if
 

end if    
    
Set rsTabella0 = ConnessioneDB.Execute(QuerySQL)	
 'response.write(QuerySQL) 
' per ogni record di rsTabella0 faccio una query per cercare tutte le riposte con ID_predomanda
  If rsTabella0.BOF=True And rsTabella0.EOF=True Then %>
      <div class="alert alert-error">
                       Domande guidate non presenti!
                        
      </div>
<%else%>
 
  
 
<% If rsTabella0.BOF=True And rsTabella0.EOF=True Then %>
  <div class="alert alert-error">
                    Domande del Test non ancora disponibili!
                     
 </div>
 <%else%>
  
 <% k=1 'inizializza la variabile i (contatore delle domande)
Do until rsTabella0.EOF %>                                   
                                    
 <%
 QuerySQL="SELECT * from PREDOMANDE2 where Id_Predomanda=" & rsTabella0("ID_Predomanda") & " order by Id_Predomanda;"  
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   'response.write("<br>"&QuerySQL) 
   i=1 'inizializza la variabile i (contatore delle domande)
 %>
  <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse2_<%=k%><%=i%>">
												 <b><%=k%>) <%=rsTabella0("Quesito")%></b> 
											</a>
										</div>
										<div id="collapse2_<%=k%><%=i%>" class="accordion-body collapse">
											<div class="accordion-inner">
 <%
 
  
 ' per scorrere tutte le risposte della domanda guidata
  Do until rsTabella.EOF
  		 
 
    ID=rsTabella("CodiceDomanda")
   url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&rsTabella("Modulo")&"_"&ID&".txt"
   url=Replace(url,"\","/")
 
              
 
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
 
sReadAll = objTextFile.ReadAll
sReadAll = url
objTextFile.Close   ' la soluzione seguente la rimuovo e dirò di copiare ed incollare la domanda plus nella spiegazione
' così da avere il livello di apprendimento comprensibile , diversamente dovrei prevedere il modo di far apparire il testo della domanda plus 
' anche nell'approfondimento di fine quiz.
'if clng(rsTabella.fields("Tipo"))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
'	    url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&rsTabella(0)&"_"&ID&".txt"
'		url=Replace(url,"\","/")
'		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
'		sReadAll1 = objTextFile.ReadAll
'		objTextFile.Close
'end if
			 
%>
                              
  
  
  
  
    
  
 <table  class="table table-hover table-nomargin table-condensed table-bordered">
		 
        <tr>
			 
			<th> <%=rsTabella("Cognome")%>&nbsp;<%=left(rsTabella("Nome"),1) &"."%> </th>
            <th>x <img src="../cSocial/img/icon_star_red.gif" width="13" height="12">
            y <img src="../cSocial/img/icon_star_black.GIF" width="13" height="12"> </th>
            <th  class='hidden-350'> <%=rsTabella("Data")%></th>
            <th class="hidden-480"><%=rsTabella("CodiceDomanda")%> </th>
			
		</tr>
	 
		
		<% if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
	    <tr><td colspan="4"><p align="center">
 <textarea rows="<%=1+round((len(domandaplus()))/50)%>" name="TestoDomandaPlus0" value="ciao" class="input-block-level"><%
			 
			 
			 Response.write(domandaplus())%> </textarea><br></td></tr><br>
        <%end if %>
   
		<tr>
			<td colspan=4>
			
			<p align="center">
			 <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" class="input-block-level"><%
			 ' if clng(rsTabella(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			 
			 Response.write(sReadAll)%> </textarea> 
             <% 
			 
			 if VotoAttivo=1 then%>
         <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse<%=k%><%=i%>">
												<center><b><i class="icon-star"></i></b></center>
											</a>
										</div>
										<div id="collapse<%=k%><%=i%>" class="accordion-body collapse">
											<div class="accordion-inner">
                                                       <center>
<a title="Fai da 1 a 5 click per esprimere quanto ti piace (Voto da 6 a 10)  " href="../vota_compito.asp?scegli=<%=scegli%>&ID=<%=iMessageId%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&MaxStelline=<%=MaxStelline%>"><img src="../cSocial/img/facebook2.jpg" width="21" height="19" align="bottom">&nbsp;Mi piace&nbsp;<img src="../cSocial/img/icon_star_red.gif" width="13" height="12"></a>   &nbsp;&nbsp;

 
<a title="Fai da 1 a 5 click per esprimere quanto non ti piace (Voto da 5 a 0) " href="../vota_compito.asp?scegli=<%=scegli%>&revoca=1&ID=<%=iMessageId%>&CodiceAllievo=<%=Session("CodiceAllievo")%>&CodiceAllievoPost=<%=CodiceAllievo%>&IDPARENT=<%=iThreadParent%>&MaxStelline=<%=MaxStelline%>">
<img src="../cSocial/img/facebook8_nonpiace_small.jpg" width="20" height="17">&nbsp;Non mi piace&nbsp;<img src="../cSocial/img/icon_star_black.GIF" width="13" height="12"></a>
</center> 
    <br>
  <% if Session("Admin")=True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
     <ul> <li><a href="#">????</a></li>
	  
	  
	  </FIELDSET>
      <% end if%>
 </ul>
</p> 
                                             
                                             
											</div>
										</div>
									</div>
             
             <%end if ' votoCompitoAbilitato%>
		      </td>
		 
		</tr>
 
     </tbody>
	</table>
    
    
         
                                             
										
    
	<br>
<%    

       i = i+ 1 
       rsTabella.MoveNext ' passa alla successiva riga della tabella contenente le domande
    Loop 

   k = k+ 1 
       rsTabella0.MoveNext
	   %>
       
       
       	</div>
										</div>
									</div>
       <%
 
Loop 
 
 rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
  rsTabella0.Close : Set rsTabella = Nothing 
 ' ConnessioneDB.Close : Set ConnessioneDB = Nothing 
 end if
 %>
<% end if%>							 
                             
    
                             
							</div>
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
                            
						</div>
					</div>
				</div>								 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		   <!-- <div class="box-content"> 
                     
                      <div class="alert alert-error">
                     KO..
                     </div>
                     
                     <div class="alert alert-success">
                     OK
                     </div>
                     -->
                      
                      
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

