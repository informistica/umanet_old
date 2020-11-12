<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica prenodo</title>   
   
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
    <script type="text/javascript" src="../js/selezionatutti.js"></script>
    
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
     
  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<%
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
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h3> <i class="icon-comments"></i> <%=Capitolo%>: <%=Paragrafo%></h3> 
                    
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
							<a href="#">Modifica prenodo</a>
                           
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
				        <h3> <i class="icon-reorder"></i>  NODI DISPONIBILI</h3>
			          </div>
				      <div class="box-content">
                      
 
 	<%
 Modifica=Request.QueryString("Modifica")
  BoxApro=Request.QueryString("BoxApro")
  
Elimina=Request.QueryString("Elimina")
NumRec=Request.QueryString("NumRec")
ID=Request.QueryString("ID")
 Id_Stud=Request.QueryString("Id_Stud")   ' se è settato vuol dire che aggiungo eccezioni per singolo stud

if Elimina<>"" then
 
 QuerySQL="Delete  " &_
"FROM preNodi WHERE preNodi.ID_Prenodo=" & ID & ";" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
QuerySQL="Delete  " &_
"FROM Nodi WHERE ID_Prenodo=" & ID & ";" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
'response.write(QuerySQL)



if Request.ServerVariables("HTTP_REFERER") <>"" then 
		response.Redirect request.serverVariables("HTTP_REFERER") 
end if

elseif Modifica="" then %>
 
				 


  <%Cartella=Request.QueryString("Cartella") 
  Capitolo=Request.QueryString("Capitolo") 
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")  
 
  if Id_Stud<>""then
     QuerySQL="SELECT *  FROM Allievi WHERE CodiceAllievo='" & Id_Stud & "'" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Cognome1=rsTabella("Cognome")
	Nome1=rsTabella("Nome")%>
	<p align="center"><font color="#FF0000" size="3">Modifica Scadenze per <%=Cognome1&" "%> <%=Nome1%> </font></p>
    <%
  end if
 
   'QuerySQL="SELECT Url, Data, Descrizione FROM VERIFICHE Where Classe='"& d &"'"
  
 QuerySQL="SELECT count(*) " &_
"FROM preNodi WHERE preNodi.Id_Paragrafo='" & CodiceTest & "'" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
NumRec=rsTabella(0)
Numero=clng(NumRec)
Dim dom()
Redim dom(Numero)

                                      ' 0			  1					2				3					4	
 QuerySQL="SELECT * " &_
"FROM preNodi WHERE preNodi.Id_Paragrafo='" & CodiceTest & "' order by Posizione" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySql & " " &Paragrafo)
 
'response.write("Numero ="&NumRec)

i=0
'paragrafo=rsTabella(2)
if rsTabella.eof and rsTabella.bof then%>
<span class="alert-error"><%=response.write("Non ci sono compiti assegnati")%></b></span> 
<%end if%>
<form method="POST" name="dati" class="form-vertical" action="2modificaprenodo.asp?BoxApro=<%=BoxApro%>&Modifica=1&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=CodiceTest%>&NumRec=<%=NumRec%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>">	
					 	
 
    <p>Scadenza: <input type="text" name="txtDataVal" id="datepicker" class="input-medium datepick" /></p>
    <input type="button" class="btn"  onClick="selezionatutti('datepicker')" value="Applica a tutti">
    <hr>
											 
 								 
<%do while not rsTabella.eof
	'if (i=0) or (StrComp(capitolo, rsTabella(0)) <> 0) then'
	dom(i)=rsTabella.fields("Quesito")
	 %>	
			<input type="text" class="hidden" value="<%=rsTabella.fields("id_Prenodo")%>" name="txtIdFrase<%=i%>" size="3" > 
			<fieldset><legend><%=i+1%> Nodo 	</legend>
           							 <div class="control-group">
										
										<div class="controls">
											<input type="text" value="<%=rsTabella.fields("Quesito")%>" name="txtFrase<%=i%>" class="input-xxlarge"> &nbsp;&nbsp; <img src="../../img/elimina.jpg" width="16" height="16"  onClick="elimina(<%=rsTabella.fields("id_Prenodo")%>);" title="Elimina"><br>
										</div>
									</div>
           
            <div class="control-group">
										
										<div class="controls">
                                        <b> <span title="Richiede caricamento immagine ?">Img</span> </b>
											  
                                             <% if (rsTabella.fields("Img")=1)  then  %>
                                         
											 <INPUT TYPE="RADIO" name="txtImg<%=i%>" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtImg<%=i%>"   checked="true" value="0">No  
                                           
										<% end if %>
										 
                                        &nbsp;&nbsp;&nbsp;
                                         <span title="Richiede caricamento file ?"><b>File</b></span> 
											 
                                             <% if (rsTabella.fields("Files")=1)  then  %>
                                            
											 <INPUT TYPE="RADIO" name="txtFile<%=i%>" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtFile<%=i%>"   checked="true" value="0">No  
                                           
										<% end if %>
                                        &nbsp;&nbsp;&nbsp;
                                        
                                             <span title="Posizione nella lista"><b>Pos</b></span> 
                                             <input  class="input-mini" title="Numero d'ordine" type="text" value="<%=rsTabella.fields("Posizione")%>" name="txtPos<%=i%>" size="1"  >
                                 &nbsp;&nbsp;&nbsp;
                                         <i title="Chiusura del compito" class="icon-calendar"></i>     <input type="text" value="<%=rsTabella.fields("Scadenza")%>" name="txtScadenza<%=i%>" id="scad<%=i%>"  class="input-small datepick"  />
    	 
										</div>
        
									</div>
                                    
            
             
           
            
            
            
			</fieldset>		 
				 
<%
	i=i+1  
	rsTabella.movenext
 
		loop%>

		<hr> 
        
        <div class="accordion" id="accordion2">
          <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse4">
												<center>Copia nodi</center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
                                            <textarea rows=<%=NumRec%> class="input-block-level">
 <% 
 for i=0 to NumRec-1
   response.write(dom(i)&chr(13))
 next   

 %>
 </textarea>
                                            </div>
                                         </div>
                                      </div>
                                   </div>
                                            
 								 
	 
 

   <input type="submit" value="Modifica" class="btn"> 
        </form>												 
			<br><br>
          <h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>#<%=BoxApro-3%>"> Torna al Libro... </a></h5> 
          
		  <% if len(Id_Stud) > 0 then %>
			
			<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>&Id_Stud=<%=Id_Stud%>">Torna al Libro (per modifica eccezioni)</a></h5> 
			
			<% end if %>
           
           	

 <br>

<% else ' aggiorno i campi ' aggiungo il test per capire se devo aggiungere eccezioni per Id_Stud  %>

 <%
 ' NumRec=Request.QueryString("NumRec")
  for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
	
   ID=Request.Form("txtIdFrase"&k)
   Quesito = Request.Form("txtFrase"&k)
   
   Img=Request.Form("txtImg"&k)
   cFile=Request.Form("txtFile"&k)
   if cFile="" then
      cFile=0
   end if	  
   Pos=Request.Form("txtPos"&k)
   Scadenza=Request.Form("txtScadenza"&k)
   if Scadenza="" then
      Scadenza=fine_anno
   end if
      Quesito = Replace(Quesito, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
	   Quesito=  Replace(Quesito,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  
   
   'TestoDomandaPlus=Request.Form("TestoDomandaPlus")
      if len(Id_Stud)>0 then ' aggiungo eccezioni
			
		   if DateDiff("D", Date(), Scadenza)>=0 then
				
			   QuerySQL="SELECT count(*) FROM Eccezioni_Nodi  WHERE  Id_Stud='"&Id_Stud&"' and Id_Prenodo='"&ID&"';" 
			   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			   ris=rsTabella(0)
				
				if ris=0 then
					QuerySQL="INSERT INTO Eccezioni_Nodi (Id_Prenodo,Id_Stud,Scadenza) SELECT '" & ID  & "','" & Id_Stud & "','" & Scadenza & "';"
					ConnessioneDB.Execute(QuerySQL)
				else
					QuerySQL = "UPDATE Eccezioni_Nodi SET Scadenza = '" &Scadenza&"' WHERE Id_Prenodo = '"&ID&"' and Id_Stud = '"&Id_Stud&"';"
					ConnessioneDB.Execute(QuerySQL)
				end if
		   end if
		   
	  else
	      QuerySQL ="UPDATE preNodi SET Quesito = '" & Quesito & "', Scadenza = '" & Scadenza & "', Img = " & Img & ", Posizione = " & Pos& ", Files = " & cFile &" WHERE Id_Prenodo =" &ID&";"
		  ConnessioneDB.Execute(QuerySQL)
	   end if	 
	     'response.Write(QuerySQL)
		  
		 
   next 
 %>
 
			<p><p>
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h5>Modifica Effettuata...</h5><br><br>
<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>">Torna al Libro</a></h5> 

<% if len(Id_Stud) > 0 then %>
	<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>&Id_Stud=<%=Id_Stud%>">Torna al Libro (per modifica eccezioni)</a></h5> 
<% end if %>	
		 
<% end if %>				 	 
				 
                 
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        
 <script language="javascript" type="text/javascript"> 
function elimina(id) {
    document.dati.action = "2modificaprenodo.asp?ID=" + id +"&Elimina=1";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
}
 </script>
			 
	</body>

 </html>

