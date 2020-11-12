<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Valutazioni nodo</title>   
   
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
       
       
       
       
<!--Controllo accesso quaderno e sessione scaduta con redirect ad index.html-->
       <script src="../js/privacy.js"></script>
  
  
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  
<% x = Request.ServerVariables("HTTP_REFERER")
if x = "" then %>
<script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")
 
 //alert("<%=x%>");
 location.href="../../../../index.html";
 
//location.href=window.history.back();
 }
 </script>
 
 <% else %>
 <script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")
 
 location.href="<%=x%>";
 
//location.href=window.history.back();
 } 
 </script><% end if%>

   
</head>

<%
  Response.Buffer = true
  On Error Resume Next  
   





 
  
 
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
 
  Dim objFSO, objTextFile 
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
 
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<% 
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceAllievo=Request.QueryString("cod")
  cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceDomanda=Request.QueryString("CodiceDomanda")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Chi=Request.QueryString("Chi")
  Cosa=Request.QueryString("Cosa")
  Dove=Request.QueryString("Dove")
  Quando=Request.QueryString("Quando")
  Come=Request.QueryString("Come")
  Perche=Request.QueryString("Perche")
  Quindi=Request.QueryString("Quindi")
  MO=Request.QueryString("MO")
  VAL=Request.QueryString("VAL")
  URL=Request.QueryString("URL")
  DATA=cdate(Request.QueryString("DATA"))
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  ID=CodiceDomanda 
  Cartella=Request.QueryString("Cartella")
   Segnalata=Request.QueryString("Segnalata")
   if Segnalata="" then
     Segnalata=0
   end if
if MO<>"" then 
 Modulo=MO
end if  


 tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 tNod=request.querystring("tNod")
  

 ' codice per permettere la visualizzazione solo delle proprie domande 
QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
  

 
Set objFSO = CreateObject("Scripting.FileSystemObject")
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Replace(url,"\","/")
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll
'sReadAll=url
'response.write(sReadAll)
objTextFile.Close
 
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine
  


Set objFSO = CreateObject("Scripting.FileSystemObject")
%>
   

<%  if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
		   <% if (CIAbilitato=0) then ' disabilito copia incolla%>
        <body class='theme-<%=session("stile")%>'  oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onLoad="Effect.toggle('dAttività','appear');Effect.toggle('dAvvisi','appear'); return false;">  
        <%else%>
         <body class='theme-<%=session("stile")%>'>
         
        <%end if%>
  <% end if %>





	<div id="navigation">
     
   
	
		 
        <!-- #include file = "../var_globali.inc" --> 
 		 
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
						<h1> <i class="glyphicon-snowflake"></i> Valutazioni nodi</h1> 
                    
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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%> : <%=Paragrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 
 
 <form method="POST" class="form-vertical" action="inserisci_modifica_nodo1.asp?davalutazione=1&VALORE=<%=VAL%>&Cartella=<%=Cartella%>&cla=<%=cla%>&cod=<%=CodiceAllievo%>&CodiceNodo=<%=CodiceDomanda%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&MO=<%=MO%>&tCap=<%=tCap%>&tSot=<%=tSot%><%=p%>&tNod<%=tNod%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
  	
	 <p align="center">
     
     
     
		<p><font size="4" color="#FF0000">Codice Nodo (<%=CodiceDomanda%>)</font><br>
  </p>

<p><input type="text"  class="input-xxlarge" name="txtChi"  value="<%=Chi%>" size="100" maxlength="250"><b> Chi</b></p> 
<p><input type="text"  class="input-xxlarge" name="txtR1Cosa" value="<%=Cosa%>" size="100" maxlength="150"><b> Cosa</b></p>
<p><input type="text"  class="input-xxlarge" name="txtR1Dove" value="<%=Dove%>" size="100" maxlength="150"><b> Dove </b></p>
<p><input type="text"  class="input-xxlarge" name="txtR1Quando" value="<%=Quando%>"  size="100" maxlength="150"><b> Quando</b></p>
  <p><input type="text"  class="input-xxlarge" name="txtR1Come" value="<%=Come%>" size="100" maxlength="150"><b> Come</b></p>
  <p><input type="text"  class="input-xxlarge" name="txtR1Perche" value="<%=Perche%>" size="100" maxlength="150"><b> Perch&egrave;</b></p> 
  <p><input type="text"  class="input-xxlarge" name="txtR1Quindi" value="<%=Quindi%>" size="100" maxlength="150"><b> Quindi </b></p>
	<b>Sintesi</b>
	<% if (ucase(session("CodiceAllievo"))<> ucase(CodiceAllievo)) and not(session("Admin")=true) then %>
	<p><textarea  class="input-block-level" rows="<%=1+round((len(sReadAll))/60)%>"  name="S1"  disabled="disabled"><%=Response.write(sReadAll)%> </textarea></p>
     <%else%>
	 <p><textarea class="input-block-level" rows="<%=1+round((len(sReadAll))/60)%>"  name="S1"   ><%=Response.write(sReadAll)%> </textarea></p>
	 <%end if%>
<%if (session("Admin")=true) then %>
<p><input type="text"  class="input-small" name="txtDATA" value="<%=DATA%>" size="8"><b> 
	Data </b> </p>
 <p><input type="text"  class="input-mini" name="txtVAl" value="<%=VAL%>" size="1"><b> 
	Valutazione </b><br>
      
      
      
          <span title="Feedback all'autore"><b>Segnalata</b></span> 
											 
                                             <% if (Segnalata=1)  then  %>
                                            
											 <INPUT TYPE="RADIO" name="txtSegnalata" checked="true" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtSegnalata"  value="0">No  	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtSegnalata" value="1">Si  
                                             <INPUT TYPE="RADIO" name="txtSegnalata"   checked="true" value="0">No  
                                           
										<% end if %> 
      
      
      
    
  <p><input type="submit" value="Invia" name="B1" class="btn-primary"> </p> <!--Definisce i due bottoni del form -->
<% else 

   'response.write(ucase(session("CodiceAllievo") &"=?" &ucase(CodiceAllievo) )
   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>
 <p><input type="text" disabled="disabled" name="txtVAl" value="<%=VAL%>" size="1"><b> 
	Valutazione </b></p>
   <p><input type="submit" value="Invia" name="B1" class="btn-primary"> </p> <!--Definisce i due bottoni del form -->
<% end if 
end if %>
<p><a target="_new" href="7_stampa_scheda_nodo.asp?CodiceAllievo=<%=CodiceAllievo%>&CodiceNodo=<%=CodiceDomanda%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&Chi=<%=Chi%>&Cosa=<%=Cosa%>&Dove=<%=Dove%>&Quando=<%=Quando%>&Come=<%=Come%>&Perche=<%=Perche%>&Quindi=<%=Quindi%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>"><img src="../../img/printer.jpg" alt="Stampa questa scheda"></a></p>

</form> <!-- Chiude l'interfaccia -->
 
 
 
  
 
   
 

 <hr>

 
<%	end if	  	%>		   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        

			 
	</body>
    <% else%> 
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if %>
   
 <script language="javascript" type="text/javascript"> 
function stampa() {
    document.dati.action = "7_stampa_schede_frasi.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&QuerySQL=<%=QueryPrima%>";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
}
 </script>

 </html>

