<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   <meta charset="UTF-8">
   <title>Valutazione domanda</title>   
   
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

	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

  
       
<!--Controllo accesso quaderno e sessione scaduta con redirect ad index.html-->
       <script src="../js/privacy.js"></script>

	   
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

 <%if session("CodiceAllievo")="" then%>
	 <BODY onLoad="showText2();"> </BODY>
  <% ' torna all'homepage
   else%>
   <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  
   <%end if %>
   
 
 
 
 
	<div id="navigation">
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
          <!--#include file="../service/gestione_errori.asp" -->
     
	</div>
    
     
   <% 
   'Response.charset="iso-8859-1"
   function ReplaceCar(sInput)
dim sAns
' l'ho implementato nella pagina chiamante in javascript , sa il cazzo perchè non funzionava

  sAns=  Replace(sInput,Chr(39),Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi  
  sAns=  Replace(sInput,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi  
  sAns = Replace(sAns,"à","a"&Chr(96)) 
  sAns = Replace(sAns,"è","e"&Chr(96))
  sAns = Replace(sAns,"é","e"&Chr(96))
 ' sAns = Replace(sAns,"i"&Chr(96)) QUESTO ANNULLA TUTTO
 ' sAns = Replace(sAns,chr(237),"i"&Chr(96))
 ' sAns = Replace(sAns,"ò","o"&Chr(96))
''  sAns = Replace(sAns,chr(243),"o"&Chr(96))
 '  sAns = Replace(sAns,"ù","u"&Chr(96))
''  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
   sAns = Replace(sAns, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto
 
ReplaceCar = sAns
end function
   
id_classe=Request.QueryString("id_classe")
CodiceDomanda=Request.QueryString("CodiceDomanda")  
daQuaderno=Request.QueryString("daQuaderno")   ' vale 1 se sono chiamata dal quaderno dello studente tramite avviso

 tCap=request.querystring("tCap")
 tSot=request.querystring("tSot")
 tDom=request.querystring("tDom")
 tFra=request.querystring("tFra")
 tNod=request.querystring("tNod")
 
 if daQuaderno="" then 
  CodiceAllievo=Request.QueryString("cod")
  cla=Request.QueryString("cla")
 
  Codice_Test=Request.QueryString("CodiceTest")

  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Quesito=Request.QueryString("Quesito")
  Tipodomanda=Request.QueryString("Tipodomanda")
  R1=Request.QueryString("R1")
  R2=Request.QueryString("R2")
  R3=Request.QueryString("R3")
  R4=Request.QueryString("R4")
  RE=Request.QueryString("RE")
  MO=Request.QueryString("MO")
  DATA=Request.QueryString("DATA")
  ORA=Request.QueryString("ORA")
  VAL=Request.QueryString("VAL")
  VALINQUIZ=Request.QueryString("VALINQUIZ")
 ' response.write("VALINQUIZ"&VALINQUIZ)
  URL=Request.QueryString("URL")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  ID=CodiceDomanda 
  Cartella=Request.QueryString("Cartella")
  INQUIZ=Request.QueryString("INQUIZ")
  Segnalata=Request.QueryString("Segnalata")
  Multiple=Request.QueryString("Multiple")  
  VF=Request.QueryString("VF")  
else 'eseguo query per caricare i parametri 

   QuerySQL="Select * from MODULO_PARAGRAFO_DOMANDE1 where CodiceDomanda=" & CodiceDomanda & ";" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	CodiceAllievo=rsTabella("CodiceAllievo")
  response.write(QuerySQL)
  Cognome=rsTabella("Cognome")
  Nome=rsTabella("Nome")
  Codice_Test=rsTabella("ID_Paragrafo")

  Capitolo=rsTabella("Tit")
  Paragrafo=rsTabella("Titolo")
  Modulo=rsTabella("ID_Mod")
  Quesito=rsTabella("Quesito")
  Tipodomanda=rsTabella("Tipo")
  R1=rsTabella("Risposta1")
  R2=rsTabella("Risposta2")
  R3=rsTabella("Risposta3")
  R4=rsTabella("Risposta4")
  
  RE=rsTabella("RispostaEsatta") 
  MO=rsTabella("ID_Mod")
  DATA=rsTabella("DATA")
  VAL=rsTabella("Voto")
  VALINQUIZ=rsTabella("In_Quiz")
 ' response.write("VALINQUIZ"&VALINQUIZ)
 ' URL=Request.QueryString("URL")
 ' Nome=Request.QueryString("Nome")
 ' Cognome=Request.QueryString("Cognome")
 
  Cartella=rsTabella("Cartella")
  INQUIZ=VALINQUIZ
  Segnalata=rsTabella("Segnalata")
  Multiple=rsTabella("Multiple")
  VF=rsTabella("VF")

end if

'Quesito="ciao"&ReplaceCar(Quesito)
 ID=CodiceDomanda 

if MO<>"" then 
 Modulo=MO
end if  
  
' codice per permettere la visualizzazione solo delle proprie domande
	QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") & "';" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	Privato=rsTabella.fields("Privato") 
	rsTabella.close
	
	
if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine del file con 'il response.redirect
'homesito="/anno_2010-2011_ITC/ECDL"   
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
        url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url3=Replace(url,"\","/")
url=url3
 
' Open file for reading.
 

'response.write(url)
on error resume next
Set objTextFile = objFSO.OpenTextFile(url, ForReading)


' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll

objTextFile.Close

'GESTIONE RRORE
If Err.Number <> 0 then
  NumeroErrore = Err.Number
  DescrizioneErrore = Err.Description
  Pagina = Request.ServerVariables("url")
  Spiegazione1="Errore, file della spiegazione mancante:"& url
  Riga=227
  response.write(Spiegazione1)
  sReadAll=Spiegazione1
''  Source=Err.Source
'  Call GestisciErrore(DescrizioneErrore,Spiegazione1,Pagina,Riga)
  Err.Number=0
End If

 Function domandaplus()	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	 url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	objTextFile.Close
End Function %>

    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserisci valutazione </h1> 
                    
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
							<a href="javascript:history.back();">Quaderno</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Domanda</a>
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
				        <h3> <i class="icon-reorder"></i>  <%Response.write (Capitolo) %> : <%Response.write (Paragrafo)%> </h3>
			          </div>
				      <div class="box-content">
                     
           <h4> <small><%=Cognome%> &nbsp;<%=left(Nome,1)&"."%></small>  </h4>
             
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			   <div class="box-content"> 
                     
                     
                       <form method="POST"  class="form-vertical" form action="inserisci_valutazione1.asp?VECCHIOVAL=<%=VAL%>&VF=<%=VF%>&Multiple=<%=Multiple%>&Voto=<%=VAL%>&Tipodomanda=<%=Tipodomanda%>&VALORE=<%=VAL%>&Cartella=<%=Cartella%>&id_classe=<%=id_classe%>&cla=<%=cla%>&cod=<%=CodiceAllievo%>&CodiceDomanda=<%=CodiceDomanda%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&DATA=<%=DATA%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&MO=<%=MO%>&INQUIZ=<%=INQUIZ%>&tCap=<%=tCap%>&tSot=<%=tSot%><%=p%>&tDom=<%=tDom%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 

     
	 

  <p><input type="text" class="input-xxlarge" name="txtDomanda"  value="<%=Quesito%>" size="100" maxlength="250"><b>Domanda <br>
	 
	<%if Tipodomanda=1 then %>
	   <br>
	   <textarea class="input-block-level" rows="6" name="TestoDomandaPlus" value="ciao" cols="100"><%=Response.write(domandaplus())%> </textarea><br>
		
	<% end if%>
	  <%if VF=0 then ' non è una domanda vero falso %>
          <p><input type="text" class="input-xxlarge" name="txtR1" value="<%=R1%>" size="100" maxlength="150"><b> 
            Risposta 1</b></p> 
          <p>
            <input type="text" class="input-xxlarge" name="txtR2" value="<%=R2%>" size="100" maxlength="150"><b> 
            Risposta 2 </b></p>
          <p>
            <input type="text" class="input-xxlarge" name="txtR3" value="<%=R3%>" size="100" maxlength="150"><b> 
            Risposta 3 </b></p>
          <p><input type="text" class="input-xxlarge" name="txtR4" value="<%=R4%>" size="100" maxlength="150"><b> 
            Risposta 4 </b></p>
              <p><input type="text" class="input-mini" name="txtRE" value="<%=RE%>" size="2"><b> 
            Risposta Esatta  </b></p>
    
         
     <%else ' è vero falso%>
            <% if (RE=1)  then  %>
                                            Risposta Esatta<br>
											 <INPUT TYPE="RADIO" name="txtRE" checked="true" value="1">Vero 
                                             <INPUT TYPE="RADIO" name="txtRE" value="0">Falso 	          
                                            <% else %>
                                             <INPUT TYPE="RADIO" name="txtRE" value="1">Vero   
                                             <INPUT TYPE="RADIO" name="txtRE"   checked="true" value="0">Falso  
                                           
										<% end if %>
           
             
	 <%end if%>
     <br><br>
   
	<b>Spiegazione</b><p><textarea class="input-block-level" rows="6" name="S1" value="ciao" cols="70"> <%=Response.write(sReadAll)%> </textarea></p>
<p><input class="input-small"  type="text" name="txtDATA" value="<%=DATA%>" size="8"><b> 
	Data </b> <input class="input-small"  type="text" name="txtOra" value="<%=ORA%>" size="8"><b> 
	Ora </b></p>
 <%if (session("Admin")=true) then %> 
              <p><input type="text" class="input-mini"  name="txtVAl" value="<%=VAL%>" size="1"><b> Valutazione </b> &nbsp;&nbsp;
              <span title="Feedback all'autore"><b>Segnalata</b></span> 											 
			 <% if (Segnalata=1)  then  %>          
                <INPUT TYPE="RADIO" name="txtSegnalata" checked="true" value="1">Si  
                <INPUT TYPE="RADIO" name="txtSegnalata"  value="0">No  	          
            <% else %>
               <INPUT TYPE="RADIO" name="txtSegnalata" value="1">Si  
               <INPUT TYPE="RADIO" name="txtSegnalata"   checked="true" value="0">No             
           <% end if %> 
	
 		<p> <input type="text" class="input-mini"  name="txtINQUIZ" value="<%=INQUIZ%>" size="1"><b> In Quiz </b> &nbsp;&nbsp;&nbsp;	<b>Codice Domanda (<%=CodiceDomanda%>)</b>  </p>
     
 		 <p><input type="submit" class="btn"  value="Invia" name="B1"> </p> <!--Definisce i due bottoni del form -->
<% else 
	   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or  (Privato=0) then %>
	  
		   <input type="text" class="input-mini"   name="txtSegnalata" value="<%=Segnalata%>" size="1"><b> Segnalata </b></p>
	   <p>
		 <input type="text" class="input-mini"    name="txtINQUIZ" value="<%=INQUIZ%>" size="1"><b> In Quiz </b>
		</p>
	   <p><input type="submit" class="btn" value="Invia" name="B1"> </p> <!--Definisce i due bottoni del form -->
	   <br><hr>
	  <% end if 
end if %>

<p><a target="_new" href="7_stampa_scheda_domanda.asp?domande=1&CodiceAllievo=<%=CodiceAllievo%>&CodiceDomanda=<%=CodiceDomanda%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>"><img src="../../img/printer.jpg" alt="Stampa questa scheda"></a></p>
   
</form> <!-- Chiude l'interfaccia -->

                         
              
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
    <% else %>
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
   'Response.Redirect "studente_domande.asp?cla="&cla
   end if %>

 </html>

