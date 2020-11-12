<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
 <%  
 
esporta=Request.QueryString("esporta")
 
  Cartella=Request.QueryString("Cartella")
	  Codice_Test=Request.QueryString("CodiceTest")
	  Modulo=Request.QueryString("Modulo")
	  Paragrafo=Request.QueryString("Paragrafo")
	  CodiceAllievo=Request.QueryString("CodiceAllievo")
	 Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
 <title>Riepilogo Topolino classe</title>
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 <title>Riepilogo Navigazione classe</title>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Riepilogo Client/Server classe</title>
							<%End Select%>
  
   
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

  


   
</head>

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
    
    
    
    
	<div class="container-fluid" id="content"  >
    
      <!-- #include file = "../include/menu_left.asp" -->
     
          
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Riepilogo metafore classe</h1> 
                      <% if session("DB")=1 then
					  scegli=3 
					  %>
                        <a title="Condividi link alla pagina" href="#" onClick="javascript:PopUpWindow(600,400,<%=scegli%>);return false;"><i class="glyphicon-share_alt"> </i> <small>Condividi</small> </a>  
                      <% end if%>
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
							<a href="#">Libro U</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Metafore </a>
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
				        <h3> <i class="icon-reorder"></i>  
						 <%Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
 									<a href="http://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Topolino_nel_Labirinto.html" target="_blank">Presentazione</a>
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 									<a href="http://www.umanetexpo.net/informistica/UWWW/Navigazione/Navigazione.html" target="_blank">Presentazione</a>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 								<a href="http://www.umanetexpo.net/informistica/UWWW/Metafore/Pagine/Terremoto_Culturale.html" target="_blank">Presentazione</a>
							<%End Select%> percorso 
						</h3>
			          </div>
				      <div class="box-content">
                      
 <% 
function ReplaceCar(sInput)
dim sAns
  
     sAns= Replace(sInput, Chr(34), Chr(96))
   sAns=  Replace(sAns,Chr(39),Chr(96))
  
ReplaceCar = sAns
end function

  
	
	 Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
  	sSQL="SELECT * " &_
" FROM Elenco_Metafore_Topolino " &_
" WHERE Id_Paragrafo='" & Cartella&"_U_2_3" & "' and Pi=0 "&_ 
 " and (Data>= CONVERT(DATETIME,'" &Session("DataClaq")  &"', 104))" &_
 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(Session("DataClaq2"))) &"', 104))"&_
 "order By Cognome, Nome;"   
' response.write(sSQL&"<br>")
						
  Case Cartella&"_U_2_5" 'Navigazione
 	sSQL="SELECT * " &_
" FROM Elenco_Metafore_Navigazione " &_
" WHERE Id_Paragrafo='" & Cartella&"_U_2_5" & "'  and Pi=0"&_ 
 " and (Data>= CONVERT(DATETIME,'" &Session("DataClaq")  &"', 104))" &_
 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(Session("DataClaq2"))) &"', 104))"&_
 "order By Cognome, Nome;"   
  Case Cartella&"_U_2_8" 'ClientServer
 End Select

'response.write(sSQL&"<br>")
	

 Set oRs = ConnessioneDB.Execute(sSQL)
 if oRs.eof then
		oRs.close 
		set oRs = Nothing
		set oCmd = nothing
		MessageChildren = ""
 end if
 
 



 Function MessageChildren(ID)
	dim oRs,oRs1, oCmd, sSQL, sAns		
	 Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
  	sSQL="SELECT * " &_
" FROM Elenco_Metafore_Topolino " &_
" WHERE CodiceMetafora=" & ID & ";"  
' response.write(sSQL&"<br>")
						
  Case Cartella&"_U_2_5" 'Navigazione
 	sSQL="SELECT * " &_
" FROM Elenco_Metafore_Navigazione " &_
" WHERE CodiceMetafora=" & ID & ";"  
' 
  Case Cartella&"_U_2_8" 'ClientServer
  End Select
'response.write(sSQL&"<br>")
 Set oRs = ConnessioneDB.Execute(sSQL)
 if oRs.eof then
		oRs.close 
		set oRs = Nothing
		set oCmd = nothing
		MessageChildren = ""
		exit function
 end if
 voti = oRs("Voto") 
 voti = voti &"+"& MessageChildren(oRs("Pf")) 		
 oRs.Close
 set oRs = nothing
 set oCmd = nothing
 MessageChildren = voti
end function%>


<%
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")
consegnato=""
do while not oRs.eof
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&oRs("Tit")&"_"&oRs("CodiceMetafora")&".txt"
url2=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&oRs("Tit")&"_simula_"&oRs("CodiceMetafora")&".txt"
url=Replace(url,"\","/")
url=Replace(url,"%20"," ")
url2=Replace(url2,"\","/")
url2=Replace(url2,"%20"," ")
'url =Server.URLEncode(url)
 If objFSO.FileExists(url) then
       ' Response.Write "Il file esiste"
	   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	   sReadAll = objTextFile.ReadAll
	   objTextFile.Close
    Else
       ' Response.Write "Il file NON esiste"
	   sReadAll="File spiegazione mancante" 
	   'sReadAll=url
    End If
	 If objFSO.FileExists(url2) then
       ' Response.Write "Il file esiste"
	   Set objTextFile = objFSO.OpenTextFile(url2, ForReading)
	   sReadAll2 = objTextFile.ReadAll
	   objTextFile.Close
    Else
       ' Response.Write "Il file NON esiste"
	   sReadAll2="File simulazione mancante" 
	   'sReadAll=url
    End If

msg="Vuoi veramente cancellare la metafora?"
if session("Admin")=true  then                                    
link_elimina="<a  onClick='return window.confirm("&msg&");'  target='_blank'  href='cancella_metafora.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&ThreadParent="&oRs("ThreadParent")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><img src='../../img/elimina_small.jpg'></a>"
else
link_elimina=""
end if 
'response.write("<br>"&sReadAll)   ThreadParent
link="<a target='_blank'  href='inserisci_valutazione_metafore.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&ThreadParent="&oRs("ThreadParent")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><i title='Esegui' class='icon-play-circle'></i></a>"
linklike="<i class='glyphicon-thumbs_up'></i>&nbsp;&nbsp;&nbsp;&nbsp;<br>"
punti=MessageChildren(oRs("CodiceMetafora"))
totpunti= eval(left(punti,len(punti)-1))

 Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%> 
																
<%	'narrazione=raccontaTopolino(oRs("CodiceMetafora"))		
	sAns = sAns &"<tr><td>"& totpunti&"<br> "&oRs("Cognome")& " " &left(oRs("Nome"),1)&".</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span><center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Topolino")&"</td>" &"<td>"& oRs("Formaggio")&"</td>" &"<td  class='hidden-480'>"& oRs("Fame")&"</td>" &"<td  class='hidden-480'>"& oRs("Labirinto")&"</td>"&"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: rgba(255, 0, 0, 0.5);'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: rgba(0, 255, 0, 0.5);'><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Testata")&"</td><td>"&link_elimina&"</td></tr>"%>

 		
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 <%	'narrazione=raccontaNavigazione(oRs("CodiceMetafora"))		
	sAns = sAns &"<tr><td>"& totpunti &"<br> "& oRs("Cognome")& " "& left(oRs("Nome"),1)&".</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span   data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span> <center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: rgba(255, 0, 0, 0.5);'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: rgba(0, 255, 0, 0.5);'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"
	
		    volere = "vuoi"
            raggiungere = "raggiungerai"
            avere = "hai"
            scegliere = "scegli"
            avvicinarsi = "ti avvicina"
            avvicinarsi1 = "avvicinarti"
            avvicinarsi2 = "avvicinarsi"
            allontanarsi = "ti allontana"
            allontanarsi1 = "ti sei allontanato troppo hai"
            scontrarsi = "e ti sei scontrato"
            continuare = "continua"
            fare = "ci sei quasi fai"
            dovere = "devi"
            ti_vi = "ti"
		narrazione=""
		narrazione =  oRs("Autista") & " " & volere & "  raggiungere " & oRs("Destinazione") & " ?"
        narrazione = narrazione & "NO!   Mancando " & oRs("Carburante") & " per raggiungere " & oRs("Destinazione") & " , " & oRs("Autista") & " nel contesto " & oRs("Luogo") & " non " & raggiungere & " l'obiettivo ! "
        narrazione = narrazione & " " & oRs("Autista") & " " & volere & " raggiungere " & oRs("Destinazione") & " ?"
        narrazione = narrazione &"   " & oRs("Autista") & "  quale  " & oRs("Strada") & " " & scegliere & " ?  "
        narrazione = narrazione &"ATTENZIONE  " & oRs("Autista") & "  la scelta  " & oRs("Strada_KO") & " " & allontanarsi & " da  " &  oRs("Destinazione")&"."
        narrazione = narrazione &" " &  oRs("Cespugli") & " " & ti_vi & " segnalano il pericolo ! "
        narrazione = narrazione &" :-(  " & oRs("Autista") & "  " & allontanarsi1 & " scelto la strada chiusa  " &  oRs("Strada_KO") & " " & scontrarsi & " con " &  oRs("Lupo") & ". "
        narrazione = narrazione &"  " & oRs("Autista") & "  per risolvere la situazione " & dovere & "  abbandonare  " & ors("Cestino") & " cosi' da " & avvicinarsi1 & " a " &  oRs("Destinazione") & ".  "
        narrazione = narrazione &"  " & oRs("Autista") & "  quale  " & oRs("Strada") & " " & scegliere & " ?  "
        narrazione = narrazione &"  " & oRs("Autista") & "  la scelta  " & oRs("Strada_OK") & " " & avvicinarsi & " a  " &  oRs("Destinazione") & "  " & continuare & " così !  "
        narrazione = narrazione &" Coraggio " & fare & " l'ultimo passo ! '"
        narrazione = narrazione &" :-) COMPLIMENTI  " & oRs("Autista") & " " & avere & " raggiunto " &  oRs("Destinazione") & "!!!"
		Set objFSO2 = CreateObject("Scripting.FileSystemObject")
	Set objCreatedFile2 = objFSO2.CreateTextFile(url2, True)
    objCreatedFile2.WriteLine(narrazione)
	objCreatedFile2.Close

	 %>
 
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Riepilogo Client/Server</title>
							<%End Select%>

 
 <% 

   consegnato=consegnato&"'"&oRs("CodiceAllievo")&"'"&","
   oRs.movenext	
  loop	
 oRs.Close
 consegnato=left(consegnato,len(consegnato)-1)
 set oRs = nothing
 set oCmd = nothing
 
 %>
 						 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                    
 
		    <div class="box-content"> 
                     
              <%
 
 
   
                  
response.write "<table class='table table-hover table-nomargin'>"
 Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
  
  response.write "<thead><th class='hidden-480'></th><th class='hidden-480'>Metafora</th><th class='hidden-480'>Topolino</th><th>Formaggio</th><th  class='hidden-480'>Fame</th><th class='hidden-480'>Labirinto</th> <th class='hidden-480'>Comportamento</th><th>KO</th>"
response.write "<th>OK</th><th >Testata</th><th >Elimina</th></tr></thead><thead><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale</th><th class='hidden-480'>Soggetto</th><th>Obiettivo</th><th  class='hidden-480'>Motivazione</th><th class='hidden-480'>Ambiente</th> <th class='hidden-480'>Comportamento</th><th>KO</th>"
response.write "<th>OK</th><th >Conseguenze</th><th >Elimina</th></tr></thead><tbody>"								
  Case Cartella&"_U_2_5" 'Navigazione
  response.write "<thead><th class='hidden-480'></th><th class='hidden-480'>Metafora</th><th class='hidden-480'>Autista</th><th>Destinazione</th><th  class='hidden-480'>Carburante</th><th  class='hidden-480'>Luogo</th><th class='hidden-480'>Strada</th> <th>Viziosa</th>"
response.write "<th>Virtuosa</th><th >Pericoli</th><th >Cestino</th><th>Lupo</th><th>&nbsp;</th></tr></thead><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale</th><th class='hidden-480'>Soggetto</th><th>Obiettivo</th><th  class='hidden-480'>Motivazione</th><th  class='hidden-480'>Ambiente</th><th class='hidden-480'>Comportamento</th> <th>KO</th>"
response.write "<th>OK</th><th >FeedBack (-)</th><th >Attaccamenti</th><th>Conseguenze</th><th>Elimina</th></tr></thead><tbody>"
  Case Cartella&"_U_2_8" 'ClientServer
 End Select



                  
 response.write sAns
 response.write "</tbody></table>"

 if consegnato="" then
	response.write("<br>Nessuna consegna")
else%>
	
<%	q="select count(*) as NC from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") ;"
 ' response.write(q)
  set rsNC= ConnessioneDB.execute(q)
  noconsegne=rsNC("NC")
 
 %>
<hr> <b>Mancata consegna: (<%=noconsegne%>)</b>
 <% q="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") order by Cognome;"
  'response.write(q)
  response.write("<br>")
  set rsTabellaNC= ConnessioneDB.execute(q)
  do while not rsTabellaNC.eof
  response.write(rsTabellaNC("Cognome")&" "&left(rsTabellaNC("Nome"),1)&".<br>")
  rsTabellaNC.movenext
  loop
end if

 %>			
              
              
              
              
                      
                      
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
<script>
function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
 
window.open('../cSocial/share.asp?scegli='+s,'share.asp?scegli='+s, 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);

}

 
    var contSi, contNo, Autista, Destinazione, Carburante, Luogo, Strada, Strada_OK, Strada_KO, Cespugli, Lupo, Cestino, Motivato, Testata, v;
    var plurale, plurale1, volere, dovere, fare, avere, raggiungere, scegliere, avvicinarsi, avvicinarsi1, avvicinarsi2, allontanarsi, allontanarsi1, scontrarsi, continuare, ti_vi;


    function raccontaNavigazione() {
        var narrazione = "";

       //leggo i valori delle variabili

        plurale = Autista.search(/ e /i); //se è presente e oppure E è >0
        plurale1 = Autista.search(","); //faccio mettere ; per indicare il prurale
        if ((plurale == -1) && (plurale1 == -1)) {
            volere = "vuoi";
            raggiungere = "raggiungerai";
            avere = "hai";
            scegliere = "scegli";
            avvicinarsi = "ti avvicina";
            avvicinarsi1 = "avvicinarti";
            avvicinarsi2 = "avvicinarsi";
            allontanarsi = "ti allontana";
            allontanarsi1 = "ti sei allontanato troppo hai";
            scontrarsi = "e ti sei scontrato";
            continuare = "continua";
            fare = "ci sei quasi fai";
            dovere = "devi";
            ti_vi = "ti";
        }
        else {
            volere = "volete";
            raggiungere = "raggiungerete";
            avere = "avete";
            scegliere = "scegliete";
            avvicinarsi = "vi avvicina";
            avvicinarsi2 = "avvicinarsi";
            avvicinarsi1 = "avvicinarvi";
            allontanarsi = "vi allontana";
            allontanarsi1 = "vi siete allontanati troppo avete";
            scontrarsi = "e vi siete scontrati";
            continuare = "continuate";
            fare = "ci siete quasi fate";
            dovere = "dovete";
            ti_vi = "vi";
        }
        narrazione = narrazione + "\n " + Autista + " " + volere + "  raggiungere " + Destinazione + " ?";
        narrazione = narrazione + "NO! <br>\n\n  Mancando " + Carburante.replace("voglia", "") + " per raggiungere " + Destinazione + " , " + Autista + " nel contesto " + Luogo + " non " + raggiungere + " l'obiettivo ! ";
        narrazione = narrazione + "<br>\n\n " + Autista + " " + volere + " raggiungere " + Destinazione + " ?";
        narrazione += "<br>\n\n   " + Autista + "  quale  " + Strada + " " + scegliere + " ?  ";
        narrazione += "<br>\n\nATTENZIONE  " + Autista + "  la scelta  " + Strada_KO + " " + allontanarsi + " da  " + Destinazione;
        narrazione += "<br>\n\n " + Cespugli + " " + ti_vi + " segnalano il pericolo ! ";
        narrazione += "<br>\n\n :-(  " + Autista + "  " + allontanarsi1 + " scelto la strada chiusa  " + Strada_KO + " " + scontrarsi + " con " + Lupo + " \n ";
        narrazione += "<br>\n\n  " + Autista + "  per risolvere la situazione " + dovere + "  abbandonare  " + Cestino + " cosi' da " + avvicinarsi1 + " a " + Destinazione + "  ";
        narrazione += "<br>\n\n  " + Autista + "  quale  " + Strada + " " + scegliere + " ?  ";
        narrazione += "<br>\n\n  " + Autista + "  la scelta  " + Strada_OK + " " + avvicinarsi + " a  " + Destinazione + "  " + continuare + " così !  ";
        narrazione += "<br>\n\n Coraggio " + fare + " l'ultimo passo ! '";
        narrazione += "<br>\n :-) COMPLIMENTI  " + Autista + " " + avere + " raggiunto " + Destinazione + "!!!";

       return narrazione;
    }

 


</script>
 </html>

