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
 <title>Riepilogo Topolino</title>
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 <title>Riepilogo Navigazione</title>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Riepilogo Client/Server</title>
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
    
    <!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>
    
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
						<h1> <i class="icon-comments"></i> Riepilogo metafore </h1> 
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
 
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Set objFSO = CreateObject("Scripting.FileSystemObject")

url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url2=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_simula_"&ID&".txt"
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


'response.write("<br>"&sReadAll)   ThreadParent
link="<a target='_blank'  href='inserisci_valutazione_metafore.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&ThreadParent="&oRs("ThreadParent")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><i style='align:center' title='Esegui' class='icon-play-circle'></i></a>"
'link="<a target='_blank'  href='sintesi_metafore.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><i title='Esegui' class='icon-play-circle'></i></a>"

msg=""


if (ucase(session("CodiceAllievo"))=ucase(oRs("CodiceAllievo"))) or Session("Admin") = true  then 
link_elimina="<a  onClick='return window.confirm("&msg&");'  href='cancella_metafora.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&ThreadParent="&oRs("ThreadParent")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><img src='../../img/elimina_small.jpg'></a>"
else 
  link_elimina=""
 end if%>

<%
if Session("Admin") = true then
			'compongo select
			isel = -2
			seleziona = "<select id='sel"&oRs("CodiceMetafora")&"' onchange='cambiavoto("&oRs("CodiceMetafora")&")' style='width:60px'>"
		    
			do while isel < 10

				if oRs("Voto")= (isel+1) then
					tipo = "selected"
				else
					tipo = ""
				end if

				seleziona = seleziona&"<option "&tipo&" value='"&(isel+1)&"'>"&(isel+1)&"</option>"
				isel = isel+1
			loop
			seleziona = seleziona & "</select>"
			'sAns1 = "<span data-placement='bottom'   rel='tooltip' title='Punti assegnati dal docente'>"&pt & ". "&seleziona&"</span>    "
			sAns1 = "<td>"&seleziona&"</td>    "

	else
		
			'sAns1 = "<span data-placement='bottom'   rel='tooltip' title='Punti assegnati dal docente'>"&pt & "." & oRs("Punti")&"</span>    "
			sAns1 = "<td>"& oRs("Voto")&"</td>    "



	end if

 
 

 Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%> 
																
								
 <%sAns = sAns &"<tr>"&sAns1&"<td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span>&nbsp;&nbsp;"&link&" </td><td  class='hidden-480'>"& oRs("Topolino")&"</td>" &"<td>"& oRs("Formaggio")&"</td>" &"<td  class='hidden-480'>"& oRs("Fame")&"</td>" &"<td  class='hidden-480'>"& oRs("Labirinto")&"</td>"&"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #f08080;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: rgba(0, 255, 0, 0.5);'><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Testata")&"</td><td>"&link_elimina&"</td></tr>"%>

 		
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 
 <%sAns = sAns &"<tr>"&sAns1&"<td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span>&nbsp;&nbsp;"&link&" </td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #f08080;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: rgba(0, 255, 0, 0.5);'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"%>
 	
 
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Riepilogo Client/Server</title>
							<%End Select%>

 
 <%sAns = sAns & MessageChildren(oRs("Pf")) 		
 oRs.Close
 set oRs = nothing
 set oCmd = nothing
 MessageChildren = sAns
end function%>
 						 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                    
 
		    <div class="box-content"> 
                     
              <%
 
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  'id_classe=Request.QueryString("id_classe")
  
 
   
                  
response.write "<table class='table table-hover table-nomargin'>"
 Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
  
  response.write "<thead><th class='hidden-480'></th><th class='hidden-480'>Metafora</th><th class='hidden-480'>Topolino</th><th>Formaggio</th><th  class='hidden-480'>Fame</th><th class='hidden-480'>Labirinto</th> <th class='hidden-480'>Comportamento</th><th>KO</th>"
response.write "<th>OK</th><th >Testata</th><th >Elimina</th></tr></thead><thead><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale</th><th class='hidden-480'>Soggetto</th><th>Obiettivo</th><th  class='hidden-480'>Motivazione</th><th class='hidden-480'>Ambiente</th> <th class='hidden-480'>Comportamento</th><th>KO</th>"
response.write "<th>OK</th><th >Conseguenze</th><th >Elimina</th></tr></thead><tbody>"								
  Case Cartella&"_U_2_5" 'Navigazione
  	sSQL="SELECT *  FROM VotiMetaforaNavigazione  WHERE CodiceMetafora=" & CodiceMetafora & ";"  
  response.write "<thead><th class='hidden-480'></th><th class='hidden-480'>Metafora</th><th class='hidden-480'>Autista</th><th>Destinazione</th><th  class='hidden-480'>Carburante</th><th  class='hidden-480'>Luogo</th><th class='hidden-480'>Strada</th> <th>Viziosa</th>"
response.write "<th>Virtuosa</th><th >Pericoli</th><th >Cestino</th><th>Lupo</th><th>&nbsp;</th></tr></thead><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale</th><th class='hidden-480'>Soggetto</th><th>Obiettivo</th><th  class='hidden-480'>Motivazione</th><th  class='hidden-480'>Ambiente</th><th class='hidden-480'>Comportamento</th> <th>KO</th>"
response.write "<th>OK</th><th >FeedBack (-)</th><th >Attaccamenti</th><th>Conseguenze</th><th>Elimina</th></tr></thead><tbody>"
  Case Cartella&"_U_2_8" 'ClientServer
 End Select


                  
 response.write MessageChildren(CodiceMetafora)
 response.write "</tbody></table>"





  
 Set rsCommenti = ConnessioneDB.Execute(sSQL)
 if not rsCommenti.eof then %>			
              <hr>
			  <table class='table table-hover table-nomargin'>
			 <thead><tr><th class='hidden-480'>Commento</th><th class='hidden-480'>Autore</th><tr></thead>
			 <%do while not rsCommenti.eof
			  autore=rsCommenti("Cognome")& " " &left(rsCommenti("Nome"),1)&"."
			  %>
			 <tr><td><%=rsCommenti("Commento")%></td><td><%=autore%></td></tr>
			 <%  rsCommenti.movenext
			  loop %>
			 
<%end if%>
			
              
              
              
                      
                      
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


 //function cambiavoto(id, codiceTest, cartella){
	 function cambiavoto(id){
	 
		 var cartella='<%=Cartella%>';
		 var codiceTest='<%=Codice_Test%>';
		 idpost = id;
		 voto = $("#sel"+id).val();
		// document.getElementById("load").style.display="block";

		  $.ajax({		method: "POST",
						url: "aggiornavoto_metafora.asp?id="+id+"&voto="+voto+"&codiceTest="+codiceTest+"&cartella="+cartella,
						dataType: "html",
						data: {  }
					}) /* .ajax */
					.done(function(ans) {
						alert(ans);
					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});

		 }


function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
 
window.open('../cSocial/share.asp?scegli='+s,'share.asp?scegli='+s, 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);

}
</script>
 </html>

