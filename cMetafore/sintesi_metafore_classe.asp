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
	  Id_Premetafora=Request.QueryString("Id_Premetafora")
	 Select Case Codice_Test%>
                              	<% Case Cartella&"_U_2_3" 'Topolino%>
 <title>Riepilogo Topolino classe</title>
								<% Case Cartella&"_U_2_5" 'Navigazione%>
 <title>Riepilogo Navigazione classe</title>
							  		<% Case Cartella&"_U_2_8" 'ClientServer%>
 <title>Riepilogo Client/Server classe</title>
							<%End Select%>
  
   
    <meta charset="utf-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

	<title>FLAT - Dynamic tables</title>

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- dataTables -->
	<link rel="stylesheet" href="../../css/plugins/datatable/TableTools.css">
	<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">
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
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.datepicker.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- dataTables -->
	<script src="../../js/plugins/datatable/jquery.dataTables.min.js"></script>
	<script src="../../js/plugins/datatable/TableTools.min.js"></script>
	<script src="../../js/plugins/datatable/ColReorderWithResize.js"></script>
	<script src="../../js/plugins/datatable/ColVis.min.js"></script>
	<script src="../../js/plugins/datatable/jquery.dataTables.columnFilter.js"></script>
	<script src="../../js/plugins/datatable/jquery.dataTables.grouping.js"></script>
	<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	<script src="../../js/demonstration.min.js"></script>

	<!--[if lte IE 9]>
		<script src="../../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->
	
<script type="text/javascript" src="../js/refresh_session.js"></script>


	<!-- Favicon -->
	<link rel="shortcut icon" href="img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="img/apple-touch-icon-precomposed.png" />

  
<style>
.table td {
  word-wrap:break-word;
}
</style>

   
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
" and Id_Premetafora="&Id_Premetafora&_
" order By Cognome, Nome;"  

 sSQL2="SELECT * " &_
" FROM preTopolino " &_
" WHERE ID_Premetafora=" & ID_Premetafora & ";"  
' " and (Data>= CONVERT(DATETIME,'" &Session("DataClaq")  &"', 104))" &_
 '" AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(Session("DataClaq2"))) &"', 104))"&_
  
' response.write(sSQL&"<br>")
						
  Case Cartella&"_U_2_5" 'Navigazione
 	sSQL="SELECT * " &_
" FROM Elenco_Metafore_Navigazione " &_
" WHERE Id_Paragrafo='" & Cartella&"_U_2_5" & "'  and Pi=0"&_ 
" and Id_Premetafora="&Id_Premetafora&_
 " order By Cognome, Nome;"   
 '" and (Data>= CONVERT(DATETIME,'" &Session("DataClaq")  &"', 104))" &_
 '" AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(Session("DataClaq2"))) &"', 104))"&_
sSQL2="SELECT * " &_
" FROM preNavigazione " &_
" WHERE ID_Premetafora=" & ID_Premetafora & ";"  
	
  Case Cartella&"_U_2_8" 'ClientServer
 End Select

'response.write(sSQL&"<br>")
	

 Set oRs = ConnessioneDB.Execute(sSQL)
 if oRs.eof then
		oRs.close 
		set oRs = Nothing
		 
 end if
Set oRs2 = ConnessioneDB.Execute(sSQL2)
 if oRs2.eof then
		oRs2.close 
		set oRs2 = Nothing
 end if
 if oRs2("Img")=1 then
 immagine=1
 else
 immagine=0
 end if
 

 Function contaLike(ID)
  Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
   
						
  Case Cartella&"_U_2_5" 'Navigazione
 	sSQL3="SELECT count(*) " &_
" FROM [VotiMetaforaNavigazione] " &_
" WHERE CodiceMetafora=" & ID & ";"  
	
' 
  Case Cartella&"_U_2_8" 'ClientServer
  
  End Select
 
 Set oRs3 = ConnessioneDB.Execute(sSQL3)
 if oRs3.eof then
		oRs3.close 
		set oRs3 = Nothingg
		contaLike=0
 end if
 contaLike=oRs3(0)
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
  	sSQL2="SELECT * " &_
" FROM preDesideri " &_
" WHERE ID_Premetafora=" & ID_Premetafora & ";"  
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
finaliste=0
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
'link="<a target='_blank'  href='inserisci_valutazione_metafore.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&ThreadParent="&oRs("ThreadParent")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><i title='Esegui' class='icon-play-circle'></i></a>"
link="<a target='_blank'  href='sintesi_metafore.asp?id_classe="&id_classe&"&Cartella="&oRs("Cartella")&"&classe="&classe&"&CodiceTest="&oRs("ID_Paragrafo")&"&CodiceMetafora="&oRs("CodiceMetafora")&"&Modulo="&oRs("ID_Mod")&"&Paragrafo="&oRs("Tit")&"&Capitolo="&oRs("Titolo")&"&CodiceAllievo="&CodiceAllievo&"'><i title='Esegui' class='icon-play-circle'></i></a>"
autore=oRs("Cognome")& " " &left(oRs("Nome"),1)&"."

%>
<%' guardo chi ha votato per la metafora
    Select Case Codice_Test 
	  
	  Case Cartella&"_U_2_3" 'Topolino 
		 
			
	  Case Cartella&"_U_2_5" 'Navigazione 
	 	 QuerySQL2="select  Cognome,Nome from [VotiMetaforaNavigazione] WHERE   CodiceMetafora="&oRs("CodiceMetafora")&";"
                                  
	 
	
	  Case Cartella&"_U_2_8" 'ClientServer 
	 

	 End Select 
votanti=""
set rsVotanti=ConnessioneDB.Execute(QuerySQL2)
do while not rsVotanti.eof
	votanti=votanti+ rsVotanti("Cognome")& " " &left(rsVotanti("Nome"),1)&". "
rsVotanti.movenext
loop

if session("Admin")=true then
mostrautore="Vota per "&autore &" hanno già votato : "&votanti
else
mostrautore="Vota"
end if
 
 solovotabili=request.QueryString("sv") ' 1 se devo far compariore solo quelle che hanno 3 livelli di profondità
if immagine=1 and solovotabili<>"" then
linklike=" <A href='#modal-1' data-toggle='modal' onClick=""vota_post("&oRs("CodiceMetafora")&",1,'"&oRs("ID_Paragrafo")&"',"&oRs("ID_Premetafora")&");"" ><i title='"&mostrautore&"' class='glyphicon-thumbs_up'></i></a>&nbsp;&nbsp;&nbsp;&nbsp;<br>"
else
linklike=""
end if
punti=MessageChildren(oRs("CodiceMetafora"))
numlike=contaLike(oRs("CodiceMetafora"))
totpunti= eval(left(punti,len(punti)-1))

if solovotabili<>"" then
 
	if totpunti=3 then ' se 
	Select Case Codice_Test%>
									<% Case Cartella&"_U_2_3" 'Topolino%> 
																	
	<%	'narrazione=raccontaTopolino(oRs("CodiceMetafora"))		
		sAns = sAns &"<tr><a onclick=""alert('ciao')""><td id=like_"&numrecord&">"&numlike&"</td></a><td>"& totpunti&"<br> "&oRs("Cognome")& " " &left(oRs("Nome"),1)&".</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'><i class='icon-bullhorn'></i></span><center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Topolino")&"</td>" &"<td>"& oRs("Formaggio")&"</td>" &"<td  class='hidden-480'>"& oRs("Fame")&"</td>" &"<td  class='hidden-480'>"& oRs("Labirinto")&"</td>"&"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #e9967a;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #2e8b57;'><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Testata")&"</td><td>"&link_elimina&"</td></tr>"%>

			
									<% Case Cartella&"_U_2_5" 'Navigazione%>
	<%	'narrazione=raccontaNavigazione(oRs("CodiceMetafora"))		
		'sAns = sAns &"<tr><td>"&numlike&"</td><td>"& oRs("Cognome") &" "& left(oRs("Nome"),1)&". "& totpunti& "</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span   data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span> <center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #e9967a;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #2e8b57;'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"%>
		
		<%'tolgo autore
		sAns = sAns &"<tr><span title=""pippo""><td id=like_"&oRs("CodiceMetafora")&">"&numlike&"</td></span><td>"& totpunti& "</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span   data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'><i class='icon-bullhorn'></i></span> <center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #e9967a;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #2e8b57;'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"%>
		
	
										<% Case Cartella&"_U_2_8" 'ClientServer%>
	<title>Riepilogo Client/Server</title>
								<%End Select%>
	<% 
	end if
else

	Select Case Codice_Test%>
									<% Case Cartella&"_U_2_3" 'Topolino%> 
																	
	<%	'narrazione=raccontaTopolino(oRs("CodiceMetafora"))		
		sAns = sAns &"<tr><td id=like_"&numrecord&">"&numlike&"</td><td>"& totpunti&"<br> "&oRs("Cognome")& " " &left(oRs("Nome"),1)&".</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span><center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Topolino")&"</td>" &"<td>"& oRs("Formaggio")&"</td>" &"<td  class='hidden-480'>"& oRs("Fame")&"</td>" &"<td  class='hidden-480'>"& oRs("Labirinto")&"</td>"&"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #ffc0cb;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #98fb98;'><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Testata")&"</td><td>"&link_elimina&"</td></tr>"%>

			
									<% Case Cartella&"_U_2_5" 'Navigazione%>
	<%	'narrazione=raccontaNavigazione(oRs("CodiceMetafora"))		
	'	sAns = sAns &"<tr><td>"& oRs("Cognome") &" "& left(oRs("Nome"),1)&". "& totpunti&"+"&numlike& "</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span   data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span> <center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #e9967a;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #2e8b57;'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"%>
		<%sAns = sAns &"<tr><td title='Piace a ..' id=like_"&oRs("CodiceMetafora")&">"&numlike&"</td><td>"& totpunti& "</td><td class='hidden-480'><span data-original-title='Spiegazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll)&"'> <i class='icon-question-sign'></i></span><span   data-original-title='Simulazione'  class='btn' rel='popover' data-trigger='hover' title='' data-placement='bottom' data-content='"&ReplaceCar(sReadAll2)&"'> <i class='icon-bullhorn'></i></span> <center>"&linklike&link&" &nbsp;&nbsp;&nbsp;&nbsp;</center></td><td  class='hidden-480'>"& oRs("Autista")&"</td>" &"<td>"& oRs("Destinazione")&"</td>" &"<td  class='hidden-480'>"& oRs("Carburante")&"</td>" &"<td  class='hidden-480'>"& oRs("Luogo")&"</td>"  &"<td  class='hidden-480'>"& oRs("Strada")&"</td>" &"<td style='background-color: #ffc0cb;'><b>"& oRs("Strada_KO")&"</b></td>" &"<td style='background-color: #98fb98;'U><b>"& oRs("Strada_OK")&"</b><td>"& oRs("Cespugli")&"</td><td>"& oRs("Cestino")&"</td><td>"&oRs("Lupo")&"</td><td>"& link_elimina&"</td></tr>"%>
		
	
										<% Case Cartella&"_U_2_8" 'ClientServer%>
	<title>Riepilogo Client/Server</title>
								<%End Select%>

<%end if


   consegnato=consegnato&"'"&oRs("CodiceAllievo")&"'"&","
   oRs.movenext	
   if solovotabili<>"" then
		finaliste=finaliste+1
  end if
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
 
 
   
                  
'response.write "<table class='table table-hover table-nomargin'>"
response.write "<table class='table table-hover table-nomargin table-bordered dataTable dataTable-noheader dataTable-nofooter'>"

' response.write "<table class='table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped'>"
                                   
 Select Case Codice_Test
  Case Cartella&"_U_2_3" 'Topolino 
  
  response.write "<thead><tr><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale</th><th class='hidden-480'>Soggetto <br> Topolino</th><th>Obiettivo Formaggio</th><th  class='hidden-480'>Motivazione Fame</th><th class='hidden-480'>Ambiente Labirinto</th> <th class='hidden-480'>Comportamento Strada</th><th>KO Viziosa</th>"
response.write "<th>OK Virtuosa</th><th >Conseguenze - Testata</th><th >Elimina</th></tr></thead><tbody>"								
  Case Cartella&"_U_2_5" 'Navigazione
  response.write "<thead><tr><th class='hidden-480'><i title='Numero di like' class='glyphicon-thumbs_up'></i></th><th class='hidden-480'>Punti</th><th class='hidden-480'>Morale<br> Metafora</th><th class='hidden-480'>Soggetto<br> Autista</th><th>Obiettivo<br> Destinazione</th><th  class='hidden-480'>Motivazione<br> Carburante</th><th  class='hidden-480'>Ambiente<br> Contesto</th><th class='hidden-480'>Comportamento<br> Strada</th> <th >KO<br> Viziosa</th>"
response.write "<th>OK <br> Virtuosa</th><th >FeedBack (-)<br> Pericoli </th><th >Attaccamenti<br> Cestino </th><th>Conseguenze<br> Lupo</th><th>Elimina</th></tr></thead><tbody>"
  Case Cartella&"_U_2_8" 'ClientServer  
 End Select



                  
 response.write sAns
 response.write "</tbody></table>"

if session("Admin") then
 if consegnato="" then
	response.write("<br>Nessuna consegna")
else%>
	
<%	q="select count(*) as NC from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") ;"
 ' response.write(q)
  set rsNC= ConnessioneDB.execute(q)
  noconsegne=rsNC("NC")
 
 if solovotabili<>"" then
 %><!--
 <hr> <b>Finaliste : (<%'=finaliste%>)</b>-->
 <% else %>
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
end if

end if

 %>			
              
       <form id="mod" action="modifica_categoria.asp" method="post">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Ti piace!</h3>
				</div>
				<div class="modal-body">
					<b>Perchè? </b><br> 
					
					<textarea class="input-block-level" rows="2" cols="40" placeholder="Scrivi un commento positivo" id="titolomodifica" name="titolomodifica"></textarea> 

				</div>
				<div class="modal-footer">
					<button class="btn" data-dismiss="modal" id="chiudimodal" aria-hidden="true">Chiudi</button>
					<button type="button" id="inviamodifica" class="btn btn-primary" onClick="controllamodifica()">Invia</button>
				</div>
			</div>
		</form>
       
             
              
                      
                      <center><h5> (<a href="<%=Request.ServerVariables("PATH_INFO")&"?"&Request.ServerVariables("QUERY_STRING")&"&sv=1"%>" target="_blank">-> <i title="Vai alla votazione metafore" class="glyphicon-thumbs_up"></i> <-</a>)</h5></center>
               <br>
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

 

 // post=metafora
 var paramglobal;
 var metaforaglobal;
function vota_post(metafora,segno,codicetest,idpremetafora) {
			var cartella="<%=cartella%>";
			var stella;
			if (segno==1)
			  stella="(+)";
			else 
			  stella="(-)";
			
//	if (window.confirm('Vuoi assegnare una stella '+ stella+' alla metafora?')) {
		 
	  	 	 paramglobal="codicemetafora="+metafora+"&Cartella="+cartella+"&segno="+segno+"&CodiceTest="+codicetest+"&ID_Premetafora="+idpremetafora;
			   
			   metaforaglobal=metafora;
}
function rimpiazza(testo){
	var pulito = new String(testo);
		pulito = pulito.replace(/&agrave;/g,"à");
		pulito = pulito.replace(/&ograve;/g,"ò");
		pulito = pulito.replace(/&ugrave;/g,"ù");
		pulito = pulito.replace(/&egrave;/g,"è");
		pulito = pulito.replace(/&igrave;/g,"ì");
		pulito = pulito.replace(/&nbsp;/g," ");
		//pulito = pulito.replace(/&/g,"e");
		pulito = pulito.replace(/&#39;/g,"`");
		
		pulito = pulito.replace("'","`");

		return pulito;
}


function controllamodifica(){
			var commento = document.getElementById("titolomodifica").value.trim();
			if(commento == ""){
				alert("Inserisci un commento alla votazione");
			}else
			{
				//document.getElementById("inviamodifica").type="submit";
			   var xhttp = new XMLHttpRequest();
			   
			    var url="8_vota_metafora_ajax.asp";
			   
				var metafora=metaforaglobal;

			//	alert("prima="+commento);
				commento= rimpiazza(commento);
				commento=encodeURIComponent(commento);
			//	alert("dopo="+commento);
				params=paramglobal+"&commento="+commento;
				
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
							var risposta=xhttp.responseText;
							var jsonrisp=JSON.parse(risposta);
								if (jsonrisp["stato"]=="1")
								{
									var res = jsonrisp["msg"].split("-");
									alert("Preferenza assegnata, ne hai ancora "+ res[1]);
									 
								  document.getElementById('like_'+metafora).innerHTML=res[0];
								 
								  // document.getElementById("post_"+post).style.display="none";
								//	$('#riga_'+riga).remove();
								  //alert("Eliminato");
								}
								else{
									$('#chiudimodal').click();
									alert(jsonrisp["msg"]);
									
								}
					}
			   };

		
		xhttp.open('POST', url) 
		xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')
		xhttp.send(params);

			 //  xhttp.open("GET", url, true);
			  // xhttp.send();
			}
		}




</script>
 </html>

