<% 'nuovo quaderno metafore -> tolgo iPost e report siccome puntava alle discussioni della classe e non a quelle di Umanet-WWW
%>

<!doctype html>
<html>
<head>
<link rel="shortcut icon" href="../favicon.ico" />

<script src="../js/google.js"></script><!--<meta charset="utf-8">-->
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

	<title>Quaderno UWWW</title>

	
 
	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- Easy pie  -->
	<link rel="stylesheet" href="../../css/plugins/easy-pie-chart/jquery.easy-pie-chart.css">
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
    
     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
	<!-- Validation -->
	<script src="../../js/plugins/validation/jquery.validate.min.js"></script>
	<script src="../../js/plugins/validation/additional-methods.min.js"></script>
	<!-- Sparkline -->
	<script src="../../js/plugins/sparklines/jquery.sparklines.min.js"></script>
	<!-- Easy pie -->
	<script src="../../js/plugins/easy-pie-chart/jquery.easy-pie-chart.min.js"></script>
	<!-- Flot -->
	<script src="../../js/plugins/flot/jquery.flot.min.js"></script>
	<script src="../../js/plugins/flot/jquery.flot.resize.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
	

	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />

</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">
     <% 
	
function ReplaceCar(sInput)
dim sAns
   
  sAns=  Replace(sInput,"è","&egrave;")
 
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
 
ReplaceCar = sAns
 
end function
	
	if Session("CodiceAllievo")="" or Session("Id_Classe")="" then %>	 
				<script language="javascript" type="text/javascript"> 
				    window.alert("Sessione  scaduta, effettua nuovamente il Login!");
                    location.href="../home.asp";
				</script>
				<%
				response.Redirect "../home.asp"
				 
				 %>
 
<% end if%>

<%


Response.AddHeader "Refresh", "600"

 ' Cartella=Request.QueryString("Cartella") 
 Cartella = Request.Cookies("Dati")("Cartella")
 
 ' TitoloCapitolo=Request.QueryString("Capitolo") 
 ' Paragrafo=Request.QueryString("Paragrafo")
  'Modulo=Request.QueryString("Modulo")
 ' CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")  
  dividA=request.QueryString("dividApro")
    'On Error Resume Next
xEstrazione=request.querystring("xEstrazione")
'id_classe=request.querystring("id_classe")
id_classe=Request.Cookies("Dati")("id_classe")

				'	if  id_classe="" then
'					id_classe=Session("Id_Classe")
'					end if
'classe=request.querystring("classe")
classe=Request.Cookies("Dati")("classe")
divid=request.querystring("divid")
if divid="" then divid=Session("divid")
divid2=request.querystring("divid")

PS=request.querystring("PS") ' vale 1 se devo mostrare anche i Punti Social chiamato da javasscript
if PS="" then ' per la prima chiamata mostrio i PS
   PS=1
end if
 
daStud=Request.QueryString("daStud")
daMenu=Request.QueryString("daMenu")
DataCla=request.form("txtData") 
DataCla2=request.form("txtData2")
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")
if daMenu<>"" then
    DataCla=request.QueryString("DataClaq") 
    DataCla2=request.QueryString("DataClaq2")
end if
if daStud<>"" then
 '  DataClaq= DataCla
  ' DataClaq2=DataCla2
end if


if DataCla="" then
   if DataClaq2<>"" then
      DataCla=DataClaq
	  DataCla2=DataClaq2
   else
     DataCla=session("DataCla")
	 DataClaq=session("DataCla")
	  DataClaq=session("DataClaq")
	 DataClaq2=session("DataClaq2")
	end if 
end if


''response.write(DataClaq & "<br>" & DataClaq2)
'if session("DataClaq")="" then
'Session("DataClaq")=DataClaq
'Session("DataClaq2")=DataClaq2
'else
' DataClaq=Session("DataClaq")
' DataClaq2=Session("DataClaq2")
' DataCla=Session("DataClaq")
' DataClaq=Session("DataClaq2")
' 
' end if
'' response.write("dopo session OK "& DataClaq & "<br>" & DataClaq2) 
'' se è la prima chiamata il valore del form sopra la classifica è nullo
'if (DataCla<>"") and (DataCla2<>"") then
'	Session("DataCla")=DataCla
'	Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
'else
'   Session("DataCla")= Session("DataClaq")
'   Session("DataCla2")= Session("DataClaq2")
'end if
'  
'  
  
  cod=Request.QueryString("cod")
  if strcomp(cod&"","")=0 then
     cod=Session("CodiceAllievo")
	
	 
  end if
  
 box_apri="toggleCapitolo"&request.querystring("tCap")
 box_apri1="toggleSottoPar"&request.querystring("tSot")
 box_apri2="toggleDomande"&request.querystring("tDom")
 box_apri3="toggleFrasi"&request.querystring("tFra")
 box_apri4="toggleNodi"&request.querystring("tNod")
 
  
  
  
  
function ReplaceCar(sInput)
dim sAns
 
  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
  sAns=  Replace(sAns,"'",Chr(96))
  
ReplaceCar = sAns
end function

   
  
  
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario
 
		%> 
        <!-- #include file = "../var_globali.inc" --> 
        
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->    
           
		<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->  		
		
		<!-- #include file = "../include/navigation.asp" --> 
            
        <!-- #include file = "../extra/test_server.asp" --> 
        
		<!-- #include file = "../include/formattaDataCla.inc" --> 

        <%
		
		
	QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL) 
	CIAbilitato=rsTabellaCI("CIAbilitato") 
	'ScalaValutaz=rsTabellaCI("ScalaValutaz")
	rsTabellaCI.close
    Dim esecuzione
    set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito


		
	' PRELEVO IN ANTICIPO IL CONGOME NOME NEL CASO LA QUERY 2 NON TROVI NULLA IN QUEL PERIODO E QUINDI RESTITUISCA NULL	
		  cod=Request.QueryString("cod")
		QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)

cognome = rsTabella("Cognome")
nome = rsTabella("Nome")  




		%>	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
   
      <!-- #include file = "../include/menu_left.asp" -->
			<div id="main">
				<div class="container-fluid">
					<div class="page-header">
					<!--  <div class="breadcrumbs">
                       <ul>
                            <li>
                                <a href="#">Home</a>
                                <i class="icon-angle-right"></i>
                            </li>
                            <li>
                                <a href="#">Studente</a>
                                <i class="icon-angle-right"></i>
                            </li>
                           
                        </ul>
						<div class="close-bread">
							<a href="#">
								<i class="icon-remove"></i>
							</a>
						</div>
					</div>-->
                     <div class="box">
				    
                     <div class="box-title">
                   
				      <h3> <a name="#top"><i class="icon-folder-open"></i> </a><b>Quaderno U-WWW di <%=ReplaceCar(cognome)%>  &nbsp;<%=ReplaceCar(left(nome,1)&".") %></h3> </b>
                     
					   <% if session("admin")=true then%>
					 
						<select id="studente" onchange="aggiorna_studente();">
					 <%  QuerySQL="Select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='" & Session("Id_Classe")&"' order by Cognome, Nome;"
						 Set rsTabellaStud = ConnessioneDB.Execute(QuerySQL) 
						 do while not rsTabellaStud.eof%>
						 <option value="<%=rsTabellaStud("CodiceAllievo")%>"><%=rsTabellaStud("Cognome")%>&nbsp;<%=left(rsTabellaStud("Nome"),1)&"."%></option>
						 <%rsTabellaStud.movenext
						 loop
					 
					 %>
					 
					 </select>&nbsp; Aggiorna
					 <%end if%>
			          </div><br>
                      <!-- #include file = "studente_domande_include/1_periodi.asp" --> 
<input type="button" class="btn"  style="width:60px;height:25px;" value="Invia" name="B1" onClick="aggiornaStud()"> 
 <input type="checkbox"  name="cbPS" value="1" checked="true" title="Deseleziona per escludere i Punti Social dalla classifica">  <b> 
	Includi PS
   </b>
</form>
                
                      
                      
					</div>	 
                         
					</div>
					
					<hr>
					 
					<div class="row-fluid">
						<div class="span12">
							
                            <div class="box-title">
				        <h4><a name="#"> <i class="icon-reorder"></i></a> <b> Attivit&agrave; U-WWW</b> </h4>
			          </div>
                      
                       
       <div class="bs-docs-example">
            <ul id="myTab2" class="nav nav-tabs">
                                  <li class="active"><a href="#profileMsg" data-toggle="tab" title="Messaggi dalla bacheca">Bacheca</a></li>
                                 
                                 
                                   <li>
                                   <a title="La mia bacheca personale"  href="../cSocial/default.asp?scegli=0&bacheca=<%=cod%>&nome=<%=nome%>&cognome=<%=cognome%>&id_classe=<%=id_classe%>&divid=<%=Session("divid")%>&cartella=<%=cartella%>" >
                                  Diario</a></li>
                                    
                                    
                                    <!-- <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown" title="I miei commenti nelle discussioni">iPost <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownPostLavagna" data-toggle="tab" title="I miei commenti">Lavagna</a></li>
                                      <li><a href="#dropdownPostForum" data-toggle="tab" title="I miei commenti">Forum</a></li>
                                      <li><a href="#dropdownPostDiario" data-toggle="tab" title="I miei commenti">Diario</a></li>  
                                      <li><a href="#dropdownPostChat" data-toggle="tab" title="I miei commenti">Chat</a></li> 
                                      
                                    </ul>
                                    </li>
                                    
                                     <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Report <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownQuiz" data-toggle="tab">Quiz</a></li>
                                      <li><a href="#dropdownCrediti" data-toggle="tab">Crediti</a></li>
                                      <li><a href="#dropdownCronologia" data-toggle="tab">Classifiche</a></li>  
                                      <li> <a href="../cMessaggi/centro_messaggi.asp?DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>  </li>              
                                              
                                    </ul>
                                    </li>-->
                                  <% if session("Admin")=true then %>
                                       
                                       
                                        <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">MYnd<b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownProfilo" data-toggle="tab">Profilo</a></li>
                                      <li><a href="#dropdownLogin" data-toggle="tab">Login</a></li>
                                      <li><a href="#dropdownContatti" data-toggle="tab">Contatti</a></li> 
                                      <li><a href="#dropdownEccezioni" data-toggle="tab">Eccezioni</a></li>
                                     
                                       
                                       
                                       <li> <A onClick="return window.confirm('Vuoi veramente cancellare questo account ?');" HREF="cancella_studente.asp?CodiceAllievo=<%=cod%>"><i class=" icon-trash" ></i> Rimuovi</a></li>
                                                          
                                                <!--                   
                                      <li><a href="#dropdownVisualizzazioni" data-toggle="tab">Visualizzzioni</a></li>
                                      -->
                                    </ul>
                                    </li>
                                       
                                       
                                       
                                       
                                  <!--     
                                       <li ><a href="#profileProfilo" data-toggle="tab">Profilo</a></li>   -->
                                     
                                       <%end if%> 
                                 
                            </ul>
                            
     
                            
                            <div id="myTabContent2" class="tab-content">
                             
                             
                             <% if session("Admin")=true then %>
                             
							 
							  <div class="tab-pane fade" id="dropdownLogin">
                             
                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Login
								</h4>
							</div>
                             
                             <%  cod=Request.QueryString("cod")
		QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)

cognome = rsTabella("Cognome")
nome = rsTabella("Nome")  %>
                             
								<!-- #include file = "studente_domande_include/2_modifica_login_1.asp" -->
                                
                                
							</div>
          
                              </div>
                              
                              
                              
                                <div class="tab-pane fade" id="dropdownContatti">
                             
                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Contatti
								</h4>
							</div>
                             
								  <!-- #include file = "studente_domande_include/2_modifica_contatti_1.asp" -->
                              
							</div>
                              </div>
                              
                              
                              <div class="tab-pane fade" id="dropdownProfilo">
                             
                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Profilo
								</h4>
							</div>
                             
								  <!-- #include file = "studente_domande_include/2_modifica_profilo_1.asp" -->
                              
							</div>
                              </div>
                              
                              
                                  <div class="tab-pane fade" id="dropdownEccezioni">
                             
                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Scadenze 
								</h4>
							</div>
                             
								  <!-- #include file = "studente_domande_include/2_modifica_eccezioni.asp" -->
                              
							</div>
                              </div>
                               
                              
                              	  
                              
                              
                              
							 <% end if%>
                             
                             
                              <div class="tab-pane fade" id="profileProfilo">
                            
                            
  
     			 <!----Inizio -->           
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Modifica Login
								</h3>
							</div>
							<div class="box-content nopadding">
<%QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)
cognome = rsTabella("Cognome")
nome = rsTabella("Nome") %> 
                           
                          
                           <!-- #include file = "studente_domande_include/2_modifica_login_1.asp" --> 
                            
                           
								
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->                 
 
                              </div>
                             
                             
                              
                              <div class="tab-pane fade in active" id="profileMsg">
                              
								<!-- #include file = "studente_domande_include/2_messaggi_1.asp" --> 
<%QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
						

					
					
				
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)%>
                                <div class="box box-color box-bordered">
								<div class="box-title">
									<h3>
										<i class="icon-reorder"></i>
										
                                         Messaggi alla Classe<a  href="../cSocial/default0.asp?scegli=2&id_classe=<%=rsTabella("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella("cartella")%>"> <i style="color:#FFF" title="Vai alla lavagna" class="icon-circle-arrow-right"></i></a> 
                                         
                                          
									</h3>
								</div>
                              
                               
                               
                               
                               
                               
									<ul class="timeline">
										
                                        
                                        
              <% If (rsTabellaAvvisi2.BOF=True And rsTabellaAvvisi2.EOF=True) and (rsTabellaDiario.BOF=True And rsTabellaDiario.EOF=True) and (rsTabellaForum.BOF=True And rsTabellaForum.EOF=True)  then %>
               							
                                        
                                    <table class="table table-hover table-nomargin">
									<thead>
										<tr>
											<th>Non ci sono messaggi</th> 
										</tr>
									</thead>
									</table>
        
             <%else%>
				 
           
            

        <%  k=0 
		     do while not rsTabellaDiario.EOF and k<1 
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon red">
											<i class="icon-bullhorn" title="Messaggi dal Diario"></i>
								</div>
								<div class="date"><%=left(rsTabellaDiario("DatePosted"),2) & ". " &monthname(mid(rsTabellaDiario("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                        <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=2&ID=" & rsTabellaDiario("ID") & "&RCount=" & rsTabellaDiario("ReplyCount")& "&TParent=" & rsTabellaDiario("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe  & "&categoria="&rtrim(rsTabellaDiario("Descrizione"))&"&id_categoria="&rsTabellaDiario("ID_Categoria")& "&bacheca="&rsTabellaDiario("Bacheca")& "&cognome="&rsTabellaDiario("Cognome")& "&nome="&rsTabellaDiario("Nome")& "'>"  & replaceCar(rsTabellaDiario("Topic")) & "</A>"%>
                                                   
                                      </div>													 
								</div>
							</div>
					<div class="line"></div>
				</li>          
				<%rsTabellaDiario.movenext
				loop
		   ' lavagna
		     k=0 
		     do while not rsTabellaAvvisi2.EOF and k<3 
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon blue">
											<i class="icon-bullhorn" title="Messaggi dalla bacheca"></i>
								</div>
								<div class="date"><%=left(rsTabellaAvvisi2("DatePosted"),2) & ". " &monthname(mid(rsTabellaAvvisi2("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                        <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=1&ID=" & rsTabellaAvvisi2("ID") & "&RCount=" & rsTabellaAvvisi2("ReplyCount")& "&TParent=" & rsTabellaAvvisi2("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe& "&categoria="&rtrim(rsTabellaAvvisi2("Descrizione"))&"&id_categoria="&rsTabellaAvvisi2("ID_Categoria")&  "&bacheca="&rsTabellaAvvisi2("Bacheca")& "&cognome="&rsTabellaAvvisi2("Cognome")& "&nome="&rsTabellaAvvisi2("Nome")&"'>"  & replaceCar(rsTabellaAvvisi2("Topic")) & "</A>"%>
                                                   
                                      </div>													 
								</div>
							</div>
					<div class="line"></div>
				</li>          
				<%rsTabellaAvvisi2.movenext
				loop
				
				
			 k=0 
		     do while not rsTabellaForum.EOF and k<3 
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon <%=session("stile")%>">
											<i class="icon-bullhorn" title="Messaggi dal forum"></i>
								</div>
								<div class="date"><%=left(rsTabellaForum("DatePosted"),2) & ". " &monthname(mid(rsTabellaForum("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                        <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=0&ID=" & rsTabellaForum("ID") & "&RCount=" & rsTabellaForum("ReplyCount")& "&TParent=" & rsTabellaForum("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe & "&categoria="&rtrim(rsTabellaForum("Descrizione"))&"&id_categoria="&rsTabellaForum("ID_Categoria")& "&bacheca="&rsTabellaForum("Bacheca")& "&cognome="&rsTabellaForum("Cognome")& "&nome="&rsTabellaForum("Nome")& "'>"  & replaceCar(rsTabellaForum("Topic")) & "</A>"%>
                                                   
                                      </div>													 
								</div>
							</div>
					<div class="line"></div>
				</li>          
				<%rsTabellaForum.movenext
				loop	
				
			 end if	
				%>
                                 
									</ul>
								</div>
                                
                                
                                
                                
                               
                               
                               
                               
                               
                               
     <% 
	 ' se  sono nel mio quaderno non visualizzo la casella per invio messaggio personale 
	 if strcomp(cod,Session("CodiceAllievo"))<>0 then %>                         
                               <div class="box box-color box-bordered">
								<div class="box-title">
									<h3>
										<i class="icon-reorder"></i>
										Invia messaggio personale
                                        <a href="../cMessaggi/centro_messaggi.asp?DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>" class='more-messages'>  <i class='icon-circle-arrow-right' style="color:#FFF"></i></a>
									</h3>
								</div>
     

   <div class="accordion" id="accordionMsg2">
									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle"  data-toggle="collapse" data-parent="#accordionMsg2" href="#collapseTwo">
												(+) messaggio
											</a>
										</div>
										<div id="collapseTwo" class="accordion-body collapse">
											<div class="accordion-inner">
												 
                                                  <form  class='form-horizontal' action="../cMessaggi/inserisci_messaggio_personale.asp?CodiceAllievo=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq%>&cbEmail=1" METHOD = "POST">
     
    <br>Messaggio : <br>
    
    <textarea class="input-block-level" name="txtMessaggio" ></textarea>
    <br> 
    <p> <input type="checkbox"  name="cbEmail" title="Selezionare per inviare un email allo studente">   Notifica per email &nbsp;&nbsp;&nbsp;<br>
      <p> <input class="btn" type="submit" value="Inserisci"><br> 
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>
                                                 
                                                 
											</div>
										</div>
									</div>
									
								</div>
      
								</div>
  <%else %>
  <br><hr>
   <%end if%>
                              </div>
                              
                              
                        <!-- tolto div ipost e report -->
                  
                
                   </div>        
                            
  <hr>              
  
              
   <div class="box-title">
				        <h3> <a name="#"><i class="icon-reorder"></i></a>  Compiti U-WWW<small title="Punti totalizzati"> (Pt.)</small></h3>
			          </div>
 <div class="row-fluid">
					 
					
</div>

 <div class="bs-docs-example">
    <!-- #include file = "../cUtenti/adovbs.inc" -->

 <%
 
 
 'per le store procedure
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd1.activeconnection = conn  
set cmd2.activeconnection = conn 
set cmd3.activeconnection = conn    
 
 
 QuerySQL="SELECT * FROM MODULI_CLASSE_UMANET " &_
" WHERE Id_Classe='" & id_classe & "';"
'response.write(QuerySQL)
  Set rsTabellaModuli = ConnessioneDB.Execute(QuerySQL)
 %>
 

  
 <% k=0 
 p=0
   compiti=0 ' serve per mettere il box se non ci sono compiti inseriti
		     do while not rsTabellaModuli.EOF  
			 ' calcolo i punteggi frase per quel modulo
			 %>
			  <!-- #include file = "studente_domande_include/3_statistica_metafore.asp" -->
			  <!-- #include file = "studente_domande_include/3_statistica_frasi.asp" --> 
              <!-- #include file = "studente_domande_include/3_statistica_nodi.asp" -->
              <!-- #include file = "studente_domande_include/3_statistica_domande.asp" -->
            
				<% ' numrsFrasi = numero di compiti inseriti da stud
                   ' numrsFrasi2 = punti ottenuti ; Pb =numrsFrasi2/numrsFrasi
                   ' numrsPreFrasi= compiti totali inseriti dal prof
                 %>
                 
                  <%
'			  	
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\745.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL=QuerySQLTOPO
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
			  %>
                 
 <% 
 ' se è stato svolto almeno un compito mostro il capitolo
 if (numrsFrasi<>0) or (numrsNodi<>0) or (numrsDomande<>0) or (numrsMetafore<>0)  then  ' devo fare anche per nodi e domande mostro solo dove ci sono compiti svolti
 %>
 
               <div class="accordion-group">            
                  <div class="accordion-heading">
                    <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%>" href="#collapsenew<%=k%>"  id="toggleCapitolo<%=k%>" title="<%=k%>">
                        <%=rsTabellaModuli("Titolo") %><small> (<% Response.write(numrsFrasi2+numrsNodi2+numrsDomande2+numrsMetafore2)%>)</small>
                    </a>
                    
                  </div>
                <div id="collapsenew<%=k%>" class="accordion-body collapse"> 
                    <div class="accordion-inner">
                    <table class="table table-hover table-nomargin table-condensed">
                                                    <thead>
                                                        <tr align="center">
                                                        <th>
                     <%
'on error resume next

 
						if numrsPreFrasi<>0 then
						percFrasi=fix((numrsFrasi/numrsPreFrasi)*10)/10*100
						else
						percFrasi=0
						end if
						if numrsPreDomande<>0 then
						percDomande=fix((numrsDomande/numrsPreDomande)*10)/10*100
						else
						percDomande=0
						end if
						if numrsPreNodi<>0 then
						percNodi=fix((numrsNodi/numrsPreNodi)*10)/10*100
						else
						percNodi=0
						end if
						if numrsPreMetafore<>0 then
						percMetafore=fix((numrsMetafore/numrsPreMetafore)*10)/10*100
						else
						percMetafore=0
						end if
						
						
						
						numrsDomandeBack=numrsDomande
						  
						'response.write("hkjhkj"&numrsDomande)
						
													
				            					   
       QuerySQL="SELECT * FROM MODULI_PARAGRAFI_CLASSE_UMANET " &_
" WHERE ID_Mod='" & rsTabellaModuli("ID_Mod") & "';"

  Set rsTabellaParagrafi = ConnessioneDB.Execute(QuerySQL) 

      ' servono solo per i parametri per aprire tutti i compiti del cap, forse si può anche fare a meno usando i parametri di rsTabellaModuli 	
				
				%>
                <!-- #include file = "studente_domande_include/2_nodi_0.asp" -->  
                <!-- #include file = "studente_domande_include/2_domande_0.asp" -->  
                <!-- #include file = "studente_domande_include/2_frasi_0.asp" -->  
                
             
 
                 <!-- #include file = "studente_domande_include/2_metafore_0.asp" -->  
            
      
              
                       <ul class="pagestats style-3">
                     					  
											<li>
												
                                                       
                                                <div class="spark">
													<div title="% di Frasi svolte" class="chart" data-percent="<%=percFrasi%>" data-color="#368ee0" data-trackcolor="#d5e7f7">
													
													<%=percFrasi%> %
                                                    
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaFrasi.eof then %>
                                                 <a style="color:#000" title="Apri tutte le frasi del capitolo" href="../cFrasi/2inserisci_valutazioni_frasi.asp?TutteCap=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>"> 
                                                 <%end if%>
													<span class="name"><%=numrsFrasi%> su <%=numrsPreFrasi%></span>
                                                    <span class="name">PF.<%=numrsFrasi2%> </span>
                                                      </a>
												</div>   
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Domande svolte" class="chart" data-percent="<%=percDomande%>" data-color="#56af45" data-trackcolor="#dcf8d7">
													<%=percDomande%> %
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaDomande.eof then %>
                                                  <a style="color:#000" title="Apri tutte le domande del capitolo" href="../cDomande/inserisci_valutazioni.asp?Tutte=1&ID_MOD=<%=rsTabellaDomande("ID_MOD")%>&ID_PAR=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_MOD")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&TitoloParagrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">
                                                   <%end if%>
													<span class="name"><%=numrsDomandeBack%> su <%=numrsPreDomande%></span>
                                                    <span class="name">PD.<%=numrsDomande2%> </span>
                                                    </a>
												</div>
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Nodi svolti" class="chart" data-percent="<%=percNodi%>" data-color="#f96d6d" data-trackcolor="#fae2e2"><%=percNodi%>%</div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaNodi.eof then %>
                                                 <a style="color:#000" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>"> 
													<%end if%>
                                                    <span class="name"><%=numrsNodi%> su <%=numrsPreNodi%></span>
                                                    <span class="name">PN.<%=numrsNodi2%> </span>
                                                    </a>
                                                   
												</div>
                                               
											</li>
                                            
                                             <li>

                                              <%
              
             %>
												<div class="spark">
													<div title="% di Metafore svolte" class="chart" data-percent="<%=percMetafore%>" data-color="#FC0" data-trackcolor="#F90"><%=percMetafore%>%</div>
												</div>
                        
												<div class="bottom">
                                                <%if not rsTabellaMetafore.eof then %>
                                                 <a style="color:#000" title="Apri tutte le Metafore del paragrafo"  href="../cMetafore/2inserisci_valutazioni_Metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaMetafore("ID_Mod")%>&CodiceAllievo=<%=rsTabellaMetafore("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore("Cartella")%>&Modulo=<%=rsTabellaMetafore("ID_Mod")%>&Capitolo=<%=rsTabellaMetafore("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore("TitoloParagrafo")%>"> 
													<%end if%>
                         
                                                    <span class="name"><%=numrsMetafore%> su <%=numrsPreMetafore%></span>
                                                    <span class="name">PM.<%=numrsMetafore2%> </span>
                                                    </a>
                                                   
												</div>
                                               
											</li>
                                            
                                            
                                     
										</ul>
                      </th>
                      
                                                             
                                                        </tr>
                                                    </thead>
                     </table>
                     
                     
                      <%
                      umanet=request.querystring("umanet") %>
                      
                     
                     
        
          
          
          <% 'per le store procedure
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd1.activeconnection = conn  
set cmd2.activeconnection = conn 
set cmd3.activeconnection = conn %>

   <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("Admin")=true) and (numrsFrasi<>0) then%>
								<!--<form name="dati" method="POST" target="_blank" action="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">
                           -->
                         <!--  <i class="icon-print"></i>-->
						      <img src="../../img/printer.jpg" title="Stampa frasi, domande, nodi">
                               <!--  <input type="submit" class="btn" value="Stampa Frasi Capitolo" >  -->
								 <a href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=1&tutto=1&CodiceAllievo=<%=cod%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>"" target="_blank">
								 <input type="button" class="btn" value="Stampa Frasi Capitolo" >
								 </a>
								 <% if session("admin")=true then%>
                                     <a title="Stampa paragrafi e domande" target="_blank" href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=<%=umanet%>&sint=1&tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">(Sint)</a>

                                       <a title="Stampa solo paragrafi" target="_blank" href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=<%=umanet%>&supersint=1&tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">(Super)</a>
                                   <a href="#"><input type="button" class="btn" value="Stampa Domande Capitolo" >
                                  <%end if%>								  
								 </a>
								 <% if session("admin")=true then%>
                                  <%end if%>

                               <!-- </form>	 -->

								<% end if%>         
                     <% p=0 
		     do while not rsTabellaParagrafi.EOF  
                %>

 							    <!-- #include file = "studente_domande_include/2_frasi_1.asp" -->   
                                <!-- #include file = "studente_domande_include/2_domande_1.asp" -->   
                                <!-- #include file = "studente_domande_include/2_nodi_1.asp" -->   
                                 
                              <%  
							 ' ID_PAR=right(rsTabellaParagrafi("ID_Paragrafo"),len((rsTabellaParagrafi("ID_Paragrafo"))- instr((rsTabellaParagrafi("ID_Paragrafo"),"_"))
							  numrsMetafore=0
							 ' response.write("PPPPPPP"&rsTabellaParagrafi("Paragrafo"))
							
							  select case rsTabellaParagrafi("Paragrafo")
							 
							 case "Topolino ed Obiettivi" 
							
							 %>
							 <!--#include file = "studente_domande_include/2_metaforeT_1.asp"-->
							 <%
							 numrsMetafore=numrsTabellaT
							 numrsMetafore2=numrsTabellaPT
							   case "Navigazione nella Rete della Vita" 
								 
							 %> 
								 <!--#include file = "studente_domande_include/2_metaforeN_1.asp"-->
							 <%
								 numrsMetafore=numrsTabellaN
								 numrsMetafore2=numrsTabellaPN	 
							  case "Relazione Cliente Servitore" 
							 %>
							 	 <!--#include file = "studente_domande_include/2_metaforeD_1.asp"-->
							 <%
							   numrsMetafore=numrsTabellaD
							   numrsMetafore2=numrsTabellaPD
							   
							   case else 
							    numrsMetafore=0
							   numrsMetafore2=0
							   
							 end select  %>
							 
                                
                           <%
						   'response.write(QuerySQL&"<br><br>")
						   %>     
                                
					 <!--Qua il controllo per vedere se ci sono compiti svolti per quel paragrafo-->    
                     <%' Response.write(rsTabellaParagrafi("ID_Paragrafo") & numrsFrasi &" " & " " & numrsNodi & " " &numrsDomande & "<br>")
					 %>
					<% if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0)  or (numrsMetafore<>0) then %>
                          
 
                                       
                          <div class="accordion-group">    
                          
                                      
                          <div class="accordion-heading">
                          
                            <a id="toggleSottoPar<%=k%><%=p%>" title="<%=k%><%=p%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%><%=p%>" href="#collapseTrenew<%=k%><%=p%>">
                            <%=rsTabellaParagrafi("Paragrafo") %> <small> (<% Response.write(numrsFrasi2+numrsNodi2+numrsDomande2+numrsMetafore2)%>)</small>
                            </a>
                            
                          </div>
                          
                          
                           
                           
                          <div id="collapseTrenew<%=k%><%=p%>" class="accordion-body collapse">       
                              <ul id="myTab3" class="nav nav-tabs">
                                <% if numrsFrasi<>0 then %>
                                  <li  class="active">
								  <%else%>
                                  <li>
								  <%end if%>
                                 <a id="toggleFrasi<%=k%><%=p%>" href="#profileFrasi<%=k%><%=p%>" data-toggle="tab">Frasi (<%=numrsFrasi2%>)</a></li>
                                
                                   
                                    <% if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  <a id="toggleDomande<%=k%><%=p%>" href="#profileDomande<%=k%><%=p%>" data-toggle="tab">Domande (<%=numrsDomande2%>)</a></li>
                                  
                                   
                                   
                                       <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  
                                  <a id="toggleNodi<%=k%><%=p%>" href="#profileNodi<%=k%><%=p%>" data-toggle="tab">Nodi (<%=numrsNodi2%>)</a></li> 
                                   <% if (numrsNodi=0 ) and (numrsFrasi=0) and (numrsDomande=0) and (numrsMetafore<>0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  
                                  <a id="toggleMetafore<%=k%><%=p%>" href="#profileMetafore<%=k%><%=p%>" data-toggle="tab">Metafore (<%=numrsMetafore2%>)</a></li> 
                                  
                                  
                                       
                            </ul>
                            
                            <div id="myTabContent2<%=k%><%=p%>" class="tab-content">
                             
                              <% if numrsFrasi<>0 then %>
                                  <div class="tab-pane fade in active" id="profileFrasi<%=k%><%=p%>">
                          
								  <%else%>
                                   <div class="tab-pane fade" id="profileFrasi<%=k%><%=p%>">
                          
								  <%end if%>
                             
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                    <a title="Apri tutte le frasi del paragrafo" style="color:#FFF"  href="../cFrasi/2inserisci_valutazioni_frasi.asp?TuttePar=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>"> 
                                                    N(<%= numrsFrasi &") Pt(" & numrsFrasi2  & ") Pb("& round( numrsFrasi2/numrsFrasi,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                        <tr>
                                                            <th>Frase</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                    </thead>
                                                    <tbody>
                  
                     <% Sottoparagrafo=""
					' p=0
					riga=0
		     do while not rsTabellaFrasi.EOF  
			   if StrComp(Sottoparagrafo, rsTabellaFrasi("SotPar")) <> 0 then
			  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
			   Sottoparagrafo=rsTabellaFrasi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaFrasi("SotPar")%></b></center></td></th>  
			 <%end if%>                        
                                                        <tr id="riga_<%=riga%>">
                                                     
                                                            <%if rsTabellaFrasi("Segnalata")=1 then%>
                                                            <td > <a style="color:#F00"  href="../cFrasi/2inserisci_valutazione_frase.asp?Cartella=<%=rsTabellaFrasi("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&MO=<%=rsTabellaFrasi("ID_Mod")%>&VAL=<%=rsTabellaFrasi("Voto")%>&id_classe=<%=id_classe%>&tCap=<%=k-1%>&tSot=<%=k-1%><%=p%>&tFra=<%=k%><%=p%>"><%=rsTabellaFrasi("Chi")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaFrasi("Voto")%></td>
                                                             <%else%>
                                                              <td> <a  href="../cFrasi/2inserisci_valutazione_frase.asp?Cartella=<%=rsTabellaFrasi("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&MO=<%=rsTabellaFrasi("ID_Mod")%>&VAL=<%=rsTabellaFrasi("Voto")%>&id_classe=<%=id_classe%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tFra=<%=k%><%=p%>">   <%=rsTabellaFrasi("Chi")%></a></td>
                                                              <td><%=rsTabellaFrasi("Voto")%></td>
                                                              <%end if%>
                                                            <td><%=rsTabellaFrasi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaFrasi("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
															 <a onClick="cancella_frase(<%=rsTabellaFrasi("CodiceFrase")%>,<%=riga%>,'<%=rsTabellaFrasi("ID_Mod")%>','<%=rsTabellaFrasi(0)%>','<%=rsTabellaFrasi("Cartella")%>','<%=rsTabellaFrasi("CodiceAllievo")%>');">
                                                                <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     
                 <% f=f+1
				 '  p=p+1
				  riga=riga+1
				    rsTabellaFrasi.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                              </div>
                              
                              
                              <% 
							'  p=0
							  if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <div class="tab-pane fade in active" id="profileDomande<%=k%><%=p%>">
                             
                                     <%else%>
                                          <div class="tab-pane fade" id="profileDomande<%=k%><%=p%>">
                             
                                     <%end if%>
                              
                                  
                                  
                                  
                                   <!-- inizio blocco frasi che diventa domande-->  
                                  

                  
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaDomande.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte le domande"  href="../cDomande/inserisci_valutazioni.asp?ID_MOD=<%=rsTabellaDomande("ID_Mod")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">  
                                                    N(<%= numrsDomande &") Pt(" & numrsDomande2  & ") Pb("& round( numrsDomande2/numrsDomande,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>         
                                                         <% if not rsTabellaDomande.eof then %>
                                                        <tr>
                                                            <th>Domanda</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessuna compito inserito</th>
                                                                                                                
                                                       </tr>
                                                    <%end if%>
                                                        
                                                        
                                                    </thead>
                                                    <tbody>
                  
                      <% Sottoparagrafo=""
					' p=0
		     do while not rsTabellaDomande.EOF  
			   if StrComp(Sottoparagrafo, rsTabellaDomande("SotPar")) <> 0 then
			   Sottoparagrafo=rsTabellaDomande("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaDomande("SotPar")%></b></center></td></th>  
			 <%end if%>      
                                                    
                                                        <tr>
                                                                         
                                                                        
                                                            
                                                             <%if rsTabellaDomande("Segnalata")=1 then%>
                                                            <td > <a style="color:red"  href="../cDomande/inserisci_valutazione.asp?Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&URL=<%=rsTabellaDomande("URL_Teoria")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>"><%=rsTabellaDomande("Quesito")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaDomande("Voto")%></td>
                                                             <%else%>
                                                              <td> <a   href="../cDomande/inserisci_valutazione.asp?Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">  <%=rsTabellaDomande("Quesito")%></a></td>
                                                              <td><%=rsTabellaDomande("Voto")%></td>
                                                              <%end if%>
                                                              
                                                            
                                                              
                                                             
                                                              
                                                            <td><%=rsTabellaDomande("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaDomande("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare la domanda?');"  href="../cDomande/cancella_domanda.asp?Verifica=0&classe=<%=classe%>&cod=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>" title="Cancella">
                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     

                 <% f=f+1
				  '  p=p+1
				    rsTabellaDomande.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->  
    
                              </div>
                              
                              
                              
                             
                              
                                <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                        <div class="tab-pane fade in active" id="profileNodi<%=k%><%=p%>">
                              
                                     <%else%>
                                          <div class="tab-pane fade" id="profileNodi<%=k%><%=p%>">
                              
                                     <%end if%>
                              
                                  <!-- inizio blocco nodi -->  
                                  
                               
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaNodi.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>"> 
                                                    N(<%= numrsNodi2 &") Pt(" & numrsNodi2  & ") Pb("& round( numrsNodi2/numrsNodi,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                           <% if not rsTabellaNodi.eof then %>
                                                        <tr>
                                                            <th>Nodi</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                           
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                  
                     <% Sottoparagrafo=""
					' p=0
		     do while not rsTabellaNodi.EOF  
			   if StrComp(Sottoparagrafo, rsTabellaNodi("SotPar")) <> 0 then
			   Sottoparagrafo=rsTabellaNodi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaNodi("SotPar")%></b></center></td></th>  
			 <%end if%> 
                                                    
                                                        <tr>
                                                                                                                       
                                                            
                                                             <%if rsTabellaNodi("Segnalata")=1 then%>
                                                   <td><a  style="color:red" title="Apri il nodo"  href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaNodi("Voto")%></td>
                                                             
                                                             <%else%>
                                                       
                                                             
                                                             <td><a title="Apri il nodo"   href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>
                                                           
                                                             <td><%=rsTabellaNodi("Voto")%></td> 
                                                             
                                                              <%end if%>
                                                              
                                                              
                                                            <td><%=rsTabellaNodi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaNodi("Ora"),5)%> </td>
                                                           
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare il nodo?');"  href="../cNodi/cancella_nodo.asp?cla=<%=d%>&cod=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&CodiceTest=<%=rsTabellaNodi("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>">
                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
                                                        </tr>
                                                     
                 <% f=f+1
				    rsTabellaNodi.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->                               </div> 
                                  
                                  <!-- fine profile nodi-->
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                   
                                <% if (numrsMetafore<>0 ) and  (numrsNodi=0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                        <div class="tab-pane fade in active" id="profileMetafore<%=k%><%=p%>">
                              
                                     <%else%>
                                          <div class="tab-pane fade" id="profileMetafore<%=k%><%=p%>">
                              
                                     <%end if%>
                              
                                  <!-- inizio blocco nodi -->  
                                  
                               
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaMetafore.eof then 
														rsTabellaMetafore.movefirst
													%>
                                                    
                                                     <%=" Nn(" & numrsMetafore &") Pt(" & numrsMetafore2 &  ") ; Pb("& round( numrsMetafore2/numrsMetafore,2) &")"%>
                                                       <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                           <% if not rsTabellaMetafore.eof then %>
                                                        <tr>
                                                             <th>Percorso</th>
                                                            <th>Metafora</th>
                                                          
                                                            <th>Data</th>
                                                            <th>Ora</th>
                                                            <th class='hidden-480'>Punti</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                           
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                  
                     <% Sottoparagrafo=""
					' p=0
					
					
				
								
					
		     do while not rsTabellaMetafore.EOF  %>
             
             
              <% select case rsTabellaParagrafi("Paragrafo")
							 
							 case "Topolino ed Obiettivi" %>
							 
  <%
     QuerySQL2="SELECT SUM(Voto) AS Pt FROM Elenco_Metafore_Topolino" &_
	 " where  ThreadParent ="& rsTabellaMetafore.fields("CodiceMetafora") & ";"  
	 'conta il numero di punti ottenuti nelle metafore topolino
	' response.write(QuerySQL2)
	 Set rsTabellaPuntiPercorso = ConnessioneDB.Execute(QuerySQL2)
	 ' numrsTabella2=rsTabella2(0)
	    puntiPercorso=rsTabellaPuntiPercorso(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabellaPuntiPercorso(0)&"" =""  then
	   puntiPercorso=0
	 end if 
  %>                             
 
									<tr>
                                    <td><a target="_blank"  href="../cMetafore/sintesi_metafore.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&CodiceAllievo=<%=cod%>"><i class="icon-play-circle"></i></a></td> 
                                    
                                    <td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&Topolino=<%=rsTabellaMetafore("Topolino")%>&Formaggio=<%=rsTabellaMetafore("Formaggio")%> &Fame=<%=rsTabellaMetafore("Fame")%>&Labirinto=<%=rsTabellaMetafore("Labirinto")%>&Strada=<%=rsTabellaMetafore("Strada")%>&Strada_OK=<%=rsTabellaMetafore("Strada_OK")%>&Strada_KO=<%=rsTabellaMetafore("Strada_KO")%>&Distanza=<%=rsTabellaMetafore.fields("Distanza")%>&Testata=<%=rsTabellaMetafore.fields("Testata")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
									
									<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">
						<%=rsTabellaMetafore.fields("Topolino")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=puntiPercorso%></td> 
                        </font>
						 <%	else %>
                         
                            <%=rsTabellaMetafore.fields("Topolino")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=puntiPercorso%></td>  
					 <% end if %>	 
						
  <td class='hidden-480'><input name="checkbox" type="checkbox"> </td><td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore("Cartella")%>&Modulo=<%=rsTabellaMetafore("Id_Mod")%>&CodiceTest=<%=rsTabellaMetafore("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore("Titolo")%>&Paragrafo=<%=rsTabellaMetafore("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>
</tr>	
							 
							 
							 
							<% case "Navigazione nella Rete della Vita" %>
							<tr>
                              <td><a  target="_blank"   href="../cMetafore/sintesi_metafore.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&CodiceAllievo=<%=cod%>"><i class="icon-play-circle"></i></a></td> 
                                    	 
							
							 
										
										<td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&Autista=<%=rsTabellaMetafore(4)%>&Destinazione=<%=rsTabellaMetafore(5)%> &Carburante=<%=rsTabellaMetafore(6)%>&Luogo=<%=rsTabellaMetafore(7)%>&Strada=<%=rsTabellaMetafore(8)%>&Strada_OK=<%=rsTabellaMetafore(9)%>&Strada_KO=<%=rsTabellaMetafore(10)%>&Cespugli=<%=rsTabellaMetafore.fields("Cespugli")%>&Cestino=<%=rsTabellaMetafore.fields("Cestino")%>&Lupo=<%=rsTabellaMetafore.fields("Lupo")%>&Distanza=<%=rsTabellaMetafore.fields("Distanza")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
	<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabellaMetafore.fields("Autista")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabellaMetafore.fields("Autista")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td>
    <%end if%>
      <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
							
							
							<% case "Relazione Cliente Servitore" %>
							<td><i class="icon-play-circle"></i></td>
							
                            
                            
										
										<td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&id_classe=<%=id_classe%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&SoggettoC=<%=rsTabellaMetafore("SoggettoC")%>&DomandaC=<%=rsTabellaMetafore("DomandaC")%>&MotivazioneC=<%=rsTabellaMetafore("MotivazioneC")%>&DesiderioC=<%=rsTabellaMetafore("DesiderioC")%>&BisognoC=<%=rsTabellaMetafore("BisognoC")%>&SoggettoS=<%=rsTabellaMetafore("SoggettoS")%>&RispostaS=<%=rsTabellaMetafore("RispostaS")%>&MotivazioneS=<%=rsTabellaMetafore("MotivazioneS")%>&DesiderioS=<%=rsTabellaMetafore.fields("DesiderioS")%>&BisognoS=<%=rsTabellaMetafore("BisognoS")%>&TipoEvento=<%=rsTabellaMetafore(14)%>&TolleranzaC=<%=rsTabellaMetafore.fields("TolleranzaC")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
	<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabellaMetafore.fields("SoggettoC")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabellaMetafore.fields("SoggettoC")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td>
    <%end if%>
      <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
							
							   
							<% end select  %>
							 
			 
                                                    
                                                       
                                                     
                 <% f=f+1
				    rsTabellaMetafore.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->                               
                                  </div> 
                              
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                  
                                                           
                            </div><!-- fine MyTabContent2-->

                          </div><!-- fine collapse(treuno)-->
                        </div> <!-- fine accordino group-- da Descrizione capitolo in giù >-->
                         <%end if %> <!--if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0) then-->
                
                         <% p=p+1
						   rsTabellaParagrafi.movenext()
						   Loop
						%>  
                        
                        
                        
                        
                    </div><!-- fine accordion inner-->
                  </div>
                </div> <!--  fine accordion group uno per ogni capitolo-->
       <%compiti=compiti+1  %>       
     <% end if  ' if numrsFrasi<>0
	 %>
			
			<% k=k+1
			   rsTabellaModuli.movenext()
			   Loop
			%>    
            <% if compiti=0 then
			' Response.write("Fine" & numrsFrasi &" " & " " & numrsNodi & " " &numrsDomande & " " &numrsMetafore & "<br>")
			 %>
            <span class="alert-error"><h5>Nessun compito UWWW inserito nel periodo selezionato</h5></span>
         <!--   <ul class="tiles">
			 
            <li class="blue">
								<a href="#"><span class='nopadding'><h5>NOn ci sono compiti</h5>
                                <span class='name'><i class="icon-twitter"></i><span class="right">1min ago</span></span></a>
							</li>
            </ul>-->
            <%
			 
			end if
			%>
 
 
 
 
  
 
 
               </div> <!--<div class="bs-docs-example"> fino blocco compiti -->
		
		 
         
       
        
        
        
         
		 <!-- #include file = "../include/colora_pagina.asp" -->
        
        
         
        
        
        
	</body>
    <br><br><br><br><hr>
      <!-- #include file = "../include/footer.asp" --> 
     <script>
function cancella_frase(CodiceFrase,riga,Modulo,Paragrafo,Cartella,CodiceAllievo) {
	if (window.confirm('Vuoi veramente cancellare la frase?')) {
	  	 	var url="../cFrasi/cancella_frase_ajax.asp?CodiceFrase="+CodiceFrase+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&Cartella="+Cartella+"&CodiceAllievo="+CodiceAllievo;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								if (risposta=="Cancellazione avvenuta!")
									$('#riga_'+riga).remove();
								else
									alert(risposta);
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 }

}
</script> 
      
 <script language="javascript" type="text/javascript">
function cancella_avviso() {
	
	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {  
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();	
	 }
}
   
   
 function aggiornaStud() {
	 // alert (DataClaq);
	 var DataClaq=document.dati.txtData.value;
	 var DataClaq2=document.dati.txtData2.value;
	// alert (DataClaq);
	 // alert (DataClaq2);
		with (document.dati) { 
		 
		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daStud=1";
		 else
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daStud=1";
	
	    }
		document.dati.submit();		
}
 

</script>

<script type="text/javascript">
	
function aggiorna_studente(){

	var stud = document.getElementById("studente").value;
	 
	window.location.href = "quaderno_metafore.asp?daStud=1&DataClaq=<%=DataCla%>&DataClaq2=<%=DataCla2%>&id_classe=<%=id_classe%>&classe=<%=classe%>&cod="+stud;
	
}		 
$(window).load(function () {
	   
	   $('#<%=box_apri%>').click();
	   $('#<%=box_apri1%>').click();
	    $('#<%=box_apri2%>').click();
		$('#<%=box_apri3%>').click();
	    $('#<%=box_apri4%>').click();
	   $("body").addClass("theme-"+"<%=stile%>").attr("data-theme","theme-"+"<%=stile%>");
  
  
	 
	  // event.stopPropagation();
	    
	});
	

/*$(".red").click(function(event){
   
   // alert("Hai cliccato sull'Elemento");
	document.location = "script/aggiorna_stile.asp?stile=red"
});
*/	
	
</script>

 
	</html>