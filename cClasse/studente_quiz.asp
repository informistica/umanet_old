<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica compito</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
		<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">

    


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

	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
     
    
	<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>
<%Response.Buffer = true
  On Error Resume Next  
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
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>&nbsp;Modifica compiti</h1> 
                    
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
							 <a href="#">Modifica</a> 
                             
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
 <% CodiceTest=Request.QueryString("CodiceTest")
 
Cartella=Request.QueryString("Cartella")
CodiceAllievo=Session("CodiceAllievo")
 
  'CodiceDomanda=Request.QueryString("CodiceDomanda")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  testnodo=Request.QueryString("testnodo") ' parametro passato da scegli_azione_test : 0 modifico il test , 1 modifico i nodi
  testmetafora=Request.QueryString("testmetafora") 
  VF=Request.QueryString("VF")  ' vale 1 se è domanda vero falso
   Multiple=Request.QueryString("Multiple")  ' vale 1 se è domanda vero falso
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
cod=Session("CodiceAllievo")
%>                
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> Compiti su "<%=Capitolo%> : <%=Paragrafo%>"</h3>
			          </div>
				      <div class="box-content">
                      
 
 		<table class="table table-hover table-nomargin"><%
if testnodo=0 then ' se non sto parlando di nodo ma di domanda o metafora%>
 <%  if testmetafora=1 then%>
 		   <center>
           <H5 align="center">Scegli la metafora da modificare: </h5>
          <% QuerySQL="SELECT * FROM Elenco_Metafore_Navigazione WHERE Allievi.CodiceAllievo='" & cod & "'" &_
           "  order by CodiceMetafora asc;"
		   response.write(QuerySQL)
          Set rsTabella = ConnessioneDB.Execute(QuerySQL)
           If rsTabella.BOF=True And rsTabella.EOF=True Then %>
              <div class="alert-error">
              <H5>Non ci sono metafore da modificare!</h5></center>
              </div>
              
        %>
         
          <%else%>
          
		  <% do while not rsTabella.eof
			' response.write(rsTabella(0))
			 %>
			<tr><td><a href="../cMetafore/inserisci_valutazione_metafore.asp?id_classe=<%=Session("Id_Classe")%>&DATA=<%=rsTabella.fields("Data")%>&Cartella=<%=rsTabella.fields("Cartella")%>&classe=<%=rsTabella.fields("Cartella")%>&cod=<%=cod%>&CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabella.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&Capitolo=<%=rsTabella.fields("Titolo")%>&TitoloParagrafo=<%=rsTabella.fields("Tit")%>&Paragrafo=<%=rsTabella.fields("Tit")%>&Autista=<%=rsTabella(4)%>&Destinazione=<%=rsTabella(5)%> &Carburante=<%=rsTabella(6)%>&Luogo=<%=rsTabella(7)%>&Strada=<%=rsTabella(8)%>&Strada_OK=<%=rsTabella(9)%>&Strada_KO=<%=rsTabella(10)%>&Cespugli=<%=rsTabella.fields("Cespugli")%>&Cestino=<%=rsTabella.fields("Cestino")%>&Lupo=<%=rsTabella.fields("Lupo")%>&Distanza=<%=rsTabella.fields("Distanza")%>&MO=<%=rsTabella.fields("ID_Mod")%>&VAL=<%=rsTabella.fields("Voto")%>&URL=<%=rsTabella.fields("URL_Teoria")%>&Segnalata=<%=rsTabella.fields("Segnalata")%>&Pippo=1 "><%=rsTabella.fields("Autista")%></a></td></tr>
			<%
			rsTabella.movenext
			loop
		   end if	
       else  ' è una domanda , devo vedere se tipo Vero Falso o Multipla  
		
		  preQuerySQL="SELECT * from MODULO_PARAGRAFO_DOMANDE1" 

		 if VF<>"" then ' domanda di tipo vero falso
			 QuerySQL=preQuerySQL &_
			" WHERE  CodiceAllievo='" & CodiceAllievo &"' AND ID_Paragrafo='"&CodiceTest &"' AND  VF=1 " 	   
		 else
		 	if Multiple<>"" then ' risposta multipla
			     QuerySQL=preQuerySQL &_
			" WHERE  CodiceAllievo='" & CodiceAllievo &"' AND ID_Paragrafo='"&CodiceTest &"' AND  Multiple=1 " 	   
			else
			    if Immagini<>"" then
				else
				QuerySQL=preQuerySQL &_ 
			" WHERE  CodiceAllievo='" & CodiceAllievo &"' AND  ID_Paragrafo='"&CodiceTest &"' AND  Multiple=0 and  VF=0     " 	
				end if
			end if
		  
	
		end if
		if CodiceSottopar<>"" then
		  QuerySQL=QuerySQL &" and ID_Sottoparagrafo='"&CodiceSottoPar&"';"
		end if
		'QuerySQL=QuerySQL &"ORDER BY Domande.CodiceDomanda;"
		 response.write(QuerySQL)
		 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		
			If rsTabella.BOF=True And rsTabella.EOF=True Then %>
		 
          <center>
          <div class="alert-error">
          <H5>Non ci sono domande da modificare!</h5></center>
           </div>
		  
		<% Else %>
         
           <H5 align="center">Scegli la domanda da modificare: </h5>
           <table class="table-condensed">
		<%    do while not rsTabella.eof
			' response.write(rsTabella(0))
			 %>
			<tr><td><a href="../cDomande/inserisci_modifica.asp?Tipodomanda=<%=rsTabella(11)%>&Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&CodiceDomanda=<%=rsTabella(9)%>&Capitolo=<%=rsTabella(7)%>&Paragrafo=<%=rsTabella(8)%>&Quesito=<%=rsTabella(1)%>&R1=<%=rsTabella(2)%>&R2=<%=rsTabella(3)%>&R3=<%=rsTabella(4)%>&R4=<%=rsTabella(5)%>&RE=<%=rsTabella(6)%>&MO=<%=rsTabella(10)%>&Multiple=<%=rsTabella.fields("Multiple")%>&VF=<%=rsTabella.fields("VF")%>"><%=rsTabella("Quesito")%></a></td></tr>
			<%
			rsTabella.movenext
			loop
		   
		   
		end if 
end if
rsTabella.close
		%>
</table>
<%else%>
<br> 

<% '    						0					1			2			3		   	4 	     5				6			7			8			9				10				11	
  QuerySQL="SELECT Allievi.CodiceAllievo, Nodi.CodiceNodo, Nodi.Chi, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Moduli.Titolo, Paragrafi.Titolo,Moduli.ID_Mod"&_
	" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg " &_
	" WHERE Nodi.Id_Stud='" & CodiceAllievo &"' AND Nodi.Id_Arg='"&CodiceTest &"' ORDER BY Nodi.CodiceNodo" 
	
	 
	 Set rsLink = ConnessioneDB.Execute(QuerySQL)
	 'response.write(QuerySql)
	If rsLink.BOF=True And rsLink.EOF=True Then %>
	<br><center>
    <div class="alert-error">
 <H5>Non ci sono nodi da modificare!</h5>
 </div>
 
	  
	<% Else%>  
<center>
	  <H5>Scegli il nodo da modificare: </h5>
	 <table class="table table-hover table-nomargin">
	 <%do while not rsLink.eof %>
	<tr><td><a href="../cNodi/inserisci_modifica_nodo.asp?Cartella=<%=Cartella%>&CodiceTest=<%=CodiceTest%>&CodiceNodo=<%=rsLink(1)%>&Capitolo=<%=rsLink(9)%>&Paragrafo=<%=rsLink(10)%>&Chi=<%=rsLink(2)%>&Cosa=<%=rsLink(3)%>&Dove=<%=rsLink(4)%>&Quando=<%=rsLink(5)%>&Come=<%=rsLink(6)%>&Perche=<%=rsLink(7)%>&Quindi=<%=rsLink(8)%>&MO=<%=rsLink(11)%>"><%=rsLink(2)%></a></td></tr>
	<%
	rsLink.movenext
	loop
	end if 
	rsLink.close
	%>
</table>
<% end if%>
						 
	 
	 
				 
				 
                   
                   
 
		  			  <div class="box-content"> 
                      <center><br>
<p><h6><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
	</H6></p>
</center>	
            			   
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

