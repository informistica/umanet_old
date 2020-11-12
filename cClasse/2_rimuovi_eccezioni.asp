<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Consulta e Rimuovi Eccezioni</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
<meta charset="utf-8">
    
    


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
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />
       
       
      
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta oppure hai cercato di leggere i dati degli altri studenti!")
location.href="../../../../"
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
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") or ((session("CodiceAllievo") <> Request.QueryString("cod")) and (session("admin") = false))then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
	<% 
	
	daRimuovi = Request.QueryString("daRimuovi")
	
	 if daRimuovi = 1 then
		provURL = session("provURL")
	 else
		 provURL = Request.ServerVariables("HTTP_REFERER")
		 session("provURL") = provURL
	end if
		
		%>
     <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  <% end if %>


	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
       
          
         
	</div>
    
    <% Id_Stud = Request.QueryString("cod")
		
		'response.write(Id_Stud)
		
		fromURL = Request.ServerVariables("HTTP_REFERER")
		
		frasi = Request.QueryString("frasi")
		nodi = Request.QueryString("nodi")
		domande = Request.QueryString("domande")

		Id_Stud = Request.QueryString("cod")

		if frasi = 1 then
			tabella = "Eccezioni_Frasi"
			nome = "Frasi"
			nomesing = "frase"
		end if 
		 
		if nodi = 1 then
			tabella = "Eccezioni_Nodi"
			nome = "Nodi"
			nomesing = "nodo"
		end if
		
		 
		if domande = 1 then
			tabella = "Eccezioni_Domande"
			nome = "Domande"
			nomesing = "domanda"
		end if
		
		if frasi="" and nodi="" and domande="" then
			Response.Redirect fromURL
		end if
		
		if Id_Stud<>""then
		 QuerySQL="SELECT *  FROM Allievi WHERE CodiceAllievo='" & Id_Stud & "'" 
		Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
		Cognome1=rsTabella("Cognome")
		Nome1=rsTabella("Nome")
		
		end if
		%>
	
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h3> <i class="icon-comments"></i> Consulta e Rimuovi Eccezioni <%=nome%> per <%=Cognome1&" "%><%=Nome1%></h3> 
                    
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
							<a href="#">Quaderno</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Rimuovi Eccezioni</a>
                           
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
				        <h3> <i class="icon-reorder"></i>  ECCEZIONI DISPONIBILI</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">
                      
 <% 
			Id_Stud = Request.QueryString("cod") 'devo riscriverlo perché perde il parametro
			
			QuerySQL="SELECT count(*) FROM "&tabella&" WHERE Id_Stud='"& Id_Stud &"';"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			numEccezioni=rsTabella(0)
			
			'response.write(Id_Stud)
			'response.write(QuerySQL)
			
			QuerySQL="SELECT TOP (50) * FROM "&tabella&" WHERE Id_Stud='"&Id_Stud&"' order by Scadenza desc;"
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			%>
			
			 
			<table  class="table table-hover table-nomargin table-condensed" style="width:100%">
			  <tr><td><b>ID Pre<%=nomesing%></b></td><td><b>Studente</b></td><td><b>Data Scadenza</b></td>
			  <% if session("admin") = true then %>
			  <td><b>Rimuovi</b></td>
			    <% else%>
				<td><b>Esegui</b></td>
			  <% end if%>
			  </tr>
			<%i=0
			do while not rsTabella.eof %>
		
			<%nomecolonna1 = "ID_Pre"&nomesing%>
		
			<% if DateDiff("D", Date(), rsTabella.fields("Scadenza")) >= 0 then
				stato = "Disponibile"
				colore = "green"
			  else
				stato = "Scaduto"
				colore = "red"
			 end if
			%>
				
				<% QuerySQL = "SELECT * FROM pre"&nome&" WHERE Id_Pre"&nomesing&" ='"&rsTabella.fields(nomecolonna1)&"';" 
				'response.write(QuerySQL)
				
				set rsTabella_new0 = ConnessioneDB.Execute(QuerySQL)
				quesito = rsTabella_new0("Quesito")
				'response.write(quesito)
				Modulo=rsTabella_new0("Id_Mod")
				CodiceTest=rsTabella_new0("Id_Paragrafo")
				Id_Sottoparagrafo=rsTabella_new0("Id_Sottoparagrafo")
			
				qsql="select Titolo from Moduli where ID_Mod='"&Modulo&"'"
				set rsTab = ConnessioneDB.Execute(qsql)
				capitolo=rsTab(0)
				qsql="select Titolo from Paragrafi where ID_Paragrafo='"&CodiceTest&"'"
				set rsTab = ConnessioneDB.Execute(qsql)
				paragrafo=rsTab(0)
				if Id_Sottoparagrafo<>"" then
				qsql="select Titolo from Sottoparagrafi where ID_Sottoparagrafo='"&Id_Sottoparagrafo&"'"
				set rsTab = ConnessioneDB.Execute(qsql)
				Sottoparagrafo=rsTab("Titolo")
				end if


				'capitolo = right(rsTabella_new0("Id_Mod"), 1)
				'response.write(capitolo)
				
				parquery = rsTabella_new0("Id_Paragrafo")
				
				QuerySQL = "SELECT Titolo, Posizione FROM Paragrafi WHERE Id_Paragrafo='"&parquery&"';"
				'response.write(QuerySQL)
				
				set rsTabella_new = ConnessioneDB.Execute(QuerySQL)
				
			'	paragrafo = rsTabella_new(1)
				titolo = rsTabella_new(0)
				
				stringadettagli = "Quesito: "&quesito&"\nTitolo del Paragrafo: "&titolo&"\nCapitolo: "&capitolo 
				'response.write(stringadettagli)
				%>
				
				 
				
				<tr><td><a style="text-decoration:none" href="#" onclick="alert('<%=stringadettagli%>')"><%=quesito%></a></td><td><%=rsTabella.fields("Id_Stud")%></td>
				<td><%=rsTabella.fields("Scadenza")%> - <span style='color:<%=colore%>' ><%=stato%></span></td>
				<% if session("admin") = true then%>
				<td><a style="text-decoration:none" href="studente_domande_include/2_rimuovi_eccezione.asp?cod=<%=Id_Stud%>&<%=nomesing%>=1&ID=<%=rsTabella.fields(nomecolonna1)%>">x</a>
				</td>
				<%else%>
				 <td><a style="text-decoration:none" href="../cFrasi/2inserisci_frase.asp?&Tipo=0&Quesito=<%=quesito%>&Cartella=<%=Session("Cartella")%>&Capitolo=<%=capitolo%>&Paragrafo=<%=paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&prefrase=1&ID_Prefrase=<%=rsTabella_new0("ID_Prefrase")%>&Scadenza=<%=rsTabella_new0("Scadenza")%>&Img=<%=rsTabella_new0("Img")%>&cFile=<%=rsTabella_new0(("Files"))%>&CodiceSottopar=<%=Id_Sottoparagrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>">+</a>
				 </td>
				<%end if%> </tr>
			
			<%
			 i=i+1
			 rsTabella.movenext
			loop%>
   </table>
      </div>
					
					
					<a style="text-decoration:none" href=<%=provURL%>><h5><% if session("admin") = true then %>Torna al Quaderno dello Studente<% else %>Torna al tuo Quaderno<% end if %></h5></a>
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
		
		<script>
		
		function dettagli(id, tipo){
		
			
			
		
		}
		
		</script>
		
	</body>

 </html>

