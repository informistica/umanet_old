<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Collega nodi</title>   
   
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
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
	
	
    
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

<script language="javascript" type="text/javascript" >
		function myLink(Id_n1,L1,Id_n2,L2,Tipo,Corso,CodiceTest,Capitolo,Paragrafo,Modulo,Stato) {
		var testo = prompt("Inserisci testo del collegamento", "");
	
		location.href="inserisci_collegamento1.asp?Id_n1="+Id_n1+"&L1="+L1+"&Id_n2="+Id_n2+"&L2="+L2+"&Tipo="+Tipo+"&Corso="+Corso+"&CodiceTest="+CodiceTest+"&Capitolo="+Capitolo+"&Paragrafo="+Paragrafo+"&Modulo="+Modulo+"&Stato="+Stato+"&T2="+testo;

	
}
	</script>
  


   
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
     
   <%
   Stato=Request.QueryString("Stato")
  CodiceTest=Request.QueryString("CodiceTest") 
  
  
  
  if CodiceTest<>"" then 
     Session("CodiceTest")=CodiceTest
  end if 
		  Capitolo=Request.QueryString("Capitolo")
		  Paragrafo=Request.QueryString("Paragrafo")
		  Modulo=Request.QueryString("Modulo")
		  Nome=Request.QueryString("Nome")
		  Cognome=Request.QueryString("Cognome")
		  Cartella=Request.QueryString("Cartella")
		  Corso=Request.QueryString("Corso") ' serve per distinguere quando è stato scelto il corso da visualizzare
		  Id_n1=Request.QueryString("Id_n1")   'id del nodo di partenza del link (href che punto all'ancora nel documento)
		  Id_n2=Request.QueryString("Id_n2")  ' 'id del nodo di arrivo del link   (ancora puntata dall'href)
		  L1=Request.QueryString("L1") ' livello del primo nodo da cui parte il link (chi, cosa, dove, ecc...)
		  L2=Request.QueryString("L2")' livello del secondo nodo a cui arriva il link (chi, cosa, dove, ecc...)
		  T2=Request.QueryString("T2") ' testo nel livello di arrivo da visualizzare sull'arco che collega i nodi
 	
   %>
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	 
          
         
	</div>
    
 <%
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	response.write(QuerySQL)
	Set rsTabellaSetting = ConnessioneDB.Execute(QuerySQL) 
	'response.write(QuerySQL)
	Nodi=rsTabellaSetting("Nodi") 
	'response.write(Nodi)
	rsTabellaSetting.close
 
 
 
 
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="glyphicon-link"></i> Collega nodi</h1> 
                    
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
							 <a href="#">Collega nodi</a> 
                             
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
                      
 
 									 
	 
	 <% 'Stato:  0 livello paragrafo, 1 livello modulo, 2 livello corso, 3 livello corsi
	   if (Stato="") then%>
  <!-- Per distinguere il tipo di query da eseguire, per capire quali nodi vanno prelevati: se quelli del singolo paragrafo, del modulo, di tutto il corso corrente, di tutti i corsi-->
	<h5>Scegli il punto di vista <%'response.write("Paragra="&Paragrafo) %></h5>
	<a href="inserisci_collegamento.asp?Stato=0&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Paragrafo</a><br>
	<a href="inserisci_collegamento.asp?Stato=1&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modulo</a><br>
	<a href="inserisci_collegamento.asp?Stato=2&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Corso</a><br>
	<a href="inserisci_collegamento.asp?Stato=3&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Corsi</a></div><br>

<%else
	  if (strcomp(Stato,"3"))=0 and (Corso="") then ' Se scelgo il punto di vista corsi allora devo selezionare il corso
		 
		   QuerySql="SELECT DISTINCT (Moduli.Cartella)  FROM Moduli WHERE Moduli.cartella<>'';"
		   Set rsCorsi = ConnessioneDB.Execute(QuerySQL)%>
		   <h5 align="center">:<br>
		   <br>
		   <table class="table table-hover table-nomargin table-bordered" width="25%">
            <tr><th>Scegli il corso che vuoi linkare al corso <%=cartella%></th></tr>
		   <%Do until (RsCorsi.eof) %>
          
								<tr><td align="center"><a href="inserisci_collegamento.asp?Corso=<%=rsCorsi(0)%>&Stato=3&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><%=rsCorsi(0)%></a></td></tr>
		  <%rsCorsi.Movenext()
		   Loop 
		   rsCorsi.close()
		   set rsCorsi=nothing
		 %>
		 </table>
		 
		 </div><br>

	 <% else
		     
		 
		  
		  
		
		  
		  
		Dim objFSO, objTextFile
		Dim sRead, sReadLine, sReadAll
		Const ForReading = 1, ForWriting = 2, ForAppending = 8
		Set objFSO = CreateObject("Scripting.FileSystemObject")
								
		%>
		 
	<!--	 
		 <table class="table table-hover table-nomargin table-bordered" align=center width="60%">
				<thead><tr>
					<th width="50%"><font color="#0022FF"><b>Paragrafo</b></font></th>
					<th width="21%"><font color="#0022FF"><b>Codice Nodo</b></font></th>
					<th width="29%"><font color="#0022FF"><b>Studente</b></font></th>
				</tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Chi    : Soggetto che compie l'Azione</b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Cosa   : Azione</b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Dove   : Spazio</b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Quando : Tempo </b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Come   : Modo </b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Perchè : Motivazione</b></font></td></tr>
				<tr><td colspan=3><p align="center"><font color="#0022FF"><b>Quindi : Conclusione</b></font></td></tr> 
		  </table>
			
		-->	
			
		</p> <!-- stampa il titolo del test -->
		
		 
			<br>
		<%   
		  if CodiceTest="" Then ' se si perde il codice lo riassegno
				   CodiceTest=Session("CodiceTest")
		   end if 
		   'in base al punto di vista scelto (paragrafo,modulo,corso,corsi) preparo la query
		Select Case Stato ' per distinguere il livello scelto (paragrafo.modulo.corso.corsi)
				  Case 0	   
			 	   'Definzione codice SQl della query per ricercare i nodi del paragrafo 
						if Nodi=1 then ' sono condivisi non cerco in base a CodiceAllievo	 
						QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Paragrafi.Posizione " &_
						" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg" &_
						" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione " &_
						" HAVING Paragrafi.ID_Paragrafo='" & CodiceTest & "' " &_   
						" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;" 
						else
						
						QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Paragrafi.Posizione " &_
						" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg" &_
						" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione " &_
						" HAVING Paragrafi.ID_Paragrafo='" & CodiceTest & "' and Nodi.Id_Stud='"& Session("CodiceAllievo") &"'" &_
						" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;" 
						
						
						end if
				  Case 1
				    'Definzione codice SQl della query per ricercare i nodi del modulo 
					if Nodi=1 then
						QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Paragrafi.Posizione " &_
					" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
					" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione, Nodi.Id_Stud " &_
					" HAVING Moduli.ID_Mod='" & Modulo & "'" &_ 
					" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo ;"
					else
													'0						1				2			 3         4               5             6            7          8          9        10          11 
						QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Paragrafi.Posizione " &_
					" FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo=Nodi.Id_Stud) ON Moduli.ID_Mod=Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo=Nodi.Id_Arg" &_
					" GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod,Nodi.Cosa,Nodi.Dove,Nodi.Quando,Nodi.Come,Nodi.Perche,Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione, Nodi.Id_Stud " &_
					" HAVING Moduli.ID_Mod='" & Modulo & "' and Nodi.Id_Stud='"& Session("CodiceAllievo") &"'" &_ 
					" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo ;"
					end if
				  Case 2
				    'Definzione codice SQl della query per ricercare i nodi del corso 
				    if Nodi=1 then
						QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Paragrafi.Posizione  "&_
					"FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg "&_
					"GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione,Moduli.Cartella, Nodi.Id_Stud   "&_
					"HAVING (Moduli.Cartella Like '"&Cartella&"%')" &_
					"ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;  " 
					else
					QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Paragrafi.Posizione  "&_
					"FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg "&_
					"GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione,Moduli.Cartella, Nodi.Id_Stud   "&_
					"HAVING (Moduli.Cartella Like '"&Cartella&"%') and (Nodi.Id_Stud='"& Session("CodiceAllievo") &"')" &_
					"ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo;  " 
					end if

			       Case 3
				    'Definzione codice SQl della query per ricercarei nodi dei corsi 
					if Nodi=1 then 
					QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Paragrafi.Posizione  "&_
					"FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg "&_
					"GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione,Moduli.Cartella, Nodi.Id_Stud   "&_			
					"HAVING (Moduli.Cartella Like '"&corso&"') And (Moduli.Cartella Like '"&cartella&"')"&_
					" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo ; " 
					else
					QuerySQL="SELECT Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Paragrafi.Posizione  "&_
					"FROM Paragrafi INNER JOIN (Moduli INNER JOIN (Allievi INNER JOIN Nodi ON Allievi.CodiceAllievo = Nodi.Id_Stud) ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg "&_
					"GROUP BY Paragrafi.Titolo, Paragrafi.ID_Paragrafo, Allievi.Cognome, Nodi.Chi, Nodi.CodiceNodo, Moduli.ID_Mod, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi,Nodi.In_Quiz,Paragrafi.Posizione,Moduli.Cartella, Nodi.Id_Stud   "&_			
					"HAVING (Moduli.Cartella Like '"&corso&"') And (Moduli.Cartella Like '"&cartella&"') and Nodi.Id_Stud='"& Session("CodiceAllievo") &"'"&_
					" ORDER BY Paragrafi.Posizione,Nodi.CodiceNodo ; " 
					end if
				  Case Else
				   ' Istruzioni di default
				   ' ho tolto da tute le query la condizione  And Nodi.In_Quiz<>0
				End Select
	
	' Set objFSO = CreateObject("Scripting.FileSystemObject")  
'   	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logSpiegazioneNodi.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL& "-"& Session("Admin"))
'				objCreatedFile.Close 
'		
		'response.write("nodi="&nodi&"<br>"&stato&"<br>"&QuerySQL)	
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		 
			  
		%>
		<% If rsTabella.BOF=True And rsTabella.EOF=True Then 
		 
		
			%><span class="alert-error">
			  <h5>Nodi della rete non ancora disponibili!<h5>
              </span>
			  <p><h5><a href="javascript:history.back()"onMouseOver="window.status='Indietro';return true;" onMouseOut="window.status=''">Indietro</a>
			</H5>
		<% Else
				if (Id_n1="") and (Id_n2="") then ' se non ho ancora selezionato nè il nodo di partenza nè il nodo di arrivo %>
						  <h5 align="center">
						  <span>Scegli il punto di partenza <i class=" icon-arrow-right"></i> del collegamento <i class="glyphicon-link"></i> </span><br><br>
			 
						  <%' visualizzo tutti i nodi con i campi href, cliccando su un livello richiamo la pagina e passo i parametri
						    Do until rsTabella.EOF
							 ID=rsTabella("CodiceNodo")
						  %>
						     <table class="table table-hover table-nomargin table-bordered" align=center width="60%">
								<thead>
                                <tr>
									<th><%=rsTabella("Titolo")  %></th>
									<th><%=rsTabella("CodiceNodo")%></th>
									<th><%=rsTabella("Cognome")%></th>
								</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
                                </thead>
                                <tbody>
								<!--in base a dove clicco stabilisco il livello semantico coinvolto nel link in partenza (href che punterà ad un ancora)-->
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=1&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Chi</a>    :<%=rsTabella(3)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=2&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Cosa</a>   : <%=rsTabella(6)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=3&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Dove</a>   : <%=rsTabella(7)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=4&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Quando</a>   : <%=rsTabella(8)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=5&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Come</a>   : <%=rsTabella(9)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=6&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Perch&egrave;</a>   : <%=rsTabella(10)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=ID%>&L1=7&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Quindi</a>   : <%=rsTabella(11)%></b></td></tr>
                                </tbody>
						     </table>
							<br>						
							  <%rsTabella.MoveNext 
							Loop 
				else 
					  if (Id_n2="") then ' se ho già scelto il  punto di partenza del collegamento ora scelgo quello di arrivo%> 
						 <h5 align="center">
						  <span class="style1">Scegli il punto di <i class=" icon-arrow-right"></i> arrivo del collegamento <i class="glyphicon-link"></i></span><br>
						  <br></h5>
			 
						  <%Do until rsTabella.EOF
							ID=rsTabella("CodiceNodo")
						  %>
						  <table class="table table-hover table-nomargin table-bordered"  align=center width="60%">
                          <thead>
								<tr>
									<th><%=rsTabella("Titolo")%></td>
									<th><%=rsTabella("CodiceNodo")%></td>
									<th><%=rsTabella("Cognome")%></td>
								</tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
                           </thead>
						   <!--in base a dove clicco stabilisco il livello semantico coinvolto nel link in arrivo (ancora puntata da href)-->
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=1&T2=<%=rsTabella(3)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Chi</a>    :<%=rsTabella(3)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=2&T2=<%=rsTabella(6)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Cosa</a>    :<%=rsTabella(6)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=3&T2=<%=rsTabella(7)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Dove</a>    :<%=rsTabella(7)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=4&T2=<%=rsTabella(8)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Quando</a>    :<%=rsTabella(8)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=5&T2=<%=rsTabella(9)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Come</a>    :<%=rsTabella(9)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=6&T2=<%=rsTabella(10)%>Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Perch&egrave;</a>    :<%=rsTabella(10)%></b></td></tr>
								<tr><td colspan=3><b><a href="inserisci_collegamento.asp?Tipo=0&Corso=<%=Corso%>&Stato=<%=Stato%>&Id_n1=<%=Id_n1%>&L1=<%=L1%>&Id_n2=<%=ID%>&L2=7&T2=<%=rsTabella(11)%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>"><i class="glyphicon-link"></i>Quindi</a>    :<%=rsTabella(11)%></b></td></tr>
						    </table>
							<br>
						<%
							   rsTabella.MoveNext  
							Loop 
					     else  ' arrivo qui solo quando ho valorizzato sia Id_n1 che Id_n2
								'Esecuzione della query per inserire il collegamento
								tipo=0
								Corso=0
								
								%>
							<script>
						 
								myLink(<%=Id_n1%>,<%=L1%>,<%=Id_n2%>,<%=L2%>,<%=tipo%>,<%=Corso%>,"<%=CodiceTest%>","<%=Capitolo%>","<%=Paragrafo%>","<%=Modulo%>",<%=Stato%>);
								 
							</script>	
							 
								<%
							' qui non ci arriva più perchè passa al javascript	 
						  ' QuerySQL="INSERT INTO Link (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2,Data) VALUES (" & clng(Id_n1) & "," &L1 & ", " & clng(Id_n2) & "," & L2 & ",'" & Session("CodiceAllievo")& "','" &T2 & "','" &Now & "');"						 
						  ' QuerySQL="INSERT INTO Link (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) VALUES (" & clng(Id_n1) & "," &L1 & ", " & clng(Id_n2) & "," & L2 & ",'" & Session("CodiceAllievo")& "','" &T2& "');"						 
						  ' response.write(QuerySql)
						 '  ConnessioneDB.Execute QuerySQL 						
					     '  response.write(QuerySql)  ' eseguita una volta ma nel db continua come per inerzia l'inserimento di record %>
					 <b><a href="inserisci_collegamento.asp?Tipo=0&Stato=<%=Stato%>&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua a collegare</a>   <br><br>
					<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h5><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h5> 
					 <% end if%>
				<%end if%>		
		<% 
		  End If 		 
		  rsTabella.Close : Set rsTabella = Nothing  ' libera le risorse chiudendo gli oggetti aperti 
		  ConnessioneDB.Close : Set ConnessioneDB = Nothing 
		 %>
 
    <% end if%>
  <% end if%>
				 
				 
                   
              <br>
     <a target="_blank" href="../cMap/spiegazione_mappa.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">Apri rete di nodi interattiva</a> <br>
      
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

