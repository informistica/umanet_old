<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Report crediti</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<meta charset="utf-8">
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

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  
	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
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
						<h1> <i class="icon-comments"></i> Report Crediti </h1> 
                    
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
							<a href="#more-files.html">Classifica</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Crediti</a>
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
                         Risultati della prova
						 
                         </h3>
			          </div>
				      <div class="box-content">
                      
 <% 

    AggiungiReport=Request.QueryString("AggiungiReport") ' <>"" se devo aggiungere nuovi crediti, metodo alternativo all'aggiunta dalla classifica
	id_classe=Session("Id_Classe") 
	if AggiungiReport="" then
			Id_Eser=Request.QueryString("ID_ESER")
			QuerySQL="SELECT count(*) FROM [2REPORT_CREDITI] Where ID_Esercitazione="& Cint(Id_Eser) &" ; "
			'response.write(QuerySQL)
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			numStud=rsTabella(0)
			QuerySQL="SELECT * FROM [2REPORT_CREDITI] Where ID_Esercitazione="& Cint(Id_Eser) &" order by  Cognome, Crediti desc "
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			
			
			dataProva=rsTabella.fields("Data")
			Response.Write("<center><h3><font color=red > '"&rsTabella.fields("Descrizione")&"' del  " & rsTabella.fields("Data") & "</font></h3></center> ")
			%>
			<center>
			<form name="AggiornaReport" class="form-horizontal" action="report_verifiche_aggiorna.asp?id_classe=<%=id_classe%>&ID_ESER=<%=Id_Eser%>&numStud=<%=numStud%>&DataTest=<%=dataProva%>" method="post">
			 <input type="hidden" name="Id_Eser1" value="<%=Id_Eser%>" >
			<table  class="table table-hover table-nomargin table-condensed" style="width:50%">
			  <tr><td><b>Cognome</b></td><td><b>Nome</b></td><td><b>Punti</b></td> </tr>
			<%i=0
			consegnato=""
			do while not rsTabella.eof %>
		
				<tr><td><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td>
				<td><input type="text" name="Punti<%=i%>" value="<%=rsTabella.fields("Crediti")%>" size="1" class="input-mini">
				<input type="hidden" name="CodiceAllievo<%=i%>" value="<%=rsTabella.fields("CodiceAllievo")%>" >
				</td> </tr>
			
			<%
			 i=i+1
			 consegnato=consegnato&"'"&rsTabella.fields("CodiceAllievo")&"'"&","
			 rsTabella.movenext
			loop
		
			
			if Session("Admin")=true then
			%> 
            <tr><td>
           <input type="text" name="txtData" value="<%=dataProva%>" class="input-small">
			</td></tr>
			<tr><td>&nbsp;</td></tr>
			   <tr><td colspan="3">
				  <input type="hidden" name="numStud" value="<%=i%>" >
				 <input type="submit" value="Aggiorna" class="btn-primary">
			   </td></tr>
			<% end if%>
			 </table>
			 <%	consegnato=left(consegnato,len(consegnato)-1) ' tolgo ,
		QuerySQL="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") order by Cognome;"
		'response.write(QuerySQL)
		set rsTabellaNew= ConnessioneDB.Execute(QuerySQL)
		response.write("<h3>Assenti :</h3>")
		do while not rsTabellaNew.eof
			response.write("<br><font color='red'> "&rsTabellaNew("Cognome") &" " & left(rsTabellaNew("Nome"),1)&"."&"</font> ")
			rsTabellaNew.movenext
		loop%>
		 
		    </form>
		    </center>
		
		 
		<br><center>
		<a href="../cGrafici/genera_grafico_report.asp?ID_ESER=<%=Id_Eser%>">Visualizza grafico</a><br></p>
		 <a href="javascript:history.back()">	Indietro </a></center>
         
      <%else ' if aggiungireport="" %>
      <%' non devo aggiornare ma aggiungere un nuovo report
			QuerySQL="Select count(*) from Allievi where Id_Classe='" & id_classe & "' and Attivo=1;" 			
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			numStud=rsTabella(0)
			QuerySQL="Select * from Allievi where Id_Classe='" & id_classe & "' and Attivo=1 order by Cognome asc;" 			
			Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			%>
			<br> <p align="center"><center>
			<form name="AggiornaReport" class="form-horizontal" action="report_verifiche_aggiorna.asp" method="post">
		    <FIELDSET style="margin:0 auto 0 auto;"><LEGEND class="sottotitoloquaderno2"><B> Aggiorna punteggio attivit&agrave;</B></LEGEND>
            <input type="hidden" name="AggiungiReport" value="1" >
             <input type="hidden" name="numStud" value="<%=numStud%>" >
            <input type="hidden" name="id_classe" value=<%=id_classe%>>
			<table class="table table-hover table-nomargin table-condensed" style="width:50%">
			  <tr><td><b>Cognome</b></td><td><b>Nome</b></td><td><b>Punti</b></td> </tr>
			<%i=0
			do while not rsTabella.eof %>
				<tr><td><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td>
				<td><input type="text" name="Punti<%=i%>" value="0" size="1" class="input-mini">
				<input type="hidden" name="CodiceAllievo<%=i%>" value="<%=rsTabella.fields("CodiceAllievo")%>" >
                
				</td> </tr>
			<%
			 i=i+1
			 rsTabella.movenext
			loop%>
           </table>
 <br><br>
 <!-- Per fare in modo che l'esercitazione svolta in una certa data compaia nell'elenco delle attività-->


<center><table  class="table table-hover table-nomargin table-condensed" style="width:50%"><tr><td>
 <p> <input type="text" name="txtVerifica" placeholder="Inserisci il titolo dell'attivit&agrave;" size="60" class="input-xxlarge"><br></p></td>
 <tr><td><select name="txtTipoVoto">
			<option selected value="S">Scritto</option>
            <option value="O">Orale</option>
            <option value="P">Pratico</option>
	 
	</select>
 <input type="text" name="txtData" value="Data" size="10" class="input-small"></td></tr>
 <tr><td>Per default il voto viene registra solo in Classifica (oppure...) <br>
 <input type="checkbox"  name="cbScrutini" title="Selezionare se il voto deve contribuire anche al calcolo della media per lo scrutinio"> 
 <b> Registra in Classifica ed in Scrutini </b>&nbsp;&nbsp;&nbsp;<br>
 <input type="checkbox"  name="cbClassifica" title="Selezionare se il voto deve contribuire solo calcolo della media per lo scrutinio">
  <b> Registra solo in Scrutini </b>&nbsp;&nbsp;&nbsp;</td></tr>
 </table>
 <p><input type="submit" value="Aggiorna" name="B1"><input type="reset" value="Azzera" name="B2"></p> <!--Definisce i due bottoni del form -->
</center>
 </p></FIELDSET>
</form> 
 
 <!-- INIZIO PARTE INSERIMENTO KAHOOT-->
 
 	<div class="box-title">
		<h3>
		<i class="icon-reorder"></i>
			Inserisci punteggi  kahoot quiz
		</h3>							 
	</div>
	<div class="box-content">
	    <form method="POST" class="form-vertical" name="dati" action="inserisci_kahoot_rapido.asp?umanet=<%=umanet%>&Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&cartella=<%=cartella%>&divid=<%=divid%>&posizione=<%=posizione%>" >
     
        
		<div class="control-group">                                        
			<label class="control-label"><b>Titolo del quiz</b></label></div>
			<div class="controls"><input type="text" name="txtTitolo" value="" class="input-large"> </div>
		</div>	
        <div class="control-group">                                        
			<label class="control-label"><b>Data</b></label></div>
			<input type="text" name="txtData" value="Data" size="10" class="input-small">
		</div>	
        <span class="alert-info"> 
        	<b>Incolla CodiceAllievo Punti (una coppia su ogni riga)</b>  
        </span>
		<div class="control-group"> 
           <textarea name="MyTextArea" rows=8 cols=70 class="input-block-level" placeholder="vedi test4_modrapid.asp" ></textarea> 
		</div>
        <p style ="text-align:center"><input class="btn btn-primary" type="submit" value="Invia" name="B2"  rel="tooltip" title="inserisci in blocco"></p>
     </form>
	</div>
 
      
      <%end if%>
      
      <%
	  rsTabella.close
	  ConnessioneDB.close
	  %>
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			   <div class="box-content"> 
                     
                      
                      
              
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

 </html>

