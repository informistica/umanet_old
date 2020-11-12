<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
	<title>Mapbook</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 <meta charset="UTF-8">

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
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
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
		<script src="js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	 <!--<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>-->
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">  

	<div id="navigation">
     
        <%
		
		' connessione al database e inclusione dei menu
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
						<h1> <i class="glyphicon-snowflake"></i> Quaderno delle Mappe </h1> 
                    
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
							<a href="#">Quaderno Mappe</a>
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
				        <h3> <i class="icon-reorder"></i>  Quaderno delle Mappe
                        
                         </h3>
			          </div>
				      <div class="box-content">
						
						
						<h4>Mappe delle tue classi</h4>
						
					<div class="accordion accordion-widget" id="accordionContenitore">
								
								
							
					  
					  <%
					  
					  Classe = Request.QueryString("classe")
					  classediv = ""
					  
					  if Session("Admin") = false then
					  
							  QuerySQL = "SELECT * FROM AssociazioniAllievi WHERE CodiceAllievo = '"&Session("CodiceAllievo")&"' OR UtenteAssociato =  '"&Session("CodiceAllievo")&"'"
							  
							  'response.write(QuerySQL)
							  set rsTab = ConnessioneDB.Execute(QuerySQL)
							  
							  if rsTab.EOF then
							  
							  k=0
							  
									 QuerySQL = "SELECT DISTINCT Classe FROM Allievi WHERE CodiceAllievo = '"&Session("CodiceAllievo")&"'"
										  'response.write(QuerySQL)
										  set rsTab2 = ConnessioneDB.Execute(QuerySQL)
										  
											  do while not rsTab2.EOF
											  
											  classediv = classediv&rsTab2("Classe")&","
											  
													  QuerySQL = "SELECT Titolo,ID_Mod FROM Moduli WHERE ID_Mod like '%"&rsTab2("Classe")&"%' AND ID_Mod IN (SELECT DISTINCT Id_Mod FROM Nodi WHERE Cartella = '"&rsTab2("Classe")&"') order by Posizione;"
													  'response.write QuerySQL
													  set rsTabella = ConnessioneDB.Execute(QuerySQL)
																
																
																
															  do while not rsTabella.EOF
															  
															  'response.write "<tr><td>"&rsTab2("Classe")&" - "&rsTabella("Titolo")&"</td><td><center><input id='"&rsTabella("Id_Mod")&"' onchange='associa("""&rsTabella("Id_Mod")&""")' value='"&rsTabella("Id_Mod")&"' type='checkbox'></center></td></tr>"
															  
															  %>
															  
															  <div class="accordion-group">
																<div class="accordion-heading">
																	<a id="cap<%=k%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionContenitore" style="text-decoration:none" href="#c<%=k%>">
																		<%=rsTab2("Classe")&" - "%><b><%=rsTabella("Titolo")%></b>
																	</a>
																</div>
																<div id="c<%=k%>" class="accordion-body collapse">
																	<div class="accordion-inner">
																	<br>
																	<table class="table table-hover table-nomargin table-bordered ">
																	<thead>
																		<th>Paragrafo</th>
																		<th>Seleziona</th>
																	</thead>
																	<tbody>
																		<tr>
																			<td style="width:90%"><b>Tutto il capitolo</b></td>
																			<td><center><input id="<%=rsTabella("Id_Mod")%>" onchange='associa("<%=rsTabella("Id_Mod")%>","C")' value='<%=rsTabella("Id_Mod")%>' type='checkbox'></center></td>
																		</tr>	
																		
																		<%
																		
																		QuerySQL = "SELECT * FROM Moduli_Paragrafi WHERE ID_Mod = '"&rsTabella("Id_Mod")&"' order by Posizione"
																		set rsTabPar = ConnessioneDB.Execute(QuerySQL)
																		
																		
																		do while not rsTabPar.EOF
																		
																		%>
																		
																		<tr>
																		<td style="width:90%"><%=rsTabPar("Paragrafo")%></td>
																		<td><center><input id="<%=rsTabPar("ID_Paragrafo")%>" onchange='associa("<%=rsTabPar("ID_Paragrafo")%>","P")' value='<%=rsTabPar("ID_Paragrafo")%>' type='checkbox'></center></td>
																		</tr>
																		
																		<%
																		
																		
																		rsTabPar.movenext
																		loop
																		
																		
																		%>
																		
																		
																		 </tbody>
																	</table>
																	</div>
																</div>
															</div>
															  
															  <%
															  
															  
															  k=k+1
															  
															  rsTabella.movenext
															  loop
											  
											  rsTab2.movenext
											  loop
							  
							  else
							  
							  k=0
							  
									  do while not rsTab.EOF
									  
										
									  
										  QuerySQL = "SELECT DISTINCT Classe FROM Allievi WHERE CodiceAllievo = '"&rsTab("CodiceAllievo")&"' OR CodiceAllievo = '"&rsTab("UtenteAssociato")&"'"
										  'response.write(QuerySQL)
										  set rsTab2 = ConnessioneDB.Execute(QuerySQL)
										  
											  do while not rsTab2.EOF
											  
											  classediv = classediv&rsTab2("Classe")&","
											  
													  QuerySQL = "SELECT Titolo,ID_Mod FROM Moduli WHERE ID_Mod like '%"&rsTab2("Classe")&"%' AND ID_Mod IN (SELECT DISTINCT Id_Mod FROM Nodi WHERE Cartella = '"&rsTab2("Classe")&"') order by Posizione;"
													  'response.write QuerySQL
													  set rsTabella = ConnessioneDB.Execute(QuerySQL)
															
															
															
															  do while not rsTabella.EOF
															  
															  'response.write "<tr><td>"&rsTab2("Classe")&" - "&rsTabella("Titolo")&"</td><td><center><input id='"&rsTabella("Id_Mod")&"' onchange='associa("""&rsTabella("Id_Mod")&""")' value='"&rsTabella("Id_Mod")&"' type='checkbox'></center></td></tr>"
															  
															  %>
															  
															   <div

															   class="accordion-group">
																<div class="accordion-heading">
																	<a id="cap<%=k%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionContenitore" style="text-decoration:none" href="#c<%=k%>">
																		<%=rsTab2("Classe")&" - "%><b><%=rsTabella("Titolo")%></b>
																	</a>
																</div>
																<div id="c<%=k%>" class="accordion-body collapse">
																	<div class="accordion-inner">
																	<br>
																	<table class="table table-hover table-nomargin table-bordered ">
																	<thead>
																		<th>Paragrafo</th>
																		<th>Seleziona</th>
																	</thead>
																	<tbody>
																		<tr>
																			<td style="width:90%"><b>Tutto il capitolo</b></td>
																			<td><center><input id="<%=rsTabella("Id_Mod")%>" onchange='associa("<%=rsTabella("Id_Mod")%>","C")' value='<%=rsTabella("Id_Mod")%>' type='checkbox'></center></td>
																		</tr>	
																		
																		<%
																		
																		QuerySQL = "SELECT * FROM Moduli_Paragrafi WHERE ID_Mod = '"&rsTabella("Id_Mod")&"' order by Posizione"
																		set rsTabPar = ConnessioneDB.Execute(QuerySQL)
																		
																		
																		do while not rsTabPar.EOF
																		
																		%>
																		
																		<tr>
																		<td style="width:90%"><%=rsTabPar("Paragrafo")%></td>
																		<td><center><input id="<%=rsTabPar("ID_Paragrafo")%>" onchange='associa("<%=rsTabPar("ID_Paragrafo")%>","P")' value='<%=rsTabPar("ID_Paragrafo")%>' type='checkbox'></center></td>
																		</tr>
																		
																		<%
																		
																		
																		rsTabPar.movenext
																		loop
																		
																		
																		%>
																		
																		
																		 </tbody>
																	</table>
																	</div>
																</div>
															</div>
															  
															  <%
															  
															  k=k+1
															  
															  rsTabella.movenext
															  loop
											  
											  rsTab2.movenext
											  loop
									  
									  rsTab.movenext
									  loop
							  
							  end if
					  
					  else
					  
					  'response.write("admin")
					  
					  QuerySQL = "SELECT DISTINCT Classe FROM Allievi WHERE Classe <> 'Expo'"
						  'response.write(QuerySQL)
						  set rsTab2 = ConnessioneDB.Execute(QuerySQL)
						  
							  do while not rsTab2.EOF
							  
							  classediv = classediv&rsTab2("Classe")&","
							  
									  QuerySQL = "SELECT Titolo,ID_Mod FROM Moduli WHERE ID_Mod like '%"&rsTab2("Classe")&"%' AND ID_Mod IN (SELECT DISTINCT Id_Mod FROM Nodi WHERE Cartella = '"&rsTab2("Classe")&"') order by Posizione;"
									  'response.write QuerySQL
									  set rsTabella = ConnessioneDB.Execute(QuerySQL)
									  
											  do while not rsTabella.EOF
															  
															  'response.write "<tr><td>"&rsTab2("Classe")&" - "&rsTabella("Titolo")&"</td><td><center><input id='"&rsTabella("Id_Mod")&"' onchange='associa("""&rsTabella("Id_Mod")&""")' value='"&rsTabella("Id_Mod")&"' type='checkbox'></center></td></tr>"
															  
															  %>
															  
															   <div class="accordion-group">
																<div class="accordion-heading">
																	<a id="cap<%=k%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionContenitore" style="text-decoration:none" href="#c<%=k%>">
																		<%=rsTab2("Classe")&" - "%><b><%=rsTabella("Titolo")%></b>
																	</a>
																</div>
																<div id="c<%=k%>" class="accordion-body collapse">
																	<div class="accordion-inner">
																	<br>
																	<table class="table table-hover table-nomargin table-bordered ">
																	<thead>
																		<th>Paragrafo</th>
																		<th>Seleziona</th>
																	</thead>
																	<tbody>
																		<tr>
																			<td style="width:90%"><b>Tutto il capitolo</b></td>
																			<td><center><input id="<%=rsTabella("Id_Mod")%>" onchange='associa("<%=rsTabella("Id_Mod")%>","C")' value='<%=rsTabella("Id_Mod")%>' type='checkbox'></center></td>
																		</tr>	
																		
																		<%
																		
																		QuerySQL = "SELECT * FROM Moduli_Paragrafi WHERE ID_Mod = '"&rsTabella("Id_Mod")&"' order by Posizione"
																		set rsTabPar = ConnessioneDB.Execute(QuerySQL)
																		
																		
																		do while not rsTabPar.EOF
																		
																		%>
																		
																		<tr>
																		<td style="width:90%"><%=rsTabPar("Paragrafo")%></td>
																		<td><center><input id="<%=rsTabPar("ID_Paragrafo")%>" onchange='associa("<%=rsTabPar("ID_Paragrafo")%>","P")' value='<%=rsTabPar("ID_Paragrafo")%>' type='checkbox'></center></td>
																		</tr>
																		
																		<%
																		
																		
																		rsTabPar.movenext
																		loop
																		
																		
																		%>
																		
																		
																		 </tbody>
																	</table>
																	</div>
																</div>
															</div>
															  
															  <%
															  
															  k=k+1
															  
															  rsTabella.movenext
															  loop
							  
							  rsTab2.movenext
							  loop
							  
					  
					  end if
					  
					  
					  %>
					  
				
					
					</div>
					
					
					<% if Session("Admin") = false then %>
					
					<br>
					<h4>Mappe delle altre classi</h4>
					
					
					<div class="accordion accordion-widget" id="accordionContenitore2">
					
					<%
					
					classediv = split(Left(classediv,Len(classediv)-1),",")
					 QuerySQL = "SELECT DISTINCT Classe FROM Allievi WHERE ("
					 i=0
					do while not i = UBound(classediv)+1
					
					QuerySQL = QuerySQL & " Classe <> '" & classediv(i) & "' AND "
					
					'if i <> UBound(classediv) then
					
					'QuerySQL = QuerySQL & " AND "
					
					'end if
					
					i=i+1
					loop
					
					QuerySQL = QuerySQL & " Classe <> 'Expo')"
					
						  'response.write(QuerySQL)
						  set rsTab2 = ConnessioneDB.Execute(QuerySQL)
						  
							  do while not rsTab2.EOF
							  
									  QuerySQL = "SELECT Titolo,ID_Mod FROM Moduli WHERE ID_Mod like '%"&rsTab2("Classe")&"%' AND ID_Mod IN (SELECT DISTINCT Id_Mod FROM Nodi WHERE Cartella = '"&rsTab2("Classe")&"') order by Posizione;"
									  'response.write QuerySQL
									  set rsTabella = ConnessioneDB.Execute(QuerySQL)
									  
											  do while not rsTabella.EOF
															  
															  'response.write "<tr><td>"&rsTab2("Classe")&" - "&rsTabella("Titolo")&"</td><td><center><input id='"&rsTabella("Id_Mod")&"' onchange='associa("""&rsTabella("Id_Mod")&""")' value='"&rsTabella("Id_Mod")&"' type='checkbox'></center></td></tr>"
															  
															  %>
															  
															   <div class="accordion-group">
																<div class="accordion-heading">
																	<a id="cap<%=k%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionContenitore2" style="text-decoration:none" href="#c<%=k%>">
																		<%=rsTab2("Classe")&" - "%><b><%=rsTabella("Titolo")%></b>
																	</a>
																</div>
																<div id="c<%=k%>" class="accordion-body collapse">
																	<div class="accordion-inner">
																	<br>
																	<table class="table table-hover table-nomargin table-bordered ">
																	<thead>
																		<th>Paragrafo</th>
																		<th>Seleziona</th>
																	</thead>
																	<tbody>
																		<tr>
																			<td style="width:90%"><b>Tutto il capitolo</b></td>
																			<td><center><input id="<%=rsTabella("Id_Mod")%>" onchange='associa("<%=rsTabella("Id_Mod")%>","C")' value='<%=rsTabella("Id_Mod")%>' type='checkbox'></center></td>
																		</tr>	
																		
																		<%
																		
																		QuerySQL = "SELECT * FROM Moduli_Paragrafi WHERE ID_Mod = '"&rsTabella("Id_Mod")&"' order by Posizione"
																		set rsTabPar = ConnessioneDB.Execute(QuerySQL)
																		
																		
																		do while not rsTabPar.EOF
																		
																		%>
																		
																		<tr>
																		<td style="width:90%"><%=rsTabPar("Paragrafo")%></td>
																		<td><center><input id="<%=rsTabPar("ID_Paragrafo")%>" onchange='associa("<%=rsTabPar("ID_Paragrafo")%>","P")' value='<%=rsTabPar("ID_Paragrafo")%>' type='checkbox'></center></td>
																		</tr>
																		
																		<%
																		
																		
																		rsTabPar.movenext
																		loop
																		
																		
																		%>
																		
																		
																		 </tbody>
																	</table>
																	</div>
																</div>
															</div>
															  
															  <%
															  
															  k=k+1
															  
															  rsTabella.movenext
															  loop
							  
							  rsTab2.movenext
							  loop
					
					
					%>
					
					
					</div>
					
					<% end if %>
					
					<center>
					<br><br>
					<input type="button" class="btn btn-primary" value="Visualizza SuperMappa" onclick="visualizzamappa()">
					
					</center>
	
               
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         
<script>

var moduli = Array();
var i = 0;
var paragrafo = false;
var capitolo = false;

function associa(id,tipo){

if(tipo=="C"){
	
	if(!paragrafo){
		
		for(var j=1; j<30; j++){
					
			if(document.getElementById(id+"_"+j) === undefined || document.getElementById(id+"_"+j) === null){
			}else{
				if(document.getElementById(id+"_"+j).checked == true){
					capitolo = true;
					document.getElementById(id+"_"+j).click();
				}
			}
			
		}
		
	}else{
		paragrafo = false;
	}
	
}else{
	
	if(!capitolo){
		var ids = id.split("_");
		var id2 = ids[0]+"_"+ids[1];
		
		if(document.getElementById(id2).checked == true){
			paragrafo = true;
			document.getElementById(id2).click();
		}
	}else{
		capitolo = false;
	}

}

if(document.getElementById(id).checked == true){
//alert("benvenuto");
moduli[i] = id;
i++;
}else{
//alert("arrivederci");
var x = moduli.indexOf(id);
moduli[x] = "null";
}

}

function visualizzamappa(){

var cartelle = ""
var moduli1 = ""
var cont = 0;

for(var i=0; i<moduli.length; i++){

if(moduli[i] != "null"){
moduli1 += moduli[i]+",";
cartella = moduli[i].split("_");
cartelle += cartella[0]+",";
cont++;
}

}

if(cont==0){
alert("Seleziona almeno un paragrafo");
}else{

var url = "https://www.umanetexpo.net/expo2015Server/UECDL/script/cMap/spiegazione_mappa.asp?super=1&cod=<%=Session("CodiceAllievo")%>&Cartella="+cartelle.substr(0,cartelle.length-1)+"&Modulo="+moduli1.substr(0,moduli1.length-1);
window.open(url,"_blank");

}

}


</script>
			 
	</body>

 </html>

