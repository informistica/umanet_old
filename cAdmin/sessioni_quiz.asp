<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Sessioni QUIZ</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	
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
    <!-- Datepicker --> 

<!-- <script src="../js/plugins/datepicker/bootstrap-datepicker.it.js"></script> -->
  
  <script src="../js/jquery-ui.js"></script>   
 <script src="../js/datapicker_it.js"></script> 




  


   
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
						<h1> <i class="icon-question-sign"></i> Sessioni Quiz</h1> 
                    
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
                         Sessioni precedenti
						 
                         </h3>
			          </div>
				      <div class="box-content">
                      
 <% 

  '  AggiungiSessione=Request.QueryString("AggiungiSessione") ' <>"" se devo aggiungere nuovi crediti, metodo alternativo all'aggiunta dalla classifica
	id_classe=Session("Id_Classe") 
	Id_Mod=request.querystring("Id_Mod")
	byUmanet=Request.QueryString("byUmanet")
	'if AggiungiSessione="" then
		'	Id_Eser=Request.QueryString("ID_ESER")
			QuerySQL="SELECT count(*) FROM [dbo].[2SESSIONI_QUIZ] Where Id_Classe='"& id_classe &"' ; "
			'response.write(QuerySQL)
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			numSessioni=rsTabella(0)
			QuerySQL="SELECT * FROM  [dbo].[2SESSIONI_QUIZ] Where Id_Classe='"& id_classe &"' order by  ID_Sessione  "
			Set rsTabella = ConnessioneDB.Execute(QuerySQL)
			
			'response.write(QuerySQL)
		'	dataProva=rsTabella.fields("Data")
		'	Response.Write("<center><font color=red ><b> '"&rsTabella.fields("Descrizione")&"' del  " & rsTabella.fields("Data") & "</font></center></b>")
			%>
			
			 
			<table  class="table table-hover table-nomargin table-condensed" style="width:75%">
			  <tr><td><b>ID</b></td><td><b>Titolo</b></td><td><b>Data</b></td><td><b>Tipo</b></td><td><b>Convalidata</b></td> </tr>
			<%i=0
			do while not rsTabella.eof %>
		
				<tr><td><%=rsTabella.fields("ID_Sessione")%></td><td><%=rsTabella.fields("Titolo")%></td>
				<td><%=rsTabella.fields("Data")%></td>
                <td>
				<% select case(rsTabella.fields("TipoQuiz")) 
				case 0:
				  response.write("Vero/Falso")
				  case 1: 
				   response.write("Singola")
				   case 2: 
				    response.write("Multipla")
				end select
				%>
                </td>
              <td>
				<% select case(rsTabella.fields("Convalidata")) 
				case 0:
				  response.write("No")
				  case 1: 
				   response.write("Si")
				   
				end select
				%>
                </td>
				 
				</td> </tr>
			
			<%
			 i=i+1
			 rsTabella.movenext
			loop%>
   </table>
      </div>
		 <div class="box-title">
				        <h3> <i class="icon-reorder"></i> 
                         Crea nuova sessione
						 
                         </h3>
			          </div>
        
        	<br>   
			<form name="AggiornaSessioni" class="form-horizontal" action="inserisci_sessione_quiz.asp" method="post">
		   

<table  class="table table-hover table-nomargin table-condensed" style="width:75%"><tr><td>
 <p> 
 <input type="hidden" value="<%=Id_Mod%>" name="txtId_Mod">
 <input type="hidden" value="<%=Id_Arg%>" name="txtId_Arg">
 <input type="text" name="txtSessione" placeholder="Inserisci il titolo dell'attivit&agrave;" size="50" class="input-xxlarge"><br></p></td>
 <tr><td><select name="txtTipo">
			<option selected value="0">Vero/Falso</option>
            <option value="1">Singola</option>
            <option value="2">Multipla</option>
	 
	</select>
      Data <input type="text" name="txtData" value="<%=FormatDateTime(now(),2) %>" id="datepicker1" class="input-medium datepick"/>
               
   <tr><td colspan="2">
      <iframe src="../cMessaggi/compilapreavviso.asp?byUmanet=<%=byUmanet%>" name="postmessage" id="postmessage" width="100%" height="60%" frameborder="0" SCROLLING="si" border="0" class="iframe">
      </iframe>
     </td></tr>            
  
 </table>
 <p><input type="submit" class="btn-primary" value="Crea" name="B1"></p> <!--Definisce i due bottoni del form -->
</center>
 </p> 
</form> 
 
 
 
 
 
      
      <%'end if%>
      
      <%
	  rsTabella.close
	  ConnessioneDB.close
	  %>
 									 
	 
	 
				 
				 
                      
                      
                      
                   
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         

			 
	</body>

 </html>

