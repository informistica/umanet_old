<%@ Language=VBScript %>

<% if session("DB") <> "1" then
	Response.Redirect "../../home.asp"
	end if
	
%>	

<!doctype html>
<html>
<head>
   
   <title>Visualizzazione Risultato Finale</title>   
	
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
    
	
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h3> <i class="icon-comments"></i> Quiz Legalità </h3> 
                    
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
							<a href="#">Admin</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Gestione Quiz Legalità</a>
                           
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
				        <h3> <i class="icon-reorder"></i>  RISULTATI SESSIONE <%=request.querystring("id")%></h3>
			          </div>
				      <div class="box-content">
                     		 	 <%
	                     		 	 
	                     		 	
	                     		 	 
								    
								
	                     		 	 
								QuerySQL = "SELECT * FROM Leg_Sessioni WHERE id = "&request.querystring("id")
								set rsSessioni = ConnessioneDB.Execute(QuerySQL)
								'response.write(QuerySQL)
								
								i=0
								do while not rsSessioni.EOF

								valore=rsSessioni("valore")
								nome=rsSessioni("nome")
								s=Split(valore, ",")
								
								i=i+1
								
								 rsSessioni.movenext
								loop
								
								'response.write(s(0))
								'response.write(UBound(s)-1)
								
								 Dim risultati()
								ReDim risultati(UBound(s)-1)
								Dim squadre()
								ReDim squadre(UBound(s)-1)
								Dim somma
								somma = 0
											
								For i = 1 To UBound(s)
									'response.write(i)
									'response.write(s(i))
									risultati(i-1) = cInt(rtrim(s(i)))
									squadre(i-1) = i
									somma = somma+cInt(rtrim(s(i)))
								Next						
								
								
								
								'response.write(s(1))
								'response.write(risultati(3))
' 								for i = UBound(risultati) - 1 To 0 Step -1
' 								        for j= 0 to i
' 								            if risultati(j)>risultati(j+1) then
' 								                temp=risultati(j+1)
' 								                'response.write(temp)
' 								                temp2=squadre(j+1)
' 								                risultati(j+1)=risultati(j)
' 								                squadre(j+1)=squadre(j)
' 								                risultati(j)=temp
' 								                squadre(j)=temp2
' 								            end if
' 								        next
' 								    next
								
								
								MyArray=risultati
								max=ubound(MyArray)
								
								For i=0 to max  
								   For j=i+1 to max  
								      if MyArray(i)<MyArray(j) then 
								          TemporalVariable=MyArray(i) 
								          TemporalVariable2=squadre(i)
								          MyArray(i)=MyArray(j)
								          squadre(i)=squadre(j) 
								          MyArray(j)=TemporalVariable 
								          squadre(j)=TemporalVariable2
								     end if 
								   next  
								next 
								
								Response.write("<script>vett = Array();vett2=Array();")
								For i=0 to max  
								  Response.write ("vett["&i&"] = "&((MyArray(i)*100)/somma)&";")
								    Response.write ("vett2["&i&"] = "&MyArray(i)&";")
								next 
								response.write("</script>")
								
								%>
								
								<style>
								.rating-histogram {
								  padding: 5px 5px 5px 44px;
								  text-align: left;
								  width: 100%;
								}
								
								.rating-histogram .rating-bar-container {
								  position: relative;
								  margin-bottom: 5px;
								}
								
								.rating-histogram .bar-label {
								  margin-right: 10px;
								  position: absolute;
								  left: -44px;
								  top: 2px;
								}
								
								.rating-histogram .bar-label:before {
/* 								  content: "\f005"; */
								  font-family: FontAwesome;
								  font-size: 16px;
								  line-height: 16px;
								  height: 16px;
								  width: 16px;
								  color: #ccc;
								  display: inline-block;
								  margin-right: 5px;
								  vertical-align: middle;
								}
								
								.rating-histogram .bar {
								  background-color: #8e70af;
								  background-image: -webkit-linear-gradient(left, #8e70af, #48bfed);
								  background-image: linear-gradient(to right, #8e70af, #48bfed);
								  -webkit-transition: width 2s ease;
								  -moz-transition: width 2s ease;
								  transition: width 2s ease;
								  opacity: .8;
								  display: inline-block;
								  vertical-align: middle;
								  width: 1%;
								  max-width:87%;
								  height: 60px;
								  margin-right: 3px;
								}
								
								.rating-histogram .bar-number {
								  font-size: 14px;
								  line-height: 1;
								  vertical-align: middle;
								}
								
								.hidden {
								  display: none;
								}
								</style>
								
								<% Dim v()
									
								ReDim v(10)
								v(1) = "one"
								v(2) = "two"
								v(3) = "three"
								v(4) = "four"
								v(5) = "five"
								v(6) = "six"
								v(7) = "seven"
								v(8) = "eight"
								v(9) = "nine"
								v(10) = "ten"
								
								%>
								
								<div class="rating-histogram">
								<% For i=0 To Ubound(risultati) %>

								
								  <div class="rating-bar-container <%=v(i+1)%>" data-id="<%=(i+1)%>">
								    <span class="bar-label" style="font-size:20px"> <br><b><%=(i+1)%></b> </span>
								    <span class="bar" style="color:white; font-size: 20px"> <br>&nbsp;&nbsp;<b>Squadra <%=squadre(i)%></b> </span>
								    <span class="bar-number"></span>
								  </div>
								  
								  <% Next %>
								  
<!--
								  <div class="rating-bar-container four" data-id="4">
								    <span class="bar-label"> 4 </span>
								    <span class="bar"></span>
								    <span class="bar-number"></span> 
								  </div>
								  <div class="rating-bar-container three" data-id="3">
								    <span class="bar-label"> 3 </span>
								    <span class="bar"></span>
								    <span class="bar-number"></span>
								  </div>
								  <div class="rating-bar-container two" data-id="2">
								    <span class="bar-label"> 2 </span>
								    <span class="bar"></span>
								    <span class="bar-number"></span> 
								  </div>
								  <div class="rating-bar-container one" data-id="1">
								    <span class="bar-label"> 1 </span>
								    <span class="bar"></span>
								    <span class="bar-number"></span>
								  </div>
-->
								</div><!-- /rating-histogram -->
								
								<div class="hidden"><!-- needs for jquery calculations -->
								  <form>
								    <input type="text" class="reviews_1star" value="5">
								    <input type="text" class="reviews_2star" value="4">
								    <input type="text" class="reviews_3star" value="2">
								    <input type="text" class="reviews_4star" value="6">
								    <input type="text" class="reviews_5star" value="3">
								    <input type="text" class="reviews_6star" value="4">
								    <input type="text" class="reviews_7star" value="2">
								    <input type="text" class="reviews_8star" value="1">
								    <input type="text" class="reviews_9star" value="8">
								    <input type="text" class="reviews_10star" value="6">
								  </form>
								</div><!-- /hidden -->
								 
								 
								 <script>
									 $(function() {
								  var stars = new Array();
								  var sum = 0;
								  var width = new Array();
								
								  for ( var i = 1; i < vett.length+1; i++ ) {
								    stars.push(parseInt($('.reviews_'+i+'star').val()));
								  }     
								
								  for ( var i = 0; i < stars.length; i++ ) {
								    sum += stars[i];       
								  }     
								
								  for ( var i = 0; i < stars.length; i++ ) {
								    w = vett[i];
									
								    width.push(w);
								    $('.rating-bar-container[data-id='+(i+1)+'] .bar').css('width', w+'%' ); 
								  }
								
								  if (sum > 0) {
								    for ( var i = 0; i < stars.length; i++ ) {
								      $('.rating-bar-container[data-id='+(i+1)+'] .bar-number').html('P. <b>'+vett2[i]+'</b>'); 
								    }
								  } else{
								    $(".rating-bar-container .bar-number").html('0%')
								  }
								});
								</script>
								
								
								
								
							
						
                      
			        
			        
			        
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
		
		<form id="mod" action="cSns/modificasessione.asp" method="post">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Modifica Sessione</h3>
				</div>
				<div class="modal-body">
					<b>Titolo</b><br><input name="titolomodifica" id="titolomodifica" type="text" class="input-xxlarge" style="width: 97%"><br>
					<b>Privata (0/1)</b><br><input name="tipomodifica" id="tipomodifica" type="text" class="input-xxlarge" style="width: 97%">
					<br>
					<b>Chiave (se privata)</b><br><input name="chiavemodifica" id="chiavemodifica" type="text" class="input-xxlarge" style="width: 97%"><br><br>
				</div>
				<div class="modal-footer">
					<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
					<button type="button" id="inviamodifica" class="btn btn-primary" onclick="controllamodifica()">Invia</button>
				</div>
			</div>
		</form>
		
		
		
		<script>
		
		function modifica(id, titolo, tipo, chiave){
			document.getElementById("titolomodifica").value=titolo;
			document.getElementById("tipomodifica").value=tipo;
			document.getElementById("chiavemodifica").value=chiave;
			document.getElementById("mod").action="cSns/modificasessione.asp?id="+id;
		}
		
		function controllamodifica(){
			var titolo = document.getElementById("titolomodifica").value.trim();
			var tipo = document.getElementById("tipomodifica").value;
			
			if(titolo == ""){
				alert("Il nome della sessione è obbligatorio");
			}else if(tipo != "0" && tipo != "1"){
				alert("Il tipo può essere 0 oppure 1");
			}else{
				document.getElementById("inviamodifica").type="submit";
			}
			
		}
		
		function controlloinvio(){
			
			var testo = document.getElementById("nomesessione").value;
			var tipo = document.getElementById("tiposessione").value;
			
			if(testo.trim()==""){
				alert("Il nome della sessione è obbligatorio");
			}else if(tipo != "0" && tipo != "1"){
				alert("Il tipo può essere 0 oppure 1");
			}else{
				document.getElementById("inviosess").type="submit";
			}
			
		}
		
		function aggiungiSessione(){
			nomesess = document.getElementById('txtNome').value.trim();
			nsquadre = document.getElementById('nsquadre').value.trim();
			
			if(!nomesess){
				alert('Nome della sessione obbligatorio');
			}else if(nsquadre < 0 && nsquadre > 100){
				alert('Il numero di squadre deve essere un intero');
			}else{
				document.getElementById('addSessione').type="submit";
			}
		}
		
		function chiudiSessione(){
			var stato = confirm("Sei sicuro di voler chiudere la sessione?");
			
			if(stato){
				window.location.href="cLegalita/chiudisessione.asp";
			}
		}
		
		</script>
		
	</body>

 </html>

