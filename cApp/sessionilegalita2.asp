<%@ Language=VBScript %>

<% 
inizio=request.querystring("inizio") ' ha valore se provengo da iisvittuone.it/cpl serve per inviare la mail
byemail=request.querystring("byemail")  'ha volore se provengo dalla mail ricevuta dopo la creazione partita serve per caricare tutti i dati del contatto
mail=request.querystring("mail")
scuola=request.querystring("scuola")
nome=request.querystring("nome")
nsquadre=request.querystring("nsquadre")
ndomande=request.querystring("ndomande")
id_app=request.querystring("id_app")   ' 1= quiz legalità; 2= quiz per cnv
if id_app="" then
id_app=1 
end if
session("db")=1
limite_sessioni=10 ' massimo numero di sessioni ammesse per un contatto
disponibili=1
'hanno valore quanto torno da chiudisessione.asp
ritorno=request.querystring("ritorno")
id_contatto=request.querystring("id_contatto")
	
	
%>	

<!doctype html>
<html>
<head>
   
   <title>Gestione App Quiz</title>   
	
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
function showText2() {window.alert("I parametri inviati non sono completi!")
location.href="https://www.iisvittuone.it/cpl/admin.php"
//location.href=window.history.back();
 }
    </script>
    <script type="text/javascript" src="../js/selezionatutti.js"></script>
    

     
  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<!-- #include file = "../cAdmin/include_mail.asp" -->
<%
  Response.Buffer = true
  'On Error Resume Next 

  if byemail<>"" then ' serve per passare il controllo, i valori vengono caricati dopo.
     scuola="scuola"
	 nome="nome"
	 nsquadre="1"
     mail="a@a.it" 
  end if
  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (mail="" or scuola="" or nsquadre="" or nome="") and (ritorno="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
	
     <body class='theme-blue' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  <% end if %>


	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		<!-- #include file = "../include/navigation_cpl.asp" -->
       
         <%
		 if ritorno<>"" and byemail<>"" then
		    QuerySQL = "SELECT email FROM Leg_Contatti where id='"&id_contatto&"'"
			set rsSessioni = ConnessioneDB.Execute(QuerySQL)
			mail=rsSessioni(0)
		 end if
		 %> 
         
	</div>
    
	
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left_cpl.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
					<% Select Case id_app
					  Case 1
						response.write("<h3> <i class='icon-comments'></i> Quiz Legalità </h3>")
					  Case 2
						response.write("<h3> <i class='icon-comments'></i> Quiz CNV </h3>")
					
					End Select %>
						 
                    
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
							<a href="#">Admin</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#"><%=mail%></a>
                           
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
				        <h3> <i class="icon-reorder"></i>  GESTIONE SESSIONI</h3>
			          </div>
				      <div class="box-content">
                     		 	 <% 
								 
								 
								 if ((ritorno="") and (byemail="")) then ' se non sono stato richiamato da chiudisessione.asp o dalla mail vuol dire che devo inserire
								 
									 QuerySQL = "SELECT * FROM Leg_Contatti where email='"&mail&"'"
									 set rsSessioni = ConnessioneDB.Execute(QuerySQL)
									 if rsSessioni.EOF then
										 ' non esiste il contatto lo aggiungo al db								 
										 QuerySQL = "INSERT INTO Leg_Contatti (email,contatto) VALUES ('"&mail&"', '"&scuola&"');" 
										 ConnessioneDB.Execute(QuerySQL)
										 QuerySQL = "SELECT max(id) FROM Leg_Contatti"
										 set maxId = ConnessioneDB.Execute(QuerySQL)
										 id_contatto=maxId(0)
									else
									i=0
										  
										  id_contatto=rsSessioni("id")
										 ' response.write(id_contatto)
										
									end if
								
									' controllo se il contatto ha raggiunto il limite di sessioni permesse
									
									QuerySQL = "SELECT count(*) FROM Leg_Sessioni where id_contatto="&id_contatto
									set rsSessioni = ConnessioneDB.Execute(QuerySQL)								
									if rsSessioni(0)< limite_sessioni then
										' creo la nuova sessione
										valore = "P,"

										'response.write nsquadre
										QuerySQL = "INSERT INTO Leg_Sessioni (nome, valore,data,id_contatto,ndomande,id_app) VALUES ('1','1','1',1,1,"&id_app&")"
										ConnessioneDB.Execute(QuerySQL)
					
										'devo ottenere l'id della partita come idmax 
										QuerySQL = "SELECT max(id) FROM Leg_Sessioni"
										set rsSessioni = ConnessioneDB.Execute(QuerySQL)
										partita=rsSessioni(0)
																
										i=0
										Randomize()
										While i < CInt(nsquadre)
											QuerySQL = "INSERT INTO Leg_Risultati (squadra, risultato,partita) VALUES ("&(i+1)&", 10,"&partita&")"
											ConnessioneDB.Execute(QuerySQL)
											numero = CInt(Rnd()*100)
											numero = numero+(100*(i+1))
											'response.write numero & "<br>"
											valore = valore & numero
											'response.write valore & "<br>"
											if i < (nsquadre-1) then
												valore = valore & ","
											end if
											'response.write valore & "<br>"
											i = i + 1
										Wend							
										QuerySQL ="UPDATE Leg_Sessioni SET nome = '" & nome & "', valore = '" &valore & "', data = '" & now() & "', id_contatto = " & id_contatto & ", ndomande = " & ndomande & "    WHERE id =" &partita &";"
										response.write(QuerySQL)
										ConnessioneDB.Execute(QuerySQL)
										
										'invio la mail con i dati
										mes = ""
										IsSuccess = false
										sMailServer ="mail.iisvittuone.it"
										Select Case id_app
										  Case 1
											sFrom = "CPL del Magentino <noreply@iisvittuone.it>"
											sSubject = "Quiz challenge sulla Legalità"
											linkAdmin="https://www.umanetexpo.net/expo2015Server/UECDL/script/cApp/sessionilegalita2.asp?byemail=1&id_contatto="&id_contatto&"&id_app="&id_app							
											linkGioca="https://www.iisvittuone.it/cpl"
											sBody=""
											sBody= sBody &"<center><img src='https://www.elexpo.net/archivio/img/CPL_small.jpg' /></center><br>"
											sBody = sBody & "<center><b><h3>Quiz challenge sulla Legalità</h3></b></center><br>"
											sBody = sBody & "Ecco i dati per accedere alla partita appena creata:<br><br>"
											cartella_partita="cpl"
										  Case 2
											sFrom = "Centro Non Violenza <noreply@iisvittuone.it>"
											sSubject = "Quiz challenge sulla comunicazione nonviolenta"
											linkAdmin="https://www.umanetexpo.net/expo2015Server/UECDL/script/cApp/sessionilegalita2.asp?byemail=1&id_contatto="&id_contatto&"&id_app="&id_app							
											linkGioca="https://www.iisvittuone.it/cnv"
											sBody=""
											sBody= sBody &"<center><img src='https://www.elexpo.net/archivio/img/CNV_small.png' /></center><br>"
											sBody = sBody & "<center><b><h3>Quiz challenge sulla Comunicazione Nonviolenta</h3></b></center><br>"
											sBody = sBody & "Ecco i dati per accedere alla partita appena creata:<br><br>"
											cartella_partita="cnv"
										
										End Select 

										 
										linkAdmin=replace(linkAdmin,"%0D","")
										linkAdmin=replace(linkAdmin,"%20","")
										linkAdmin=replace(linkAdmin," ","")
										linkGioca=replace(linkGioca,"%0D","")
										linkGioca=replace(linkGioca,"%20","")
										linkGioca=replace(linkGioca," ","")
										
										
										
										s=Split(valore, ",") 
										sBody = sBody & "Codice Partita: <b>"&partita&"</b>"&_
										"<br>Nome sessione: <b>"&nome&"</b>"&_
										"<br>Numero squadre: "&Ubound(s)&"<br>"&_
										"<br>Numero domande: "&ndomande&"<br>"
									
										For i = 1 To UBound(s)
										sBody = sBody & "Codice Squadra "&i&": <b>"&rtrim(s(i))&"</b>"
										if i < Ubound(s) then
											sBody = sBody & "<br>"
										end if
										Next
										sBody = sBody & "<br><h4> Indirizzo per accedere al gioco inserendo il codice partita e il codice di squadra : </h4>"
										sBody = sBody &"<img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' />&nbsp;&nbsp;<a title 'Entra nel gioco ' href='"& linkGioca&"'> iisvittuone.it/"&cartella_partita&"</a>"
										sBody = sBody & "<br><h4>Link per il docente:</h4>"										
										sBody = sBody &" <a title 'Entra in Umanet per il CPL ' href='"& linkAdmin&"'> Gestisci le tue sessioni</a>  "
		   	   
										sTo=mail
										TestEMail()
										
									else
									  disponibili=0
									  response.write("<center><span class='error'><h5>Hai raggiunto il limite massimo di sessioni, cancella qualcuna di quelle svolte</h5></span></center>")
									end if
							end if ' if ritorno=""	 
							
							
							QuerySQL = "SELECT * FROM Leg_Sessioni where id_contatto="&id_contatto & " and id_app="&id_app						
							set rsSessioni = ConnessioneDB.Execute(QuerySQL) 
							
							nSess=0							 
							do while not rsSessioni.EOF
								valore=rsSessioni("valore")								
								s=Split(valore, ",")
								if  s(0)="P" then
								  nSess=nSess+1
								end if
								rsSessioni.movenext
							loop
							if nSess=0 then
								errore = "<center>Non ci sono sessioni aperte.</center>"
							end if
							
							valore="R,1,2,3"
							s = Split(valore, ",") 'inizializzo a R per non andare in errore
							
							if nSess>0 then
								QuerySQL = "SELECT * FROM Leg_Sessioni  where id_contatto="&id_contatto & " and id_app="&id_app
								set rsSessioni = ConnessioneDB.Execute(QuerySQL)
								
								i=0
								rsSessioni.movefirst
								do while not rsSessioni.EOF 'and i<1								
								valore=rsSessioni("valore")
								nome=rsSessioni("nome")
								partita=rsSessioni("id")  '***
								ndomande=rsSessioni("ndomande")
								s=Split(valore, ",")
								if  s(0)="P" then  ' visualizzo solo le sessioni aperte
								%>
								
									<center>
									Codice Partita: <b><%=partita%></b>
									<br>Nome sessione: <b><%=nome%></b>
									<br>Numero squadre: <%=Ubound(s)%>
									<br>Numero domande: <%=ndomande%>
									<br><br>
									
									<%
									
									For i = 1 To UBound(s)
									response.write("<b>Squadra "&i&"</b>: "&rtrim(s(i)))
									if i < Ubound(s) then
										response.write "<br>"
									end if
									Next
									
									%>
									
									<br><br><input onclick='chiudiSessione(<%=partita%>,<%=id_contatto%>)' id="chiudiSessione<%=partita%>" type="button" class="btn" value="Chiudi Sessione" style='vertical-align:top'>
									<hr>
									</center>
								
								<%
								end if
								i=i+1							
								 rsSessioni.movenext
								loop
								
							end if
							 								
							 if disponibili<>0 then
								%>	
								<br>
								<center>
								<form action="cLegalita/inseriscisessione.asp?da=1&contatto=<%=id_contatto%>&id_zpp=<%=id_app%>" id="newSession" method='post'>
								<input type="hidden" name="txtContatto" id="txtContatto" value="<%=id_contatto%>"><br>
								<input type="text" name="txtNome" id="txtNome" placeholder="Nome della sessione" class="input-xlarge"><br>
								<input type="text" name="nsquadre" id="nsquadre" placeholder="Numero delle squadre" class="input-xlarge"><br>
								<input type="text" name="ndomande" id="ndomande" placeholder="Numero delle domande" class="input-xlarge"><br>
								
								<input onclick='aggiungiSessione()' id="addSessione" type="button" class="btn" value="Aggiungi Sessione" style='vertical-align:top'>
								</form>
								</center>
							 <%end if%>
							
						<div class="box-content">
							
							
							
							
							</div>
                      
                      </div>
			        </div>
							
					 
							
					
					<div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i> SESSIONI CHIUSE</h3>
			          </div>
				      <div class="box-content">
                     		 	 
				 
						<div class="box-content">					
							
							<% QuerySQL1 = "SELECT * FROM Leg_Sessioni WHERE valore like '%R%' and id_contatto="&id_contatto & "and id_app="&id_app
							set rsSessioniChiuse = ConnessioneDB.Execute(QuerySQL1) %>
							
							<table class="table table-hover table-nomargin">
								<tr>
									<th><b>ID</b></th><th><b>Titolo<b></th><th><b>Data</b></th><th><b>Azione</b></th>
								</tr>
								
								<% do while not rsSessioniChiuse.EOF %>
								
								<tr>
								<td><%=rsSessioniChiuse("id")%></td>
								<td><a href="risultatofinale2.asp?id=<%=rsSessioniChiuse("id")%>&mail=<%=mail%>"><%=rsSessioniChiuse("nome")%></a></td>
								<td><%=rsSessioniChiuse("data")%></td>
								<td>
								<a style="text-decoration:none" href="cLegalita/eliminasessione.asp?da=1&id_contatto=<%=id_contatto%>&id=<%=rsSessioniChiuse("id")%>"><i class="icon-remove"></i></a></td>
								</tr>
								
								<% rsSessioniChiuse.movenext
								loop %>
								
							</table>
							
							
							</div>
                      
                      </div>
			        </div>
			        
			        
			        
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
		
		 
		
		
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
			}else if(nsquadre < 0 && nsquadre > 100) {
				alert('Il numero di squadre deve essere un intero');
			}else{
				document.getElementById('addSessione').type="submit";
			}
		}
		
		function chiudiSessione(codice,id_contatto){
			var stato = confirm("Sei sicuro di voler chiudere la sessione?");
			
			if(stato){
				window.location.href="cLegalita/chiudisessione.asp?da=1&partita="+codice+"&id_contatto="+id_contatto;
			}
		}
		
		</script>
		
	</body>

 </html>

