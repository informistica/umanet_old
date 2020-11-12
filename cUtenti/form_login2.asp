<%@ Language=VBScript %>

<% 'effettuo una verifica: se l'utente è già connesso su db expo allora setto variabile a true (nel body mostro alert) -> evita conflitto sessioni

	connessione = Request.QueryString("connessione")

	if Session("DBLogin") = 1 and session("loggato") = true and connessione = 2 then
		Response.Redirect "../include/altrodb.asp?DB=Expo&Opposto=Classi"
	else if Session("DBLogin") = 2 and session("loggato") = true and connessione = 1 then
		Response.Redirect "../include/altrodb.asp?DB=Classi&Opposto=Expo"
	else
	pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
  if (left(pathEnd1,10)="c:\inetpub") then
     locale=1
  else
     locale=0
  end if 	
end if  
  
  ' per rirpristinare le materie multiple togliere questa parte cablata 
  
  
  if connessione = 1 then 'expo

app=1
  materia="Umanet 1"
  cartella="Expo" 
  id_materia=1
  DBCopiatestonline="Copiaditestonline.mdb" ' lo faccio puntare al Cdb per la didattica ordinaria
  DBClassifica="DBClassifica.mdb"
  DBForum="forum.mdb"
  DBLavagna="lavagna.mdb"
  DBDiario="diario.mdb"
  DBDesideri="desideri.mdb"
  
 CodiceAllievo = Session("CodiceAllievo")
	Session("Materia")=materia
	Session("ID_Materia")="materia_"&id_materia
	Session("ID_Matsint")=id_materia ' mi serve la chiave numerica per il DBMatprof per recuperare la login dell'admin
	Session("idxMat")=id_materia
	Session("Cartella")=cartella
	Session("DBCopiatestonline")=DBCopiatestonline
	Session("DBClassifica")=DBClassifica
	Session("DBForum")=DBForum
	Session("DBLavagna")=DBLavagna
	Session("DBDiario")=DBDiario
	Session("DBDesideri")=DBDesideri
    Session("DB")=1




else ' classi
  
  app=1
  materia="Umanet 1"
  cartella="Expo" 
  id_materia=1
  DBCopiatestonline="Copiaditestonline2.mdb" ' lo faccio puntare al Cdb per la didattica ordinaria
  DBClassifica="DBClassifica.mdb"
  DBForum="forum.mdb"
  DBLavagna="lavagna.mdb"
  DBDiario="diario.mdb"
  DBDesideri="desideri.mdb"
  
 CodiceAllievo = Session("CodiceAllievo")
	Session("Materia")=materia
	Session("ID_Materia")="materia_"&id_materia
	Session("ID_Matsint")=id_materia ' mi serve la chiave numerica per il DBMatprof per recuperare la login dell'admin
	Session("idxMat")=id_materia
	Session("Cartella")=cartella
	Session("DBCopiatestonline")=DBCopiatestonline
	Session("DBClassifica")=DBClassifica
	Session("DBForum")=DBForum
	Session("DBLavagna")=DBLavagna
	Session("DBDiario")=DBDiario
	Session("DBDesideri")=DBDesideri
    Session("DB")=2
	
	end if
	
	end if
	
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    

 %>

  <!-- #include file = "../var_globali.inc" --> 
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
<%
 ' modifico procedure per login diretto in base alla classe a cui appartiene l'utente che si logga, senza specificare in anticipo la classe
 
 if Session("CodiceAllievo") <> "" then
 
 QuerySQL = "SELECT * FROM AssociazioniAllievi WHERE CodiceAllievo = '"&Session("CodiceAllievo")&"' OR UtenteAssociato = '"&Session("CodiceAllievo")&"'"
 'response.write(QuerySQL)
 set rsTabellaA = ConnessioneDB.Execute(QuerySQL)
 'response.write(rsTabellaA("CodiceAllievo"))
 
 do while not rsTabellaA.EOF
 
 QuerySQL = "SELECT count(*) FROM Allievi WHERE ((CodiceAllievo = '"&rsTabellaA("UtenteAssociato")&"' OR CodiceAllievo = '"&rsTabellaA("CodiceAllievo")&"') AND Id_Classe = '"&Request.QueryString("id_classe")&"') "
 'response.write("<br>"&QuerySQL)
 set rsTabellaA1 = ConnessioneDB.Execute(QuerySQL)
 
 num = rsTabellaA1(0)
 
 if num = 1 then
	UA = rsTabellaA("UtenteAssociato")
	CA = rsTabellaA("CodiceAllievo")
end if
 
 rsTabellaA.movenext
 loop
 
 
 if num = 1 then
 'response.write("Ciao")
 QuerySQL = "SELECT * FROM Allievi WHERE ((CodiceAllievo = '"&UA&"' OR CodiceAllievo = '"&CA&"') AND Id_Classe = '"&Request.QueryString("id_classe")&"') "
 'response.write("<br>Query:"&QuerySQL)
 set rsTabellaA2 = ConnessioneDB.Execute(QuerySQL)


 
 Session("CodiceAllievo") = rsTabellaA2("CodiceAllievo")
 Session("stile")=rsTabellaA2("Stile")

  Response.Cookies("Dati")("CodiceAllievo")= Session("CodiceAllievo")
  Response.Cookies("Dati")("Stile")= Session("Stile")
 if CodiceAllievo <> Session("CodiceAllievo") then
		Session("cambio")=1
		'per la foto profilo che si perde  
		QuerySQL = "SELECT Cartella FROM Classi WHERE Id_Classe = '"&Request.QueryString("id_classe")&"' "
 
		set rsCartella = ConnessioneDB.Execute(QuerySQL)
		session("id_classe_img")= rsCartella(0)
		
		else
		Session("cambio")=0
		end if
 ' commento perchè altrimenti non si vede la classifica perchè mancano i parametri
 'Session("DataCla") = ""
 'Session("DataClaq") = ""
 'Session("DataCla2") = ""
 'Session("DataClaq2") = ""

 end if
 
' response.write("ciao")
 
 end if
 
if id_classe="" then 
 id_classe=Request.QueryString("id_classe")
end if
 app=Request.QueryString("app") ' vale 1 se sono stato chiamata da apprendimento
 doc=Request.QueryString("doc")  'vale 1 se provengo da home_doc.asp
 ' cartella=Request.QueryString("cartella")
  'id_materia=Request.QueryString("id_materia")
  
  id_materia=1
 session("ID_Materia")="materia_"&id_materia
 
 id_as=Request.QueryString("id_as") ' anno scolastico
  
 session("id_as")=id_as
  id_scuola=Request.QueryString("id_scuola") ' anno scolastico
 session("id_scuola")=id_scuola
  
'  divid=request.querystring("divid")
  logadmin=request.querystring("logadmin")
 
  
%> 
  
<%


if doc<>""then 
id_classe="8COM"
end if

QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
 registra=rsTabella1("Registra")
 Set rsTabella1 = nothing
 
'if strcomp(ucase(session("Loggato")), "TRUE")=0 then
if session("loggato") = true then
       if strcomp(cartella,"ECDL")=0 then
		   Session("CodiceAllievo")="ospite" 
		   Session("Id_Classe")="7COM"
		   response.redirect "https://www.umanet.net/anno_2012-2013_2/UECDL/index_ecdl_youtube.htm" 
	  else  
	   ' distinguo se sono stato chiamato da app o ver ed in base a ciò scewlgo il redirect
	   id_classe=Request.QueryString("id_classe")
	   cartella=Request.QueryString("cartella")
	   
	   QuerySQL="SELECT count(*) FROM [dbo].[3PERIODI] Where ID_Classe='"& id_classe &"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
numPeriodi=rsTabella(0)+1' +1 per Oggi  
redim periodi(numPeriodi)  
' faccio la query per prelevare i periodi di valutazione per questa classe 
QuerySQL="SELECT * FROM [dbo].[3PERIODI] Where Id_Classe='"& id_classe &"';"
		
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
   ' carico il vettore delle date di valutazione
    periodi(0)=inizio_anno
    idperiodo=1
	selezionato=0
	do while not rsTabella.eof
           periodi(idperiodo)=rsTabella.fields("Data")
		   if  rsTabella.fields("Iniziale")=1 then
		     selezionato=idperiodo ' periodo da cui deve partire la classifica
		     DataIniz=rsTabella.fields("Data")
		   end if
		   
		   idperiodo=idperiodo+1
		   rsTabella.movenext()
    loop 	
	' se il giorno o il mese hanno na sola cifra devo aggiungere lo 0 davanti
	giorno=day(date())
	mese= month(date())
	anno=year(date())
	if len(giorno)=1 then
	   giorno="0" &	day(date())
	end if
	if len(mese)=1 then
	   mese="0" &	month(date())
	end if
	DataOggi=giorno&"/"&mese&"/"&year(date())
	periodi(idperiodo)= DataOggi 
	if not rsTabella.eof then 
	rsTabella.movefirst()
	end if
	
	'url="C:\Inetpub\umanetroot\logPeriodi.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(idperiodo)
				'objCreatedFile.Close	
'response.write(QuerySQL)
'response.write(numPeriodi)	
	
%>


    <% for i=0 to numPeriodi 
		  ' response.write(left(DataCla,10)& "---"& left(periodi(i),10) & i & "di" & numPeriodi  &"<br>")
		   
		next %>	

<%
		  for i=0 to numPeriodi

	   %>
	   
			   <% if numPeriodi = 1 then %>
			   
						  
			   <% else %>
			   
						   <% if i = numPeriodi-1 then %>
<% Session("DataCla") = periodi(i)
Session("DataClaq") = periodi(i) 
DataCla = periodi(i)
DataClaq = periodi(i)
%>							
						   <%else %>
						   
								
						   
						   <%end if%>
			   
			   <% end if %>
			   
	   <% next %>
	   
	 


	   <%
		  for i=0 to numPeriodi

	   %>
	   
			   <% if numPeriodi = 1 then %>
			   
						   
			   
			   <% else %>
			   
						   <% if i = numPeriodi then %>
<% DataCla2New = DateAdd("d",1,periodi(i))
Session("DataCla2") = DataCla2New
DataCla2 = DataCla2New
DataClaq2 = DataCla2New
Session("DataClaq2") = DataCla2New 
Session("DataClaOld") = periodi(i)
%>							
				
						   
						   
						   <%end if%>
			   
			   <% end if %>
			   
	   <% next %>
	   
	 
	   <%
	   
		if num <> 1 then	   
	   stringa_redirect_app="../cClasse/home_app.asp?id_classe=" &  id_classe  & "&cartella=" & cartella  & "&id_materia=" & session("idxMat")  
	   else
	   Session("Id_Classe") = id_classe
	   ' commento perchè altrimenti non si vede la foto profilo nel cambio classe associata
	   'Session("id_classe_img") = Request.QueryString("cartella")
	   Session("CartellaIniz") = Request.QueryString("cartella")
	    stringa_redirect_app="../cClasse/quaderno.asp?umanet=0&Cognome="&session("Cognome")&"&Nome="&session("Nome")&"&stile="&session("stile")&"&id_classe="&id_classe&"&classe="&Request.QueryString("cartella")&"&cod="&Session("CodiceAllievo")&"&DataClaq="&Session("DataClaq")&"&DataClaq2="&Session("DataClaq2")
		
	   end if
	   Response.Redirect stringa_redirect_app
	 ' Response.write stringa_redirect_app
	 
	  end if 
	   
End IF

 ' Response.Redirect stringa_redirect_app


Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
 

%>
<!doctype html>
<html>
<head>
<title>Login Umanet</title>
	<meta charset="utf8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">


	<!-- jQuery INCLUDERE PER login asincrono -->
	<script src="../../js/jquery.min.js"></script>
	<script src="https://ajax.googleapis.com/ajax/libs/jqueryui/1.8/jquery-ui.min.js"></script>
    
    
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<script src="../../js/eakroko.js"></script>

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
   
   
   
   <!--<script src="../../js/sha256.js">/* SHA-256 JavaScript implementation */</script>-->
   
   
    <script type="text/javascript">
		$(document).ready(function(){
			// parametri che vengono passati alla pagina che effettua il login e restituisce in modo asincrono il risultato al chiamante
			// nel tuo caso gli unici parametri che devi passare tramite la funzione login() alla pagina login.asp sono CodiceAllievo e PwdAllievo 
			//quindi ignora le tre righe seguenti	
			var app = '<% =app %>';
			var id_materia = '<% =id_materia %>';
			var id_as = '<% =id_as %>';
			var logadmin = '<% =logadmin %>';
			var id_classe = '<% =id_classe %>';
			
			//quando si clicca sul bottone di login viene invocata la funzione
			$('#btnLogin').click(function(){
				login();
			});
			
			//la funzione login può essere chiamata anche premedo il tasto invio 
			$(document).bind('keypress', function (e){
				if(e.keyCode=="13"){
					login();
				}
			});
				
			function login(){
				//vengono valorizzate le variabili da passare a login.asp leggendo i valori dal form
				var CodiceAllievo = $('#loginCodiceAllievo').val();
				var PwdAllievo = $('#loginPwd').val();
				var PwdAllievoSHA256 = Sha256.hash(PwdAllievo)
				
				//alert(PwdAllievoSHA256)
				$.post(
					'login2.asp', 
					{   //passaggio dei parametri alla pagina login.asp
						id_classe:id_classe,
						app:app, 
						id_materia:id_materia,
						id_as:id_as,
						CodiceAllievo:CodiceAllievo,
						//PwdAllievo:PwdAllievo
						PwdAllievoSHA256:PwdAllievoSHA256
					}, 
					function(data){
						//la funzione login.asp restituisce i parametri semplicemente scrivendo (response.write riga 37 nel caso di errore o riga 118 in caso di login corretto) 
						if(data=="errore"){
							//alert(data);
							$('.ribbon').css('background', 'red');
							$('#erroreLogin').fadeIn(1500);
							setTimeout(function(){
								$('#erroreLogin').fadeOut(1500);
							},6000);
						}else{
							$('.blurred').show();
							$('.ribbon').css('background', 'green');
							setTimeout(function(){
								if (id_classe == '6COM')
								  window.location.href="../cSocial/default0.asp?scegli=1&id_classe=6COM&divid=&cartella=Expo";
							 //	a=4
								else
								//  a=3
								//se il login ha successo viene effettuato un redirect alla pagina 
								window.location.href="../cClasse/quaderno.asp?"+data;
								//window.location.href="../cSocial/default0.asp?"+data;
							 								 
							},1000);
					 		
						}
					}// end function(data)
				);
			}
		});
		
		 function loginospite() {
		with (document.dati) { 
		    txtCodiceAllievo.value="ospite"
			txtPwdAllievo.value="ospite"
		 }
		  $('#btnLogin').click();
	  
	 
	    event.stopPropagation();
}

//serve per far "lampeggiare" il bottone per il login come ospite
$(document).ready(function(){
	$('#logospite').animate( { backgroundColor: "green" }, 2000 ).animate( { backgroundColor: "#00CC00" }, 2000 );
	$('#passdim').animate( { backgroundColor: "red" }, 2000 ).animate( { backgroundColor: "#ff0000", color: "white" }, 2000 );
});
	</script>

<script src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>
      

 <script language="javascript" type="text/javascript"> 
  
 function crittapwd() {
 var PwdAllievo=login.PwdAllievo.value;
 var PwdAllievoSHA256 = Sha256.hash(PwdAllievo)
 //  alert(PwdAllievoSHA256);
 
 if (PwdAllievoSHA256=="")
	{
	   alert("Password non crittografata");
	   return 0;
	}
 else
	{
    document.login.action = "login256.asp?cartella=<%=cartella%>&divid=<%=divid%>&app=<%=app%>&id_classe=<%=id_classe%>&logadmin=<%=logadmin%>&PwdAllievoSHA256="+PwdAllievoSHA256;
	document.login.submit();
		
	 
    }
	
}
  </script>

</head>

	<body class="login">
	
		<div class="wrap">
			<div class="container">
				<div class="row">
				
					<div class="span6 offset3">
						<div class="ribbon"></div>
						<div class="login-body">
                        	<div class="blurred"><i class="icon icon-spin icon-spinner"></i></div>
							<h1> Umanet&nbsp;<i class="glyphicon-user_add"></i>&nbsp;3.0</h1>
                            
							<!--<br>
							<div class="alert alert-danger">
            					A causa di problemi tecnici è stato necessario ripristinare un backup del database del 13/12. Sei hai inserito compiti tra il 13/12 e il 16/12 contatta l'amministratore di sistema.
           					</div>-->
							<% if session("PwdRecuperata") = false and session("ProvenienzaRecuperata") = "/expo2015Server/UECDL/script/cUtenti/recuperapass.asp" then %>
							<br>
							<div class="alert alert-danger">
            					Impossibile recuperare la password:<br>
								utente non presente nel DataBase
           					</div>
							<% else if session("PwdRecuperata") = true and session("ProvenienzaRecuperata") = "/expo2015Server/UECDL/script/cUtenti/recuperapass.asp" then %>
							<br>
							<div class="alert alert-success">
            					Password recuperata correttamente:<br>
								riceverai una Email con le istruzioni
           					</div>
							<%end if%>
							<%end if%>
							
							<%session("PwdRecuperata") = ""
							session("ProvenienzaRecuperata") = "" %>
							
							<form name="dati">
							
								<div class="email">
									<input id="loginCodiceAllievo" type="text" name='txtCodiceAllievo' placeholder="Username" class='input-block-level' style="border-bottom: #fff">
								</div>
								
								<div class="pw">
									<input id="loginPwd" type="password" name="txtPwdAllievo" class="input-block-level" placeholder="Password">
								</div>
								
								<div class="submit">
								
									<button type="button" class="btn" onclick="window.location.href='recuperapass.asp?DB=2&user='+prompt('Inserire username dell\'utente in questione')" id="passdim" style="width:48%">
										<i class="icon-refresh"></i>
										Rec. Password                       
								    </button>
									&nbsp;
									<button type="button" id="btnLogin" class="btn btn-primary" style="width:48%" value="LOGIN">LOGIN</button>
								</div>
								                               
							</form>
                            
                            <!-- <a href="../cSocial/default0.asp?scegli=0&id_classe=6COM&cartella=Expo&CodiceAllievo=ospite&by_email=1&DB=1&id_materia=materia_1">-->
                            
							<% if connessione = 1 then %>
							
                            <button class="btn" onClick="loginospite();" id="logospite">
                                <i class="icon-user"></i>
                                Login come ospite                       
                          	  </button> <!--</a> -->
							 
							<% end if %>							 
                                
                            <div id="erroreLogin" class="alert alert-danger" style="display:none">
            					Codice allievo o password errati
           					</div>
						 <!-- TOGLIERE I COMMENTI PER APRIRE ISCRIZIONE-->
							<%
							registra=1
							if registra=1 then%>
							<div class="forget">
								<a href="form_registrati.asp?id_classe=<%=id_classe%>&divid=<%=divid%>">
									<span title="Registrati per avere accesso al corso">
										Registrati<br>
										<img class="img-rounded" src="../../img/umanet3_small.png" alt="">
                                    <!--<img class="img-rounded" src="../../img/EvolutionExpo.png" alt="">-->
                                    </span>
								</a>
							</div>
                            <% end if%>
                            <div class="forget">
								 
									<span title="Entra nel percorso per gli Eletti di Expo">
										
										<img class="img-rounded" src="../../img/logoelexposolo.png" alt="" width="50%" height="80%">
									 
                                    </span>
								 
							</div>
					
                    
						</div>
					</div>
				
				</div><!-- end row -->
				<!--
				<div class="box-cookie">
					<p>Per utilizzare correttamente il sito attiva i Cookie</p>
				</div>
				-->
			</div><!-- end container -->
			
				
		</div>
		
	
	</body>
    
</html>
