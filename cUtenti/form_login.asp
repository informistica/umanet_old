<%@ Language=VBScript %>
<%
 ' modifico procedure per login diretto in base alla classe a cui appartiene l'utente che si logga, senza specificare in anticipo la classe
 
 if Session("DBLogin") = 2 and session("loggato") = true then
		Response.Redirect "../include/altrodb.asp?DB=Classi&Opposto=Expo"
	end if
 
 
if id_classe="" then 
 id_classe=Request.QueryString("id_classe")
end if
 app=Request.QueryString("app") ' vale 1 se sono stato chiamata da apprendimento
 ' cartella=Request.QueryString("cartella")
  'id_materia=Request.QueryString("id_materia")
  
  id_materia=1
 session("ID_Materia")="materia_"&id_materia
  
'  divid=request.querystring("divid")
  logadmin=request.querystring("logadmin")
 
  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
%> 
   <!-- #include file = "../var_globali.inc" --> 
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
'QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
'Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

if strcomp(ucase(session("Loggato")), "TRUE")=0 then
       if strcomp(cartella,"ECDL")=0 then
		   Session("CodiceAllievo")="ospite" 
		   Session("Id_Classe")="7COM"
		   response.redirect "https://www.umanet.net/anno_2012-2013_2/UECDL/index_ecdl_youtube.htm" 
	  else  
	   ' distinguo se sono stato chiamato da app o ver ed in base a ciò scewlgo il redirect
	   id_classe=Request.QueryString("id_classe")
	   cartella=Request.QueryString("cartella")
	  ' stringa_redirect_app="../cClasse/home_app.asp?id_classe=" &  id_classe  & "&cartella=" & cartella  & "&id_materia=" & session("idxMat")  
	   
	   stringa_redirect_app = "../cSocial/default0.asp?scegli=1&id_classe=" &  id_classe  & "&cartella=" & cartella  & "&id_materia=" & session("idxMat")  
	   
	   Response.Redirect stringa_redirect_app
	 
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


	<!-- jQuery -->
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
   
    <script type="text/javascript">
		$(document).ready(function(){
			//var id_classe = '<% =id_classe %>';
			var app = '<% =app %>';
		//	var cartella = '<% =cartella %>';
			var id_materia = '<% =id_materia %>';
			var logadmin = '<% =logadmin %>';
			
			$('#btnLogin').click(function(){
				login();
			});
			
			$(document).bind('keypress', function (e){
				if(e.keyCode=="13"){
					login();
				}
			});
				
			function login(){
				var CodiceAllievo = $('#loginCodiceAllievo').val();
				var PwdAllievo = $('#loginPwd').val();
				$.post(
					'login.asp', 
					{
					//	id_classe:id_classe, 
						app:app, 
					//	cartella:cartella, 
						id_materia:id_materia,
						CodiceAllievo:CodiceAllievo,
						PwdAllievo:PwdAllievo
						//Nome:Nome
						//Cognome:Cognome
					}, 
					function(data){
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
								//window.location.href="../cSocial/quaderno.asp?"+data;   10/03/15 Ho aggiunto Cognome e Nome sotto come parametri costanti 
								window.location.href="../cSocial/default0.asp?scegli=1&id_classe=6COM&divid=&cartella=Expo&Nome=Ospite&Cognome=Ospite";
								
								
								 
							},100);
					 		
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

  

$(document).ready(function(){
	 $('#logospite').click();
	//$('#logospite').animate( { backgroundColor: "green" }, 2000 ).animate( { backgroundColor: "#00CC00" }, 2000 );
});
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
                            
							<form name="dati">
							
								<div class="email">
									<input id="loginCodiceAllievo" type="text" name='txtCodiceAllievo' placeholder="Username" class='input-block-level' style="border-bottom: #fff">
								</div>
								
								<div class="pw">
									<input id="loginPwd" type="password" name="txtPwdAllievo" class="input-block-level" placeholder="Password">
								</div>
								
								<div class="submit">
									<input type="button" id="btnLogin" class="btn btn-primary" value="LOGIN">
								</div>
								                               
							</form>
                            
                            <!-- <a href="../cSocial/default0.asp?scegli=0&id_classe=6COM&cartella=Expo&CodiceAllievo=ospite&by_email=1&DB=1&id_materia=materia_1">-->
                            
                            <button class="btn" onClick="loginospite();" id="logospite">
                                <i class="icon-user"></i>
                                Login come ospite                       
                          	  </button> <!--</a> -->
                                
                            <div id="erroreLogin" class="alert alert-danger" style="display:none">
            					Codice allievo o password errati
           					</div>
							
							<div class="forget">
								<a href="form_registrati.asp?id_classe=6COM">
									<span title="Registrati per avere accesso al corso">
										Registrati<br>
										<img class="img-rounded" src="../../img/umanet3_small.png" alt="">
									<!-- da sistemare trasparenze img 
                                    <img class="img-rounded" src="../../img/EvolutionExpo.png" alt="">-->
                                    </span>
								</a>
							</div>
					
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
