<%@ Language=VBScript %>
<%
	Response.charset="utf-8"
	Function newpassword()
	  ' Creo la variabile "caratteri" contenente tutti i
	  ' numeri da 0 a 9 e tutte le lettere dalla A alla Z
	  caratteri = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" 
	  Randomize()
	  Do Until len(password) = 4 
		' Genero un valore casuale compreso tra 1 e 37
		' dove 1 corrisponde al numero 0 e 37 alla lettera Z 
		carattere = Int((37 * Rnd) + 1)
		' Aggiorno la variabile "password" usando Mid per individuare
		' all'interno della stringa "caratteri" il numero o la lettera
		' che corrisponde al numero memorizzato nella variabile "carattere"
		password = password & Mid(caratteri,carattere,1) 
	  Loop 
	  newpassword = password
	End Function

	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
 
			   <!-- #include file = "include/stringa_connessione.inc" -->
			  <!-- #include file = "include/sha256.asp" -->
			  <!-- #include file = "include/var_globali.inc" --> 
			  
	<%  
    errore=1
	CodiceAllievo=request("CodiceAllievo")
	PwdAllievo=request("hash")
	QuerySQL ="Select CodiceAllievo from Allievi " &_
		" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo & "' and PasswordSHA256='"&PwdAllievo&"' ;"
	set rsTabella=ConnessioneDB.Execute(QuerySQL)
	if not rsTabella.eof then
	    errore=0
		NewPwdAllievo= newpassword
		msg1="Password: " & NewPwdAllievo
		linkAvviso="https://"&dominio&homesito&"/script/cUtenti/form_login3.asp?app=1&id_materia=1&DB=1&id_matlong=materia_1"
		msg2=" <a href='"&linkAvviso&"'> Modifica in Umanet Evolution </a>"	
		NewPwdAllievo= SHA256(newpassword)
		QuerySQL ="UPDATE Allievi SET   Allievi.PasswordSHA256 = '" &NewPwdAllievo& "'" &_
			" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo & "'"
			
			'response.write("<br>"&QuerySQL)
		  '  ConnessioneDB.Execute(QuerySQL)
	else
	  msg3= "La password non è stata modificata"
	end if

%>		  
<html lang="en">
	<head>
		<meta https-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
		<meta charset="utf-8" />
		<title>Cambio password</title>

		<meta name="description" content="overview &amp; stats" />
		<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0" />

		<!-- bootstrap & fontawesome -->
		<link rel="stylesheet" href="include/assets/css/bootstrap.min.css" />
		<link rel="stylesheet" href="include/assets/font-awesome/4.2.0/css/font-awesome.min.css" />

		<!-- page specific plugin styles -->

		<!-- text fonts -->
		<link rel="stylesheet" href="include/assets/fonts/fonts.googleapis.com.css" />

		<!-- ace styles -->
		<link rel="stylesheet" href="include/assets/css/ace.min.css" class="ace-main-stylesheet" id="main-ace-style" />

		<!--[if lte IE 9]>
			<link rel="stylesheet" href="include/assets/css/ace-part2.min.css" class="ace-main-stylesheet" />
		<![endif]-->

		<!--[if lte IE 9]>
		  <link rel="stylesheet" href="include/assets/css/ace-ie.min.css" />
		<![endif]-->

		<!-- inline styles related to this page -->

		<!-- ace settings handler -->
		<script src="include/assets/js/ace-extra.min.js"></script>
        <script type="text/javascript" src="js/sha256.js">/* SHA-256 JavaScript implementation */</script>
        
		<script src="include/assets/js/jquery.min.js"></script>
		<script src="include/assets/js/jquery-ui.min.js"></script>
        

		<!-- HTML5shiv and Respond.js for IE8 to support HTML5 elements and media queries -->

		<!--[if lte IE 8]>
		<script src="include/assets/js/html5shiv.min.js"></script>
		<script src="include/assets/js/respond.min.js"></script>
		<![endif]-->
	</head>


	<body class="login-layout light-login" onload="controllodati(window.localStorage.getItem('usernamesalvata'), window.localStorage.getItem('passwordsalvata'))">
		<div class="main-container">
			<div class="main-content">
				<div class="row">
					<div class="col-sm-10 col-sm-offset-1">
						<div class="login-container">
							<div class="center">
								<h1>									
									<span class="red">ElexpoApp</span>	
                                    <img src="img/icon.png" alt="" width="40" height="40">							
								</h1>
								<h4 class="blue" id="id-company-text">&copy; Umanet Evolution Technologies</h4>
							</div>

							<div class="space-6"></div>

							<div class="position-relative">
								<div id="login-box" class="login-box visible widget-box no-border">
									<div class="widget-body">
										<div class="widget-main">
											<h4 class="header blue lighter bigger">
												 <img src="img/logo.jpg" alt="" width="50" height="50">
												<%if errore=0 then%>
													E' stata generata una nuova password per l'utente <b> <%=CodiceAllievo%></b>
												<%else%>
													Sessione non valida, la password non è stata modificata
												<%end if%>
											</h4>
											<div class="space-6"></div>							 
										</div><!-- /.widget-main -->

										 <div class="toolbar clearfix">
										 <%if errore=0 then%>
											<div>
												<a href="#"  class="forgot-password-link">
													<i class="ace-icon fa fa-arrow-left"></i><b>
														 <%=msg1%>													 
													 </b>
												</a>
											</div>											
											<div>
												<a href="<%=linkAvviso%>" class="forgot-password-link">
												    Puoi cambiarla in Umanet														 
													<i class="ace-icon fa fa-arrow-right"></i>
												</a>
											</div>
											<%end if%>											
										</div>                                       
										</div>
									</div><!-- /.widget-body -->
								</div><!-- /.login-box -->
							</div><!-- /.position-relative --> 
				</div><!-- /.row -->
			</div><!-- /.main-content -->   
	</body>
</html>
