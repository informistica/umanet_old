<%@ Language=VBScript %>

<%  
    On Error Resume Next
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
   
%>

<%
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

       %>  
   
   <!-- #include file = "../var_globali.inc" -->
    <!-- #include file = "include/sha256.asp" -->
   
   <%  
   
   DB = Request.QueryString("DB")  
	
	D = Split(DB, ",")
	DB = D(0)
	
   if DB=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
  
else
ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
  
end if

%>
   
   
   <%  
    errore=1
	
	CodiceAllievo=Request.QueryString("CodiceAllievo")
	PwdAllievo=Request.QueryString("hash")
	
	CodiceA = Split(CodiceAllievo, ",")
	CodiceAllievo = CodiceA(0)
	
	PwdA = Split(PwdAllievo, ",")
	PwdAllievo = PwdA(0)
	
	
	
	QuerySQL ="Select CodiceAllievo from Allievi " &_
		" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo & "' and PasswordSHA256='"&PwdAllievo&"' ;"
		'response.write(QuerySQL)
	set rsTabella=ConnessioneDB.Execute(QuerySQL)
	
	if not rsTabella.eof then
	    errore=0
		NewPwdAllievo= newpassword
		msg1="Password: " & NewPwdAllievo
		if DB = 1 then
			linkAvviso="https://"&dominio&homesito&"/script/cUtenti/form_login3.asp?app=1&id_materia=1&DB=1&id_matlong=materia_1"
		else
			linkAvviso="https://"&dominio&homesito&"/home.asp?classi=1"
		end if
		msg2=" <a href='"&linkAvviso&"'> Modifica in Umanet Evolution </a>"	
		NewPwdAllievo= SHA256(NewPwdAllievo)
		QuerySQL ="UPDATE Allievi SET   Allievi.PasswordSHA256 = '" &NewPwdAllievo& "'" &_
			" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo & "'"
			
			'response.write("<br>"&QuerySQL)
		    ConnessioneDB.Execute(QuerySQL)
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
		<link rel="stylesheet" href="include/assets/css/font-awesome.min.css" />

		<!-- page specific plugin styles -->

		<!-- text fonts -->

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
        <script type="text/javascript" src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>
        
		<script src="include/assets/js/jquery.min.js"></script>
		<script src="include/assets/js/jquery-ui.min.js"></script>
        

		<!-- HTML5shiv and Respond.js for IE8 to support HTML5 elements and media queries -->

		<!--[if lte IE 8]>
		<script src="include/assets/js/html5shiv.min.js"></script>
		<script src="include/assets/js/respond.min.js"></script>
		<![endif]-->
	</head>


	<body class="login-layout light-login">
		<div class="main-container">
			<div class="main-content">
				<div class="row">
					<div class="col-sm-10 col-sm-offset-1">
						<div class="login-container">
							<div class="center">
								<h1>									
									<span class="red">UmanetExpo</span>	
								</h1>
								<h4 class="blue" id="id-company-text">&copy; Umanet Evolution Technologies</h4>
							</div>

							<div class="space-6"></div>

							<div class="position-relative">
								<div id="login-box" class="login-box visible widget-box no-border">
									<div class="widget-body">
										<div class="widget-main">
											<h4 class="header blue lighter bigger">
												<%if errore=0 then%>
													E' stata generata una nuova password per l'utente <b> <%=CodiceAllievo%></b><br><br>
													<center><%=msg1%></center>
												<%else%>
													Sessione non valida, la password non è stata modificata
												<%end if%>
											</h4>
											<div class="space-6"></div>							 
										</div><!-- /.widget-main -->

										 <div class="toolbar clearfix">
										 <%if errore=0 then%>
											<div></div>											
											<div>
												<a href="<%=linkAvviso%>" class="forgot-password-link">
												    Cambiala in Umanet&nbsp;<i class="ace-icon fa fa-arrow-right"></i>
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