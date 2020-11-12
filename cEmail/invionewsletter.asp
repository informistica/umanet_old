<%@ Language=VBScript %>
<%
	
	
	
	'invio email newsletter: prof può darci un'occhiata?
	
	'sTo = Request.Form("destinatario")
	sFrom = Request.Form("mittente")
	'sFrom = "info@evo.elexpo.net"
	sBody = Request.Form("messaggio")
	sSubject = Request.Form("oggetto")
	sMailServer = "mail.iisvittuone.net"
	
	
	Sub TestEMail()

	

  Set objMail = Server.CreateObject("CDO.Message")
  Set objConf = Server.CreateObject("CDO.Configuration")
  Set objFields = objConf.Fields
SMTP_SERVER_PICKUP_DIRECTORY="C:\inetpub\mailroot\Pickup"
  With objFields
    .Item("https://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("https://schemas.microsoft.com/cdo/configuration/smtpserver")  = sMailServer
    .Item("https://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10
    .Item("https://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
	.Item("https://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory")=SMTP_SERVER_PICKUP_DIRECTORY
    
    .Update
  End With
	
	
	
  With objMail
    Set .Configuration = objConf
    .From = sFrom
    .To = sTo
    .Subject = sSubject
    .HTMLBody = sBody
  End With

    Err.Clear 
 ' on error resume next
 
    objMail.Send
		
  if len(Err.Description) = 0 then
        mes = " MESSAGGIO INVIATO a " + sTo
     '   mes = mes + " TESTS COMPLETED SUCCESSFULLY!"
        IsSuccess = true
    else
    mes = " " + Err.Description + " INVIO NON RIUSCITO!"
  end if
  Set objFields = Nothing
  Set objConf = Nothing
  Set objMail = Nothing
End sub
	
	
	
%>

<!-- #include file = "../stringhe_connessione/login_ospite_expo.asp" -->

<!doctype html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Invio Email Singolo</title>
	</head>
	<body>
		
		<% 
			
			'effettuo query per ottenere tutti gli utenti di umanetexpo
			
			Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>
   
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		
			
		<%
				
			'effettuo invio email
			'response.write(destinatario&"<br>"&oggetto&"<br>"&mittente&"<br>"&messaggio)
			'TestEMail()
			
			QuerySQL = "SELECT Email FROM Allievi WHERE 1=1;"
			set rsTabella = ConnessioneDB.Execute(QuerySQL)
			'response.write(QuerySQL)
			
			'QuerySQL0 = "SELECT count(*) FROM Allievi WHERE 1=1;"
			'set rsTabellaC = ConnessioneDB.Execute(QuerySQL0)
			'response.write(rsTabellaC(0))
			
			On Error Resume Next
			
			do while not rsTabella.EOF 
				'response.write(rsTabella("Email")&"<br>")
				sTo = rsTabella("Email")
				TestEMail()
				rsTabella.MoveNext
			loop
			
			
			Session.Abandon
          Response.Cookies("Dati")("Loggato")= ""
		  Response.Cookies("Dati")("Cognome")= ""
		  Response.Cookies("Dati")("Nome")=""
		  Response.Cookies("Dati")("CodiceAllievo")= ""
		  Response.Cookies("Dati")("Username")=""
		  Response.Cookies("Dati")("DataTest")= "" 
		  Response.Cookies("Dati")("Id_Classe")=""
		  Response.Cookies("Dati")("cartella")=""
		  Response.Cookies("Dati")("Cartella")=""
		  Response.Cookies("Dati")("CartellaAdmin")= ""
	      Response.Cookies("Dati")("In_Quiz")= ""
	      Response.Cookies("Dati")("CodAdmin")= ""
		  ' impostate in home.asp
		  
     Response.Cookies("Dati")("Materia")= ""
	 Response.Cookies("Dati")("ID_Materia")= ""
	 Response.Cookies("Dati")("ID_Matsint")= ""  ' mi serve la chiave numerica per il DBMatprof per recuperare la login dell'admin
	 Response.Cookies("Dati")("idxMat")= ""
	 
	 Response.Cookies("Dati")("DBCopiatestonline")= ""
	 Response.Cookies("Dati")("DBForum")= ""
	 Response.Cookies("Dati")("DBLavagna")= ""
	 Response.Cookies("Dati")("DBDiario")= ""
	 Response.Cookies("Dati")("DBDesideri")= ""
			
						
		%>
		
		
		
		<script>
			alert("Email inviata correttamente"); window.location.href="https://evo.elexpo.net/portale/admin/compilaemail.php"; 
		</script>
		
	</body>
</html>