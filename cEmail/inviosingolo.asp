<%
	
	'invio email singolo
	
	sTo = Request.Form("destinatario")
	sFrom = Request.Form("mittente")
	'sFrom = "info@evo.elexpo.net"
	sBody = Request.Form("messaggio")
	sSubject = Request.Form("oggetto")
	'sMailServer = "mail.iisvittuone.net"
	
	
	Sub TestEMail()

emailmit = sFrom
email = sTo
oggetto = sSubject
testo = sBody
set config = CreateObject("CDO.Configuration")
sch = "http://schemas.microsoft.com/cdo/configuration/"
with config.Fields
.item(sch & "sendusing") = 2 ' cdoSendUsingPort
.item(sch & "smtpserver") = "mail.iisvittuone.net" 'application("smtpserver")
.item(sch & "smtpserverport") = 587 'application("smtpserverport")
.item(sch & "smtpauthenticate") = 1 'basic auth
.item(sch & "sendusername") = "umanet" 'application("sendusername")
.item(sch & "sendpassword") = "Inform1stic@" 'application("sendpassword")
.update
end with
with CreateObject("CDO.Message")
  .configuration = config
  .to = email
  .from = emailmit
  .subject = oggetto
  .HTMLBody = testo
  call .send()
end with

End sub
	
	
%>

<!doctype html>
<html>
	<head>
		<meta charset="utf-8">
		<title>Invio Email Singolo</title>
	</head>
	<body>
		
		
		<% 
			'effettuo invio email
			'response.write(destinatario&"<br>"&oggetto&"<br>"&mittente&"<br>"&messaggio)
			TestEMail()
		
		%>
		
		<script>
			alert("Email inviata correttamente"); window.location.href="https://evo.elexpo.net/portale/admin/compilaemail.php"; 
		</script>
		
	</body>
</html>