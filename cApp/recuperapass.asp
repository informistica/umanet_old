
<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
 
<%
Response.charset="utf-8"
Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
		 
 
	 
		%> 
        <!-- #include file = "include/var_globali.inc" --> 
		<!-- #include file = "include/stringa_connessione.inc" --> 
 	     
	 
                 
<%  
	mes = ""
	IsSuccess = false

Sub TestEMail()
	Dim sch, cdoConfig, cdoMessage
	sch = "https://schemas.microsoft.com/cdo/configuration/"
	SMTP_SERVER_PICKUP_DIRECTORY="C:\inetpub\mailroot\Pickup"
	Set cdoConfig = CreateObject("CDO.Configuration")
	With cdoConfig.Fields
	.Item(sch & "sendusing") = 1 ' cdoSendUsingPort
	.Item(sch & "smtpserver") = "127.0.0.1"
	.Item(sch & "smtpserverport") = 25
	.Item("https://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory")=SMTP_SERVER_PICKUP_DIRECTORY    
	.update
	End With
	Set cdoMessage = CreateObject("CDO.Message")
	With cdoMessage
	Set .Configuration = cdoConfig
	.From = sFrom
	.To = sTo
	.Subject = sSubject 
	'.TextBody = sBody
	.HTMLBody=sBody
	.Send
	End With
	Set cdoMessage = Nothing
	Set cdoConfig = Nothing
End sub
    
  QuerySQL="SELECT CodiceAllievo,PasswordSHA256 FROM Allievi where Email='"&request("email")&"';"
  set rsTabella=ConnessioneDB.Execute(QuerySQL) 	 
  sSubject="Nuova password ElexpoApp"
  sBody= "Messaggio da : Umanet Evolution Technologies"  
  sBody = Server.HTMLEncode(sBody)
  linkAvviso=dominio&homesito&"/script/cApp/nuovapass.asp?CodiceAllievo="&rsTabella("CodiceAllievo")&"&hash="&rsTabella("PasswordSHA256")
  sBody = sBody &"  <br> <a title 'Cambia password' href='"& linkAvviso&"'> Clicca qui per cambiare la tua password</a> <img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' /> "
  sFrom="info@umanetexpo.net"
  sTo=request("email")
  TestEMail()
  response.write("Ti è stata inviata un email per generare una nuova password all'indirizzo: "& sTo)
 ' response.write(linkAvviso)
 

%>
 