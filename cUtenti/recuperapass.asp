<%@ Language=VBScript %>

<%  

	'On Error Resume Next

   DB = Request.QueryString("DB")
   user = Request.QueryString("user")
   
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
   
%>
   
   <!-- #include file = "../var_globali.inc" -->
   
 <%  if DB=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
  
else
ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
  
end if

%>
   
   
<%  
	mes = ""
	IsSuccess = false
    
	
	Sub TestEMail()

	

  Set objMail = Server.CreateObject("CDO.Message")
  Set objConf = Server.CreateObject("CDO.Configuration")
  Set objFields = objConf.Fields

sch = "http://schemas.microsoft.com/cdo/configuration/"
with objConf.Fields
 .item(sch & "sendusing") = 2 ' cdoSendUsingPort
 .item(sch & "smtpserver") = "mail.iisvittuone.it" 'application("smtpserver")
 .item(sch & "smtpserverport") = 587 'application("smtpserverport")
 .item(sch & "smtpauthenticate") = 1 'basic auth
 .item(sch & "sendusername") = "umanet" 'application("sendusername")
 .item(sch & "sendpassword") = "Inform1stic@" 'application("sendpassword")
 .update
end with
	
	
	
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
	
	
  QuerySQL="SELECT count(*) FROM Allievi where CodiceAllievo='"&user&"';"
  set rsNumero=ConnessioneDB.Execute(QuerySQL)
	num = rsNumero(0)

	
	if num <> 0 and user <> "" then
	
	  QuerySQL="SELECT CodiceAllievo,PasswordSHA256,Email FROM Allievi where CodiceAllievo='"&user&"';"
	  set rsTabella=ConnessioneDB.Execute(QuerySQL)
	
		'response.write(QuerySQL)
	  
	  sMailServer = "mail.iisvittuone.it"
	  sSubject="Nuova password Umanet"
	  'sBody= "Messaggio da: Umanet Evolution Technologies"  
	  'sBody = Server.HTMLEncode(sBody)
	  linkAvviso=dominio&homesito&"/script/cUtenti/nuovapass.asp?CodiceAllievo="&rsTabella("CodiceAllievo")&"&hash="&rsTabella("PasswordSHA256")&"&DB="&DB
	  'sBody = sBody &"  <br> <a title 'Cambia password' href='"& linkAvviso&"'> Clicca qui per cambiare la tua password</a> <img alt='enlightened' height='20' src='https://www.umanetexpo.net/expo2015Server/UECDL/js/plugins/ckeditor/plugins/smiley/images/lightbulb.gif' title='Idee per evolvere' width='20' /> "
	  
	  
	  sBody = "<style type=""text/css"">"&_ 	
	  "body {"&_
		"width: 100%;"&_
		"margin: 0;"&_
		"padding: 0;"&_
		"-webkit-font-smoothing: antialiased;"&_
	"}"&_
	"@media only screen and (max-width: 600px) {"&_
		"table[class=""table-row""] {"&_
			"float: none !important;"&_
			"width: 98% !important;"&_
			"padding-left: 20px !important;"&_
			"padding-right: 20px !important;"&_
		"}"&_
		"table[class=""table-row-fixed""] {"&_
			"float: none !important;"&_
			"width: 98% !important;"&_
		"}"&_
		"table[class=""table-col""], table[class=""table-col-border""] {"&_
			"float: none !important;"&_
			"width: 100% !important;"&_
			"padding-left: 0 !important;"&_
			"padding-right: 0 !important;"&_
			"table-layout: fixed;"&_
		"}"&_
		"td[class=""table-col-td""] {"&_
			"width: 100% !important;"&_
		"}"&_
		"table[class=""table-col-border""] + table[class=""table-col-border""] {"&_
			"padding-top: 12px;"&_
			"margin-top: 12px;"&_
			"border-top: 1px solid #E8E8E8;"&_
		"}"&_
		"table[class=""table-col""] + table[class=""table-col""] {"&_
			"margin-top: 15px;"&_
		"}"&_
		"td[class=""table-row-td""] {"&_
			"padding-left: 0 !important;"&_
			"padding-right: 0 !important;"&_
		"}"&_
		"table[class=""navbar-row""] , td[class=""navbar-row-td""] {"&_
			"width: 100% !important;"&_
		"}"&_
		"img {"&_
			"max-width: 100% !important;"&_
			"display: inline !important;"&_
		"}"&_
		"img[class=""pull-right""] {"&_
			"float: right;"&_
			"margin-left: 11px;"&_
            "max-width: 125px !important;"&_
			"padding-bottom: 0 !important;"&_
		"}"&_
		"img[class=""pull-left""] {"&_
			"float: left;"&_
			"margin-right: 11px;"&_
			"max-width: 125px !important;"&_
			"padding-bottom: 0 !important;"&_
		"}"&_
		"table[class=""table-space""], table[class=""header-row""] {"&_
			"float: none !important;"&_
			"width: 98% !important;"&_
		"}"&_
		"td[class=""header-row-td""] {"&_
			"width: 100% !important;"&_
		"}"&_
	"}"&_
	"@media only screen and (max-width: 480px) {"&_
		"table[class=""table-row""] {"&_
			"padding-left: 16px !important;"&_
			"padding-right: 16px !important;"&_
		"}"&_
	"}"&_
	"@media only screen and (max-width: 320px) {"&_
		"table[class=""table-row""] {"&_
			"padding-left: 12px !important;"&_
			"padding-right: 12px !important;"&_
		"}"&_
	"}"&_
	"@media only screen and (max-width: 458px) {"&_
		"td[class=""table-td-wrap""] {"&_
			"width: 100% !important;"&_
		"}"&_
	"}"&_
	"</style>"&_
	  "<table width=""100%"" height=""100%"" bgcolor=""#E4E6E9"" cellspacing=""0"" cellpadding=""0"" border=""0"">"&_
	  "<tbody><tr><td width=""100%"" align=""center"" valign=""top"" bgcolor=""#E4E6E9"" style=""background-color:#E4E6E9; min-height: 200px;"">"&_
	  "<table><tbody><tr><td class=""table-td-wrap"" align=""center"" width=""458""><table class=""table-space"" height=""18"" style=""height: 18px; font-size: 0px; line-height: 0; width: 450px; background-color: #e4e6e9;"" width=""450"" bgcolor=""#E4E6E9"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""18"" style=""height: 18px; width: 450px; background-color: #e4e6e9;"" width=""450"" bgcolor=""#E4E6E9"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
	  "<table class=""table-space"" height=""8"" style=""height: 8px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""8"" style=""height: 8px; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
	  "<table class=""table-row"" width=""450"" bgcolor=""#FFFFFF"" style=""table-layout: fixed; background-color: #ffffff;"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-row-td"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal; padding-left: 36px; padding-right: 36px;"" valign=""top"" align=""left"">"&_
		  "<table class=""table-col"" align=""left"" width=""378"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""table-layout: fixed;""><tbody><tr><td class=""table-col-td"" width=""378"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal; width: 378px;"" valign=""top"" align=""left"">"&_
			"<table class=""header-row"" width=""378"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""table-layout: fixed;""><tbody><tr><td class=""header-row-td"" width=""378"" style=""font-family: Arial, sans-serif; font-weight: normal; line-height: 19px; color: #478fca; margin: 0px; font-size: 18px; padding-bottom: 10px; padding-top: 15px;"" valign=""top"" align=""left"">Grazie per aver richiesto una nuova password</td></tr></tbody></table>"&_
			"<div style=""font-family: Arial, sans-serif; line-height: 20px; color: #444444; font-size: 13px;"">"&_
			 "<b style=""color: #777777;"">Segui le istruzioni del sistema</b>"&_
			  "<br>"&_
			  "Clicca sul link qui sotto per poter reimpostare la password"&_
			"</div>"&_
		  "</td></tr></tbody></table>"&_
		"</td></tr></tbody></table>"&_
		"<table class=""table-space"" height=""12"" style=""height: 12px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""12"" style=""height: 12px; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
		"<table class=""table-space"" height=""12"" style=""height: 12px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""12"" style=""height: 12px; width: 450px; padding-left: 16px; padding-right: 16px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""center"">&nbsp;<table bgcolor=""#E8E8E8"" height=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td bgcolor=""#E8E8E8"" height=""1"" width=""100%"" style=""height: 1px; font-size:0;"" valign=""top"" align=""left"">&nbsp;</td></tr></tbody></table></td></tr></tbody></table>"&_
		"<table class=""table-space"" height=""16"" style=""height: 16px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""16"" style=""height: 16px; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_

		"<table class=""table-row"" width=""450"" bgcolor=""#FFFFFF"" style=""table-layout: fixed; background-color: #ffffff;"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-row-td"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal; padding-left: 36px; padding-right: 36px;"" valign=""top"" align=""left"">"&_
		 " <table class=""table-col"" align=""left"" width=""378"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""table-layout: fixed;""><tbody><tr><td class=""table-col-td"" width=""378"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal; width: 378px;"" valign=""top"" align=""left"">"&_
			"<div style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; text-align: center;"">"&_
			  "<a href="""&linkAvviso&""" style=""color: #ffffff; text-decoration: none; margin: 0px; text-align: center; vertical-align: baseline; border: 4px solid #6fb3e0; padding: 4px 9px; font-size: 15px; line-height: 21px; background-color: #6fb3e0;"">&nbsp; Genera &nbsp;</a>"&_
			"</div><br><br><center>Oppure clicca <a href="""&linkAvviso&""">qui</a></center>"&_
			"<table class=""table-space"" height=""16"" style=""height: 16px; font-size: 0px; line-height: 0; width: 378px; background-color: #ffffff;"" width=""378"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""16"" style=""height: 16px; width: 378px; background-color: #ffffff;"" width=""378"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
		  "</td></tr></tbody></table>"&_
		"</td></tr></tbody></table>"&_
	"<table class=""table-space"" height=""6"" style=""height: 6px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""6"" style=""height: 6px; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
	"<table class=""table-row-fixed"" width=""450"" bgcolor=""#FFFFFF"" style=""table-layout: fixed; background-color: #ffffff;"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-row-fixed-td"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal; padding-left: 1px; padding-right: 1px;"" valign=""top"" align=""left"">"&_
	  "<table class=""table-col"" align=""left"" width=""448"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""table-layout: fixed;""><tbody><tr><td class=""table-col-td"" width=""448"" style=""font-family: Arial, sans-serif; line-height: 19px; color: #444444; font-size: 13px; font-weight: normal;"" valign=""top"" align=""left"">"&_
		"<table width=""100%"" cellspacing=""0"" cellpadding=""0"" border=""0"" style=""table-layout: fixed;""><tbody><tr><td width=""100%"" align=""center"" bgcolor=""#f5f5f5"" style=""font-family: Arial, sans-serif; line-height: 24px; color: #bbbbbb; font-size: 13px; font-weight: normal; text-align: center; padding: 9px; border-width: 1px 0px 0px; border-style: solid; border-color: #e3e3e3; background-color: #f5f5f5;"" valign=""top"">"&_
		 " <a href=""#"" style=""color: #428bca; text-decoration: none; background-color: transparent;"">Umanet Evolution Technologies © "&right(Date(),4)&"</a>"&_
		"</td></tr></tbody></table>"&_
	 " </td></tr></tbody></table>"&_
	"</td></tr></tbody></table>"&_
	"<table class=""table-space"" height=""1"" style=""height: 1px; font-size: 0px; line-height: 0; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""1"" style=""height: 1px; width: 450px; background-color: #ffffff;"" width=""450"" bgcolor=""#FFFFFF"" align=""left"">&nbsp;</td></tr></tbody></table>"&_
	"<table class=""table-space"" height=""36"" style=""height: 36px; font-size: 0px; line-height: 0; width: 450px; background-color: #e4e6e9;"" width=""450"" bgcolor=""#E4E6E9"" cellspacing=""0"" cellpadding=""0"" border=""0""><tbody><tr><td class=""table-space-td"" valign=""middle"" height=""36"" style=""height: 36px; width: 450px; background-color: #e4e6e9;"" width=""450"" bgcolor=""#E4E6E9"" align=""left"">&nbsp;</td></tr></tbody></table></td></tr></tbody></table>"&_
	"</td></tr>"&_
	" </tbody></table>"
	  
	  
	  sFrom="Umanet Expo <noreply@iisvittuone.it>"
	  sTo=trim(rsTabella("Email"))
	  
	  
	  TestEMail() 'invio effettivo della mail
	  
	  'response.write("Abbiamo inviato un'email per generare una nuova password all'indirizzo: "& sTo)
	 ' response.write(linkAvviso)
	 
	 session("PwdRecuperata") = true
	 session("ProvenienzaRecuperata") = Request.ServerVariables("SCRIPT_NAME")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
		
	else
		'response.write("Utente non ancora registrato") 'mando in errore 500 la pagina volutamente! l'app è già pronta per gestire l'errore 500
		session("PwdRecuperata") = false
		session("ProvenienzaRecuperata") = Request.ServerVariables("SCRIPT_NAME")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
		
	end if
%>