 
<%


smtpserver = "mail.iisvittuone.it" 
smtpserverport = 587  
smtpauthenticate = 1 'basic auth
sendusername = "umanet"  
sendpassword = "Inform1stic@" 

Sub TestEMailOld()

Dim sch, cdoConfig, cdoMessageU
sch = "https://schemas.microsoft.com/cdo/configuration/"
SMTP_SERVER_PICKUP_DIRECTORY="C:\inetpub\mailroot\Pickup"
Set cdoConfig = CreateObject("CDO.Configuration")
With cdoConfig.Fields
.Item(sch & "sendusing") = 1 ' cdoSendUsingPort

'funziona con tutti e tre gli indirizzi ip
.Item(sch & "smtpserver") = smtpserver
.Item(sch & "smtpserverport") = smtpserverport
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
'.AddAttachment "C:\inetpub\umanetroot\expo2015Server\GrazieMike.pdf"
.Send
End With
Set cdoMessage = Nothing
Set cdoConfig = Nothing


End sub


Sub TestEMail()

	

  Set objMail = Server.CreateObject("CDO.Message")
  Set objConf = Server.CreateObject("CDO.Configuration")
  Set objFields = objConf.Fields
SMTP_SERVER_PICKUP_DIRECTORY="C:\inetpub\mailroot\Pickup"
 
  
sch = "http://schemas.microsoft.com/cdo/configuration/"
with objFields
 .item(sch & "sendusing") = 2 ' cdoSendUsingPort
 .item(sch & "smtpserver") = smtpserver
 .item(sch & "smtpserverport") = smtpserverport
 .item(sch & "smtpauthenticate") = 1 'basic auth
 .item(sch & "sendusername") = sendusername
 .item(sch & "sendpassword") = sendpassword
 .update
end with
	
	
	
	
  With objMail
    Set .Configuration = objConf
	.BodyPart.Charset = "utf-8"
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
 