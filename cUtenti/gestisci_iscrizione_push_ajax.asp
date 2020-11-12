<%@ Language=VBScript %>
<% Response.Buffer=True
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%

on error resume next
api=CInt(Request.QueryString("api"))
id=Request.QueryString("id")
CodiceAllievo=Request.QueryString("codiceallievo")

	if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then
		Select Case api
  		Case 1 'cancello iscrizione	
			QuerySQL ="DELETE  FROM push_subscriptions WHERE CodiceAllievo ='" &session("CodiceAllievo")&"' and id="&id&";"
			ConnessioneDB.Execute(QuerySQL)
			If Err.Number = 0 Then
				msg="Cancellazione avvenuta!"
			Else
				msg=Err.Description
				Err.Number = 0
			End If
		Case 2 ' attivo/disattivo iscrizione
		    stato=Request.QueryString("stato")
			QuerySQL ="UPDATE push_subscriptions Set Attiva ='"&stato&"' WHERE CodiceAllievo ='" &session("CodiceAllievo")&"' and id="&id&";"
			ConnessioneDB.Execute(QuerySQL)
			if (strcomp(stato,"0")=0) then
			  msg="Disattivazione effettuata!"
			else
			  msg="Attivazione effettuata!"
			end if
		Case Else
		    msg="Chiamata non valida!"
		End Select
		 response.write(msg)
	else
		response.write("Non puoi modificare i dati degli altri studenti!")
	end if
  
 


%>
