<%@ Language=VBScript %>
 <%  if session("admin") = false then
		response.redirect("../../../../index.html")
		end if
		%>
<%
  Response.Buffer = true
  'On Error Resume Next
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 Sessione scaduta, rieffettuare il login.
  <% else %>
   <%

		 CI = Request.QueryString("CI")
		 cod = Request.QueryString("cod")
     prob=Request.QueryString("probabilita")
   %>

		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

 						<%

						   QuerySQL="UPDATE Allievi SET CIAbilitato = "&CI&", Probabilita = "&prob&"  WHERE CodiceAllievo = '"&cod&"';"
						   'response.write(QuerySql&"<br>")
						   ConnessioneDB.Execute QuerySQL

						%>
<script>alert("Aggiornamento effettuato correttamente"); window.location.href="<%=Request.ServerVariables("HTTP_REFERER")%>";</script>
				<%
				'Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo

				'Response.Redirect Session("urlmappa")

				%>
<% end if %>
