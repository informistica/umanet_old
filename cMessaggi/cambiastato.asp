<%@Language=VBScript%>
<!doctype html>
<html>
<head>
<title>Cambia Stato Notifica</title>
</head>
<body>

<% if session("loggato") <> true then %>

<script>alert("Devi essere loggato per accedere a questa sezione"); window.location.href="../../../../"</script>

<%else%>

<% Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>

<!-- #include file = "../var_globali.inc" --> 
<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" --> 
<!-- #include file = "../service/controllo_sessione.asp" --> 

<%  

tutte = Request.QueryString("Tutte") 
lette = Request.QueryString("Lette")
cod = Request.QueryString("cod")

if tutte <> 1 then

' notifiche specifiche (entro nel secondo if)

	if lette = 1 then
	
	' tutte quelle lette
	
	rimuovi = Request.QueryString("Rimuovi")
	ripristina = Request.QueryString("Ripristina")
	
		if rimuovi = 1 then
			QuerySQL = "DELETE FROM AVVISI WHERE CodiceAllievo = '"&cod&"' AND Visto = 1;"
		else
			QuerySQL = "UPDATE AVVISI SET Visto = 0 WHERE CodiceAllievo = '"&cod&"' AND Visto = 1;"
		end if
	
	else

	' una notifica in particolare
	
	Id_Notifica = Request.QueryString("IdNotifica")
	rimuovi = Request.QueryString("Rimuovi")
	leggi = Request.QueryString("Leggi")
	ripristina = Request.QueryString("Ripristina")
	
		if rimuovi = 1 then
			QuerySQL = "DELETE FROM AVVISI WHERE ID_Avviso = '"&Id_Notifica&"' AND Visto = 1;"
		else
			if ripristina = 1 then
				QuerySQL = "UPDATE AVVISI SET Visto = 0 WHERE ID_Avviso = '"&Id_Notifica&"';"
			else
				QuerySQL = "UPDATE AVVISI SET Visto = 1 WHERE ID_Avviso = '"&Id_Notifica&"';"
			end if
		end if
	
	end if

else

' tutte le notifiche

rimuovi = Request.QueryString("Rimuovi")
leggi = Request.QueryString("Leggi")

	if rimuovi = 1 then
		QuerySQL = "DELETE FROM AVVISI WHERE CodiceAllievo = '"&cod&"';"
	else
		QuerySQL = "UPDATE AVVISI SET Visto = 1 WHERE CodiceAllievo = '"&cod&"';"
	end if

	
end if

ConnessioneDB.Execute(QuerySQL)

%>

<script>alert("Modifica effettuata correttamente"); window.location.href="centro_messaggi.asp"</script>

<%end if%>

</body>
</html>