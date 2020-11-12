<%@ Language=VBScript %>

<% db = Request.QueryString("DB") 'questo è il database a cui è connesso attualmente l'utente
opposto = Request.QueryString("Opposto")
%>

<!doctype html>
<html>
<head>
<title>Errore DB</title>
</head>
<body>

<script>
alert("Stai tentando di accedere alla Sezione <%=opposto%>, ma risulti connesso nella Sezione <%=db%>. Effettua il Logout da quest'utlima e riprova");
window.location.href = "../../home.asp"
</script>

</body>
</html>