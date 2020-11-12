<!-- login.asp -->

<html>
<head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<style>
<!--
 li.MsoNormal
	{mso-style-parent:"";
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	margin-left:0cm; margin-right:0cm; margin-top:0cm}
-->
</style>

<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Login</title>

</head>

<body>
<div id="container">
<div id="bloc_destra_cont">
<div id="bloc_sinistra_login">
<div class="contenuti_login" >

<H3><font size="5">Modifica email  allievo</font></H3>
<%  
    Dim ConnessioneDB, rsTabella, QuerySQL, CodiceTest, CodiceAllievo,CodiceAllievo1,PwdAllievo, CodiceCorso, CodiceCap,i,StringaConnessione,TitoloTest
    Dim Cognome,Nome
     
    
 
     
    'CodiceAllievo = Request.Form("txtCodiceAllievo")
'	CodiceAllievo="" ' per ripristinare l'utente vuoto
	CodiceAllievo = Request("txtCodiceAllievo")
	'NewEmail = Request.Form("txtNewEm")
	NewEmail = Request("txtNewEm")
	 
 '  StringaConnessione=Request.QueryString("StringaConnessione")
  '  Response.Cookies("Dati")("StrConn") = StringaConnessione
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") ' Assegna alla variabile connessione il risultato del metodo CreateObject("Tipo di connessione") dell'oggetto Server
   	 
   
     %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   
 
 <%
  CodiceAllievo1 = Replace(CodiceAllievo, "'", "''")  

		QuerySQL ="UPDATE Allievi SET Allievi.Email = '" &NewEmail& "'" &_
		" WHERE Allievi.CodiceAllievo= '"&CodiceAllievo & "'"
	    ConnessioneDB.Execute(QuerySQL)
		 
		'response.write(QuerySQL)
		if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
	 
 
%>
 
  </div>
  </div>
  </div>
  </div>
 </body>
</html>