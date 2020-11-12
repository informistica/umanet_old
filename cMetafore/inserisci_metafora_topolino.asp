<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
  
  'Request.Cookies("Dati")("CodiceTest")= Codice_Test
  
  Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito")
  prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella")

  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = Request.QueryString("Num")
  Num=Num+1
  
    
%>
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
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inserisci Metafora Topolino</title>

<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>

</head>

<%
  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
<div id="container">

<form method="POST" form action="inserisci_metafora_topolino1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <div id="bloc_destra_cont">
 <b><font color=#FF0000 size="4"><%Response.write (Cognome) %>&nbsp<%Response.write (Nome)%></font></b><br><br>
  <b><font color=#FF0000 size="4">Inserisci metafora  :</font></b>
  <br>
  <p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Capitolo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (Paragrafo) %></b></font> <!-- stampa il titolo del test -->
	
    <p>
	<!--<div id="bloc_sinistra_login">-->
		<div class="contenuti_login" style="width: 700px; height: auto;">	
	<p align="left">&nbsp;</p>&nbsp;<p><font size="4" color="#FF0000">Nodo (N : 
	    <%response.write(Num)%>)</font><br>
  </p>
 
 <p><input type="text" name="txtTopolino"  size="100" maxlength="250" value="">
  <b> Topolino</b>
  </p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
   
  <p><input type="text" name="txtFormaggio" value="" size="100" maxlength="150">
  <b> Formaggio</b></p> 
	
  <p>
	<input type="text" name="txtFame" value="" size="100" maxlength="150">
	<b> Fame </b></p>
  
  <p>
	<input type="text" name="txtLabirinto" value="" size="100" maxlength="150">
	<b> Labirinto </b></p>
  
   <p>
<input type="text" name="txtStrada" value="" size="100" maxlength="150">
	<b> Strada </b></p>
   <p>
<input type="text" name="txtStrada_ko" value="" size="100" maxlength="150">
	<b> Strada KO </b></p>
    
 <p>
	<input type="text" name="txtStrada_ok" value="" size="100" maxlength="150">
	<b> Strada OK </b></p>
	 <p>
	<input type="text" name="txtTestata" value="" size="100" maxlength="150">
	<b> Testata </b></p>
	 <p>
	<input type="text" name="txtDistanza" value="" size="1" maxlength="4">
	<b> Distanza </b>(numero da 1 a 5)</p>
	 
	<b>Sintesi</b>
	<p><textarea rows="6" name="S1" cols="75">Inserisci una descrizione a parole della situazione. La descrizione deve includere le parole chiave inserite come parametri.</textarea></p>
  <p><input type="submit" value="Invia" name="B1"><input type="reset" value="Reimposta" name="B2"></p> <!--Definisce i due bottoni del form -->
</form> <!-- Chiude l'interfaccia -->
</div>
</div>
</div>
</div>
</body>
</html>