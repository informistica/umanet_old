<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
 TipoEvento=Request.QueryString("TipoEvento")
  if strcomp(TipoEvento&"","")=0 then
	     TipoEvento=1
	  end if
 ' TipoEvento=1
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
<title>Inserisci Metafora Data Base dei Desideri</title>

<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 
 state=1
 function scambia() {
	// alert("ciao");
	// alert(document.inserisci.txtTipoEvento.value);
	if (document.inserisci.txtTipoEvento.value = 1 ) 
	{
		document["rappresentazione"].src = "../img/clienteostesi.jpg";
		document.inserisci.txtTipoEvento.value=0;
	} 
	else 
	{
		document["rappresentazione"].src = "../img/clienteosteno.jpg";  
		document.inserisci.txtTipoEvento.value=1;
	}
}



immagini = new Array(2); 
immagini[0]="../img/clienteosteno.jpg";
immagini[1]="../img/clienteostesi.jpg";
function scambia_n(cont) {
	if (cont = 1) alert ("Attenzione per simulare il terremoto devi inserire un evento paradossale che abbia significati in contrasto, la risposta del client deve deludere l'aspettativa del server!");
	document["rappresentazione"].src = immagini[cont];
}
 </script>

</head>
<!-- #include file = "../service/controllo_sessione.asp" -->
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

<form method="POST" name="document" form action="inserisci_metafora_dbdesideri1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <div id="bloc_destra_cont">
 <b><font color=#FF0000 size="4"><%Response.write (Cognome) %>&nbsp<%Response.write (Nome)%></font></b><br><br>
  <b><font color=#FF0000 size="4">Inserisci metafora  :</font></b>
  <br>
  <p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Capitolo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (Paragrafo) %></b></font> <!-- stampa il titolo del test -->
	
    <p>
	<!--<div id="bloc_sinistra_login">-->
		<div class="contenuti_login" style="width: 850px; height: auto;">	
	<p align="left">&nbsp;</p>&nbsp;<p><font size="4" color="#FF0000">Metafora (N : 
	    <%response.write(Num)%>)</font><br>
  </p>
   
     <fieldset><legend>Rappresentazione</legend>
     <br>
     <b> Tipo di Evento </b></p>
     <select name="txtTipoEvento" style="width:auto" onChange="scambia_n(txtTipoEvento.value);" disabled="true"> 
		<%if TipoEvento=0 then %>
           <option selected  value="1">Coerente
         <option   value="0">Paradossale
         <%else%>
          <option   value="1">Coerente
         <option selected  value="0">Paradossale
        
         <%end if%>
	 </select> <br><br>
    
    <center>
<% if TipoEvento=0 then%>
      <img src="../../img/clienteostesi.jpg" name="rappresentazione" width="500px" height="300px">
  <%else %>
   <img src="../../img/clienteosteno.jpg" name="rappresentazione" width="500px" height="300px">
     
  <%end if%>
</center><br>
  
 <fieldset><legend>Client</legend>
 <p><input type="text" name="txtSoggettoC"  size="100" maxlength="250" value="">
  <b> Soggetto</b>
  </p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
   
  <p><input type="text" name="txtDomandaC" value="" size="100" maxlength="150">
  <b> Domanda</b></p> 
	
  <p>
	<input type="text" name="txtMotivazioneC" value="" size="100" maxlength="150">
	<b> Motivazione </b></p>
  
  <p>
	<input type="text" name="txtDesiderioC" value="" size="100" maxlength="150">
	<b> Desiderio </b></p>
  
   <p>
<input type="text" name="txtBisognoC" value="" size="100" maxlength="150">
	<b> Bisogno </b></p>
   <p>
</fieldset>
<fieldset><legend>Server</legend>
 <p><input type="text" name="txtSoggettoS"  size="100" maxlength="250" value="">
  <b> Soggetto</b>
  </p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
   
  <p><input type="text" name="txtRispostaS" value="" size="100" maxlength="150">
  <b> Risposta</b></p> 
	
  <p>
	<input type="text" name="txtMotivazioneS" value="" size="100" maxlength="150">
	<b> Motivazione </b></p>
  
  <p>
	<input type="text" name="txtDesiderioS" value="" size="100" maxlength="150">
	<b> Desiderio </b></p>
  
   <p>
<input type="text" name="txtBisognoS" value="" size="100" maxlength="150">
	<b> Bisogno </b></p>
   <p>
</fieldset>
	<br> 
    <br>
    <input type="text" name="txtTolleranzaC" value="" size="2" maxlength="1">
	<b> Tolleranza del Client (n. da 0 a 9) </b></p>
    <br>
    
    
    
	 
	<b>Sintesi</b>
	<p><textarea rows="6" name="S1" cols="75"></textarea></p>
  <p><input type="submit" value="Invia" name="B1"><input type="reset" value="Reimposta" name="B2"></p> <!--Definisce i due bottoni del form -->
</form> <!-- Chiude l'interfaccia -->
</div>
</div>
</div>
</div>
</body>
</html>