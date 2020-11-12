<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  						
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
  
  'Request.Cookies("Dati")("CodiceTest")= Codice_Test
  
  Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo") ' TOPOLINO ED OBIETTIVI
  ParagrafoNuovo="Navigazione nella Rete della Vita" ' VALORE CABLATO
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  'Quesito=Request.QueryString("Quesito")
  'prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella")
  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = cint(Request.QueryString("Num"))
  Num=Num+1
  daTopolino=Request.QueryString("daTopolino") ' vale 1 se sono stata chiamata da spiegazione_metafora_topolino tramite invia a -->
  CodiceMetafora=Request.QueryString("CodiceMetafora")
  if daTopolino=1 then
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <%  
   '' 
   ' eseguo la query con codiemetafora per ricreare tutti i parametri sopra che non ho tramite QueryString
    QuerySQL="Select * from M_Topolino where CodiceMetafora=" & cint(CodiceMetafora)& ";"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	'Privato=rsTabella.fields("Privato") 
	'rsTabella.close
     Cartella=rsTabella.fields("Cartella")
	 'Codice_Test=rsTabella.fields("Id_Arg")
	 Modulo=rsTabella.fields("Id_Mod")
	   
  end if
  
    
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
<title>Inserisci Metafora Patente</title>

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
	 <BODY onLoad="showText2();">
	
     </BODY>
<% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
<div id="container">

<form method="POST" form action="inserisci_metafora_patente1.asp?prenodo=<%=prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=ParagrafoNuovo%>&Modulo=<%=Modulo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <div id="bloc_destra_cont">
 <b><font color=#FF0000 size="4"><%Response.write (Cognome) %>&nbsp<%Response.write (Nome)%></font></b><br><br>
  <b><font color=#FF0000 size="4">Inserisci metafora  :</font></b>
  <br>
  <p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Capitolo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (ParagrafoNuovo) %></b></font> <!-- stampa il titolo del test -->
	
    
	<!--<div id="bloc_sinistra_login">-->
		<div class="contenuti_login" style="width: 690px; height: auto;">	
           <img class="imground" src="../../img/patente_sint/metafora_1.jpg" width="130px" height="130px">
           <img class="imground"  src="../../img/patente_sint/metafora_2.jpg" width="130px" height="130px">
           <img class="imground" src="../../img/patente_sint/metafora_3.jpg" width="130px" height="130px">
           <img class="imground" src="../../img/patente_sint/metafora_4.jpg" width="130px" height="130px">
           <img class="imground" src="../../img/patente_sint/metafora_5.jpg" width="130px" height="130px"><br><br>
           <center>
           <img class="imground"  src="../../img/patente_sint/metafora_6.jpg" width="130px" height="130px">
           </center>
           
  
	<p align="left">&nbsp;<font size="4" color="#FF0000">Metafora (N : 
	    <%response.write(Num)%>)</font><br>
  </p>
 
 <% Lupo="" 
   if daTopolino=1 then %>
   <% Lupo=rsTabella.fields("Testata")%>
    <p><input type="text" name="txtAutista"  size="100" maxlength="250" value="<%=rsTabella.fields("Topolino")%>">
  <b> Autista</b>
  </p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
   
  <p><input type="text" name="txtDestinazione" value="<%=rsTabella.fields("Formaggio")%>" size="100" maxlength="150">
  <b> Destinazione</b></p> 
	
  <p>
	<input type="text" name="txtCarburante" value="<%=rsTabella.fields("Fame")%>" size="100" maxlength="150">
	<b> Carburante </b></p>
  
  <p>
	<input type="text" name="txtLuogo" value="<%=rsTabella.fields("Labirinto")%>" size="100" maxlength="150">
	<b> Luogo </b></p>
  
   <p>
<input type="text" name="txtStrada" value="<%=rsTabella.fields("Strada")%>" size="100" maxlength="150">
	<b> Strada </b></p>
   <p>
<input type="text" name="txtStrada_ko" value="<%=rsTabella.fields("Strada_KO")%>"size="100" maxlength="150">
	<b> Strada KO </b></p>
    
 <p>
	<input type="text" name="txtStrada_ok" value="<%=rsTabella.fields("Strada_OK")%>" size="100" maxlength="150">
	<b> Strada OK </b></p>
	 
   <% rsTabella.close %>
 <%else%>
 
 <p><input type="text" name="txtAutista"  size="100" maxlength="250" value="">
  <b> Autista</b>
  </p> 
  <!-- crea la variabile di tipo inputbox avente un certo nome -->
   
   
  <p><input type="text" name="txtDestinazione" value="" size="100" maxlength="150">
  <b> Destinazione</b></p> 
	
  <p>
	<input type="text" name="txtCarburante" value="" size="100" maxlength="150">
	<b> Carburante </b></p>
  
  <p>
	<input type="text" name="txtLuogo" value="" size="100" maxlength="150">
	<b> Luogo </b></p>
  
   <p>
<input type="text" name="txtStrada" value="" size="100" maxlength="150">
	<b> Strada </b></p>
   <p>
<input type="text" name="txtStrada_ko" value="" size="100" maxlength="150">
	<b> Strada KO </b></p>
    
 <p>
	<input type="text" name="txtStrada_ok" value="" size="100" maxlength="150">
	<b> Strada OK </b></p>
	
	<%end if ' chiudo l'if che distingue se devo riepire il form con i dati della query%>
	 
	 <p>
	<input type="text" name="txtCespugli" value="" size="100" maxlength="150">
	<b> Cespugli </b></p>
	
	 <p>
	<input type="text" name="txtLupo" value="<%=Lupo%>" size="100" maxlength="150">
	<b> Lupo </b></p>
	 <p>
	<input type="text" name="txtCestino" value="" size="100" maxlength="150">
	 <b> Cestino </b></p>
	 <p>
	<input type="text" name="txtDistanza" value="" size="1" maxlength="4">
	<b> Distanza </b></p>
	<% 'Prelovo la spiegazione della metafora topolino, che qua andrà estesa 
	      'ID=CodiceMetafora
  			  'CARTA=rsTabella.fields("Cartella")
				url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&CodiceMetafora&".txt" 'per il server on line
				 url=Replace(url,"\","/")
	 		   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
				on error resume next
				 If Err.Number <> 0 Then
					Response.Write Err.Description 
					Err.Number = 0
				 sReadAll="File della spiegazione mancante" & "<br>" & url
				 else
				' Use different methods to read contents of file.
				sReadAll = objTextFile.ReadAll
				'sReadAll=url
				    Err.Number = 0
				End If
				objTextFile.Close
   
	%> 
	 
	<b>Sintesi</b>
	<p><textarea rows="6" name="S1" cols="75"><%=sReadAll%></textarea></p>
  <p><input type="submit" value="Invia" name="B1"><input type="reset" value="Reimposta" name="B2"></p> <!--Definisce i due bottoni del form -->
</form> <!-- Chiude l'interfaccia -->
</div>
</div>
</div>
</div>
</body>
</html>