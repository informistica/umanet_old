<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%
 
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag
  Id_Classe=Request.QueryString("Id_Classe")
  divid=request.QueryString("divid")
  mod_umanet=request.QueryString("mod_umanet")
  Classe=Request.QueryString("Classe")
  
 ' posizione= Request.QueryString("posizione")
 ' response.Write("jhjhk="&posizione)
  Titolo = Request.Form("TxtTitolo")
  Num = Request.Form("TxtNum") ' numero di paragrafi che si vogliono inserire
  ID_Mod=Request.Form("txtID_Mod")
  Dim ConnessioneDB , rsTabella,QuerySQL

   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")   
'Session("DB2")=1   
%> 
 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
 
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
<title>Inserisci Modulo Umanet </title>

</head>
<body bgcolor="#FFFFFF">
<div id="container">

<% if num<>"" then %>
<form method="POST" form action="inserisci_modulo_umanet1.asp?Id_Classe=<%=Id_Classe%>&ID_Mod=<%=ID_Mod%>&Titolo=<%=Titolo%>&Classe=<%=Classe%>&Num=<%=Num%>&divid=<%=divid%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
		 
		 <br>
		   <FIELDSET style="width:auto"><LEGEND><b>
           Inserisci i titoli dei paragrafi del modulo:<font color=#FF0000 size="4"> <%Response.write (Titolo & " " & ID_Mod) %> 
           </font></b></LEGEND>
		  
			<p>
			  
		 <% for k=1 to Num%>
		  ID <input type="text" name="txtId<%=k%>" value="<%=ID_Mod%>_<%=k%>" size="9" maxlength="7" >
          <input type="text" name="txtDomanda<%=k%>" value="" size="135" maxlength="150" tabindex="<%=k%>">
		  <b>Paragrafo <%=k%> </b></p> 
		   Pagine
          <%for j=1 to 18%>
          <input type="text" name="txtPg<%=k%>_<%=j%>" value="" size="3" maxlength="5" >
          <%next%>
          <p>
			<%next %>
		  <p><input type="submit" value="Invia" name="B1"><input type="reset" value="Reimposta" name="B2"></p> <!--Definisce i due bottoni del form -->
		</form> <!-- Chiude l'interfaccia -->
		 </fieldset>
		 

<% else%>

 
<p class="titolo">

   <form method="POST" form action="inserisci_modulo_umanet.asp?Id_Classe=<%=Id_Classe%>&Classe=<%=Classe%>&divid=<%=divid%>" >
    <FIELDSET><LEGEND><b><%  response.write("Configurazione moduli umanet  ") %> </b></LEGEND><br>
   
    
    
     <b>Inserisci Nuovo Modulo</b><br><br>
	<%response.write("ID del modulo ?")
	  QuerySQL="SELECT max(posizione) FROM MODULI_UMANET1 where Cartella='"&Session("Cartella")&"';"
	  set rsTabella1=ConnessioneDB.Execute(QuerySQL)  
	  if isnull(rsTabella1(0)) then
	    maxPos=0
	  else
	      maxPos=rsTabella1(0)
	  end if
	  posizione=maxPos+1
	'InStrRev([inizio,]stringa1,stringa2[,compara])
	%> 
   <input type="text"  name="txtID_Mod" size="7" value="<%=classe%>_U_<%=posizione%>" title="Incrementare di 1 l'indice del modulo"> 
	<p><%response.write("Titolo del modulo ?")%> 
    <input type="text" name="txtTitolo" size="50"></p>
	<% response.write("Quanti paragrafi vuoi inserire in questo modulo ?") %> 
    <input type="text" name="txtNum" size="1"></p>
    <p><input type="submit" value="Invia" name="B1"></p>
 
	</p>
 
    
    </FIELDSET>
    </form>
    
    <!-- Per utilizzare un modulo già esistente in altra classe -->
    <fieldset><legend>Trasferisci modulo</legend> 
    <!-- Per utilizzare un modulo già esistente in altra classe -->
    <div style="overflow:scroll; height:300px;">
    <iframe src="seleziona_origine.asp?byUmanet=1" name="postmessage" id="postmessage" width="100%" height="100%" frameborder="0" SCROLLING="no" border="0" class="iframe"></iframe>
    </div>
  </fieldset>  
    
   <%end if%> 
	 
</div>
</body>
</html>