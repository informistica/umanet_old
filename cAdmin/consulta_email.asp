<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../stile.css">
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
<title>Email degli studenti</title>
 <script src="../lib/prototype.js" type="text/javascript"></script> 
  <script src="../src/scriptaculous.js" type="text/javascript"></script> 
  <script src="../src/unittest.js" type="text/javascript"></script> 
<script language="../javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

location.href="studente_domande.asp?Classe=<%=Session("Classe")%>&Id_Classe=<%=Session("Id_Classe")%>"

//location.href=window.history.back();
 }
 </script>
  <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")

location.href="../../home.asp"
//location.href=window.history.back();
 }
 </script>
<link href="../../stile.css" rel="stylesheet" type="text/css">
</head>

<%
 
  Dim ConnessioneDB,rsTabella, QuerySQL
 
 'StringaConnessione= Response.Cookies("Dati")("StrConn")
  
  <!--#include file="../admin/gestione_errori.asp" -->
  
Response.Buffer = true 
'On Error Resume Next 
 
 
  if session("CodiceAllievo")="" then%>
	 <BODY onLoad="showText2();"> </BODY>
  <% end if %>
 
   <%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>
    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	
   <body bgcolor="#FFFFFF">
 
<div id="container" class="contenuti_login"><br><br>

  <p align="center">
  
 
 
 
  <form method="POST" form action="consulta_profili1.asp?NumRec=<%=i%>&id_classe=<%=id_classe%>&divid=<%=divid%>" > 
    <p>
      
      <div id="bloc_destra_cont">
      <br>
   
     
	
	<div class="contenuti" style="width: 90%; height: auto" >	
	<p align="center">
    <b><p class="sottotitoloquaderno" style="font-size:18px; font-weight:100" align="center">Classe <%=Session("Cartella")%> </b></font><br><br> 
     <%
  
  id_classe=request.QueryString("id_classe")
  divid=request.QueryString("divid")
  
	QuerySQL="Select Url_img from Classi where ID_Classe='" & id_classe & "';" 
	
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	if not rsTabella.eof then
		urlimg=rsTabella(0)
	else
		urlimg=""
	end if
	urlC= "../../"&"Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/img" ' vuole il percorso relativo della cartella
    urlC=Replace(urlC,"\","/")

  if strcomp(urlimg&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è = a ""
    
  	urlimgclasse=urlC&"/"&"profilo_vuoto.png" %>	
     <img class="imground" src="<%=urlimgclasse%>" > <br>
<%else%>
    <% urlimgclasse=urlC&"/"& urlimg ' aggiungo al percorso il nome del file%>
     <img class="imground" src="<%=urlimgclasse%>" >  <br>
<%end if %>
 
  

<%	

    QuerySQL="Select * from Allievi where Id_Classe='" & id_classe & "' order by cognome asc;" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	classe=rsTabella("Classe")
	 
	  %>
     
      

  <b><p class="sottotitoloquaderno" style="font-size:18px; font-weight:100" align="center">Email degli studenti </b></font></b><p>

    
    
       
     <table border="2" align="center">
 <tr>
 <td align="center">
<a href="#" onClick="Effect.toggle('dCSV','BLIND'); return false;">File CSV per esportare i contatti</a> 
<div id="dCSV" style="display:none;"><div style="background-color:#ffff00;width:570px;border:1px solid red;padding:10px;"> 
 

<% 
  response.write("First Name, Last Name, E-mail Address <br>")
  do while not rsTabella.eof  
    response.write(rtrim(rsTabella("Nome"))& "," & rtrim(rsTabella("Cognome"))&","& rtrim(rsTabella("Email"))&"<br>")
   rsTabella.movenext
  loop 
  
 %>
 
</div></div> 
 
 </td>
 
 </table>
     
     
    
    
    	<div id="bloc_sinistra_login">
    
<div class="contenuti" >	
<br>	
	
   
   


    
    
    
	<%
	
	' aggiungo foto classe
	
	
	
	i=0
	url= "../../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
    url=Replace(url,"\","/")
	rsTabella.movefirst
   do while not rsTabella.eof %>
   
   <hr>
   <p align="left" class="sottotitolo" style="text-align:left"> <%=rsTabella("Cognome") & "  " & rsTabella("Nome")%></p><br>
      
<% if strcomp(rsTabella("Url_img")&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è = a ""
    
  	urlimg=url&"/"&"profilo_vuoto_thumb.png" %>	
    <fieldset style="width:15%"><img class="imground" src="<%=urlimg%>" ></fieldset><br>
<%else%>
    <% urlimg=url&"/"& rsTabella("Url_img") ' aggiungo al percorso il nome del file%>
    <fieldset style="width:24%"> <img class="imground" src="<%=urlimg%>" > </fieldset> <br>
<%end if %>


	 

    
    <b>Email </b></p>
    <input type="text" size="50" name="email<%=i%>" value="<%=rsTabella("Email")%>" > 
   
  <br>
 
  
 <%
   rsTabella.movenext
   i=i+1
  loop 
 %>
    
    <%if session("Admin")=true then%>
    <p><input type="submit" value="Invia" name="B1"> 
 <%  end if%>
</form> <!-- Chiude l'interfaccia -->
   

 
   
</div>


</body>
 
</html>
  
    