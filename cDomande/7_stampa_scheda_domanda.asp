<%@LANGUAGE="VBSCRIPT"%>
 <% Set ConnessioneDB = Server.CreateObject("ADODB.Connection")%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
      
 
<html>
<head>
<title>Stampa compiti</title>
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<META https-equiv=Content-Type content="text/html; charset=iso-8859-1">
 <META https-equiv=Content-Type content="text/html; charset=iso-8859-1">
<LINK media=screen href="../../stile.css" type=text/css rel=stylesheet>
<LINK media=print href="../../stile.css" type=text/css rel=stylesheet>
 
<script type="text/javascript">
window.onload=function() {
window.print();
}
</script>
<!--<link rel="stylesheet" type="text/css" href="../stile.css">-->
</head>

<body>
<%
Function domandaplus()	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	    url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	objTextFile.Close
End Function
 
domande=Request.QueryString("domande")
CodiceAllievo=request.QueryString("CodiceAllievo")
CodiceDomanda=request.QueryString("CodiceDomanda")
Capitolo=request.QueryString("Capitolo")
Paragrafo=request.QueryString("Paragrafo")
Cartella=request.QueryString("Cartella")
Dim risposte(5)



 QuerySQL="SELECT Domande.CodiceDomanda, Domande.Quesito, Domande.Risposta1, Domande.Risposta2, Domande.Risposta3, Domande.Risposta4, Domande.RispostaEsatta, Allievi.Cognome, Allievi.Nome, Moduli.Titolo as [Mod], Paragrafi.Titolo as [Par], Domande.Voto,Moduli.ID_Mod, Paragrafi.ID_Paragrafo,Domande.Data,Domande.Tipo " &_
" FROM Moduli INNER JOIN (Paragrafi INNER JOIN (Allievi INNER JOIN Domande ON Allievi.CodiceAllievo = Domande.Id_Stud) ON Paragrafi.ID_Paragrafo = Domande.Id_Arg) ON Moduli.ID_Mod = Domande.Id_Mod " &_
" WHERE (((Domande.CodiceDomanda)=" & CodiceDomanda & "));"
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
risposte(1)=rsTabella("Risposta1")
risposte(2)=rsTabella("Risposta2")
risposte(3)=rsTabella("Risposta3")
risposte(4)=rsTabella("Risposta4")
  
  %>
 
<body bgcolor="#FFFFFF">
<div id="container">

 <form method="POST"> <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 <div id="bloc_destra_cont">
 
 
	<!--<div id="bloc_sinistra_login">-->
	<div class="contenuti_login" style="width: auto; height: auto" >	
	<p align="center"> 
  
  
  
    
    <%
	
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
    url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &rsTabella("ID_Mod")&"_Spiegazioni/"&rsTabella("ID_Mod")&"_"&rsTabella("Par")&"_"&CodiceDomanda&".txt"
    url=Replace(url,"\","/")
    Set objTextFile = objFSO.OpenTextFile(url, ForReading)
    sReadAll = objTextFile.ReadAll
	'sReadAll=url
	'response.write(sReadAll)
	 objTextFile.Close
	
	%>
     
	 



  <table border="1"  align=center width="60%">
		<tr>
			<td><font size="-2"><%=rsTabella("Mod") %></font></td>
            <td><font size="-2"><%=rsTabella("Par")  %></font></td>
			<td><font size="-2"><%=rsTabella("Cognome")%></font></td>
			<td><font size="-2"><%=rsTabella("CodiceDomanda")%></font></td>
            <td><font size="-2"><%=rsTabella("Data")%></font></td>
		</tr>
		<tr>
			<td colspan=5>
			<p align="center"><b><%=rsTabella("Quesito")%></b></td>
		</tr>
        <tr>
			<td colspan=5>
			<p align="center"><b><%=risposte(rsTabella("RispostaEsatta"))%></b></td>
		</tr>
        
		
		<% if (rsTabella.Fields("Tipo")=1 ) then ' inserisco domanda plus leggendola dal file  altrimenti domanda normale %>
	    <tr><td colspan="3"><p align="center">
			 <textarea rows="<%=1+round((len(domandaplus()))/50)%>" name="TestoDomandaPlus0" value="ciao" cols="100"><%
			 
			 
			 Response.write(domandaplus())%> </textarea><br></td></tr><br>
        <%end if %>
   
		<tr>
			<td colspan=3>
			
			<p align="center">
			 <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" cols="100"><%
			 ' if clng(rsTabella(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			 
			 Response.write(sReadAll)%> </textarea><br>
		      </td>
		 
		</tr>
	</table>

 </div>
 </div>
<%rsTabella.close()
set rsTabella=nothing%>
</body>
</html>
