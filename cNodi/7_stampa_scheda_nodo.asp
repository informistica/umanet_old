<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
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
<title>Stampa Nodo</title>
<script type="text/javascript">
window.onload=function() {
window.print();
}
</script>
<script language="javascript" type="text/javascript"> 
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

location.href="studente_domande.asp"
//location.href=window.history.back();
 }
 </script>
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

<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO
   Dim ConnessioneDB, rsTabella,QuerySQL,Privato ' varrà 1 se ogni stud legge solo il suo materiale
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <%  
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  
  
  
  
  
  
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  'cla=Request.QueryString("cla")
 ' Codice_Test=Request.QueryString("CodiceTest")
  CodiceNodo=Request.QueryString("CodiceNodo")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Chi=Request.QueryString("Chi")
  Cosa=Request.QueryString("Cosa")
  Dove=Request.QueryString("Dove")
  Quando=Request.QueryString("Quando")
  Come=Request.QueryString("Come")
  Perche=Request.QueryString("Perche")
  Quindi=Request.QueryString("Quindi")
 ' MO=Request.QueryString("MO")
  VAL=Request.QueryString("VAL")
 ' URL=Request.QueryString("URL")
  DATA=cdate(Request.QueryString("DATA"))
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  ID=CodiceNodo
  Cartella=Request.QueryString("Cartella")
  

Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
url=Replace(url,"\","/")
Set objTextFile = objFSO.OpenTextFile(url, ForReading)
sReadAll = objTextFile.ReadAll
'sReadAll=url
'response.write(sReadAll)
objTextFile.Close

 

   
%>


<div id="container">  
 <div id="bloc_destra_cont" class="contenuti_login">

	
    
	 
 	
 <table border="1"  align="center" width="60%" id="blugradient1">
    <tr> <td colspan="3" align="center"><%=Cognome & " " & Nome%></td></tr>
    <tr>
      <td width="13%"><b>Nodo n</b>.<%=CodiceNodo%></td>
      <td width="18%"><%=Capitolo%></td><td width="18%"><%=Paragrafo%></td>
    </tr> <!-- visualizzo i livelli cosa,dove,quando,...-->
    <tr><td><b> Chi </b></th><td colspan=3><p align="center"><b><%=Chi%></b></td></tr>
    <tr><td><b> Cosa </b></th><td colspan=2><p align="center"><%=Cosa %></td></tr>
    <tr><td><b> Dove </b></th><td colspan=3><p align="center"><%=Dove %></td></tr>
    <tr><td><b> Quando </b></th><td colspan=3><p align="center"><%=Quando %></td></tr>
    <tr><td><b> Come </b></th><td colspan=3><p align="center"><%=Come %></td></tr>
    <tr><td><b> Perchè </b></th><td colspan=3><p align="center"><%=Perche%></td></tr>
    <tr><td><b> Quindi </b></th><td colspan=3><p align="center"><%=Quindi%></td></tr>
    <tr>
    <td colspan=3>
    <p align="center">
     <textarea rows="<%=1+round((len(sReadAll))/50)%>" name="TestoDomandaPlus" value="ciao" cols="80"><% 
     Response.write(sReadAll)%> </textarea><br>
    </td>
    </tr>
  </table>
</div>
</div>
</div>
</div>
</body>
 
</html>