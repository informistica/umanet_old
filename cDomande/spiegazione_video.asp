<%@ Language=VBScript %>
<% dim video,CodiceAllievo,ID_Mod,ID_Paragrafo,CodiceDomanda,QuerySQL,ConnessioneDB, StringaConnessione
 
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
   %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <%


' inserisco nella tabella la visualizzazione effettuata dall'utente
video=Request.QueryString("video") 
ID_Mod=Request.QueryString("ID_Mod") 
ID_Paragrafo=Request.QueryString("ID_Paragrafo") 
CodiceDomanda=Request.QueryString("CodiceDomanda") 
davisualizzazioni=Request.QueryString("davisualizzazioni")
if davisualizzazioni="" then
QuerySQL="  INSERT INTO Visualizzazioni (CodiceAllievo, ID_Mod, ID_Paragrafo,CodiceDomanda,Data) SELECT '" & Session("CodiceAllievo") & "','" & ID_Mod & "', '" & ID_Paragrafo & "','" & CodiceDomanda & "','" & now() & "';"
ConnessioneDB.Execute QuerySQL
end if 
   

%>
<HTML>
<head>
 <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
<TITLE>VIDEO DELLA RETE</TITLE>
</HEAD>
<%Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
<BODY> 
<center><OBJECT CLASSID="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" WIDTH="919" HEIGHT="582" CODEBASE="https://active.macromedia.com/flash5/cabs/swflash.cab#version=7,0,0,0">
<PARAM NAME=movie VALUE="<%=video&".swf"%>">
<PARAM NAME=play VALUE=true>
<PARAM NAME=loop VALUE=false>
<PARAM NAME=wmode VALUE=transparent>
<PARAM NAME=quality VALUE=low>
<EMBED SRC="<%=video&".swf"%>" WIDTH=919 HEIGHT=582 quality=low loop=false wmode=transparent TYPE="application/x-shockwave-flash" PLUGINSPAGE="https://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash">
</EMBED>
</OBJECT></center>
<SCRIPT src='video.js'></script>
</BODY>
</HTML>
