<!-- modifica_domande.asp -->
<%@ Language=VBScript %>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag,MO
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceDomanda=Request.QueryString("CodiceDomanda")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("cartella")
  ID=CodiceDomanda 
  Tipodomanda=Request.QueryString("Tipodomanda")
  Quesito=Request.QueryString("Quesito")
  R1=Request.QueryString("R1")
  R2=Request.QueryString("R2")
  R3=Request.QueryString("R3")
  R4=Request.QueryString("R4")
  RE=Request.QueryString("RE")
  MO=Request.QueryString("MO")
  Multiple=Request.QueryString("Multiple") 
  VF=Request.QueryString("VF")  
  if MO<>"" then 
 Modulo=MO
end if  
 
  
Dim objFSO, objTextFile
Dim sRead, sReadLine, sReadAll
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")


   %><!-- #include file = "../var_globali.inc" --><%
  
 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" ' PER ONLINE *******************
 url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"

url=Replace(url,"\","/")
'url=url3

				'dim objFSO,objCreatedFile
				'Const ForReading = 1, ForWriting = 2, ForAppending = 8
				'Dim sRead, sReadLine, sReadAll, objTextFile
	'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url4=Server.MapPath("/anno_2012-2013/log_modifica.txt")
'				Set objCreatedFile = objFSO.CreateTextFile(url4, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.Close 


'response.write(url1)
'response.write(url)
' Open file for reading.
Set objTextFile = objFSO.OpenTextFile(url, ForReading)

' Use different methods to read contents of file.
sReadAll = objTextFile.ReadAll
'response.write(sReadAll)
objTextFile.Close

Function domandaplus()	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url=Replace(url,"\","/")
	 
	' Open file for reading.
	Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	domandaplus = objTextFile.ReadAll
	'domandaplus=url
	'response.write(sReadAll)
	objTextFile.Close
End Function 
 

   
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
<title>Modifica Domanda</title>

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

<form method="POST" form action="inserisci_modifica1.asp?VF=<%=VF%>&Tipodomanda=<%=Tipodomanda%>&Cartella=<%=Cartella%>&CodiceDomanda=<%=CodiceDomanda%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&MO=<%=MO%>&Multiple=<%=Multiple%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
  <div id="bloc_destra_cont"> 
 <b><font color=#FF0000 size="4"><%Response.write (Cognome) %>&nbsp<%Response.write (Nome)%></font></b><br><br>
  <font size="4" color="#FF0000"><b>Modifica la</b></font><b><font color=#FF0000 size="4"> domanda del test :</font></b>
  <br><p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Capitolo) %></b></font>  <!-- stampa il titolo del test -->
	<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (Paragrafo) %></b></font> <!-- stampa il titolo del test -->
	
    <p>
	<!--<div id="bloc_sinistra_login">-->
	<div class="contenuti_login" style="width: 960px; height: auto;" >	
	 
		<p><font size="4" color="#FF0000">Codice Domanda (<%=CodiceDomanda%>)</font><br>
  </p>

  <p><input type="text" name="txtDomanda"  value="<%=Quesito%>" size="135" maxlength="250"><b> 
	Domanda 
	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
	<%if Tipodomanda=1 then %>
	   <br>
	   <textarea rows="6" name="TestoDomandaPlus" value="ciao" cols="96"><%=Response.write(domandaplus())%> </textarea><br>		
	<% end if%>
    <%if VF=0 then ' non è una domanda vero falso %>
          <p><input type="text" name="txtR1" value="<%=R1%>" size="135" maxlength="150"><b> 
            Risposta 1</b></p> 
          <p>
            <input type="text" name="txtR2" value="<%=R2%>" size="135" maxlength="150"><b> 
            Risposta 2 </b></p>
          <p>
            <input type="text" name="txtR3" value="<%=R3%>" size="135" maxlength="150"><b> 
            Risposta 3 </b></p>
          <p><input type="text" name="txtR4" value="<%=R4%>" size="135" maxlength="150"><b> 
            Risposta 4 </b></p>
              <p><input type="text" name="txtRE" value="<%=RE%>" size="1"><b> 
            Risposta Esatta  </b></p>
    
         
     <%else ' è vero falso%>
           <p><input type="text" name="txtRE" value="<%=RE%>" size="1"><b> 
            Risposta Esatta (0=Falso/1=Vero) </b></p>
    
             
	 <%end if%>
     
	<b>Spiegazione</b><p><textarea rows="6" name="S1" value="ciao" cols="96"><%=Response.write(sReadAll)%> </textarea></p>

  <p><input type="submit" value="Invia" name="B1"></p> <!--Definisce i due bottoni del form -->
</form> <!-- Chiude l'interfaccia -->
</div>
</div>
</div>
</div>
</body>
</html>