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
    <body>
  <% end if %>

    <div id="container">
	<div class="contenuti_login" style="height:auto; width:auto">
	<font color=#FF0000 size="4">


  
   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,url4,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   Cartella=Request.QueryString("Cartella")
   TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  'response.write("TP:"&TitoloParagrafo)
   Modulo=Request.QueryString("Modulo")
   

   'StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
    
   Paragrafo=Request.QueryString("Paragrafo")
   NumRec=clng(Request.Form("TxtNUMREC"))
   Titolo=Request.Form("TitoloParagrafo")
  ' response.write("NUMREC" & numrec)
for k=0 to NumRec ' per scorrere tutto il form e fare un update ad ogni ciclo
   ID=Request.Form("txtCodiceNodo"&k)
   Chi = Request.Form("txtChi"&k)
   Chi = Replace(Chi, Chr(34), "'")
   Chi=Replace(Chi,"'","''")
   Cosa = Request.Form("txtR1Cosa"&k)
   Cosa = Replace(Cosa, Chr(34), "'")
   Cosa=Replace(Cosa,"'","''")
   Dove = Request.Form("txtR1Dove"&k)
   Dove = Replace(Dove, Chr(34), "'")
   Dove=Replace(Dove,"'","''")
   Quando = Request.Form("txtR1Quando"&k)
   Quando = Replace(Quando, Chr(34), "'")
   Quando=Replace(Quando,"'","''")
   Come = Request.Form("txtR1Come"&k)
   Come = Replace(Come, Chr(34), "'")
   Come=Replace(Come,"'","''")
   Perche = Request.Form("txtR1Perche"&k)
   Perche = Replace(Perche, Chr(34), "'")
   Perche=Replace(Perche,"'","''")
   Quindi = Request.Form("txtR1Quindi"&k)
   Quindi = Replace(Quindi, Chr(34), "'")
   Quindi=Replace(Quindi,"'","''")
   Sintesi=Request.Form("S1"&k)
   Sintesi= Replace(Sintesi, Chr(34), "'")
   Sintesi=Replace(Sintesi,"'","''")
   Spiegazione=Request.Form("S1"&k)
   DATA=cdate(Request.Form("txtDATA"&k))
   VAL=clng(Request.Form("txtVAL"&k))
   Voto=VAL
   Segnalata=clng(Request.Form("txtSegnalata"&k))
   if (Request.Form("txtINQUIZ"&k)<>"") then
      INQUIZ=clng(Request.Form("txtINQUIZ"&k))
   end if 
   
 
 
	 if (session("Admin")=True)  then 
  QuerySQL ="UPDATE Nodi SET Chi = '" & Chi & "', Cosa= '" & Cosa & "',Dove= '" & Dove & "',Quando= '" & Quando & "', Come= '" & Come & "', Perche= '" & Perche & "', Quindi = '" & Quindi & "', Voto = '" & voto & "',Segnalata = '"&Segnalata&"', In_Quiz = '" & INQUIZ &"',Data='" & DATA & "'  WHERE CodiceNodo =" &ID&";"
 
  	
else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
   QuerySQL ="UPDATE Nodi SET Chi = '" & Chi & "', Cosa= '" & Cosa & "',Dove= '" & Dove & "',Quando= '" & Quando & "', Come= '" & Come & "', Perche= '" & Perche & "', Quindi = '" & Quindi & "'  WHERE CodiceNodo =" &ID&";"
  
   end if 
   
end if
	 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&TitoloParagrafo&"_"&ID&".txt"  'per server on-line
	' url=Server.MapPath("/ECDL/")& "/" & Cartella &"/" &Modulo&"_Nodi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" ' per localhost
  
		'url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	
	url3=Replace(url,"\","/")
	url=url3
 
 ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySQL) %> <br> <%
'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

	Dim objFSO,objCreatedFile
	'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'CANCELLA LA VECCHIA VERSIONE DEL FILE11
	'response.write(Cartella)
	'response.write(url)
	'objFSO.DeleteFile url
	'Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
	'objCreatedFile.WriteLine(Spiegazione)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	'objCreatedFile.Close



next 
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Modifica avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If



Response.Redirect "../cClasse/home_app.asp?id_classe="&Session("Id_Classe")&"&divid="&Session("divid")
 




   %>
	</font>   
	 
		
      <h4><a href="../cClasse/studente_domande.asp?DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>&cla=<%=cla%>&CodiceAllievo=<%=CodiceAllievo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua 
		a valutare o modificare le domande...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Apprendimento... </a></h3> 
			
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	