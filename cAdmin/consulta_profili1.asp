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
<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
<%if session("DB")=1 then %>
Location.href="../../home.asp"
	
  <%else%>
Location.href="../../home.asp"
  <%end if%>

//location.href=window.history.back();
 }
 </script>

</head>

<%
  Response.Buffer = true
  On Error Resume Next  
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body>
  <% end if %>
  
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">


  
   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,url4,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Dim objFSO,objCreatedFile
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
  Dim sRead, sReadLine, sReadAll, objTextFile
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   MO=Request.QueryString("MO")
   Cartella=Request.QueryString("Cartella")
  ' CodiceAllievo=Request.QueryString("CodiceAllievo")
 
   'StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   'homesito="/anno_2010-2011_ITC/ECDL"    
   Paragrafo=Request.QueryString("Paragrafo")
   NumRec=cint(Request.Form("TxtNUMREC"))
   
  ' response.write(numrec)
  for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
   Domanda = Request.Form("txtDomanda"&k)
   ID=Request.Form("txtCodiceDomanda"&k)
   R11 = Request.Form("txtR1"&k)
   R1=Replace(R11,"'","''")
   R22 = Request.Form("txtR2"&k)
   R2=Replace(R22,"'","''")
   R33 = Request.Form("txtR3"&k)
   R3=Replace(R33,"'","''")
   R44 = Request.Form("txtR4"&k)
   R4 = Replace(R44,"'","''")
  
  
   'Spiegazione=Request.Form("S1")
   'TestoDomandaPlus=Request.Form("TestoDomandaPlus")
     
	 
'	 Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
' 	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logInQuiz.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(Request.Form("txtINQUIZ"&k))
'				objCreatedFile.Close
'	 
'	 
	 
	  
   RE = cint(Request.Form("txtRE"&k))
   VAL=cint(Request.Form("txtVAL"&k))
   INQUIZ=cint(Request.Form("txtINQUIZ"&k))
   DATA=cdate(Request.Form("txtDATA"&k))
    Segnalata=Request.Form("txtSegnalata"&k)
   if Segnalata="" then
     Segnalata=0
   end if
   ' per la spiegazione della domanda 
   ' url=Server.MapPath(homesito)&"/"& Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
   ' url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
	'url3=Replace(url,"\","/")
	'url=url3

  ' per il testo della domanda plus
    ' url4=Server.MapPath(homesito)& "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
    ' url4=Replace(url4,"\","/")
	 
    
      QuerySQL ="UPDATE Domande SET Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE &  "', Voto = '" & VAL & "', In_Quiz = " & INQUIZ &", Data= '" & DATA & "', Segnalata= '" & Segnalata & "' WHERE CodiceDomanda =" &ID&";"
	
	 ConnessioneDB.Execute(QuerySQL)
	' response.write("<br>"&QuerySQL)

 url=Request.Form("url"&k)    
if Cint(Segnalata)=1 then
 Spiegazione=Request.Form("txtSpiegazione"&k)
' Aggiorno la spiegazione
	' se è segnalata aggiorno file di spiegazione
	objFSO.DeleteFile url
	'  response.Write("<br>Cancello : " &url)
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	'' Write a line with a newline character.
	objCreatedFile.WriteLine(Spiegazione)
		'  response.Write("<br>Creo : " &Spiegazione)
	''Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	'objCreatedFile.Close
end if 

'response.write(QuerySQL) %> <br> <%
'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus

'Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
 '	url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logStud.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
'Create the FSO.
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'response.write(Cartella)
'response.write(url)
'objFSO.DeleteFile url
'Set objCreatedFile = objFSO.CreateTextFile(url, True)
' Write a line with a newline character.
'objCreatedFile.WriteLine(Spiegazione)
'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
'objCreatedFile.Close
' per aggiornare la domanda plus
'if Tipodomanda=1 then
'	objFSO.DeleteFile url4
'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
'	' Write a line with a newline character.
'	objCreatedFile.WriteLine(TestoDomandaPlus)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
'	objCreatedFile.Close
'end if 
next 
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Modifica avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If




 




   %>
	</font>   
	 
		
      <h4><a href="../admin/studente_domande.asp?DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>&cla=<%=cla%>&CodiceAllievo=<%=CodiceAllievo%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua 
		a valutare o modificare le domande...</a></h4>
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 
			
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	