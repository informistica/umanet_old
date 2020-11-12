<!-- calcola_risultato_MODBC3.asp -->
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
    <body bgcolor="#FFFFFF">
  <% end if %>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">


  
   <% 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,url4,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   Codice_Test = Request.QueryString("CodiceTest")
   MO=Request.QueryString("ID_MOD")
   Cartella=Request.QueryString("Cartella")
   Paragrafo=Request.QueryString("Paragrafo")
   TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  'response.write("TP:"&TitoloParagrafo)
   Modulo=Request.QueryString("Modulo")
   CodiceAllievo=Request.QueryString("CodiceAllievo")
      INQUIZ=0 ' non gestisco INQUIZ per le metafore
   
   'StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../var_globali.inc" -->
  
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   ' ho messo include  ma non funziona ??? lo devo mettere qua altrimenti mi da errore boh ?
      
   Paragrafo=Request.QueryString("Paragrafo")
   NumRec=cint(Request.Form("TxtNUMREC"))
   Titolo=Request.Form("TitoloParagrafo")
  'response.write("NUMREC=" & NumRec)
'response.write("Codice_Test=" & Codice_Test)
    


 Select Case Codice_Test
	Case Cartella&"_U_3_3" ' metafora topolino			
'if Codice_Test= "U_3_3" then  ' METAFORA TOPOLINO

		for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
		   ID=Request.Form("txtCodiceMetafora"&k)
		   Topolino = Request.Form("txtTopolino"&k)
		   Topolino = Replace(Topolino, Chr(34), "'")
		   Topolino=Replace(Topolino,"'","''")
		   Formaggio = Request.Form("txtR1Formaggio"&k)
		   Formaggio = Replace(Formaggio, Chr(34), "'")
		   Formaggio=Replace(Formaggio,"'","''")
		   Fame = Request.Form("txtR1Fame"&k)
		   Fame = Replace(Fame, Chr(34), "'")
		   Fame=Replace(Fame,"'","''")
		   Labirinto = Request.Form("txtR1Labirinto"&k)
		   Labirinto = Replace(Labirinto, Chr(34), "'")
		   Labirinto=Replace(Labirinto,"'","''")
		   Strada = Request.Form("txtR1Strada"&k)
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request.Form("txtR1Strada_OK"&k)
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request.Form("txtR1Strada_KO"&k)
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Testata = Request.Form("txtR1Testata"&k)
		   Testata = Replace(Testata, Chr(34), "'")
		   Testata=Replace(Testata,"'","''")
		   Distanza=Request.Form("txtR1Distanza"&k)
		   Sintesi=Request.Form("S1"&k)
		   Sintesi= Replace(Sintesi, Chr(34), "'")
		   Sintesi=Replace(Sintesi,"'","''")
		   Spiegazione=Request.Form("S1"&k)
		   DATA=cdate(Request.Form("txtDATA"&k))
		   VAL=cint(Request.Form("txtVAL"&k))
		   Voto=VAL
		   if (Request.Form("txtINQUIZ"&k)<>"") then
			  INQUIZ=cint(Request.Form("txtINQUIZ"&k))
		   end if 
		   if strcomp(Request.Form("cb"&k),"on")= 0 then ' se è selezionata
			Segnalata=1
			'response.write("Segnalate"&k) %><br><br><%
		    else
			Segnalata=0	
     end if
		
      '  response.write("<br>non Stampo="&i)
		 
		 
			 if (session("Admin")=True)  then 
		  QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza & "',Voto = '" & voto & "', In_Quiz = '" & INQUIZ &"',Data='" & DATA &"',Segnalata=" & Segnalata & "  WHERE CodiceMetafora =" &ID&";"
		 
			
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		   QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza &"',Segnalata=" & Segnalata &  "'  WHERE CodiceMetafora =" &ID&";"
		  
		   end if 
		   
		end if
			 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &MO&"_Metafore/"&MO&"_"&TitoloParagrafo&"_"&ID&".txt"   
			 url=Replace(url,"\","/")
		
		
		      
		
				
		 ConnessioneDB.Execute(QuerySQL)
		'response.write(QuerySQL) %> <br> <%
		'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus
		
		
			' if strcomp(Request.Form("cb"&k),"on")= 0 then ' se è selezionata aggiorno il file
				
				Dim objFSO,objCreatedFile
				Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				 
				'Create the FSO.
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'CANCELLA LA VECCHIA VERSIONE DEL FILE11
				'response.write(Cartella)
				'response.write(url)
				objFSO.DeleteFile url
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				' Write a line with a newline character.
				objCreatedFile.WriteLine(Spiegazione)
				'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
				objCreatedFile.Close
				
			
			'end if
		next 
  
  
 
 Case Cartella&"_U_3_5" ' metafora navigazione			
 
		for k=0 to NumRec-1 ' per scorrere tutto il form e fare un update ad ogni ciclo
		   ID=Request.Form("txtCodiceMetafora"&k)
		   Autista = Request.Form("txtAutista"&k)
		   Autista = Replace(Autista, Chr(34), "'")
		   Autista=Replace(Autista,"'","''")
		   Destinazione = Request.Form("txtR1Destinazione"&k)
		   Destinazione = Replace(Destinazione, Chr(34), "'")
		   Destinazione=Replace(Destinazione,"'","''")
		   Carburante = Request.Form("txtR1Carburante"&k)
		   Carburante = Replace(Carburante, Chr(34), "'")
		   Carburante=Replace(Carburante,"'","''")
		   Luogo = Request.Form("txtR1Luogo"&k)
		   Luogo = Replace(Luogo, Chr(34), "'")
		   Luogo=Replace(Luogo,"'","''")
		   Strada = Request.Form("txtR1Strada"&k)
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request.Form("txtR1Strada_OK"&k)
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request.Form("txtR1Strada_KO"&k)
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Cespugli = Request.Form("txtR1Cespugli"&k)
		   Cespugli = Replace(Cespugli, Chr(34), "'")
		   Cespugli=Replace(Cespugli,"'","''")
		   Lupo = Request.Form("txtR1Lupo"&k)
		   Lupo = Replace(Lupo, Chr(34), "'")
		   Lupo=Replace(Lupo,"'","''")
		   Cestino = Request.Form("txtR1Cestino"&k)
		   Cestino = Replace(Cestino, Chr(34), "'")
		   Cestino=Replace(Cestino,"'","''")
		   Distanza=Request.Form("txtR1Distanza"&k)
		   Sintesi=Request.Form("S1"&k)
		   Sintesi= Replace(Sintesi, Chr(34), "'")
		   Sintesi=Replace(Sintesi,"'","''")
		   Spiegazione=Request.Form("S1"&k)
		   DATA=cdate(Request.Form("txtDATA"&k))
		   VAL=cint(Request.Form("txtVAL"&k))
		   Voto=VAL
		   if (Request.Form("txtINQUIZ"&k)<>"") then
			  INQUIZ=cint(Request.Form("txtINQUIZ"&k))
		   end if 
		    if strcomp(Request.Form("cb"&k),"on")= 0 then ' se è selezionata
			Segnalata=1
			'response.write("Segnalate"&k) %><br><br><%
		    else
			Segnalata=0	
     end if
		' response.Write("CESPUGLI"&Cespugli)
		 
			 if (session("Admin")=True)  then 
		  QuerySQL ="UPDATE M_Navigazione SET Autista = '" & Autista & "', Destinazione= '" & Destinazione & "',Carburante= '" & Carburante & "',Luogo= '" & Luogo & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Cespugli = '" & Cespugli & "',Lupo = '" & Lupo & "',Cestino = '" & Cestino & "',Distanza = '" & Distanza & "',Voto = '" & voto & "', In_Quiz = '" & INQUIZ &"',Data='" & DATA &"',Segnalata=" & Segnalata &  "   WHERE CodiceMetafora =" &ID&";"
		 
			
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		   QuerySQL ="UPDATE M_Navigazione SET Autista = '" & Autista & "', Destinazione= '" & Destinazione & "',Carburante= '" & Carburante & "',Luogo= '" & Luogo & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Cespugli = '" & Cespugli & "',Lupo = '" & Lupo & "',Cestino = '" & Cestino & "',Distanza = '" & Distanza &"',Segnalata=" & Segnalata & "    WHERE CodiceMetafora =" &ID&";"
		  
		   end if 
		   
		end if
			 url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &MO&"_Metafore/"&MO&"_"&TitoloParagrafo&"_"&ID&".txt"   
			 url=Replace(url,"\","/")
			' url=url3
		 
		 ConnessioneDB.Execute(QuerySQL)
		'response.write(QuerySQL) %> <br> <%
		'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus
		
		 
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			 
			'Create the FSO.
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			'CANCELLA LA VECCHIA VERSIONE DEL FILE11
			objFSO.DeleteFile url
			Set objCreatedFile = objFSO.CreateTextFile(url, True)
			' Write a line with a newline character.
			objCreatedFile.WriteLine(Spiegazione)
			'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
			objCreatedFile.Close
		next 
   
  
  
  
  
  
  
  Case Else
				   ' Istruzioni di default
  End Select
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Modifica avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If




 




   %>
	</font>   
	 
		
      <h3 class="sottotitolo"><a href="../cClasse/studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Torna allo Studente </a></h3> 
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<div class="contenuti_test">
<!--
<h3 class="sottotitolo"><a href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>"> Torna al Quaderno dello Studente </a></h3> 
-->
<!--#include file="../include/tornaquaderno.html" --> 

 <h3 class="sottotitolo"><a href="../cClasse/studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Torna alla Classifica </a></h3> 

<!--			     
<h3 class="sottotitolo"><a href="../U-ECDL/home_uecdl_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina U-ECDL Verifica... </a></h3> 
<h3 class="sottotitolo"><a href="../U-ECDL/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina U-ECDL Apprendimento... </a></h3> 
<h3 class="sottotitolo"><a href="../home_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla Home Page Apprendimento... </a></h3> 
 
-->
</div>
			
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	