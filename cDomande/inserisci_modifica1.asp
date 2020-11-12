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

   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
	<!-- #include file = "tabella_corrispondenze.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    <!--#include file="../service/gestione_errori.asp" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")
TestoDomandaPlus=Request.Form("TestoDomandaPlus")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
Multiple=Request.QueryString("Multiple")  
  VF=Request.QueryString("VF") 
   
   Domanda1 = Request.Form("txtDomanda")
   Domanda=Replace(Domanda1,"'","''")
   Tipodomanda=Request.QueryString("Tipodomanda")
   ID=Request.QueryString("CodiceDomanda")

   R11 = Request.Form("txtR1")
   R1=Replace(R11,"'","''")


   R22 = Request.Form("txtR2")
   R2=Replace(R22,"'","''")

   R33 = Request.Form("txtR3")
   R3=Replace(R33,"'","''")

   R44 = Request.Form("txtR4")
   R4 = Replace(R44,"'","''")
 
   Spiegazione=Request.Form("S1")
 
  function controlla(RisposteEsatte)
	 controlla=0
	 i=0
	 while (i<=16) and not(esiste)
		if v2(i)= RisposteEsatte then 
		   controlla=1
		end if
		i=i+1
	 wend
 end function
 
 
     if (len(Request.Form("txtRE"))=0)  or (len(Spiegazione)=0)  then 
			   errore=2
	 else
	       RE = clng(Request.Form("txtRE"))
		   errore=0
	 end if 
 '  if (len(Request.Form("txtRE"))=0) then 
'      Response.Redirect("inserisci_modifica1.asp")
'   end if 
'   
  
  
   
if VF<>"" then ' se ì vero o falso devo fare meno controlli
  if  (len(Domanda)=0)  then 
	  errore=2
  
   elseif (RE<>0) and (RE<>1) then ' risposta vero falso 0 o 1
     errore=4
   elseif  (IsNumeric(RE)=0) then
      errore=5
  end if

else  
	  
		 function controlla(RisposteEsatte)
			 controlla=0
			 i=0
			 while (i<=16) and not(esiste)
				if v2(i)= RisposteEsatte then 
				   controlla=1
				end if
				i=i+1
			 wend
		 end function
		 
		   if Multiple<>"0" then
			   ' controllo validità numero che indica la risposta esatta deve appartenere alla tabella di corrispondenza
			   esiste=controlla(RE)
			   if esiste = 0 then
				  errore = 3
			   end if 	   
			   
		   else
			   if ((RE<1) or (RE>4)) then 
				  errore=1
			   end if
		   end if 
		   
		   if ( (len(Domanda)=0) or (len(R1)=0) or (len(R2)=0) or (len(R3)=0) or (len(R4)=0) or (len(Spiegazione)=0) ) then 
			   errore=2
		   end if
		 
		   'Domanda1=Domanda
		   'response.write("Domanda="&Domanda1)
		 if Multiple<>"" then
			' se non devo inserire domanda multipla pongo a 0 il campo 
			Multiple=1
		 else
			Multiple=0
		 end if 
 end if
 
  if (errore<>0) then
	  if (errore=1) then
		 response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4, RE="&RE)
	  end if 
	  if (errore=2) then
		response.write("Controlla che non ci siano campi lasciati vuoti")
	  end if 
	  if (errore=3) then
		response.write("Controlla le risposte esatte (max 3 vere)")
	  end if 
	   if (errore=4) then
		response.write("Controlla le risposte esatte, valori ammessi 0 per (Falso) o 1 per (Vero)")
	  end if 
	   if (errore=5) then
		response.write("Controlla le risposte esatte, valori ammessi 0 per (Falso) o 1 per (Vero)")
	  end if 
 
   %>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
   else
   
   
    
    

   url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
    url1= "../" & Cartella & "/" &Modulo&"_Spiegazioni/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"

url3=Replace(url,"\","/")
url=url3
' per il testo della domanda plus
     url4=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&  "/" & Cartella &"/" &Modulo&"_Domande/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
 
	' url_file=Server.MapPath("/ECDL/")& "/"& url ' per localhost
     url4=Replace(url4,"\","/")
	 



     QuerySQL ="UPDATE Domande SET Quesito = '" & Domanda & "', Risposta1= '" & R1 & "',Risposta2= '" & R2 & "',Risposta3= '" & R3 & "', Risposta4= '" & R4 & "', RispostaEsatta= '" & RE & "', Data = '" & DataTest & "'  WHERE CodiceDomanda =" &ID&";"
	 ConnessioneDB.Execute(QuerySQL)



'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA DOMANDA

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
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

'GESTIONE RRORE
If Err.Number <> 0 then
'  NumeroErrore = Err.Number
  DescrizioneErrore = Err.Description
  Pagina = Request.ServerVariables("url")
  Spiegazione1="Impossibile modificare il file contenente la spiegazione"
  Riga=120
'  Source=Err.Source
  Call GestisciErrore(DescrizioneErrore,Spiegazione1,Pagina,Riga)
  Err.Number=0
End If

'response.write(url)
' per aggiornare la domanda plus
if Tipodomanda=1 then

	objFSO.DeleteFile url4
	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
	' Write a line with a newline character.
	objCreatedFile.WriteLine(TestoDomandaPlus)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
end if 
On Error Resume Next
If Err.Number = 0 Then

Response.Write "Modifica avvenuta! "
Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>
	</font>   
	 
		
      <h4><a href="../cClasse/studente_quiz.asp?Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Continua a modificare...</a></h4>
	<p>&nbsp;</p>
<% end if%>
	<div id=piede_pagina>
			<p><p>
			
	<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna alla pagina Verifica... </a></h3> 					</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	