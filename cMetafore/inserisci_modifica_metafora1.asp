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
	<div class="contenuti_login" >
	<font color=#FF0000 size="4">

 
   <% 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,voto
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   davalutazione=Request.QueryString("davalutazione")
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
    CodiceMetafora = Request.QueryString("CodiceMetafora")
	 
   MO=Request.QueryString("MO")
   voto=Request.Form("txtVAL")
    DATA = cdate(Request.Form("txtDATA"))
	
	 if strcomp(Request.Form("cb1"),"on")= 0 then ' se è selezionata
			Segnalata=1
			'response.write("Segnalate"&k) %><br><br><%
		    else
			Segnalata=0	
	 end if
		   
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../service/controllo_sessione.asp" -->
<%  
                            'Lettura dei dati memorizzati nei cookie. 
   'CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
   
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   Cartella=Request.QueryString("Cartella")
  
  ' CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  CodiceAllievo = Request.QueryString("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")

Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
ID=Request.QueryString("CodiceNodo")

 
 Select case CodiceTest
   Case Cartella&"_U_2_3" 		
  		   ID=Request.QueryString("CodiceMetafora")
		   Topolino = Request.Form("txtTopolino")
		   
		   Topolino = Replace(Topolino, Chr(34), "'")
		 
		   Topolino=Replace(Topolino,"'","''")
		   Formaggio = Request.Form("txtR1Formaggio")
		   Formaggio = Replace(Formaggio, Chr(34), "'")
		   Formaggio=Replace(Formaggio,"'","''")
		    
		   Fame = Request.Form("txtR2Fame")
		   Fame = Replace(Fame, Chr(34), "'")
		   Fame=Replace(Fame,"'","''")
		   Labirinto = Request.Form("txtR3Labirinto")
		   Labirinto = Replace(Labirinto, Chr(34), "'")
		   Labirinto=Replace(Labirinto,"'","''")
		   Strada = Request.Form("txtR4Strada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request.Form("txtR5Strada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request.Form("txtREStrada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Testata = Request.Form("txtRETestata")
		   Testata = Replace(Testata, Chr(34), "'")
		   Testata=Replace(Testata,"'","''")
		   Distanza=Request.Form("txtREDistanza")
		   Sintesi=Request.Form("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request.Form("S1")
		   
		   DATA=cdate(Request.Form("txtDATA"))
		   if Request.Form("txtVAL")<>"" then
		   VAL=cint(Request.Form("txtVAL"))
		   else
		   VAL=1
		   end if
		    response.write("134")
		   Voto=VAL
		    
	errore=0
	
    if ((len(Topolino)=0) or (len(Formaggio)=0) or (len(Fame)=0) or (len(Labirinto)=0) or (len(Strada)=0) or (len(Strada_OK)=0) or(len(Strada_KO)=0) or(len(Distanza)=0) or(len(Testata)=0)) then
  
   errore=2
  ' response.write(errore&"<br>")
 '  response.write(len(Topolino)&"-" &len(Formaggio)&"-" &len(Fame)&"-" &len(Labirinto)&"-" &len(Strada)&"-" &len(Strada_OK)&"-" &len(Strada_KO)&"-" &len(Distanza)&"-" &len(Testata))
   
  end if 
   
   
	 Case Cartella&"_U_2_5"  
           ID=Request.QueryString("CodiceMetafora")
		   Autista = Request.Form("txtAutista")
		   Autista = Replace(Autista, Chr(34), "'")
		   Autista=Replace(Autista,"'","''")
		   Destinazione = Request.Form("txtR1Destinazione")
		   Destinazione = Replace(Destinazione, Chr(34), "'")
		   Destinazione=Replace(Destinazione,"'","''")
		   Carburante = Request.Form("txtR1Carburante")
		   Carburante = Replace(Carburante, Chr(34), "'")
		   Carburante=Replace(Carburante,"'","''")
		   Luogo = Request.Form("txtR1Luogo")
		   Luogo = Replace(Luogo, Chr(34), "'")
		   Luogo=Replace(Luogo,"'","''")
		   Strada = Request.Form("txtR1Strada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request.Form("txtR1Strada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request.Form("txtR1Strada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Cespugli = Request.Form("txtR1Cespugli")
		   Cespugli = Replace(Cespugli, Chr(34), "'")
		   Cespugli=Replace(Cespugli,"'","''")
		   Lupo = Request.Form("txtR1Lupo")
		   Lupo = Replace(Lupo, Chr(34), "'")
		   Lupo=Replace(Lupo,"'","''")
		   Cestino = Request.Form("txtR1Cestino")
		   Cestino = Replace(Cestino, Chr(34), "'")
		   Cestino=Replace(Cestino,"'","''")
		   Distanza=Request.Form("txtR1Distanza")
		   If Distanza="" then
		   Distanza=Request.Form("txtREDistanza")
		   end if
		   Sintesi=Request.Form("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request.Form("S1")
		   DATA=cdate(Request.Form("txtDATA"))
		   if Request.Form("txtVAL")<>"" then
		   VAL=cint(Request.Form("txtVAL"))
		   else
		   VAL=1
		   end if
		   Voto=VAL
		   errore=0
		   
		   response.write(len(Autista) &" " & len(Destinazione)&" "  & len(Carburante)&" " & len(Luogo)&" " & len(Strada)&" " & len(Strada_OK)&" " & len(Strada_KO)&" " & len(Distanza)&" " & len(Cespugli)&" "& len(Lupo)&" "& len(Luogo)&" "& len(Cestino)&" ")
    if ((len(Autista)=0) or (len(Destinazione)=0) or (len(Carburante)=0) or (len(Luogo)=0) or (len(Strada)=0) or (len(Strada_OK)=0) or(len(Strada_KO)=0) or(len(Distanza)=0) or(len(Cespugli)=0) or(len(Lupo)=0) or(len(Luogo)=0) or(len(Cestino)=0)) then
   errore=2
   end if 
   
   Case Cartella&"_U_2_8"  
  	 ID=Request.QueryString("CodiceMetafora")
       SoggettoC = ucase(Request.Form("txtSoggettoC"))
	   SoggettoC = Replace(SoggettoC, Chr(34), "'")
	   SoggettoC=  Replace(SoggettoC,"'",Chr(96))
  
	   DomandaC = ucase(Request.Form("txtDomandaC"))
	   DomandaC = Replace(DomandaC, Chr(34), "'")
	   DomandaC =  Replace(DomandaC,"'",Chr(96))
	
	
	   MotivazioneC = ucase(Request.Form("txtMotivazioneC"))
	   MotivazioneC = Replace(MotivazioneC, Chr(34), "'")
	   MotivazioneC =  Replace(MotivazioneC,"'",Chr(96))
	
	   DesiderioC = ucase(Request.Form("txtDesiderioC"))
	   DesiderioC = Replace(DesiderioC, Chr(34), "'")
	   DesiderioC=  Replace(DesiderioC,"'",Chr(96))
	   BisognoC = ucase(Request.Form("txtBisognoC"))
	   BisognoC = Replace(BisognoC, Chr(34), "'")
	   BisognoC =  Replace(BisognoC,"'",Chr(96))
	
	   SoggettoS = ucase(Request.Form("txtSoggettoS"))
	   SoggettoS = Replace(SoggettoS, Chr(34), "'")
	   SoggettoS =  Replace(SoggettoS,"'",Chr(96))
	   
	   RispostaS = ucase(Request.Form("txtRispostaS"))
	   RispostaS = Replace(RispostaS, Chr(34), "'")
	   RispostaS=  Replace(RispostaS,"'",Chr(96))
	   
	   MotivazioneS = ucase(Request.Form("txtMotivazioneS"))
	   MotivazioneS = Replace(MotivazioneS, Chr(34), "'")
	   MotivazioneS =  Replace(MotivazioneS,"'",Chr(96))
	   
	   
	   DesiderioS = ucase(Request.Form("txtDesiderioS"))
	   DesiderioS = Replace(DesiderioS, Chr(34), "'")
	   DesiderioS=  Replace(DesiderioS,"'",Chr(96))
	   
		BisognoS = ucase(Request.Form("txtBisognoS"))
	   BisognoS = Replace(BisognoS, Chr(34), "'")
	   BisognoS=  Replace(BisognoS,"'",Chr(96))
		
	   TipoEvento = Request.Form("txtTipoEvento")
	  
	       
		   TolleranzaC=Request.Form("txtTolleranzaC")
	   	   Sintesi=Request.Form("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request.Form("S1")
		   DATA=cdate(Request.Form("txtDATA"))
		   if Request.Form("txtVAL")<>"" then
		   VAL=cint(Request.Form("txtVAL"))
		   else
		   VAL=1
		   end if
		   Voto=VAL
end select

  
  
   
 
if (errore=0) then 
     

	   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"  
	   url=Replace(url,"\","/")
	
	response.write(CodiceTest)
         Select case CodiceTest
		 Case Cartella&"_U_2_3"  
 	 
			 if (session("Admin")=True)  then 
			 response.write("ciao1")	
		  QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza & "',Voto = '" & voto & "', Data='" & DATA & "', Segnalata=" & Segnalata & "  WHERE CodiceMetafora =" &ID&";"
		
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		response.write("ciao2")	
		   QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza & "', Segnalata=" & Segnalata & "   WHERE CodiceMetafora =" &ID&";" 
		   
		   end if 
		   response.write(ucase(session("CodiceAllievo"))&"!<br>?"&ucase(CodiceAllievo))	
		end if
		
		Case Cartella&"_U_2_5"  
		     if (session("Admin")=True)  then 
		  QuerySQL ="UPDATE M_Navigazione SET Autista = '" & Autista & "', Destinazione= '" & Destinazione & "',Carburante= '" & Carburante & "',Luogo= '" & Luogo & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Cespugli = '" & Cespugli & "',Lupo = '" & Lupo & "',Cestino = '" & Cestino & "',Distanza = '" & Distanza & "',Voto = '" & voto &"',Data='" & DATA & "', Segnalata=" & Segnalata & "   WHERE CodiceMetafora =" &ID&";"		
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		   QuerySQL ="UPDATE M_Navigazione SET Autista = '" & Autista & "', Destinazione= '" & Destinazione & "',Carburante= '" & Carburante & "',Luogo= '" & Luogo & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Cespugli = '" & Cespugli & "',Lupo = '" & Lupo & "',Cestino = '" & Cestino & "',Distanza = '" & Distanza & "', Segnalata=" & Segnalata & "   WHERE CodiceMetafora =" &ID&";"
		   end if 
		end if
		
		Case Cartella&"_U_2_8"  
		          if (session("Admin")=True)  then 
		  
		  QuerySQL ="UPDATE M_Desideri SET SoggettoC = '" & SoggettoC & "', DomandaC= '" & DomandaC & "',MotivazioneC= '" & MotivazioneC & "',DesiderioC= '" & DesiderioC & "', BisognoC= '" & BisognoC & "', SoggettoS= '" & SoggettoS &  "', RispostaS = '" & RispostaS & "', MotivazioneS = '" & MotivazioneS &"', DesiderioS= '" & DesiderioS & "', BisognoS= '" & BisognoS &  "', TolleranzaC= " & TolleranzaC &  ", TipoEvento= '" & TipoEvento & "' ,Voto = '" & voto &"',Data='" & DATA & "', Segnalata=" & Segnalata & "    WHERE CodiceMetafora =" &ID&";"
			
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		    QuerySQL ="UPDATE M_Desideri SET SoggettoC = '" & SoggettoC & "', DomandaC= '" & DomandaC & "',MotivazioneC= '" & MotivazioneC & "',DesiderioC= '" & DesiderioC & "', BisognoC= '" & BisognoC & "', SoggettoS= '" & SoggettoS &  "', RispostaS = '" & RispostaS & "', MotivazioneS = '" & MotivazioneS &"', DesiderioS= '" & DesiderioS & "', BisognoS= '" & BisognoS &  "', TolleranzaC= " & TolleranzaC &  ", TipoEvento= '" & TipoEvento & "',Data='" & DATA & "', Segnalata=" & Segnalata & "    WHERE CodiceMetafora =" &ID&";"
			
		   end if 
		end if
		
		
		end select
	
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logModMeta.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	response.write(QuerySQL)
	ConnessioneDB.Execute(QuerySQL)
	 
	
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA METAFORA
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Create the FSO.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'CANCELLA LA VECCHIA VERSIONE DEL FILE11
	response.write("<br>"&url)
	objFSO.DeleteFile url
	
	'On Error Resume Next
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
    objCreatedFile.WriteLine(Spiegazione)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	response.write(Spiegazione)
	objCreatedFile.Close
	response.write(url)
	'On Error Resume Next
	If Err.Number = 0 Then
	
	Response.Write "Modifica avvenuta! "
	Else
	Response.Write Err.Description 
	Err.Number = 0
	End If
	' torna alla pagina chiamante
  ' response.Redirect request.serverVariables("HTTP_REFERER") lo commento e metto l'altro per far vedere le modifiche nel form
   response.Redirect "inserisci_valutazione_metafore.asp?damodifica=1&Cartella="& Session("Cartella")&"&Modulo="&Modulo&"&CodiceTest="&CodiceTest&"&CodiceMetafora="&ID&"&Capitolo="&Capitolo&"&TitoloParagrafo="&Paragrafo


   else
   
    response.write(errore&") Controlla che non ci siano campi lasciati vuoti")%>
 
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%

end if 





   %>
	</font>   
	<% if davalutazione<>"" then ' se sono stata chimata da inserisci_valutazione devo tornare alla pagina studente.asp altrimenti no %> 
	 
        <h3 class="sottotitolo"><a href="../cClasse/studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Torna allo Studente </a></h3> 
	      <% else %>
		
		
	<!--	
      <h4><a href="studente_quiz.asp?Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&testnodo=1">Continua ...</a></h4>
	-->
	<%end if %>
	
	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->


<div class="contenuti_test">

<!-- <h3 class="sottotitolo"><a href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>&cod=<%=CodiceAllievo%>"> Torna al Quaderno dello Studente </a></h3> -->

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
	