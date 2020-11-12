<!-- calcola_risultato_MODBC3.asp -->
<%@ Language=VBScript %>
 
<%
  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    
 
   <% 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Spiegazione
   Dim RispostaData,RispostaEsatta,RisposteOK, RisposteKO, RecordModificati,inbianco,voto
   Dim RispDate(),RispEsatte(),Errori(),NumDom 
   Dim RispDate1(),RispEsatte1()
   Dim Risultato_relativo 
   
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
		   Topolino = Request("txtTopolino")
		   
		   Topolino = Replace(Topolino, Chr(34), "'")
		 
		   Topolino=Replace(Topolino,"'","''")
		   Formaggio = Request("txtFormaggio")
		   Formaggio = Replace(Formaggio, Chr(34), "'")
		   Formaggio=Replace(Formaggio,"'","''")
		    
		   Fame = Request("txtR2Fame")
		   Fame = Replace(Fame, Chr(34), "'")
		   Fame=Replace(Fame,"'","''")
		   Labirinto = Request("txtR3Labirinto")
		   Labirinto = Replace(Labirinto, Chr(34), "'")
		   Labirinto=Replace(Labirinto,"'","''")
		   Strada = Request("txtR4Strada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request("txtStrada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request("txtStrada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Testata = Request("txtTestata")
		   Testata = Replace(Testata, Chr(34), "'")
		   Testata=Replace(Testata,"'","''")
		   Distanza=Request("txtDistanza")
		   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		   
		   DATA=cdate(Request("txtDATA"))
		   if Request("txtVAL")<>"" then
		   VAL=cint(Request("txtVAL"))
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
		   Autista = Request("txtAutista")
		   Autista = Replace(Autista, Chr(34), "'")
		   Autista=Replace(Autista,"'","''")
		   Destinazione = Request("txtDestinazione")
		   Destinazione = Replace(Destinazione, Chr(34), "'")
		   Destinazione=Replace(Destinazione,"'","''")
		   Carburante = Request("txtCarburante")
		   Carburante = Replace(Carburante, Chr(34), "'")
		   Carburante=Replace(Carburante,"'","''")
		   Luogo = Request("txtLuogo")
		   Luogo = Replace(Luogo, Chr(34), "'")
		   Luogo=Replace(Luogo,"'","''")
		   Strada = Request("txtStrada")
		   Strada = Replace(Strada, Chr(34), "'")
		   Strada=Replace(Strada,"'","''")
		   Strada_OK = Request("txtStrada_OK")
		   Strada_OK = Replace(Strada_OK, Chr(34), "'")
		   Strada_OK=Replace(Strada_OK,"'","''")
		   Strada_KO = Request("txtStrada_KO")
		   Strada_KO = Replace(Strada_KO, Chr(34), "'")
		   Strada_KO=Replace(Strada_KO,"'","''")
		   Cespugli = Request("txtCespugli")
		   Cespugli = Replace(Cespugli, Chr(34), "'")
		   Cespugli=Replace(Cespugli,"'","''")
		   Lupo = Request("txtLupo")
		   Lupo = Replace(Lupo, Chr(34), "'")
		   Lupo=Replace(Lupo,"'","''")
		   Cestino = Request("txtCestino")
		   Cestino = Replace(Cestino, Chr(34), "'")
		   Cestino=Replace(Cestino,"'","''")
		   Distanza=Request("txtDistanza")
		   If Distanza="" then
		   Distanza=Request("txtDistanza")
		   end if
		   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		   DATA=cdate(Request("txtDATA"))
		   if Request("txtVAL")<>"" then
		   VAL=cint(Request("txtVAL"))
		   else
		   VAL=1
		   end if
		   Voto=VAL
		   errore=0
		   
		  ' response.write(len(Autista) &" " & len(Destinazione)&" "  & len(Carburante)&" " & len(Luogo)&" " & len(Strada)&" " & len(Strada_OK)&" " & len(Strada_KO)&" " & len(Distanza)&" " & len(Cespugli)&" "& len(Lupo)&" "& len(Luogo)&" "& len(Cestino)&" ")
    if ((len(Autista)=0) or (len(Destinazione)=0) or (len(Carburante)=0) or (len(Luogo)=0) or (len(Strada)=0) or (len(Strada_OK)=0) or(len(Strada_KO)=0) or(len(Distanza)=0) or(len(Cespugli)=0) or(len(Lupo)=0) or(len(Luogo)=0) or(len(Cestino)=0)) then
   errore=2
   end if 
   
   Case Cartella&"_U_2_8"  
  	 ID=Request.QueryString("CodiceMetafora")
       SoggettoC = ucase(Request("txtSoggettoC"))
	   SoggettoC = Replace(SoggettoC, Chr(34), "'")
	   SoggettoC=  Replace(SoggettoC,"'",Chr(96))
  
	   DomandaC = ucase(Request("txtDomandaC"))
	   DomandaC = Replace(DomandaC, Chr(34), "'")
	   DomandaC =  Replace(DomandaC,"'",Chr(96))
	
	
	   MotivazioneC = ucase(Request("txtMotivazioneC"))
	   MotivazioneC = Replace(MotivazioneC, Chr(34), "'")
	   MotivazioneC =  Replace(MotivazioneC,"'",Chr(96))
	
	   DesiderioC = ucase(Request("txtDesiderioC"))
	   DesiderioC = Replace(DesiderioC, Chr(34), "'")
	   DesiderioC=  Replace(DesiderioC,"'",Chr(96))
	   BisognoC = ucase(Request("txtBisognoC"))
	   BisognoC = Replace(BisognoC, Chr(34), "'")
	   BisognoC =  Replace(BisognoC,"'",Chr(96))
	
	   SoggettoS = ucase(Request("txtSoggettoS"))
	   SoggettoS = Replace(SoggettoS, Chr(34), "'")
	   SoggettoS =  Replace(SoggettoS,"'",Chr(96))
	   
	   RispostaS = ucase(Request("txtRispostaS"))
	   RispostaS = Replace(RispostaS, Chr(34), "'")
	   RispostaS=  Replace(RispostaS,"'",Chr(96))
	   
	   MotivazioneS = ucase(Request("txtMotivazioneS"))
	   MotivazioneS = Replace(MotivazioneS, Chr(34), "'")
	   MotivazioneS =  Replace(MotivazioneS,"'",Chr(96))
	   
	   
	   DesiderioS = ucase(Request("txtDesiderioS"))
	   DesiderioS = Replace(DesiderioS, Chr(34), "'")
	   DesiderioS=  Replace(DesiderioS,"'",Chr(96))
	   
		BisognoS = ucase(Request("txtBisognoS"))
	   BisognoS = Replace(BisognoS, Chr(34), "'")
	   BisognoS=  Replace(BisognoS,"'",Chr(96))
		
	   TipoEvento = Request("txtTipoEvento")
	  
	       
		   TolleranzaC=Request("txtTolleranzaC")
	   	   Sintesi=Request("S1")
		   Sintesi= Replace(Sintesi, Chr(34), chr(96))
		   Sintesi=Replace(Sintesi,Chr(39),chr(96))
		   Spiegazione=Request("S1")
		   DATA=cdate(Request("txtDATA"))
		   if Request("txtVAL")<>"" then
		   VAL=cint(Request("txtVAL"))
		   else
		   VAL=1
		   end if
		   Voto=VAL
end select

  
  
   
 
if (errore=0) then 
     

	   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"  
	   url=Replace(url,"\","/")
	
	'response.write(CodiceTest)
         Select case CodiceTest
		 Case Cartella&"_U_2_3"  
 	 
			 if (session("Admin")=True)  then 
			' response.write("ciao1")	
		  QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza & "',Voto = '" & voto & "', Data='" & DATA & "', Segnalata=" & Segnalata & "  WHERE CodiceMetafora =" &ID&";"
		
		else if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then 
		'response.write("ciao2")	
		   QuerySQL ="UPDATE M_Topolino SET Topolino = '" & Topolino & "', Formaggio= '" & Formaggio & "',Fame= '" & Fame & "',Labirinto= '" & Labirinto & "', Strada= '" & Strada & "', Strada_OK= '" & Strada_OK & "', Strada_KO = '" & Strada_KO & "', Testata = '" & Testata & "',Distanza = '" & Distanza & "', Segnalata=" & Segnalata & "   WHERE CodiceMetafora =" &ID&";" 
		   
		   end if 
		 '  response.write(ucase(session("CodiceAllievo"))&"!<br>?"&ucase(CodiceAllievo))	
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
'	response.write(QuerySQL)
	ConnessioneDB.Execute(QuerySQL)
	 
	
	
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SPIEGAZIONE DELLA METAFORA
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Create the FSO.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'CANCELLA LA VECCHIA VERSIONE DEL FILE11
'	response.write("<br>"&url)
	objFSO.DeleteFile url
	
	'On Error Resume Next
	Set objCreatedFile = objFSO.CreateTextFile(url, True)
	' Write a line with a newline character.
    objCreatedFile.WriteLine(Spiegazione)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	response.write(Spiegazione)
	objCreatedFile.Close
	'response.write(url)
	'On Error Resume Next
	If Err.Number = 0 Then
	'RESPONSE.WRITE(querysql)
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
	