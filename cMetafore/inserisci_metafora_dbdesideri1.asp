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
function showText2() {window.alert("La sessione � scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>

</head>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validit� della sessione, se � scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
    <body bgcolor="#FFFFFF">
  <% end if %>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,QuerySQLD,QuerySQL1,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   Dim i,Domanda,Domanda1,R1,R11,R2,R22,R3,R33,R4,R44,RE,CodCap,CodiceCap,Sintesi
   Dim SoggettoC,DomandaC,MotivazioneC,DesiderioC,BisognoC,RispostaS,SoggettoS,MotivazioneS,DesiderioS,BisognoS
   
   
   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   daSimulazione = Request.QueryString("daSimulazione")
   daDesideri = Request.QueryString("daDesideri") ' settato se sono chiamato per lo sviluppo della metafora
   CodiceTest = Request.QueryString("CodiceTest")
   Li = cint(Request.QueryString("Li"))
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   prenodo=Request.QueryString("prenodo") ' serve per capire il chiamante e quindi sapere se alla fine devo redirectare ad home_ver o home_app
   
   
    
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	' Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
       %>   

   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../service/controllo_sessione.asp" -->
  
   
<%  
                            'Lettura dei dati memorizzati nei cookie. 
  ' CodiceTest = Request.Cookies("Dati")("CodiceTest")
   
  ' homesito="/anno_2010-2011_ITC/ECDL"
   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
   DataTest = Request.Cookies("Dati")("DataTest")
   
   CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")
	CodiceMetafora=Request.QueryString("CodiceMetafora")
	Paragrafo=Request.QueryString("Paragrafo")
	Modulo=Request.QueryString("Modulo")
	DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
 
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
    
   TipoEvento = ucase(Request.Form("txtTipoEvento"))
   TipoEvento = Replace(TipoEvento, Chr(34), "'")
   TipoEvento=  Replace(TipoEvento,"'",Chr(96))	  
   TolleranzaC = cint(Request.Form("txtTolleranzaC"))
   Sintesi=ucase(Request.Form("S1"))
   Sintesi= Replace(Sintesi, Chr(34), "'")
 '  Sintesi=  Replace(Sintesi,"'",Chr(96))

   
   
  
    
    
   
   
   
   
'response.write("<br>"&SoggettoC&"<br>"& DomandaC &"<br>"&MotivazioneC&"<br>"&DesiderioC &"<br>"&BisognoC&"<br>"&SoggettoS&"<br>"&RispostaS&"<br>"& MotivazioneS&"<br>"& DesiderioS &"<br>"&BisognoS) 
   
   
   
   if ( (len(SoggettoC)=0) or (len(DomandaC)=0) or (len(MotivazioneC)=0) or (len(DesiderioC)=0) or (len(BisognoC)=0) or (len(SoggettoS)=0) or(len(RispostaS)=0) or(len(MotivazioneS)=0) or(len(DesiderioS)=0) or (len(BisognoS)=0)   ) then


'  Response.Redirect("inserisci_test.asp?Cartella=Cartella&Num=0&Cognome=Cognome&Nome=Nome&CodiceTest=CodiceTest&Capitolo=Capitolo&Paragrafo=Paragrafo&Modulo=Modulo") 
   ' Response.Redirect("inserisci_test.asp") 
   errore=2
  
   end if
   
 if (errore=0) then
   
   ' devo vedere se il setting � tale da richiedere voto=1 come default oppure no  
    QuerySQL1="Select * from Setting"
	Set rsTabella = ConnessioneDB.Execute(QuerySQL1) 
	Valutato=rsTabella.fields("Valutato") 
	rsTabella.close
	if Valutato=1 then
        Voto=1 ' valore di default 
	else
	    Voto=0
	end if
	
	if daSimulazione=1 then ' aggiorno 
 
 QuerySQL1 ="UPDATE M_Desideri SET SoggettoC = '" & SoggettoC & "', DomandaC= '" & DomandaC & "',MotivazioneC= '" & MotivazioneC & "',DesiderioC= '" & DesiderioC & "', BisognoC= '" & BisognoC & "', SoggettoS= '" & SoggettoS &  "', RispostaS = '" & RispostaS & "', MotivazioneS = '" & MotivazioneS &"', DesiderioS= '" & DesiderioS & "', BisognoS= '" & BisognoS &  "', TipoEvento= '" & TipoEvento &  "', TolleranzaC= '" & TolleranzaC & "' WHERE CodiceMetafora =" &CodiceMetafora&";"
	 ConnessioneDB.Execute QuerySQL1 
			
 else	
  QuerySQL1="INSERT INTO M_Desideri (SoggettoC, DomandaC, MotivazioneC,DesiderioC,BisognoC,SoggettoS,RispostaS,MotivazioneS,DesiderioS,BisognoS,TipoEvento,TolleranzaC,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora,Pi) SELECT '" & SoggettoC & "','" & DomandaC & "', '" & MotivazioneC & "','" & DesiderioC & "','" & BisognoC & "','" & SoggettoS & "','" & RispostaS & "','" & MotivazioneS & "','" & DesiderioS & "','" & BisognoS & "','" & TipoEvento & "'," & TolleranzaC & ",'" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) & "'," & Li &";" 
 '  end if 
 
 response.write(QuerySQL1)
 ConnessioneDB.Execute QuerySQL1 
 ' collego la metafora padre alla figlio"
   QuerySQL="select max(CodiceMetafora) from M_Desideri"
   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   MaxID=rsTabella(0)
   Session("CodiceMetafora")=MaxID
    ' per ritornare il valore al chiamante per visualizzare alert verde con link diretto a metafora 
	 Session("Capitolo")=Capitolo
	 Session("Paragrafo")=Paragrafo
	 Session("CodiceTest")=CodiceTest
   QuerySQL ="UPDATE M_Desideri SET Pf = " & MaxID & " WHERE CodiceMetafora =" &Li&";"
   ConnessioneDB.Execute QuerySQL 
  
 end if 
   
  QuerySQLD = "SELECT CodiceMetafora,Cartella FROM M_Desideri WHERE CodiceMetafora=(Select Max(CodiceMetafora) FROM M_Desideri);" 
    Set rsTabella = ConnessioneDB.Execute(QuerySQLD)
    ID=rsTabella(0)
 if daSimulazione<>1 then
  
  ' ' provo a metterle fuori perch� va in errore update m_desideri con ID nullo 
  ' QuerySQL = "SELECT CodiceMetafora,Cartella FROM M_Desideri WHERE CodiceMetafora=(Select Max(CodiceMetafora) FROM M_Desideri);" 
'    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'    ID=rsTabella(0)
    CARTA=rsTabella(1)
	url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & CARTA &"/" &Modulo&"_Metafore/"&Modulo&"_"&Paragrafo&"_"&ID&".txt" 'per il server on line
   
	'CREAZIONE FILE DI TESTO PER INSERIRE LA SINTESI DELLA METAFORA
	
	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	 
	'Create the FSO.
	 
	url=Replace(url,"\","/")
	  
		
					'url2="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logTopo.txt"
	'				Set objCreatedFile = objFSO.CreateTextFile(url2, True)
	'				objCreatedFile.WriteLine(url)
	'				objCreatedFile.Close
	
	'response.write(url)
	Set objCreatedFile = objFSO.CreateTextFile(url, True)

  if instr(Sintesi,"<script>")<>0 then
	   Sintesi=Replace(Sintesi,"<script>","")
	   Sintesi=Replace(Sintesi,"</script>","")
	end if
	' Write a line with a newline character.
	objCreatedFile.WriteLine(Sintesi)
	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	objCreatedFile.Close
	'response.write(url)
	
	'if Tipo="1" then 'CREAZIONE FILE DI TESTO PER INSERIRE LA DOMANDA
	'
	'	url4=Replace(url4,"\","/")
	'	 
	'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
	'	' Write a line with a newline character.
	'	objCreatedFile.WriteLine(Domanda)
	'	'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
	'	objCreatedFile.Close
	'end if 
	'response.write("<br>" & url)
	
	'On Error Resume Next
	If Err.Number = 0 Then
	session("inserita")=true
	Response.Write "Inserimento avvenuto! "
	
	Else
	Response.Write Err.Description 
	Err.Number = 0
	End If
	




   %>
	</font>   
	 
		
      <h4><a href="inserisci_metafora_dbdesideri.asp?Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>">Continua ...</a></h4>

<%end if ' chiudo if daSimulazione<>1 then%>

	<p>&nbsp;</p>
	<div id=piede_pagina>
			<p><p>
			<div class="contenuti_test">
			 <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
                <h3 class="sottotitolo"><a href="../cClasse/quaderno_metafore.asp?id_classe=<%=Session("Id_Classe")%>&cod=<%=CodiceAllievo%>&DataClaq=<%=Session("DataClaq")%>&DataClaq2=<%=Session("DataClaq2")%>"> Vai al Quaderno </a></h3> 
 

			<a href="#" onClick="history.go(-1);return false;" class="sottotitolo">Indietro</a>
             <%' response.Redirect request.serverVariables("HTTP_REFERER") ' torno indietro direttamente senza chiedere%>
<%else
  if (errore=1) then
     'response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
  end if 
  if (errore=2) then
    response.write("Controlla che non ci siano campi lasciati vuoti")
  end if %>
	<a href="#" onClick="history.go(-1);return false;">Indietro</a>
  <%
  
end if 			


if daDesideri<>"" then
QuerySQL ="UPDATE M_Desideri SET Pf="&ID&" WHERE CodiceMetafora =" &CodiceMetafora&";"
ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL& "<br>")
QuerySQL ="UPDATE M_Desideri SET Pi="&CodiceMetafora&" WHERE CodiceMetafora =" &ID&";"
ConnessioneDB.Execute QuerySQL


' metto anche i link nella tabella link forse serve solo uno dei due modi
LP=Li ' Livello Partenza Strada_OK nella metafora prima, messo IN BASE AI RADIO BOX
LD=2 ' Livello Arrivo Destinazione nella metafora 
T2=""
QuerySQL="INSERT INTO LinkDesideri (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(CodiceMetafora) & "','" &LP & "', '" & CInt(ID) & "','" & LD & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL&"<br>")
' non lo metto perch� non prevedo di tornare indietro ma solo in avanti, altrimenti mi mette blu anche il livello di arrivo del link
'QuerySQL="INSERT INTO LinkNavigazione (Id_n1, L1, Id_n2,L2,Id_Stud,Testo2) SELECT '" & CInt(ID) & "','" &LD & "', '" & CInt(CodiceMetafora) & "','" & LP & "','" & Session("CodiceAllievo")& "','" &T2 & "';"
'ConnessioneDB.Execute QuerySQL 
'response.write(QuerySQL)

end if
session("inserita")=true
response.Redirect request.serverVariables("HTTP_REFERER") 
'Response.Redirect "quaderno_metafore.asp?id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&Session("CodiceAllievo")&"&DataClaq="&Session("DataClaq")&"&DataClaq2="&Session("DataClaq")&"&damenu=1"
	  



%>
            </div>
			</div>
 <!-- se il login � corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	