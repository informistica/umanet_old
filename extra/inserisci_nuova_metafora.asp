<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Documento senza titolo</title>
</head>
<%
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
Dim QuerySQL
%>
<!-- #include file = "stringa_connessione.inc" -->
<body>
<%

   				   
				    CodiceTest="6Classe"
					Modulo="6C" 
					DataTest="12/12/2112"  
				    Voto=0
				   SoggettoC =  "?"
				   DomandaC =  "?"
				   MotivazioneC = "?" 
				   DesiderioC = "?"
				   BisognoC="?"
				   SoggettoS =  "?"
				   RispostaS = "?"    
				   MotivazioneS = "?"    
				   DesiderioS =  "?" 
				   BisognoS="?"
				   TipoEvento = 1 
				   TolleranzaC = 3 
				   URL_teoria="?"
				   Cartella="?"
	
	QuerySQL="Select Classe from Classi;"
	set rsTabella0=ConnessioneDB.Execute(QuerySQL)
	
	do while not rsTabella0.eof 
	
		QuerySQL="Select CodiceAllievo from ALLIEVI_CLASSE where Classe='"&rsTabella0(0) &"';"
	    set rsTabella=ConnessioneDB.Execute(QuerySQL)
		response.write("<br>"&QuerySQL)
	
	'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logfrasi.txt"
	'Set objCreatedFile = objFSO.CreateTextFile(url, True)
	 i=0
	 do while not rsTabella.eof 
	 CodiceAllievo=rsTabella("CodiceAllievo")

QuerySQL="INSERT INTO M_Desideri (SoggettoC, DomandaC, MotivazioneC,DesiderioC,BisognoC,SoggettoS,RispostaS,MotivazioneS,DesiderioS,BisognoS,TipoEvento,TolleranzaC,URL_Teoria,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,Ora) SELECT '" & SoggettoC & "','" & DomandaC & "', '" & MotivazioneC & "','" & DesiderioC & "','" & BisognoC & "','" & SoggettoS & "','" & RispostaS & "','" & MotivazioneS & "','" & DesiderioS & "','" & BisognoS  & "'," & TipoEvento & "," & TolleranzaC &",'"  & URL_teoria &"','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella & "','" & FormatDateTime(now, 4) &"';" 
	'response.write("<br>"&QuerySQL)
    ConnessioneDB.Execute(QuerySQL)
    rsTabella.movenext()
	loop
 rsTabella0.movenext()
loop

%>
</body>
</html>
