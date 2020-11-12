<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</head>

<body>
Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  


Set RecSet = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"'"
RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic

' CONTROLLA SE L'USERNAME INSERITO E' GIA' STATO USATO

 

' Chiude la connessione al DB

RecSet.Close
Set RecSet = Nothing

' FA LA CONDIZIONE PER VERIFICARE SE L'USERNAME
' IMMESSO E' GIA' STATO USATO...

IF usato = True then

' USERNAME GIA' USATO.
%>
<hr>
<p align="center"><b><font face="Verdana" size="2" color="#FF0000">Codice allievo inserito già in uso!</font></b></p>
<hr>
<%
Else

' NICK NON USATO...
' PROCEDE ALLA SUA REGISTRAZIONE...

Set RecSet = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM Allievi Order By CodiceAllievo Desc"
RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic

RecSet.Addnew

RecSet("CodiceAllievo") = username
RecSet("Password") = password
RecSet("Cognome") = cognome
RecSet("Nome") = nome
RecSet("Email") = email
RecSet("Classe") = classe
RecSet("Sezione") = sezione
RecSet("Anno")="2009-2010"
RecSet.Update

' CHIUDE LA CONNESSIONE AL DB

RecSet.Close
Set RecSet = Nothing
'if (cint(classe)=6) or (cint(classe)=7) then
		Set RecSet = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"' and Password='"&password&"';"
		RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
		id=RecSet("CodiceAllievo")
		RecSet.Close
		Set RecSet = Nothing
		
		Domanda="?"
		R1="?"
		R2="?"
		R3="?"
		R4="?"
		Chi="?"
		Cosa="?"
		Dove="?"
		Quando="?"
		Come="?"
		Perche="?"
		Quindi="?"
		RE=1
		Spiegazione="?"
		CodiceAllievo=id
		CodiceTest="6Classe"
		Modulo="6C" 
		DataTest="12/12/2112"  ' NB DOPO QUESTA DATA SE ESISTEREMO ANCORA DOVRO' METTERE UNA DATA SEMPRE MAGGIORE DELL'A/S in CORSO IN MODO DA FAR FUNZIONARE IL LEFT JOIN  NELLE QUERY PER LE CLASSIFICHE CON DATA VARIABILE
		' dovrò inserire aNCHE I CREDITI INIZIZIALI !
		Cartella="?"
		Voto=0
		In_Quiz=0
		
		QuerySQL="INSERT INTO Domande (Quesito, Risposta1, Risposta2,Risposta3,Risposta4,RispostaEsatta,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Domanda & "','" & R1 & "', '" & R2 & "','" & R3 & "','" & R4 & "','" & RE & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','"& Cartella &  "','"& In_Quiz &"';"
		ConnessioneDB.Execute QuerySQL 
		
		
		    QuerySQL="INSERT INTO Nodi (Chi, Cosa, Dove,Quando,Come,Perche,Quindi,Id_Stud,Id_Arg,Id_Mod,Data,Cartella) SELECT '" & Chi & "','" & Cosa & "', '" & Dove & "','" & Quando & "','" & Come & "','" & Perche & "','" & Quindi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Cartella & "';"
 
   ConnessioneDB.Execute QuerySQL 
		   
QuerySQL="INSERT INTO Frasi (Chi,Id_Stud,Id_Arg,Id_Mod,Data,Voto,Cartella,In_Quiz) SELECT '" & Chi & "','" & CodiceAllievo & "','" & CodiceTest & "','" & Modulo & "','"  & DataTest & "','" & Voto & "','" & Cartella & "','" & In_Quiz & "';"
 
   ConnessioneDB.Execute QuerySQL 
   
QuerySQL="INSERT INTO 2CREDITI (Id_Esercitazione,Id_Stud,Crediti) SELECT '" & Classe & "','" & CodiceAllievo & "','" & 1 & "';"
 
   ConnessioneDB.Execute QuerySQL    

'response.write(QuerySQL)		 
		 
		ConnessioneDB.Close
		Set ConnessioneDB = Nothing


</body>
</html>
