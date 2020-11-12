<!-- calcola_risultato_MODBC3.asp -->
<html>
<head>	
	 
</head>
<body>
     

<%@ Language=VBScript %>

<% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet,k
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "stringa_connessione.inc" -->
   <!-- #include file = "stringa_connessione_forum.inc" -->
     <%  
   
Id_Classe=Request.QueryString("id_classe")
 
 dim objFSO,objCreatedFile
				Const ForReading = 1, ForWriting = 2, ForAppending = 8
				Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")

' per incrementare il num di pos  della predomanda/nodo/frase
QuerySQL="Select * from Allievi where Id_Classe='"&Id_Classe&"';"
set rsTabella=ConnessioneDB.Execute (QuerySQL) 
  response.write("<br>" & QuerySQL)
 
 cont=0
 DataTest="12/12/2112"
do while not rsTabella.eof		
			Messaggio="InizializzaDB"
	QuerySQL="INSERT INTO FORUM_MESSAGES (comments,CodiceAllievo,Id_Classe,Punti,DatePosted) SELECT '" &Messaggio & "','" & rsTabella("CodiceAllievo") & "','" & Id_Classe & "',1,'"&DataTest &"';"
 
'' 
url="C:\Inetpub\umanetroot\anno_2012-2013\logFORUM.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(QuerySQL)
				objCreatedFile.Close
  response.write("<br>" & QuerySQL)
   ConnessioneDB1.Execute QuerySQL 
		    
    
	rsTabella.movenext  
loop 

	On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento avvenuto! "
	Else
		Response.Write Err.Description 
		Err.Number = 0
	End If


   %>
	</font>   
	 
	
	</body>
	</html>
	