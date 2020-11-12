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
function showText() {window.alert("Non puoi cancellare i test degli altri studenti!")

location.href="studente_domande.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>"
//location.href=window.history.back();
 }
 
 function showText2() {
	 window.alert("Test cancellati da tutte le bacheche!")

 
 }
 </script>
</head>

     

   <% Response.Buffer=True 
   Dim ConnessioneDB,  rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo,idA
  
   CodiceTest=request.QueryString("CodiceTest")
   tipoTest=request.QueryString("tipoTest") ' per la tabella Risultati o Risultati1
   DataTest= Request.QueryString("DataTest")
   SessioneQuiz=request.QueryString("SessioneQuiz")
   aggiorna=request.QueryString("aggiorna")
  


' 
		Set ConnessioneDB  = Server.CreateObject("ADODB.Connection")
		 %>   
		 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		 <% 
		if tipoTest=0 then '
		  QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati.Risultato, Risultati.Data,Risultati.Ora, Risultati.CodiceTest,Risultati.Risultato*8/100 as [PUNTI],Risultati.ID_R " &_
" FROM Allievi INNER JOIN Risultati ON Allievi.CodiceAllievo = Risultati.CodiceAllievo " &_
" WHERE   Risultati.CodiceTest='"&  CodiceTest & "' and Risultati.Sessione="&SessioneQuiz &_
" ORDER BY Allievi.Cognome Asc, Risultati.Ora; "
		 
		else
		
		
		 QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati1.Risultato, Risultati1.Data,Risultati1.Ora, Risultati1.CodiceTest,Risultati1.Risultato*8/100 as [PUNTI],Risultati1.ID_R " &_
" FROM Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo " &_
" WHERE  Risultati1.CodiceTest='"&  CodiceTest & "' and Risultati1.Sessione="&SessioneQuiz &_
" ORDER BY Allievi.Cognome Asc, Risultati1.Ora; "
		 end if
		 
		Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
			i=1
			Do until rsTabella.EOF
			 idR=rsTabella("ID_R")
		   if clng(Request.Form("cbDelete"&i)<>0) then ' se  lo devo cancellare lo canncello  , basta che non sono stato chiamato per aggiornare
				 if aggiorna="" then
				  if tipoTest=0 then '
						QuerySQL ="DELETE   FROM Risultati where ID_R =" &idR&";"	
				  else
						QuerySQL ="DELETE   FROM Risultati1 where ID_R =" &idR&";"	
				  end if
		     end if
			   ConnessioneDB.Execute(QuerySQL)
			 'response.write(QuerySQL)
			end if 
			' aggiorno il risultato 
			
				if( Request.Form("txtCambia"&i)<> "") then 
				    if tipoTest=0 then '
						 QuerySQL ="UPDATE Risultati SET Risultato = " & clng(Request.Form("txtCambia"&i)) & " where ID_R =" &idR&";"	
		 
				  else
						 QuerySQL ="UPDATE Risultati1 SET Risultato = " & clng(Request.Form("txtCambia"&i)) & " where ID_R =" &idR&";"	
		 
				  end if
				 ConnessioneDB.Execute(QuerySQL) 
				' response.write( QuerySQL )
				end if
			 
		   	
			i=i+1
			rsTabella.movenext()
		   loop	
		  set ConnessioneDB = nothing
		
'	   
 
 %>
<body>
      
<%	    

 

   if Request.ServerVariables("HTTP_REFERER") <>"" then 
									response.Redirect request.serverVariables("HTTP_REFERER") 
								end if %>
   
 
	 
		
      
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	