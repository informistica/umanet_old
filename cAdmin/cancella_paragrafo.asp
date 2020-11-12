<%@ Language=VBScript %>
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../stile.css">
	 
 <script language="javascript" type="text/javascript"> 
function showText2() {//window.alert("Cancellazione effettuata ")
//location.href="../Home.asp"
//location.href=window.history.back();
 }
 </script>
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella,rsTabella1, QuerySQL,StringaConnessione,URL,RecSet
   
   
   Id_Mod=Request.QueryString("Id_Mod")
   Id_Classe=request.querystring("Id_Classe")
   Classe=Request.QueryString("Classe")

   Id_Par=Request.QueryString("Id_Par")
   Id_SotPar=Request.QueryString("Id_SotPar")
   
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%  

 
 'response.write("Pippo="&Id_SotPar&"<br>")
  
 ' prima devo cancellare  i compiti del paragrafo 
 if Id_SotPar="" then  ' cancello paragrafo
 
     QuerySQL ="DELETE   FROM preFrasi WHERE Id_Paragrafo ='" &Id_Par&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preNodi WHERE Id_Paragrafo ='" &Id_Par&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preDomande WHERE Id_Paragrafo ='" &Id_Par&"';"
	 ConnessioneDB.Execute(QuerySQL)
     QuerySQL ="DELETE   FROM Classi_Moduli_Paragrafi WHERE Id_Paragrafo ='" &Id_Par&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM Paragrafi WHERE Id_Paragrafo ='" &Id_Par&"';"
	 ConnessioneDB.Execute(QuerySQL)
Else ' cancello sottoparagrafo
	 QuerySQL ="DELETE   FROM preFrasi WHERE Id_Sottoparagrafo ='" &Id_SotPar&"';"
	 response.write(QuerySQL&"<br>")
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preNodi WHERE Id_Sottoparagrafo ='" &Id_SotPar&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM preDomande WHERE Id_Sottoparagrafo ='" &Id_SotPar&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE   FROM ParagrafiSottoparagrafi WHERE Id_Paragrafo ='" &Id_SotPar&"';"
	 ConnessioneDB.Execute(QuerySQL)
	QuerySQL ="DELETE  FROM Sottoparagrafi WHERE Id_Sottoparagrafo ='" &Id_SotPar&"';"
	 ConnessioneDB.Execute(QuerySQL)


end if

	'response.write(url)
	On Error Resume Next
	If Err.Number = 0 Then
		if Id_SotPar="" then
			Response.Write "<script>alert('Cancellazione del Paragrafo avvenuta!')</script> "
		Else
			Response.Write "<script>alert('Cancellazione del Sottoparagrafo avvenuta!')</script> "
		end if
	Else
	Response.Write Err.Description 
	Err.Number = 0
	End If


	if Request.ServerVariables("HTTP_REFERER") <>"" then 	 
		response.Redirect request.serverVariables("HTTP_REFERER") 
	end if 
	%>						 
						
 
	</body>
	</html>
	