<%
Session.Timeout=60
 if session("DBLavagna")="" then
	  response.redirect "../redirect.asp"
 end if
'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
'ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
 '             "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/UECDL/database/" & Session("DBLavagna")    


	'			Set objFSO = CreateObject("Scripting.FileSystemObject")
	'			url="C:\Inetpub\umanetroot\anno_2012-2013\logForum.txt"
	'			Set objCreatedFile = objFSO.CreateTextFile(url, True)
	'			objCreatedFile.WriteLine(Server.MapPath("../../database/" & Session("DBLavagna")))
	'			objCreatedFile.Close
 
ConnessioneDB2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
             "DBQ=" & Server.MapPath("../../database/" & Session("DBLavagna"))

 	  

%>
