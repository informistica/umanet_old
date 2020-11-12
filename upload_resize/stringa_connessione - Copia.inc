<%'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 Session.Timeout=60
 if session("DBCopiatestonline")="" then
	  response.redirect "../redirect.asp"
 end if
	
'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
 '             "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/UECDL/database/" & Session("DBCopiaditestonline")
    
	
	
	            'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_connessione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(Server.MapPath("database/Copiaditestonline.mdb"))
'				objCreatedFile.Close
	
	
	ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../../database/" & Session("DBCopiaditestonline"))


 homesito="/anno_2013-2014/UECDL"   

 
%>
