<%
 	Session.Timeout=60
	 if (Session("DBCopiatestonline") ="") or (session("Id_Classe")="") or (session("CodiceAllievo")="")  then%>
	<!-- #include file = "stringa_connessione_refresh.asp" -->  
	
	      <% 
		  if (Session("DBLavagna")="") or (session("Id_Classe")="") or (session("CodiceAllievo")="")  then
		     response.redirect "redirect.asp"
	      end if
	 
	 end if
'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
' ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
 '             "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBDesideri")
    
	
	
	            'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\log_connessione.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(Server.MapPath("database/Copiaditestonline.mdb"))
'				objCreatedFile.Close

	 
    
	ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../database/" & Session("DBDesideri"))

 homesito="/expo2015/UECDL"   
 
%>
