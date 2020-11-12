 
 <% Session("Loggato") = True
    Session("DB")=2
	
		  Session("Cognome") = "Ospite"
		  Session("Nome") = "Ospite"
		  Session("CodiceAllievo") ="ospite1"
		  Session("Username")= "ospite1" ' per la chat dopo disastro 
		 ' Session("DataTest") = DataTest
		  Session("stile")="darkblue"  ' posso metterlo blue per fare test e vedere quando refresh acced a db
		   session("Id_Classe")="8COM"
		   Session("DataCla")="10/09/2013"
		   Session("DataCla2")="10/09/2017"
		    Session("DataClaq")="10/09/2013"
		   Session("DataClaq2")="10/09/2017"
		  
		  Session("cartella")="DOC"
		  Session("Admin")=False
		  session("ID_Materia")="materia_1"
		  
		    app=1
  			materia="Docenti"
 			' cartella="Expo" 
 			 id_materia=1
 			 Session("idxMat") =id_materia
			 Session("Materia")="Docenti"
			 session("DBCopiatestonline")="ok"
			 
			 Session("CodAdmin")=codAdmin
  
    
 'dim objFSO,objCreatedFile
			'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
			'	Dim sRead, sReadLine, sReadAll, objTextFile
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				 
					url="C:\Inetpub\umanetroot\expo2015Server\log\log_page1.txt"
				Set objCreatedFile = objFSO.CreateTextFile(url, True)
				objCreatedFile.WriteLine(session("Id_Classe") & "-" & session("Id_Classe") & "-" &session("DB") & "-" &Session("CodiceAllievo")&"-"&stringa_redirect_app)
				objCreatedFile.Close
 
			  
			  %>
  