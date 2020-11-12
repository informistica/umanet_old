 
 <% Session("Loggato") = True
    Session("DB")=1
	
		  Session("Cognome") = "Ospite"
		  Session("Nome") = "Ospite"
		  Session("CodiceAllievo") ="ospite"
		  Session("Username")= "ospite" ' per la chat dopo disastro 
		 ' Session("DataTest") = DataTest
		  Session("stile")="darkblue"  ' posso metterlo blue per fare test e vedere quando refresh acced a db
		   session("Id_Classe")="6COM"
		   Session("DataCla")="10/09/2013"
		   Session("DataCla2")="10/09/2016"
		    Session("DataClaq")="10/09/2013"
		   Session("DataClaq2")="10/09/2016"
		  
		  Session("cartella")="Expo"
		  Session("Admin")=False
		  session("ID_Materia")="materia_1"
		  
		    app=1
  			materia="Umanet 1"
 			' cartella="Expo" 
 			 id_materia=1
 			 Session("idxMat") =id_materia
			 Session("Materia")="Umanet 1"
			 session("DBCopiatestonline")="ok"
			 
			 Session("CodAdmin")=codAdmin
  
    'dim objFSO,objCreatedFile
			'	Const ForReading = 1, ForWriting = 2, ForAppending = 8
			'	Dim sRead, sReadLine, sReadAll, objTextFile
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
				 
				'	url="C:\Inetpub\umanetroot\expo2015Server\log\log_page_expo.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(session("Id_Classe") & "-" & session("Id_Classe") & "-" &session("DB") & "-" &Session("CodiceAllievo")&"-"&stringa_redirect_app)
 
 
			  
			  %>
  