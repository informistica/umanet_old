<%@ Language=VBScript %>


<%  if session("admin") = false and Session("identita") <> true then
		response.redirect("../../../../index.html") 
		end if
		%>
		
		
<%  Session.LCID=1040	
   id_classe=Request("id_classe")
   id_classe1=id_classe
   'cartella=Request("cartella")
   app=Request("app") ' vale 1 se sono stato chiamata da apprendimento
   logadmin=request("logadmin")
   CodiceAllievo = request("CodiceAllievo")
   PwdAllievoSHA256 = request("PwdAllievoSHA256")
   id_as=Request("id_as")
   identita=Request("identita") ' valorizzato se provengo dal quaderno e sono admin che assume identità/ruolo studente
   

				
   'response.Write(id_classe&" "&cartella&" "&app&" "&logadmin&" "&CodiceAllievo&" "&PwdAllievo)
   
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
%>
   
   <!-- #include file = "../var_globali.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_registrati.inc" -->
   
<%  

q="select Cartella from Classi where ID_Classe='"&id_classe&"'"
Set rsTabella = ConnessioneDB.Execute(q)
cartella=rsTabella(0)

	'memorizzazione dei parametri 
   Session("CodiceAllievo")=CodiceAllievo
   CodiceAllievo= Replace(CodiceAllievo, "'", chr(96))  ' DA SISTEMARE IMPEDENDO INSERIMENTO CARATTERI SPECIALI
   Response.Cookies("Dati")("CodiceAllievo") = CodiceAllievo
    if identita<>"" then
	    CodiceAllievo=Request.querystring("CodiceAllievo")
	    QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_ 
			  " WHERE CodiceAllievo='" & CodiceAllievo& "';"
	else
   QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_ 
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievoSHA256 & "';"
		end if
		response.write(QuerySQL) 
		
	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	  if not rsTabella.eof then
	   PwdAllievoDB=rsTabella("PasswordSHA256")
	   end if
	   if identita=1 then
		PwdAllievoSHA256=rsTabella("PasswordSHA256")
	   end if
	   
	   '			  	
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\loginpwdcritta.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				 
'				' chiudere recrodset del file in fondo alla pagina
'				objCreatedFile.WriteLine(QuerySQL)
'				'objCreatedFile.Close
'				'
	   
		'response.write(QuerySQL) 
  
  'se il risultato della query è nullo allora vuol dire che non è stato trovato nessun studente avente il codice 		specificato nella query
 ' objCreatedFile.WriteLine("<br>2"&strcomp(ucase(PwdAllievoDB),ucase(PwdAllievoMD5)))	
	If (rsTabella.EOF) or (strcomp(ucase(PwdAllievoDB),ucase(PwdAllievoSHA256))<>0)   Then 
	'If  (strcomp(PwdAllievoNew,PwdAllievo)<>0) Then
	  Session("Loggato") = False
	  'Response.write(QuerySQL)
	 
	response.write("errore")
	else
	' loggatto aggirno accessi 
	QuerySQLU="UPDATE Allievi SET Accessi = Accessi + 1 WHERE CodiceAllievo='" & CodiceAllievo & "';"
	ConnessioneDB.Execute(QuerySQLU)
	'objCreatedFile.WriteLine("<br>3"&QuerySQLU)	
	
	  if not((PwdAllievoDB=pwdAdmin) and (CodiceAllievo=codAdmin)) then
	' ho trovato lo studente e non è l'admin adesso devo stabilire a quale id_classe appartiene quindi eseguo query su stud_as_classe (precedentemente popolata) con id_as e id_stud
		' QuerySQL="SELECT Id_Classe FROM [dbo].[stud_as_classe] Where Id_Stud='"& CodiceAllievo &"' and Id_As=2;"
		 QuerySQL="SELECT Id_Classe FROM [dbo].[stud_as_classe] Where Id_Stud='"& CodiceAllievo &"' and Id_Classe='"&id_classe&"';"
		 
			'response.write("<br>"&QuerySQL) 
			'objCreatedFile.WriteLine(QuerySQL)
		 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' id_classe=rsTabella("Id_Classe")
	 end if
	 'response.write("<br>"&id_classe) 
	 
	 ' lo ripeto non è un errore, serve per ricaricare il record set
	  if identita<>"" then
	     
	     QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_ 
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievoSHA256 & "';"
	  else
	  QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_ 
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievoSHA256 & "';"
			  
	  end if
	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	 'objCreatedFile.WriteLine(QuerySQL)
		   if (PwdAllievoMD5=pwdAdmin) and (CodiceAllievo=codAdmin) then	
			  cartellaAdmin=rsTabella("Classe")
			  session("Admin")=true
		  end if 
		  QuerySQL="SELECT Data FROM [dbo].[3PERIODI] Where Id_Classe='"& id_classe &"' and Iniziale=1;"
 		  Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
		  if not rsTabella1.eof then
		 	 Session("DataCla")=rsTabella1(0)
			  DataClaq=cdate(inizio_anno)
		  else
		    DataClaq=cdate(inizio_anno)
		 	 Session("DataCla")=inizio_anno
		  end if
		  Session("DataCla2")=DataCla2Default
          Session("Loggato") = True
		'  DataClaq2= left(now(),10)
		  DataClaq2= FormatDateTime(now(),2)

		  Cognome=rtrim(rsTabella.Fields("Cognome")) 
		  Nome=rtrim(rsTabella.Fields("Nome"))
		  In_Quiz=rsTabella.Fields("In_Quiz")
		 
	  	  if (PwdAllievoDB=pwdAdmin) and (CodiceAllievo=codAdmin) then 
			  ' se sono Admin lascio id_classe quella della clsse in cui entro
			  ' quindi non faccio niente prendo da querystring
			  Session("Admin")=true	
			 ' id_classe="6COM"	
			 	 id_classe=id_classe1
			Session("identita")=false
		  else
			 Session("Admin")=false
			  Session("identita")=true
			 id_classe=rsTabella.Fields("Id_Classe")
		  end if
		  
		  'objCreatedFile.WriteLine(PwdAllievoDB&"=?"&pwdAdmin)
			
		  classe=rsTabella.Fields("Classe")
			
		  Session("stile")=rsTabella.Fields("Stile") 
		  Session("Cognome") = Cognome
		  Session("Nome") = Nome
		  Session("CodiceAllievo") = CodiceAllievo
		  Session("Username")= CodiceAllievo ' per la chat dopo disastro    
		  Session("Id_Classe")=id_classe
		  Session("cartella")=cartella
		  Session("Cartella")=cartella
		  Session("CartellaIniz")=cartella
		  Session("CartellaAdmin")=cartellaAdmin
		  Session("In_Quiz")=In_Quiz
		  Session("CodAdmin")=codAdmin	
		  Session("id_as")=id_as
		  Session("id_scuola")=id_scuola	 
		  Session("id_classe_img")=cartella ' prima era classe 17/09/2018 per il problema della foto profilo che non si vede cambiando classe associata
		  Session("CartellaAdmin") = "Admin"
		  Session("DBCopiatestonline") = True
		 
		  
		  Response.Cookies("Dati")("id_scuola")= id_scuola	
		  Response.Cookies("Dati")("id_as")= id_as	   
		  Response.Cookies("Dati")("Loggato")= Session("Loggato") 
		  Response.Cookies("Dati")("Cognome")= Session("Cognome") 
		  Response.Cookies("Dati")("Nome")=Session("Nome") 
		  Response.Cookies("Dati")("CodiceAllievo")= Session("CodiceAllievo") 
		  Response.Cookies("Dati")("Username")=Session("Username")  ' per la chat dopo disastro 
		  Response.Cookies("Dati")("DataTest")= Session("DataTest") 
		  Response.Cookies("Dati")("Id_Classe")=Session("Id_Classe") 
		  Response.Cookies("Dati")("cartella")=Session("cartella") 
		  Response.Cookies("Dati")("Cartella")=Session("Cartella") 
		  Response.Cookies("Dati")("CartellaAdmin")= Session("CartellaAdmin")
	      Response.Cookies("Dati")("In_Quiz")= Session("In_Quiz")
	      Response.Cookies("Dati")("CodAdmin")= Session("CodAdmin")
		  Response.Cookies("Dati")("Admin")= Session("Admin")
		  Response.Cookies("Dati")("Admin2")= Session("Admin2")	   
		  Response.Cookies("Dati")("stile")= Session("stile")
		  ' impostate in home.asp	  
		 Response.Cookies("Dati")("Materia")= Session("Materia") 
		 Response.Cookies("Dati")("ID_Materia")= Session("ID_Materia") 
		 Response.Cookies("Dati")("ID_Matsint")= Session("ID_Matsint")  
		 ' mi serve la chiave numerica per il DBMatprof per recuperare la login dell'admin
		 Response.Cookies("Dati")("idxMat")= Session("idxMat") 
		 Response.Cookies("Dati")("Cartella")= Session("Cartella") 
		 Response.Cookies("Dati")("DBCopiatestonline")= Session("DBCopiatestonline") 
		 Response.Cookies("Dati")("DBForum")= Session("DBForum") 
		 Response.Cookies("Dati")("DBLavagna")= Session("DBLavagna") 
		 Response.Cookies("Dati")("DBDiario")= Session("DBDiario") 
		 Response.Cookies("Dati")("DBDesideri")= Session("DBDesideri") 		
		 
		  Response.Cookies("Dati")("DataCla")=  Session("DataCla")		
		  Response.Cookies("Dati")("DataCla2")=  Session("DataCla2") 
		  
		  Response.Cookies("Dati")("DB")=Session("DB")
		  	'Session("Admin")=true	 
		  response.write("umanet=0&Cognome="&session("Cognome")&"&Nome="&session("Nome")&"&stile="&session("stile")&"&id_classe="&id_classe&"&classe="&classe&"&cod="&CodiceAllievo&"&DataClaq2="&DataClaq2&"&DataClaq="& DataClaq)
		 
		  		  
		 if identita<>"" then
						if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 
		 end if
	end if
	'objCreatedFile.Close
%>
