<%@ Language=VBScript %>

<%  Session.LCID=1040
   

  ' cartella=Request("cartella")
   app=Request("app") ' vale 1 se sono stato chiamata da apprendimento
   logadmin=request("logadmin")
   CodiceAllievo = request("CodiceAllievo")
   PwdAllievoSHA256 = request("PwdAllievoSHA256")
   id_as=Request("id_as")



   'response.Write(id_classe&" "&cartella&" "&app&" "&logadmin&" "&CodiceAllievo&" "&PwdAllievo)

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")

%>

   <!-- #include file = "../var_globali.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_registrati.inc" -->

<%


 
	'memorizzazione dei parametri
   Session("CodiceAllievo")=CodiceAllievo
   CodiceAllievo= Replace(CodiceAllievo, "'", chr(96))  ' DA SISTEMARE IMPEDENDO INSERIMENTO CARATTERI SPECIALI
   Response.Cookies("Dati")("CodiceAllievo") = CodiceAllievo

   QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievoSHA256 & "';"

	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	  if not rsTabella.eof then
	   PwdAllievoDB=rsTabella("PasswordSHA256")
	   end if

	


	' Set objFSO = CreateObject("Scripting.FileSystemObject")

				' url="C:\inetpub\umanetroot\expo2015Server\loginpwdcritta.txt"
				' Set objCreatedFile = objFSO.CreateTextFile(url, True)

				' ' chiudere recrodset del file in fondo alla pagina
				' objCreatedFile.WriteLine(QuerySQL)
				' 'objCreatedFile.Close
' '				'

		'response.write(QuerySQL)

  'se il risultato della query è nullo allora vuol dire che non è stato trovato nessun studente avente il codice 		specificato nella query
 ' objCreatedFile.WriteLine(strcomp(ucase(PwdAllievoDB),ucase(PwdAllievoMD5)))	 'all'inizio c'era <br>2 (da lasciare?)
	If (rsTabella.EOF) or (strcomp(ucase(PwdAllievoDB),ucase(PwdAllievoSHA256)) <> 0)   Then
	'If  (strcomp(PwdAllievoNew,PwdAllievo)<>0) Then
	  Session("Loggato") = False
	  'Response.write(QuerySQL)


	response.write("errore")

	'response.write(QuerySQL&"<br>"&PwdAllievoDB&"<br>"&PwdAllievoSHA256&"<br>")
	else

	  ' id_classe=Request("id_classe")
	  id_classe=rsTabella("Id_Classe")
      id_new = id_classe


	q="select Cartella from Classi where ID_Classe='"&id_classe&"'"
	Set rsTabellaC = ConnessioneDB.Execute(q)
	cartella=rsTabellaC(0)
	' loggatto aggiorno accessi
	QuerySQLU="UPDATE Allievi SET Accessi = Accessi + 1 WHERE CodiceAllievo='" & CodiceAllievo & "';"
	ConnessioneDB.Execute(QuerySQLU)
	'objCreatedFile.WriteLine("<br>3"&QuerySQLU)

	  if not((PwdAllievoDB=pwdAdmin) and (CodiceAllievo=codAdmin)) then
	' deprecato perchè adesso è una classe unica per tutti i tre anni
	' ho trovato lo studente e non è l'admin adesso devo stabilire a quale id_classe appartiene quindi eseguo query su stud_as_classe (precedentemente popolata) con id_as e id_stud
		' QuerySQL="SELECT Id_Classe FROM [dbo].[stud_as_classe] Where Id_Stud='"& CodiceAllievo &"' and Id_As=2;"
		' QuerySQL="SELECT Id_Classe FROM [dbo].[stud_as_classe] Where Id_Stud='"& CodiceAllievo &"' and Id_Classe='"&id_classe&"';"

			'response.write("<br>"&QuerySQL)
			'objCreatedFile.WriteLine(QuerySQL)
		' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		' id_classe=rsTabella("Id_Classe")

	 end if
	 'response.write("<br>"&id_classe)

	 ' lo ripeto non è un errore, serve per ricaricare il record set
	  QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievoSHA256 & "';"
	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	   'id_new = rsTabella("Id_Classe")
	 'objCreatedFile.WriteLine(QuerySQL)
		   if (PwdAllievoSHA256=pwdAdmin) and (CodiceAllievo=codAdmin) then
			  cartellaAdmin=rsTabella("Classe")
			  'id_new = "6COM"
			  session("Admin")=true
		  end if

		  QuerySQL="SELECT Data FROM [dbo].[3PERIODI] Where Id_Classe='"& rsTabella("Id_Classe") &"' and Iniziale=1;"
		  Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

		  'Set objFSO = CreateObject("Scripting.FileSystemObject")

				'url="C:\inetpub\umanetroot\expo2015Server\loginpwdcritta.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)

				' chiudere recrodset del file in fondo alla pagina
				'objCreatedFile.WriteLine(QuerySQL)
				'objCreatedFile.WriteLine(rsTabella1(0))
				'objCreatedFile.Close

		  if not rsTabella1.eof then
		 	 Session("DataCla")=rsTabella1(0)
			  Session("DataClaq")=rsTabella1(0)
			  'DataClaq=cdate(inizio_anno)
			  DataClaq=cdate(Session("DataCla"))
		  else
		    DataClaq=cdate(inizio_anno)
		 	 Session("DataCla")=inizio_anno
			  Session("DataClaq")=inizio_anno
			 'DataClaq="12/09/2017"
		  end if

		  ' Session("DataCla")=rsTabella1(0)
			  ' 'DataClaq=cdate(inizio_anno)
			  ' DataClaq=cdate(Session("DataCla"))

		  ' if DataClaq = "" then
		  ' DataClaq=cdate(inizio_anno)
		 	 ' Session("DataCla")=inizio_anno
		  ' end if

		  Session("DataCla2")=DataCla2Default
		  Session("DataClaq2")=DataCla2Default
          Session("Loggato") = True
		' DataClaq2= left(now(),10)
		  DataClaq2= DateAdd("d",1,FormatDateTime(now(),2))
		  Session("DataClaOld") = FormatDateTime(now(),2)

		  Cognome=rtrim(rsTabella.Fields("Cognome"))
		  Nome=rtrim(rsTabella.Fields("Nome"))
		  In_Quiz=rsTabella.Fields("In_Quiz")

	  	  if (PwdAllievoDB=pwdAdmin) and (CodiceAllievo=codAdmin) then
			  ' se sono Admin lascio id_classe quella della clsse in cui entro
			  ' quindi non faccio niente prendo da querystring
			  Session("Admin")=true
			  id_classe="6COM"
			 '' id_classe=Request("id_classe")
		  else
			 Session("Admin")=false
			 id_classe=rsTabella.Fields("Id_Classe")
			 idclasseimg=rsTabella("Classe")

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
		 ' Session("CartellaIniz")=classe ' modificato il 19/9/19 per il problema "con questo utente non puoi inserire"
		  Session("CartellaIniz")=cartella
		  Session("CartellaAdmin")=cartellaAdmin
		  Session("In_Quiz")=In_Quiz
		  Session("CodAdmin")=codAdmin
		  Session("id_as")=id_as
		  Session("id_scuola")=id_scuola
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
		 Response.Cookies("Dati")("id_classe_img") = idclasseimg
		 session("id_classe_img") = idclasseimg
		  Response.Cookies("Dati")("DataCla")=  Session("DataCla")
		  Response.Cookies("Dati")("DataCla2")=  Session("DataCla2")

		  if session("DBLogin") = "" then
			session("DBLogin") = session("DB")
		end if

		  Response.Cookies("Dati")("DB")=Session("DB")


		  	' dim objFSO,objCreatedFile
				' Const ForReading = 1, ForWriting = 2, ForAppending = 8
				' Dim sRead, sReadLine, sReadAll, objTextFile
				' Set objFSO = CreateObject("Scripting.FileSystemObject")
				' 'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
				' url="C:\inetpub\umanetroot\expo2015Server\logLOGINNEW.txt"
				' Set objCreatedFile = objFSO.CreateTextFile(url, True)
				' objCreatedFile.WriteLine("Cognome="&session("Cognome")&"&Nome="&session("Nome")&"&stile="&session("stile")&"&id_classe="&id_classe&"&classe="&classe&"&cod="&CodiceAllievo&"&DataClaq2="&DataClaq2&"&DataClaq="& DataClaq)
				' objCreatedFile.Close


		  	'Session("Admin")=true
		  response.write("umanet=0&Cognome="&session("Cognome")&"&Nome="&session("Nome")&"&stile="&session("stile")&"&id_classe="&id_classe&"&classe="&classe&"&cod="&CodiceAllievo&"&DataClaq2="&DataClaq2&"&DataClaq="& DataClaq)
		  'response.write("Cognome="&session("Cognome")&"&Nome="&session("Nome")&"&stile="&session("stile")&"&id_classe="&id_classe&"&classe="&classe&"&cod="&CodiceAllievo&"&DataClaq2="&DataClaq2&"&DataClaq="& DataClaq&"&scegli=2")

	end if
	'objCreatedFile.Close
%>
