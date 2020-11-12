<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
<% Session.CodePage = 65001 %>

<%  'Session.LCID=1040	
  ' id_classe=Request("id_classe")
'   id_as=Request("id_as")
'   cartella=Request("cartella")
'   app=Request("app") ' vale 1 se sono stato chiamata da apprendimento
'   logadmin=request("logadmin")
   DB=request.querystring("DB")
   CodiceAllievo = request.querystring("CodiceAllievo")
   PwdAllievo = request.querystring("PwdAllievo")
   
   ' Session("DB")=DB
  ' Response.Cookies("Dati")("DB")=DB
  ' Session("DBCopiatestonline")=DB	
    id_materia=1
' session("ID_Materia")="materia_"&id_materia		
   'response.Write(id_classe&" "&cartella&" "&app&" "&logadmin&" "&CodiceAllievo&" "&PwdAllievo)
   
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  
%>
   
   <!-- #include file = "../../var_globali.inc" --> 
   <!-- #include file = "../include/stringa_connessione.inc" -->
               
  
   
<%  
 
				  
				  
				  
'ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
'" Initial Catalog=Copiaditestonline; User Id=informistica; Password=123Maurosho;"

 homesito="/expo2015/UECDL"   
	'memorizzazione dei parametri 
  ' Session("CodiceAllievo")=CodiceAllievo
   'CodiceAllievo= Replace(CodiceAllievo, "'", chr(96))  ' DA SISTEMARE IMPEDENDO INSERIMENTO CARATTERI SPECIALI
  ' Response.Cookies("Dati")("CodiceAllievo") = CodiceAllievo
  
   QuerySQL="SELECT Cognome,Nome,PasswordSHA256,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		  " FROM [dbo].[Allievi]" &_ 
			  " WHERE CodiceAllievo='" & CodiceAllievo& "' and PasswordSHA256 = '" & PwdAllievo & "';"
	   Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	     if not rsTabella.eof then
	   PwdAllievoDB=rsTabella("PasswordSHA256")
	   end if
	 
		 ' QuerySQL="SELECT Cognome,Nome,Password,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
	 		'  " FROM [dbo].[Allievi]" &_ 
			'  " WHERE CodiceAllievo='" & rsTabella1("CodiceAllievo") & "' and Password = '" & PwdAllievo & "';"
			
			'  QuerySQL="SELECT Cognome,Nome,Password,In_Quiz,Id_Classe,Classe,CodiceAllievo,Stile"&_
'	 		  " FROM [dbo].[Allievi]" &_ 
'			  " WHERE Password='" & PwdAllievoNew & "';"
			  
			  
	  ' Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		 
  
  'se il risultato della query è nullo allora vuol dire che non è stato trovato nessun studente avente il codice 		specificato nella query
	If (rsTabella.EOF) or (strcomp(ucase(PwdAllievoDB),ucase(PwdAllievo))<>0)   Then 
	'If  (strcomp(PwdAllievoNew,PwdAllievo)<>0) Then
	  Session("Loggato") = False
	  'Response.write(QuerySQL)
	response.write("errore")
	else
	' ho trovato lo studente adesso devo stabilire a quale id_classe appartiene quindi eseguo query su stud_as_classe (precedentemente popolata) con id_as e id_stud
	
		   if (PwdAllievo=pwdAdmin) and (CodiceAllievo=codAdmin) then	 
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
		 
	  	  if (PwdAllievo=pwdAdmin) and (CodiceAllievo=codAdmin) then 
			  ' se sono Admin lascio id_classe quella della clsse in cui entro
			  ' quindi non faccio niente prendo da querystring
			  Session("Admin")=true	
			  id_classe="6COM"		  
		  else
			 Session("Admin")=false
			 id_classe=rsTabella.Fields("Id_Classe")
		  end if
			
		  classe=rsTabella.Fields("Classe")
			utente= Cognome & " " & left(Nome,1)&"."
		   
		   response.write(" { "  &_
 """classe"": """ & classe& """," &_
 """utente"": """ & utente & """," &_
 """titolo"": ""Test di prova"","  &_
 """totale"": ""2"","  &_
 """domanda1"": ""Testo domanda 1"","  &_
 """risposta1.1"": ""Testo R11"","  &_
 """risposta1.2"": ""Testo R12"","  &_
 """risposta1.3"": ""Testo R13"","  &_
 """risposta1.4"": ""Testo R14"","  &_
  """domanda2"": ""Testo domanda 2"","  &_
 """risposta2.1"": ""Testo R21"","  &_
 """risposta2.2"": ""Testo R22"","  &_
 """risposta2.3"": ""Testo R23"","  &_
 """risposta2.4"": ""Testo R24"""  &_
"}")		  
		   
		   
		   
		   
	end if
%>

<% Session.CodePage = 1252 %>