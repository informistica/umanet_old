<%@ Language=VBScript %>

<% Call Response.AddHeader("Access-Control-Allow-Origin", "*") %>
 
<!-- #include file = "include/adovbs.inc" -->
<%
' VENGONO SOSTITUITI GLI APICI (') CON DUE APICI ('')
' PER EVITARE IL PROBLEMA "SQL INJECTION"

nome=Replace(Request("nome"), "'","") 
nome=ucase(left(nome,1))&lcase(right(nome,len(nome)-1))
cognome=Replace(Request("cognome"), "'","")
cognome=ucase(left(cognome,1))&lcase(right(cognome,len(cognome)-1))
username = Replace(Request("username"), "'","")
password = Replace(Request("password"), "'","")
password_conferma = Replace(Request("password_conferma"), "'","")
'mipiace = Replace(Request("mipiace"), "'","")
'nonmipiace=Request("nonmipiace")
'descriviti=Request("descriviti")
'classe=Request("classe")
'sezione=Request("sezione")
id_classe=Request("id_classe")
email=Request("email")
DB=Request("DB")
Session("DB")=DB
 

 
		' PERCORSO DEL DATABASE
		   StringaConnessione = Request.Cookies("Dati")("StrConn")
		 
			Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
		 
			
			%>   
                <!-- #include file = "include/var_globali.inc" --> 
			   <!-- #include file = "include/stringa_connessione.inc" -->
               
			<%  
		
		
		Set RecSet = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"'"
		RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
		' CONTROLLA SE L'USERNAME INSERITO E' GIA' STATO USATO	
		IF Not RecSet.Eof Then	
			usato = True
			Else
			usato = False
		End IF						
		RecSet.Close
		Set RecSet = Nothing
		
		' FA LA CONDIZIONE PER VERIFICARE SE L'USERNAME
		' IMMESSO E' GIA' STATO USATO...%>
        <%
		if strcomp(password,password_conferma)<>0 then
		  password_ko=1
		else
		  password_ko=0  
		end if
		%>
       
                      
 
 	<%	IF usato = True then
		
		' USERNAME GIA' USATO.
		response.write("{""stato"": ""0"","  &_
 """messaggio"": """&"Username non disponibile"""&"}")
		Else
		   if password_ko=1 then
		     response.write("{""stato"": ""0"","  &_
 """messaggio"": """&"Password non corrispondenti"""&"}")
		   else
		   
		
		 
				' NICK NON USATO...
				' PROCEDE ALLA SUA REGISTRAZIONE...
				    QuerySQL="Select Classe from Classi where Id_Classe='" & id_classe & "';" 
					Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
					classe=rsTabella("Classe")
					QuerySQL="Select * from Setting where Id_Classe='" & id_classe & "';" 
					Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
					' se raggiungo il limite ricomncio
					in_quiz=cint(rsTabella("In_Quiz"))
					max_in_quiz=cint(rsTabella("Max_In_Quiz"))
					if (in_quiz=max_in_quiz+1) then
					   in_quiz=1
					end if   
				
				
				Set RecSet = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM Allievi Order By CodiceAllievo Desc"
				RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
				
				RecSet.Addnew
				
				RecSet("CodiceAllievo") = username
				RecSet("PasswordSHA256") = password
				RecSet("Cognome") = cognome
				RecSet("Nome") = nome

				RecSet("Classe") = classe
				'RecSet("Sezione") = sezione
				RecSet("Anno")="2015-2016"
				RecSet("Id_Classe") = id_classe
				RecSet("In_Quiz") = in_quiz
				
			'	RecSet("Mipiace") = mipiace
			'	RecSet("Nonmipiace") = nonmipiace
			'	RecSet("Descriviti") = descriviti
				RecSet("Stile") = "blue"
				RecSet("Email") = email
				RecSet("Tag") = 2 ' indica gli utenti registrati dall'app
				
				
				RecSet.Update
				
				' CHIUDE LA CONNESSIONE AL DB
				RecSet.Close
				Set RecSet = Nothing
				 
				QuerySQL ="UPDATE Setting SET In_Quiz = " & cint(in_quiz)+1 & "  WHERE Id_Classe ='" &id_classe &"';"
					 ConnessioneDB.Execute(QuerySQL)
				
				'if (cint(classe)=6) or (cint(classe)=7) then
						Set RecSet = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT * FROM Allievi where CodiceAllievo= '" & username &"' and PasswordSHA256='"&password&"';"
						RecSet.Open SQL, ConnessioneDB, adOpenStatic, adLockOptimistic
						id=RecSet("CodiceAllievo")
						RecSet.Close
						Set RecSet = Nothing
						
				session("id_as")=2 ' poi farò query per persacer anno attivo		
				 QuerySQL="INSERT INTO stud_as_classe (Id_Stud,Id_As,Id_Classe) SELECT '" & username & "'," &  session("id_as") & ",'" & id_classe & "';"
				 
				' response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL 
				 
						
						'trasferisco in un file include usato anche da cClasse/promuoviti.asp
			
   
   %>
   
   <!-- #include file = "../include/inizializzaDB.asp" -->  
				 
				
		<%		
		
		  response.write("{""stato"": ""1"","  &_
 """messaggio"": """&"Registrazione effettuata"&"""," &_
 """classe"": """&classe&""""&"}")
				
				
			 
				
				  end if
				END IF
				%>	 
	 
	 
				 
			 
                      
                      
                      
                      
              
