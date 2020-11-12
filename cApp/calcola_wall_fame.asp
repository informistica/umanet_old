
<%
Call Response.AddHeader("Access-Control-Allow-Origin", "*")
paragrafo = Request.QueryString("paragrafo")
%>

 <% Response.Buffer=True 
   ' On Error Resume Next

  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
 
			   <!-- #include file = "include/stringa_connessione.inc" -->
               
<%  
	
	if paragrafo <> 1 then
		vistadb = "RISULTATI_ALLIEVI1"
	else
		vistadb = "RISULTATI_ALLIEVI"
	end if
	

'  response.write("Default LCID is: " & Session.LCID & "<br>")
'response.write("Date format is: " & date() & "<br>")
'response.write("Date forma2t is: " &FormatDateTime(now(),2) & "<br>") 
Session.LCID=1040
'response.write("Date format is: " & date() & "<br>")
'response.write("Date forma2t is: " &FormatDateTime(now(),2) & "<br>") 	



   CodiceTest = Request.QueryString("CodiceTest")
   SessioneQuiz = Request.QueryString("SessioneQuiz")
   Tipo=Request.QueryString("Tipo")
   
    QuerySQL=" Select count(*) from "&vistadb&" where CodiceTest='"&CodiceTest&"'  and  Sessione='"&SessioneQuiz&"' and  Tipo="&Tipo&";"
	set rsTabella=ConnessioneDB.Execute(QuerySQL)
	NumQuiz=rsTabella(0)
		
	if NumQuiz>0 then
	
	QuerySQL=" Select * from "&vistadb&" where CodiceTest='"&CodiceTest&"' and  Tipo="&Tipo&" and  Sessione='"&SessioneQuiz&"' order by Risultato desc, Cognome asc"
	set rsTabella=ConnessioneDB.Execute(QuerySQL)
	
	' response.write("<br>""titolo"": """ & Replace(rsTabella("Titolo"),",","")& """," &_ """totale"": """&NumQuiz&""",")
    if paragrafo <> 1 then
		response.write(rsTabella("Titolo")&"," & NumQuiz &"$<br>")
	else
		response.write(rsTabella("Expr1")&"," & NumQuiz &"$<br>")
	end if
	
	 i=1
	 do while not rsTabella.eof 
	  datashort=(Day(rsTabella("Data"))) &"/"& month(rsTabella("Data"))& "/"&right(rsTabella("Data"),2)
	 utente=rtrim(rsTabella("Cognome")) & " " & left(rsTabella("Nome"),1)&"."
    ' response.write(i &","& utente &","& rsTabella("Risultato") &"," & rsTabella("Data") &","& rsTabella("Tentativi")&"$<br>"  )
	'response.write(","&i &","& utente &","& rsTabella("Risultato") &"," & datashort &","& rsTabella("Tentativi")&"$<br>"  )
	response.write(","&i &","& utente &","& rsTabella("Risultato") &"," & datashort &","&"$<br>"  ) 'tolgo i tentativi perché nella vista RISULTATI_ALLIEVI non c'è la colonna "Tentativi"
	 
	 ' response.write("<br>""posizione"&i&""":""" &i &""","  &_
'	  	"<br>""utente"&i&""":""" &utente &""","  &_	  	  
' 		"<br>""risultato"&i&""":""" &rsTabella("Risultato") &""","  &_
' 		"<br>""data"&i&""":""" &rsTabella("Data") &""","  &_
' 		"<br>""ora"&i&""":""" &left(rsTabella("Ora"),8) &""""&",")

	 i=i+1
	 rsTabella.movenext
	 loop
	 
	 else
		QuerySQL="Select Titolo from [2SESSIONI_QUIZ] where ID_Sessione="&SessioneQuiz&";"
		set rsTabella=ConnessioneDB.Execute(QuerySQL)
		TitoloQuiz=rsTabella(0)
		response.write(TitoloQuiz&"$")
	 end if

   %>
	 