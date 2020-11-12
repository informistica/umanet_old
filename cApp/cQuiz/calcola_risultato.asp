 

<% 
Call Response.AddHeader("Access-Control-Allow-Origin", "*") 
paragrafo = Request.QueryString("paragrafo")
%>



   <% Response.Buffer=True 
   ' On Error Resume Next
  
  Session.LCID=1040
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
  
			   <!-- #include file = "../include/stringa_connessione.inc" -->
               
<%  
   
	if paragrafo <> 1 then
		tabelladb = "Risultati1"
	else
		tabelladb = "Risultati"
	end if
  
   CodiceTest = Request.QueryString("CodiceTest")
   CodiceAllievo = Request.QueryString("CodiceAllievo")
   Risultato = Request.QueryString("Risultato")
   Quiz = Request.QueryString("Quiz") 
   SessioneQuiz = Request.QueryString("SessioneQuiz") ' codice sessione, torneo
   OraTest=FormatDateTime(now, 4)
   DataTest=FormatDateTime(now, 2)
   Tipo=Request.QueryString("Tipo") ' 0 v/f, 1 singola, 2 multipla
   if Tipo="" then 
      Tipo=1
   end if	  

 QuerySQL=" Select * from "&tabelladb&" where CodiceTest='"&CodiceTest&"' and Sessione="&SessioneQuiz&"  and CodiceAllievo='"&CodiceAllievo&"' and Tipo="&Tipo&" order by Risultato desc,  Tentativi asc"
 set rsTabella=ConnessioneDB.Execute(QuerySQL)


 if rsTabella.eof then ' se è il primo lo inserisco altrimenti aggiorno
   posizione_old=0 
   messaggio="E' il tuo primo test"
   QuerySQL="  INSERT INTO "&tabelladb&" (CodiceAllievo, CodiceTest, Data,Ora,Risultato,In_Quiz,Sessione,Tipo,Tentativi) SELECT '" & CodiceAllievo & "','" & CodiceTest & "', '" & DataTest & "', '" & OraTest & "','" & Round(Risultato,0)    & "'," &Quiz & "," &SessioneQuiz  & ","&Tipo&",1;"
   ConnessioneDB.Execute(QuerySQL)
   
   
else
 ' calcolo la vecchia posizione 
 QuerySQL=" Select * from "&tabelladb&" where CodiceTest='"&CodiceTest&"' and Sessione="&SessioneQuiz&" and Tipo="&Tipo&" order by Risultato desc"
 set rsTabella=ConnessioneDB.Execute(QuerySQL)
 
	 posizione_old=1
	 do while not rsTabella.eof and (strcomp(rsTabella("CodiceAllievo"),CodiceAllievo)<>0)
	 posizione_old=posizione_old+1
	 rsTabella.movenext
	 loop
 
 ID_R=rsTabella("ID_R")
 QuerySQL="UPDATE "&tabelladb&" SET Tentativi = Tentativi +1, Risultato="&  Round(Risultato,0)&", Data ='"&DataTest&"', Ora='"&OraTest&"' where ID_R="&ID_R
 ConnessioneDB.Execute(QuerySQL)
end if	

 QuerySQL=" Select * from "&tabelladb&" where CodiceTest='"&CodiceTest&"' and Sessione="&SessioneQuiz&" and Tipo="&Tipo&" order by Risultato desc, Tentativi asc"
 set rsTabella=ConnessioneDB.Execute(QuerySQL)
 
	 posizione=1
	 do while not rsTabella.eof and (strcomp(rsTabella("CodiceAllievo"),CodiceAllievo)<>0)
	 posizione=posizione+1
	 rsTabella.movenext
	 loop 
if posizione_old<>0 then
 if posizione_old>posizione then
 messaggio="Hai guadagnato " & posizione_old- posizione &" posizioni"
 else if posizione_old<posizione then
         messaggio="Hai perso " & posizione-posizione_old &" posizioni"
      else
	     messaggio="Posizione invariata"
	  end if
end if
end if

response.write("{")

If Err.Number = 0 Then
	response.write("""stato"": ""1"","  &_
 """posizione"": """&posizione&"""," &_ 
 """messaggio"": """ & messaggio & """") 
Else
response.write("stato"": ""0"","  &_
 """errore"": """&Err.Description&"""") 
  Err.Number = 0
End If

response.write("}")





   %>
	 