<!-- calcola_risultato_MODBC3.asp -->
<html>
<head>	
	<link rel="stylesheet" type="text/css" href="../../../stile.css">
	 
   
  
   
</head>
<body>
    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

<%@ Language=VBScript %>
 
     
<% Response.Buffer=True 
   Dim ConnessioneDB, rsTabella, QuerySQL
 
   
   'Apertura della connessione al database  
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../cAdmin/include_mail.asp" -->
   <!-- #include file = "../var_globali.inc" -->
     <%  
'Id_Classe=Request.QueryString("Id_Classe")
'divid=Request.QueryString("divid")
CodiceAllievo=Request.QueryString("CodiceAllievo")  
Messaggio=Request.Form("txtMessaggio")
 
Messaggio = Replace(Messaggio, Chr(34), Chr(96))' sostituisco gli apici " con l'apice storto
Messaggio = Replace(Messaggio, Chr(39), Chr(96))' sostituisco gli apici ' con l'apice storto
'Messaggio=  Replace(Messaggio,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
Commentatore= Session("Cognome") & " " & left(Session("Nome"),1) &"."   
Azione=Messaggio

'cbEmail=request.QueryString("cbEmail")

cbEmail=request.QueryString("cbEmail")
DataClaq=request.QueryString("DataClaq") 
DataClaq2=request.QueryString("DataClaq2")

 %>

<% if day(date()) < 10 then
  giorno="0" & day(date()) 
	else
	giorno=day(date())
   end if

if len(year(date()) ) = 2 then
anno="20"& year(date())
elseif len(year(date()) ) =  3 then
anno="2"& year(date())
else
anno=year(date())
end if

'giornosettimana=giorniset(weekday(date()))
mese= month(date())

DataAvviso = giorno & "/" & mese& "/" & anno 

%>


<%

'DataAvviso=mid(FormatDateTime(Now,2),4,2)&"/" &left(FormatDateTime(Now,2),2)&"/"& right(FormatDateTime(Now,2),4)
   'QuerySQL="  INSERT INTO AVVISI (CodiceAllievo,Testo,Data,Azione,CodiceAllievo2,Commentatore)  SELECT '" & CodiceAllievo & "','" & Messaggio & "','" &  DataAvviso   & "', '" & Azione & "', '" & Session("CodiceAllievo") & "', '" & Commentatore & "';"   
   
   
   if len(Messaggio) < 250 then
	   QuerySQL="  INSERT INTO AVVISI (CodiceAllievo,Testo,Data,Azione,CodiceAllievo2,Commentatore)  SELECT '" & CodiceAllievo & "','" & Messaggio & "','" &  now()   & "', '" & Azione & "', '" & Session("CodiceAllievo") & "', '" & Commentatore & "';"   
	   ConnessioneDB.Execute (QuerySQL) 
	   response.write (QuerySQL)
	   response.Redirect request.serverVariables("HTTP_REFERER")
   else
		response.write("<script>alert('Impossibile inserire il messaggio: hai superato il numero massimo di caratteri!'); window.location.href='"&request.serverVariables("HTTP_REFERER")&"'</script>")
end if

  
' if (cbEmail<>"")  then
	
	' mes = ""
	' IsSuccess = false
	' sFrom = "info@umanet.net"
	' sMailServer = "127.0.0.1"
	' 'sBody = Trim(Request.Form("txtBody"))
	' sSubject = "Messaggio personale da " & Session("Cognome") & " " & left(Session("Nome"),1)& "." 
    ' QuerySQL="Select CodiceAllievo,Email from Allievi where CodiceAllievo='"&CodiceAllievo&"' and Email<>'' ;"
    ' set rsTabella=ConnessioneDB.Execute(QuerySQL) 
   
	' if  (strcomp(rsTabella("Email")&"","")<>0) then
	 ' sTo=rsTabella("Email")
	  ' sBody= Request("txtMessaggio")
	
	    ' ' response.write("Email prof")
		  ' linkAvviso=dominio&homesito&"/script/studente_domande.asp?divid="& Session("divid")&"&DataClaq="& DataClaq &"&DataClaq2="&DataClaq2&"&id_classe="&Session("Id_Classe")&"&classe="&Session("Cartella")&"&cod="&CodiceAllievo&"&by_email=1&DBCopiatestonline="&Session("DBCopiatestonline")&"&DBForum="&Session("DBForum")&"&DBLavagna="&Session("DBLavagna")&"&DBDiario="&Session("DBDiario")&"&id_materia="&Session("Id_Materia")&"&materia="&Session("Materia")
 ' sBody = sBody &"     "& linkAvviso

 
	   ' TestEMail()
       ' response.write("<br>Inviata mail a " & sTo)
	' else
	      ' response.write("<br>Non è presente l'indirizzo email")

     ' end if
' end if
	
	
	
	
	
	
	
	
	
	
	
	
	

	' On Error Resume Next
	' If Err.Number = 0 Then
		' Response.Write "Inserimento dell'avviso avvenuto! "
		
	' Else
		' Response.Write Err.Description 
		' Err.Number = 0
	' End If
     
 			' if Request.ServerVariables("HTTP_REFERER") <>"" then 
							' response.Redirect request.serverVariables("HTTP_REFERER") 
		 ' end if 
' connessioneDB.close()
' set connessioneDB = nothing
   %>
	 


 
 
    
 
		  
          
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../../home_app.asp?id_classe=<%=Id_Classe%>&divid=<%=divid%>"> Torna alla pagina Apprendimento... </a></h3> 
	
  
			</div>
 <!-- se il login è corretto richima la pagina per inserire le domande del test -->

	
	</body>
	</html>
	