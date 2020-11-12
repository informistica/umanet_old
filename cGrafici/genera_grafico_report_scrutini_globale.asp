<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 

'   https://docs.fusioncharts.com/charts/

Id_Eser=request.querystring("ID_ESER")
classifiche=Request.QueryString("classifiche") ' vale 1 se mostro le classifiche
verifiche=Request.QueryString("verifiche")

  Dim ConnessioneDB,ConnessioneDB1, rsTabella,rsTabella1, QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione
   
   'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 
	%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   </body>
<center>
<p>
 
</p>
<div class="citazioni" ><div> <span style="font-style: normal">

<b><font size="3">GESTIONE</font>&nbsp;</b> </span></div>
<hr>
<%
'   
Data=request.QueryString("Data")   
Data2=request.QueryString("Data2")  
ID_Classe=request.QueryString("ID_Classe")  
Ordina=request.QueryString("Ordina")

QuerySQL="SELECT count(*)  FROM Allievi Where Id_Classe='"&Id_Classe &"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
numRec=rsTabella(0)

QuerySQL="SELECT Cognome,Nome,CodiceAllievo,Classe FROM Allievi Where Id_Classe='"&Id_Classe &"' order by Cognome, Nome  "
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)

 QuerySQL="select count(*) from 4PERIODI_CLASSIFICA where CodiceAllievo='" &rsTabella("CodiceAllievo") &"'"&_
			  " AND  (Dal >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Al <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
 			   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
   
    num_periodi_calcola=rsTabella1(0)
	
	QuerySQL= "SELECT count(*) FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo") &"' and Scrutini=1 "&_
		  " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);"
		 'response.write(QuerySQL)
 			   Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)
    num_voti=rsTabella2(0)
 
Dim Studenti()
Dim Medie()
ReDim Studenti(numRec+1)
ReDim Medie(numRec+1)
Dim app,app2

Dim FileObject,riga
Set FileObject=CreateObject("Scripting.FileSystemObject")
esiste=FileObject.FileExists("data.xml")
'Response.Write("<p>il file esiste? "&esiste&"</p>")
Set FileObject=Nothing

Dim objFSO,objCreatedFile
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim sRead, sReadLine, sReadAll, objTextFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
url=Server.MapPath(homesito & "/Grafici")& "/data.xml"   
Set objCreatedFile = objFSO.CreateTextFile(url, True)
' Write a line with a newline character.
gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())


' carico i vettori dal form chiamante per ordinarli 
if Ordina<>"" then
riga="<chart caption='MEDIA DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Media' numberPrefix='V.'>"
objCreatedFile.WriteLine(riga)
 'response.Write(numRec)
 
for i=1 to numRec
   Studenti(i)=request.form("Studente"&i)
   Medie(i)=request.form("MediaT"&i)
   response.Write("media"& Medie(i) & "<br>")
next 

 
'ordino
for i=1 to numRec
   for j=1 to numRec
      if Medie(i)>Medie(j) then ' scambia
	     app=Medie(i)
		 Medie(i)=Medie(j)
		 Medie(j)=app
		 app2=Studenti(i)
		 Studenti(i)=Studenti(j)
		 Studenti(j)=app2
	  end if
    next
  next

  for i=1 to numRec-1
   ' response.Write( Studenti(i)) 
	'response.Write( abs(fix(Medie(i)*10)) & "<br>")
	
	
            riga="<set label='" &Studenti(i) &"' value='" &abs(fix(Medie(i)*10)) & "'/>"
			 objCreatedFile.WriteLine(riga)
			 response.write(riga&"<br>")
  next

else ' se non devo ordinare e quindi prendere i dati dal form che ha chiamato devo ricalcolare le medie, quindi o verifiche o classifiche
    QuerySQL="SELECT Cognome,Nome,CodiceAllievo,Classe FROM Allievi Where Id_Classe='"&Id_Classe &"' order by Cognome, Nome  "
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
   ' CodiceAllievoCampione=rsTabella(0)  
 
  
    riga="<chart caption='MEDIA DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Voti' numberPrefix='V.'>"
    objCreatedFile.WriteLine(riga)%>
   
    
	
	
	
	<%' per ogni studente vado a vedere i voti memorizzati nei vari periodi, come nella cronologia del quaderfno
    numRec=1
    do while not rsTabella.eof %>
			<% QuerySQL="select * from 4PERIODI_CLASSIFICA where CodiceAllievo='" &rsTabella("CodiceAllievo") &"'"&_
			  " AND  (Dal >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Al <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
 			   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
               QuerySQL= "SELECT distinct [2CREDITI].Id_Esercitazione, Descrizione,Data,Scrutini,Crediti  FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo") &"' and Scrutini=1 "&_
		 " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);" 		  
 		 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)%>

 
             <% media=0 
			  for k=1 to num_periodi_calcola%> 
                  
                  <% media=media+rsTabella1("Vv") %>
                  <% rsTabella1.movenext()%>
             <%next%>
             <% media=media/(num_periodi_calcola) %>
              
               
                <% media2=0 
			    assente=0
			  for k=1 to num_voti%> 
              <% if not rsTabella2.eof then%>
			    
                  
                  <% media2=media2+rsTabella2("Crediti") %>
                  <% rsTabella2.movenext()%>
               <%else%>
                 <%  assente=assente+1%>
                   
               <%end if%>
             <%next%>
             <% 
			    if (num_voti-assente>0) then %>
                  <% media2=media2/(num_voti-assente) %>
                 
				<%else
				   media=0%>
				   
				<% end if  
				  
			%>
             
                
               <%
			   mediaT=((fix(media*10)/10) + (fix(media2*10)/10))/2
			   response.write(fix(mediaT*10))
			   
		 	  riga="set label='" &rsTabella("Cognome") &"' value='" & fix(mediaT*10) & "'/"
			'  riga= "<set label='"  
			  response.write(riga &"<br>")
			  'response.write(cognome & " " & media&"<br>") 
			 
			 objCreatedFile.WriteLine("<"&riga&">")
			 numRec=numRec+1
	    rsTabella.movenext
	loop
	rsTabella.close
	ConnessioneDB.close
			   %>
        
             
<%
end if ' if Ordina<>""

riga="</chart>"
objCreatedFile.WriteLine(riga)
objCreatedFile.Close
On Error Resume Next
If Err.Number = 0 Then

Response.redirect("../Grafici/genera.html")
Else
Response.Write Err.Description 
Err.Number = 0
End If

%>
<h4 style="text-align: center"><i><a href="../../home.asp" >Vai all'HomePage</a> </h4>
</center>
</html>