<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
</head>
<link rel="stylesheet" type="text/css" href="../../stile.css">
<body>

<% 

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
   Medie(i)=request.form("Media"&i)
   response.Write( Medie(i) & "<br>")
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
 
  if (classifiche<>"") then
    riga="<chart caption='PUNTEGGI DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Voti' numberPrefix='V.'>"
    objCreatedFile.WriteLine(riga)
   
    QuerySQL="SELECT count(*) FROM [dbo].[3PERIODI] Where ID_Classe='"& ID_Classe &"'" &_	
	" AND  (Data>=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
	" AND Data<=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);"
	
   else ' verifiche<>""
      riga="<chart caption='VERIFICHE DEGLI STUDENTI' subcaption='" & gira_data &"' xAxisName='Studente' yAxisName='Voti' numberPrefix='V.'>"
      objCreatedFile.WriteLine(riga)
      QuerySQL= "SELECT count(*) FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo")&"' and Scrutini=1 "&_
		  " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);"	
   end if 
	Set rsTabella2 = ConnessioneDB.Execute(QuerySQL) 
	num_periodi_calcola=rsTabella2(0)
	
	
   'response.Write(QuerySQL & "<br>")
   
  do while not rsTabella.eof %>
			<% cognome=rsTabella("Cognome")
			 if (classifiche<>"") then
					QuerySQL="select * from 4PERIODI_CLASSIFICA where CodiceAllievo='" &rsTabella("CodiceAllievo") &"'"&_
				  " AND  (Dal >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
				  " AND Al <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
				  ");"
				'  response.Write(QuerySQL & "<br>")
				   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)%>
				 <% media=0 
				  for k=1 to num_periodi_calcola%> 
					  <% media=media+rsTabella1("Vv") %>
				  
					  <% rsTabella1.movenext()%>
				 <%next%>
				 <% media=fix(media/(k-1)*10)  
			else
	QuerySQL= "SELECT distinct [2CREDITI].Id_Esercitazione, Descrizione,Data,Scrutini,Crediti  FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo") &"' and Scrutini=1 "&_
		 " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);" 	
		   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
		        media=0 
			    assente=0
			  for k=1 to num_periodi_calcola%> 
              <% if not rsTabella1.eof then%>
                  <% media=media+rsTabella1("Crediti") %>
                  <% rsTabella1.movenext()%>
               <%else%>
                 <%  assente=assente+1%>
               <%end if%>
             <%next%>
             <% 
			    if (num_periodi_calcola-assente>0) then %>
                 <%media=fix((media/(num_periodi_calcola-assente))*10)/10 %>
                     
				<%else
				   media=0%>
				<% end if  
		   end if
			       
             
             
		 	  riga="set label='" &cognome &"' value='" & media & "'/"
			'  riga= "<set label='"  
			  response.write(riga &"<br>")
			  'response.write(cognome & " " & media&"<br>") 
			 
			 objCreatedFile.WriteLine("<"&riga&">")
			 
	    rsTabella.movenext
	loop
	rsTabella.close
	ConnessioneDB.close

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