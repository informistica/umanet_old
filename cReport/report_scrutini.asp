<html>

<head>
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Scegli </title>
<link rel="stylesheet" type="text/css" href="../../stile.css">
</head>
<!--
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
<body>
<div class="citazioni" ><div> <span style="font-style: normal">
<%
Dim Id_Classe
Dim ConnessioneDB,rsTabella,rsTabella1,QuerySQL
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")

%>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<% 

    Id_Classe=Request.QueryString("Id_Classe")
	num_periodi=Request.QueryString("num_periodi")
	classifiche=Request.QueryString("classifiche") ' vale 1 se mostro le classifiche
	verifiche=Request.QueryString("verifiche")
	Data=Request.form("txtData")
	Data2=Request.form("txtData2")
	
	QuerySQL="SELECT Cognome,Nome,CodiceAllievo,Classe FROM Allievi Where Id_Classe='"&Id_Classe &"' order by Cognome, Nome  "
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	' serve per le classifiche
	 QuerySQL="select count(*) from 4PERIODI_CLASSIFICA where CodiceAllievo='" &rsTabella("CodiceAllievo") &"'"&_
			  " AND  (Dal >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Al <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
 			   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
   
    num_periodi_calcola=rsTabella1(0)
   'num_periodi_calcola=2
	dim periodi()
	redim periodi(num_periodi_calcola)
	QuerySQL="select Data from [dbo].[3PERIODI]  where Id_Classe='" & Id_Classe &"'" &_	
	 " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
 	Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
	for i=0 to num_periodi_calcola-1 
	   periodi(i)=rsTabella1(0)
	   rsTabella1.movenext()
	next   
	
	
	QuerySQL= "SELECT count(*) FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo") &"' and Scrutini=1 "&_
		  " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);"
		 'response.write(QuerySQL)
 			   Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)
    num_voti=rsTabella2(0)
   'num_periodi_calcola=2
	dim valutazioni()
	redim valutazioni(num_voti)
	 
	 QuerySQL="select Data from 2ESERCITAZIONI_SINGOLI  where Id_Classe='" & Id_Classe &"' and Scrutini=1" &_	
	 " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
			' response.write(QuerySQL)
 	Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)
	
	for i=0 to num_voti-1 
	   valutazioni(i)=rsTabella2(0)
	   rsTabella2.movenext()
	next   
	
	
if (classifiche<>"") and (verifiche<>"") then%>


  <center><br><br>
   <table border=1  id="zebra_stud">
    <form name="dati" method="post" action="genera_grafico_report_scrutini_globale.asp?Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>&Ordina=1">
	 <tr><td align="center" colspan="<%=num_voti+num_periodi_calcola+7%>"><font color="#FF0000"><b>REPORT GENERALE</b></font></td></tr>
      <tr><td rowspan="2" colspan="2"><b>Cognome &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Nome</b></td>
    <b><td align="center" colspan="<%=num_periodi_calcola+2%>"><b>Classifiche</b></td>
       
       
      
      <td align="center"  colspan="<%=num_voti+2%>"><b>Verifiche</b></td> <td rowspan="2" align="center" ><b>MEDIA</b> </td> 
     </tr>
     
     <tr>
        
      <% for k=1 to num_periodi_calcola%> 
           <td align="center"><b>Voto <%=k%><br><font size="-2"><%=periodi(k-1)%></font></b></td>
          <%next%>
      <td align="center"><b>Media</b></td> <td align="center"><b>Voto</b></td>
      <% for k=1 to num_voti%> 
           <td align="center"><b>Voto <%=k%><br><font size="-2"><%=valutazioni(k-1)%></font></b></td>
          <%next%>
      <td align="center"><b>Media</b></td> <td align="center"><b>Voto</b></td> 
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


	    <tr><td><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td>
             <% media=0 
			    assente=0
			  for k=1 to num_periodi_calcola%> 
              
               <% if not rsTabella1.eof then%>
			    
                  <td align="center"> <%=rsTabella1("Vv")%></td>
                  <% media=media+rsTabella1("Vv") %>
                  <% rsTabella1.movenext()%>
               <%else%>
                 <%  assente=assente+1%>
                  <td align="center"><b>A</b></td>
               <%end if%>
             <%next%>
             
             <%if (num_periodi_calcola-assente>0) then %>
                  <% media=media/(num_periodi_calcola-assente) %>
                 <td align="center"><%=fix(media*10)/10%></td>  
				<%else
				   media=0%>
				    <td align="center"><b>NC</b></td>
				<% end if %> 
             <td>&nbsp;</td>
             
             
               
               
               
                <% media2=0 
			    assente=0
			  for k=1 to num_voti%> 
              <% if not rsTabella2.eof then%>		    
                  <td align="center"> <%=rsTabella2("Crediti")%></td>
                  <% media2=media2+rsTabella2("Crediti") %>
                  <% rsTabella2.movenext()%>
               <%else%>
                 <%  assente=assente+1%>
                  <td align="center"><b>A</b></td>
               <%end if%>
             <%next%>
             <% 
			    if (num_voti-assente>0) then %>
                  <% media2=media2/(num_voti-assente) %>
                 <td align="center"><%=fix(media2*10)/10%></td>  
				<%else
				   media=0%>
				    <td align="center"><b>NC</b></td>
				<% end if  
				  
			%>
             
                <td>&nbsp;</td>
                <td align="center"><%= fix(10*(media+media2)/2)/10%></td>
               <%
			   mediaT=((fix(media*10)/10) + (fix(media2*10)/10))/2
			  ' response.write(fix(mediaT*10))
			   %>
         </tr>
	   <input type="hidden" name="Studente<%=numRec%>" value="<%=rsTabella.fields("Cognome")%>">
       <input type="hidden" name="MediaT<%=numRec%>" value="<%=fix(mediaT) %>">
       <input type="hidden" name="Media<%=numRec%>" value="<%=media %>">
       <input type="hidden" name="Media2<%=numRec%>" value="<%=media2 %>">
	<% numRec=numRec+1
	   rsTabella.movenext%>
	  
       
	  <%
	loop
	rsTabella.close
	ConnessioneDB.close
%>  </table></center>
<br>
<a href="../cGrafici/genera_grafico_report_scrutini_globale.asp?classifiche=1&verifiche=1&numRec=<%=numRec%>&Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>">Visualizza grafico della MEDIA</a><br></p>
<input type="submit" value="In ordine di voto">  
</form>
<%













  else	
	   if classifiche<>"" then
	
	

	
	
 
	Response.Write("<font color=red size=2em><b>" &rsTabella("Classe")&" <br>Report voti classifica <br><b> dal " &Data & " al " & Data2 &"<br></b> </font></b>")
	%>
	<br> <p align="center">
    
	<table border=1  id="zebra_stud">
    <form name="dati" method="post" action="genera_grafico_report_scrutini.asp?Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>&Ordina=1">
	  <tr><td><b>Cognome</b></td><td><b>Nome</b></td>
      <% for k=1 to num_periodi_calcola%> 
           <td align="center"><b>Voto <%=k%><br><font size="-2"><%=periodi(k-1)%></font></b></td>
          <%next%>
      <td><b>Media</b></td> <td><b>Voto</b></td>
     </tr>
	<%' per ogni studente vado a vedere i voti memorizzati nei vari periodi, come nella cronologia del quaderfno
    numRec=1
    do while not rsTabella.eof %>
			<% QuerySQL="select * from 4PERIODI_CLASSIFICA where CodiceAllievo='" &rsTabella("CodiceAllievo") &"'"&_
			  " AND  (Dal >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
			  " AND Al <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#" &_
			  ");"
 			   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)%>

	    <tr><td><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td>
             <% media=0 
			  for k=1 to num_periodi_calcola%> 
                  <td> <%=rsTabella1("Vv")%></td>
                  <% media=media+rsTabella1("Vv") %>
                  <% rsTabella1.movenext()%>
             <%next%>
             <% media=media/(num_periodi_calcola) %>
              <td><%=fix(media*10)/10%></td>
               <td><input type="text" name="voto" size="1"></td>
               
             
              
         </tr>
	   <input type="hidden" name="Studente<%=numRec%>" value="<%=rsTabella.fields("Cognome")%>">
       <input type="hidden" name="Media<%=numRec%>" value="<%=fix(media*10)/10%>">
	<% numRec=numRec+1
	   rsTabella.movenext%>
	  
       
	  <%
	loop
	rsTabella.close
	ConnessioneDB.close
%>  </table>
<br>
<a href="../cGrafici/genera_grafico_report_scrutini.asp?classifiche=1&numRec=<%=numRec%>&Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>">Visualizza grafico della media classifica</a><br></p>
<input type="submit" value="In ordine di voto">  
</form>

<%else  'if classifiche<>"" vuol dire che è  verifiche<>"" %>
   
 <%
 
 
	
	 
	 
	Response.Write("<font color=red size=2em><b>" &rsTabella("Classe")&" <br>Report voti verifiche <br><b> dal " &Data & " al " & Data2 &"<br></b> </font></b>")
	%>
	<br> <p align="center">
    
	<table border=1  id="zebra_stud">
    <form name="dati" method="post" action="genera_grafico_report_scrutini.asp?Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>&Ordina=1">
	  <tr><td><b>Cognome</b></td><td><b>Nome</b></td>
      <% for k=1 to num_voti%> 
           <td align="center"><b>Voto <%=k%><br><font size="-2"><%=valutazioni(k-1)%></font></b></td>
          <%next%>
      <td><b>Media</b></td> <td><b>Voto</b></td>
     </tr>
	<%' per ogni studente vado a vedere i voti memorizzati nei vari periodi, come nella cronologia del quaderfno
    numRec=1
    do while not rsTabella.eof %>
			<% 
			
	QuerySQL= "SELECT distinct [2CREDITI].Id_Esercitazione, Descrizione,Data,Scrutini,Crediti  FROM 2ESERCITAZIONI_SINGOLI INNER JOIN 2CREDITI ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione' and Id_Stud='" &rsTabella("CodiceAllievo") &"' and Scrutini=1 "&_
		 " AND  (Data >=#" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"# " &_
		 " AND Data <=#" & mid(Data2,4,2)&"/" &left(Data2,2)&"/"& right(Data2,4)  &"#);" 		  
 		 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL)%>

	    <tr><td><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td>
             <% media=0 
			    assente=0
			  for k=1 to num_voti%> 
              <% if not rsTabella2.eof then%>
			    
                  <td> <%=rsTabella2("Crediti")%></td>
                  <% media=media+rsTabella2("Crediti") %>
                  <% rsTabella2.movenext()%>
               <%else%>
                 <%  assente=assente+1%>
                  <td><b>A</b></td>
               <%end if%>
             <%next%>
             <% 
			    if (num_voti-assente>0) then %>
                  <% media=media/(num_voti-assente) %>
                 <td><%=fix(media*10)/10%></td>  
				<%else
				   media=0%>
				    <td><b>NC</b></td>
				<% end if  
				  
			%>
             
               <td><input type="text" name="voto" size="1"></td>
         </tr>
	   <input type="hidden" name="Studente<%=numRec%>" value="<%=rsTabella.fields("Cognome")%>">
       <input type="hidden" name="Media<%=numRec%>" value="<%=fix(media*10)/10%>">
	<% numRec=numRec+1
	   rsTabella.movenext%>
	  
       
	  <%
	loop
	rsTabella.close
	ConnessioneDB.close
%>  </table>
<br>
<a href="../cGrafici/genera_grafico_report_scrutini.asp?verifiche=1&numRec=<%=numRec%>&Data=<%=Data%>&Data2=<%=Data2%>&ID_Classe=<%=ID_Classe%>">Visualizza grafico della media verifiche</a><br></p>
<input type="submit" value="In ordine di voto">  
</form>


 
<% end if
end if '' if classifiche<>"" and verifiche <>"" %>
</div>
</body>
</html>
