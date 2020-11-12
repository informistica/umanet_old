 


<%

' conto quanti periodi di valutazione ci sono , mi serve per caricare il vettore delle date
QuerySQL="SELECT count(*) FROM [dbo].[3PERIODI] Where ID_Classe='"& id_classe &"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
numPeriodi=rsTabella(0)+1' +1 per Oggi  
redim periodi(numPeriodi)  
' faccio la query per prelevare i periodi di valutazione per questa classe 
QuerySQL="SELECT * FROM [dbo].[3PERIODI] Where Id_Classe='"& id_classe &"' order by Data;"
		
Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
   ' carico il vettore delle date di valutazione
    periodi(0)=inizio_anno
    idperiodo=1
	selezionato=0
	do while not rsTabella.eof
           periodi(idperiodo)=rsTabella.fields("Data")
		   if  rsTabella.fields("Iniziale")=1 then
		     selezionato=idperiodo ' periodo da cui deve partire la classifica
		     DataIniz=rsTabella.fields("Data")
		   end if
		   
		   idperiodo=idperiodo+1
		   rsTabella.movenext()
    loop 	
	' se il giorno o il mese hanno na sola cifra devo aggiungere lo 0 davanti
	giorno=day(date())
	mese= month(date())
	anno=year(date())
	if len(giorno)=1 then
	   giorno="0" &	day(date())
	end if
	if len(mese)=1 then
	   mese="0" &	month(date())
	end if
	DataOggi=giorno&"/"&mese&"/"&year(date())
	periodi(idperiodo)= DataOggi 
	if not rsTabella.eof then 
	rsTabella.movefirst()
	end if
	
	'url="C:\Inetpub\umanetroot\logPeriodi.txt"
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'objCreatedFile.WriteLine(idperiodo)
				'objCreatedFile.Close	
'response.write(QuerySQL)
'response.write(numPeriodi)	
	
%>


<!-- Form che serve per calcolare la classifica a partire da una data -->	
	<form method="POST" name="dati" id="dati">
    <% for i=0 to numPeriodi 
		  ' response.write(left(DataCla,10)& "---"& left(periodi(i),10) & i & "di" & numPeriodi  &"<br>")
		   
		next %>	
<h4><i class="icon-calendar"></i>&nbsp;&nbsp;Periodo </h4>
<%if (DataCla<>"") then %>
     <b>Dal </b> <select name="txtData" id="txtData" class="input-large">
	   <%
		  for i=0 to numPeriodi

		  'response.write periodi(i) & "  " & DataCla
		  
	   %>
	   
		<% if cDate(periodi(i)) = cDate(DataCla) then 
		selected = "selected='selected'" 
		Session("DataCla") = periodi(i)
Session("DataClaq") = periodi(i)
DataCla = periodi(i)
DataClaq = periodi(i) 
else
selected = ""
		end if %>
	   
			   
			   
						   <% if i = 0 then %>
							<option <%=selected%>  value="<%=periodi(i)%>">Inizio a/s</option>
						   <% else %>
								  <option <%=selected%>  value="<%=periodi(i)%>"><%=periodi(i)%></option> 
						   <%end if%>
			   
			   
					  
			   
			   
	   <% next %>
	   
	</select>
	     
<% else  ' devo far partire la classifica in base alla data iniziale impostat nella tabella
   
   
   
  
%>
     <b>Dal </b> <select name="txtData" id="txtData" class="input-large">
	   <%
		  for i=0 to numPeriodi

	   %>
	   
			   <% if numPeriodi = 1 then %>
			   
						   <% if i = 0 then %>
							<option  value="<%=periodi(i)%>">Inizio a/s</option>
						   <% else %>
								  <option  value="<%=periodi(i)%>"><%=periodi(i)%></option>  
						   <%end if%>
			   
			   <% else %>
			   
						   <% if i = numPeriodi-1 then %>
							<option selected value="<%=periodi(i)%>"><%=periodi(i)%></option> 
<% Session("DataCla") = periodi(i)
Session("DataClaq") = periodi(i) 
DataCla = periodi(i)
DataClaq = periodi(i)
%>							
						   <%else %>
						   
									<% if i = 0 then %>
									<option value="<%=periodi(i)%>">Inizio a/s</option>
								   <%else %>
								   <option  value="<%=periodi(i)%>"><%=periodi(i)%></option>  
								   <% end if %>
						   
						   
						   <%end if%>
			   
			   <% end if %>
			   
	   <% next %>
	   
	</select>
	 
<% end if%>

<%'Session("DataCla")=DataCla
'Session("DataClaq")=DataCla 
%>

<%' inserisco casella data al  

			' Set objFSO = CreateObject("Scripting.FileSystemObject")
				' 'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
					' url="C:\Inetpub\umanetroot\expo2015Server\logPeriodi1.txt"
				' Set objCreatedFile = objFSO.CreateTextFile(url, True)
				' objCreatedFile.WriteLine("DataCla="&DataCla)
					' objCreatedFile.WriteLine("DataCla2="&DataCla2)
						' objCreatedFile.WriteLine("DataClaq="&DataClaq)
							' objCreatedFile.WriteLine("DataClaq2="&DataClaq2)
				' 'objCreatedFile.Close
		
'response.write "fine: "&DataCla2&"<br> btn: "&Request.QueryString("daForm")
'response.write Request.ServerVariables("HTTP_REFERER")&"<br>"&Request.ServerVariables("SCRIPT_NAME") 	

		
	
	if Request.QueryString("daForm") = 1 then
		Session("DataClaOld") = DataCla2
	end if
				
	'Response.write Session("DataClaOld")
	 if (DataCla2<>"") then %>
	  
	  <b>al </b> <select name="txtData2" id="txtData2" class="input-large">
	   <%
	   
		  for i=0 to numPeriodi

		  'response.write periodi(i) & "  " & DataCla2
		
		
		
	   %>
	   
		<% if cDate(periodi(i)) = cDate(Session("DataClaOld")) then 
		selected = "selected='selected'" 
		
		if i = UBound(periodi) then
		DataCla2New = DateAdd("d",1,periodi(i))
		else
		DataCla2New = periodi(i)
		end if
Session("DataCla2") = DataCla2New
DataCla2 = DataCla2New
DataClaq2 = DataCla2New
Session("DataClaq2") = DataCla2New 
Session("DataClaOld") = periodi(i)
else
selected = ""
		end if %>
	   
			 
			   
						   <% if i = 0 then %>
							<option <%=selected%>  value="<%=periodi(i)%>">Inizio a/s</option>
						   <% else %>
								  <option <%=selected%>  value="<%=periodi(i)%>"><%=periodi(i)%></option> 
						   <%end if%>
			   
			   
					  
			   
			   
	   <% next %>
	   
	</select>
	 
	 <%else %>

<b>al </b> <select name="txtData2" id="txtData2" class="input-large">
	   <%
		  for i=0 to numPeriodi

	   %>
	   
			   <% if numPeriodi = 1 then %>
			   
						   <% if i = 0 then %>
							<option  value="<%=periodi(i)%>">Inizio a/s</option>
						   <% else %>
								  <option  value="<%=periodi(i)%>"><%=periodi(i)%></option>  
						   <%end if%>
			   
			   <% else %>
			   
						   <% if i = numPeriodi then %>
							<option selected value="<%=periodi(i)%>"><%=periodi(i)%></option> 
<% 

if i = UBound(periodi) then
		DataCla2New = DateAdd("d",1,periodi(i))
		else
		DataCla2New = periodi(i)
		end if
		
Session("DataCla2") = DataCla2New
DataCla2 = DataCla2New
DataClaq2 = DataCla2New
Session("DataClaq2") = DataCla2New 
Session("DataClaOld") = periodi(i)


%>							
						   <%else %>
						   
									<% if i = 0 then %>
									<option value="<%=periodi(i)%>">Inizio a/s</option>
								   <%else %>
								   <option  value="<%=periodi(i)%>"><%=periodi(i)%></option>  
								   <% end if %>
						   
						   
						   <%end if%>
			   
			   <% end if %>
			   
	   <% next %>
	   
	</select>
	 
<% end if%>	

  <%' Response.write "fine2: "&DataCla2 %>	  
		  <% 'response.write(Request.ServerVariables("HTTP_REFERER")) 
		  %>
 
	<%'Session("DataCla2")=DataCla2 
	%>	  
		  
		  
 
<!-- ------------------- -->	
	