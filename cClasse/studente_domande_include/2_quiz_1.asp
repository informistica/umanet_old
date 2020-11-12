<%	
QuerySQL="SELECT * FROM RISULTATI_ALLIEVI WHERE CodiceAllievo='" & cod & "' "&_
" and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_  
" ORDER By Data asc;"
 
'url="C:\inetpub\umanetroot\expo2015Server\logAllieviRisultati.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
' response.write(QuerySQL)
%>
 
<!-- Div tendina per i risultatinei quiz -->
 <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Risultati nei Quiz
                                                </h3>
                                            </div> 
                                          <div class="box-content nopadding">
                                          
 <table class="table table-hover table-nomargin"> 
 
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			 <thead>
              <tr><th align="center"><center>Nessun quiz svolto nei Paragrafi</center></th></tr>
			  </thead>
<% Else%>
         <thead>
		<tr><th colspan=7><center>Quiz svolti sui Paragrafi </center> </th></tr>
		 </thead>	
         <tbody>
			<%if (strcomp(request.cookies("Dati")("Admin"),"true")=0) then %>
				<tr><td><b><center>Libro</center></b></td><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td class='hidden-480'><b>Data</b><td class='hidden-480'><b>Ora</b></td><td ><b>Risultato</b></td><td><b>Tipo</b><td><b>N.</b></td><td class='hidden-480'><b>Sessione</b></td><td class='hidden-480'><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Libro</center></b></td><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td class='hidden-480'><b>Data</b></td><td class='hidden-480'><b>Ora</b><td><b>Risultato</b></td><td><b>Tipo</b><td><td><b>N.</b></td><td class='hidden-480'><b>Sessione</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (strcomp(request.cookies("Dati")("Admin"),"true")=0) then %>
				<tr><td><%=rsTabella("Classe")%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td class='hidden-480'><%=rsTabella(4)%></td><td class='hidden-480'><%=rsTabella(8)%></td> <td><%=rsTabella(5)%></td>
                <td class='hidden-480'>
                <% select case(rsTabella.fields("Tipo")) 
				case 0:
				  response.write("Vero/Falso")
				  case 1: 
				   response.write("Singola")
				   case 2: 
				    response.write("Multipla")
				end select
				%>
                </td>
                
                <td class='hidden-480'><%=rsTabella("In_Quiz")%></td> <td class='hidden-480'><%=rsTabella("Sessione")%></td><td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&IdR=<%=rsTabella(7)%>&id_classe=<%=id_classe%>"title="Cancella"><img src="../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=rsTabella("Classe")%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td class='hidden-480'><%=rsTabella(4)%></td><td class='hidden-480'><%=rsTabella(8)%></td><td><%=rsTabella(5)%></td>
                <td class='hidden-480'>
                <% select case(rsTabella.fields("Tipo")) 
				case 0:
				  response.write("Vero/Falso")
				  case 1: 
				   response.write("Singola")
				   case 2: 
				    response.write("Multipla")
				end select
				%>
                </td>
                
                <td class='hidden-480'><%=rsTabella("In_Quiz")%></td><td class='hidden-480'><%=rsTabella("Sessione")%></td>
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>
 



<% ' logica per mostrare i risultati nei quiz dello studente relativi ai singoli MODULI
     
QuerySQL="SELECT * FROM RISULTATI_ALLIEVI1 WHERE CodiceAllievo='" & cod & "'" &_
" and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_  
" ORDER By Data asc;"
  
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
 <table class="table table-hover table-nomargin"> 
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <thead>
              <tr><th colspan="7" align="center"><center>Nessun quiz svolto nei Moduli</center></th></tr>
			  </thead>
<% Else%>
		<tr><th colspan="8" align="center"><center>Risultati nei quiz sui Moduli </center> </th></tr>
			<%if (strcomp(request.cookies("Dati")("Admin"),"true")=0) then %>
				<tr><td><b><center>Libro</center></b></td><td><b><center>Modulo</center></b></td><td><b>Data</b><td class='hidden-480'><b>Ora</b></td><td><b>Risultato</b></td><td><b>Tipo</b></td><td><b>N.</b></td><td class='hidden-480'><b>Sessione</b></td><td class='hidden-480'><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Libro</center></b></td><td><b><center>Modulo</center></b></td><td><b>Data</b></td><td class='hidden-480'><b>Ora</b><td><b>Risultato</b></td><td><b>Tipo</b></td><td><b>N.</b></td><td class='hidden-480'><b>Sessione</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (request.cookies("Dati")("Admin")=true) then %>
				<tr><td><%=rsTabella("Classe")%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(4)%></td><td class='hidden-480'><%=rsTabella(8)%></td><td><%=rsTabella(5)%></td>
                 <td class='hidden-480'>
                <% select case(rsTabella.fields("Tipo")) 
				case 0:
				  response.write("Vero/Falso")
				  case 1: 
				   response.write("Singola")
				   case 2: 
				    response.write("Multipla")
				end select
				%>
                </td>
                <td class='hidden-480'><%=rsTabella("In_Quiz")%></td><td class='hidden-480'><%=rsTabella("Sessione")%></td><td class='hidden-480'><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&amp;IdR=<%=rsTabella(7)%>&amp;id_classe=<%=id_classe%>" title="Cancella"><img src="../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=rsTabella("Classe")%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(4)%></td><td class='hidden-480'><%=rsTabella(8)%></td> <td><%=rsTabella(5)%></td>
           <td class='hidden-480'>
                <% select case(rsTabella.fields("Tipo")) 
				case 0:
				  response.write("Vero/Falso")
				  case 1: 
				   response.write("Singola")
				   case 2: 
				    response.write("Multipla")
				end select
				%>
                </td>
          <td class='hidden-480'><%=rsTabella("In_Quiz")%></td><td class='hidden-480'><%=rsTabella("Sessione")%></td>
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>

 