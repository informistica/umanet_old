<% QuerySQL="SELECT CodiceAllievo, Descrizione, Data, [2CREDITI].Crediti, ID_Credito FROM (Allievi INNER JOIN [2CREDITI] ON CodiceAllievo = Id_Stud) INNER JOIN [2ESERCITAZIONI_SINGOLI] ON  [2CREDITI].Id_Esercitazione = [2ESERCITAZIONI_SINGOLI].ID_Esercitazione  WHERE CodiceAllievo='" & cod & "' And Descrizione<> 'Iscrizione' " &_
 " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &CDATE(DataClaq2) &"', 104)) ORDER By Data asc;"

  'response.write(QuerySQL)

'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logAllieviRisultati.txt"
	'			Set objCreatedFile = objFSO.CreateTextFile(url, True)
	'			objCreatedFile.WriteLine(QuerySQL)
	'			objCreatedFile.Close
	
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
 

<!-- Div tendina per i risultatinei quiz -->
 
 
 <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Crediti extra
                                                </h3>
                                            </div> 
                                          <div class="box-content nopadding">
 

<table class="table table-hover table-nomargin"> 

<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			<thead>
			  <tr><th align="center">Non ci sono crediti !</th></tr>
              </thead>
			  
<% Else%>
	<thead>
		<tr><th colspan=5><center>Crediti nelle attivit&agrave; </center> </th></tr>
          </thead>
			<%if (strcomp(request.cookies("Dati")("Admin"),"true")=0) then %>
				<tr><td><b>Attivit&agrave;</b></td><td><b>Data</b></td><td><b>Crediti</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b>Attivit&agrave;</b></td><td><b>Data</b></td><td><b>Crediti</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
		<%if (strcomp(request.cookies("Dati")("Admin"),"true")=0) then %>
				<tr><td><%=rsTabella(1)%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&amp;IdR=<%=rsTabella(4)%>" title="Cancella"><i class=" icon-trash" ></i></a></td>
			<%else %>
		  <tr><td><%=rsTabella(1)%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td>  
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>
 

 