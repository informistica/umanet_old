
    <%
	id = Request.QueryString("id")
	titolo=Request.QueryString("titolo")
    titolo=Replace(titolo,"'",chr(96))
	testo=Request.QueryString("testo")
    testo=Replace(testo,"'",chr(96))
    
	on error resume next
		  QuerySQL="SELECT * FROM [push_subscriptions] WHERE (CodiceAllievo='" & session("CodiceAllievo") & "');"
		  'RESPONSE.WRITE(querysql)
		  set rsTabella=ConnessioneDB.Execute(QuerySQL)
		  If Err.Number = 0 Then
				stato=1
				messaggio="Modifica avvenuta"
			Else
				stato=0
				messaggio=Err.Description
				Err.Number = 0
			End If
            %>
    <table class="table table-hover table-nomargin table-bordered  dataTable-scroll-x table-striped">
	<thead><tr><th>CodiceAllievo</th><th>end_point</th><th>expirationTime</th><th>Attiva</th><th>Cancella</th></tr></thead>
        <% i=1
        do while not rsTabella.eof%>
            <tr id="riga_<%=i%>"> 
                <td><%=rsTabella("CodiceAllievo")%></td>
                
                <td><a href="#modal-<%=i%>" role="button" class="btn" data-toggle="modal">Dettagli</a>
				<div id="modal-<%=i%>" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
					<div class="modal-header">
						<button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
						<h3 id="myModalLabel">End point</h3>
					</div>
					<div class="modal-body">
						<p><%=rsTabella("end_point")%></p>
					</div>
					<div class="modal-footer">
						<button class="btn" data-dismiss="modal" aria-hidden="true">Close</button>	 
					</div>
				</div>
				</td>
				<td><%=rsTabella("expirationTime")%></td>
				<td>
				
			    <% 
				if (strcomp(rsTabella("Attiva"),"1")=0)  then  %>

					<INPUT onClick="modifica_iscrizione(<%=rsTabella("id")%>,'<%=rsTabella("CodiceAllievo")%>','1');" TYPE="RADIO" name="txtAttiva<%=i%>" checked="true" value="1">Si
                    <INPUT onClick="modifica_iscrizione(<%=rsTabella("id")%>,'<%=rsTabella("CodiceAllievo")%>','0');" TYPE="RADIO" name="txtAttiva<%=i%>"  value="0">No
                <% else %>
                    <INPUT onClick="modifica_iscrizione(<%=rsTabella("id")%>,'<%=rsTabella("CodiceAllievo")%>','1');" TYPE="RADIO" name="txtAttiva<%=i%>" value="1">Si
                    <INPUT onClick="modifica_iscrizione(<%=rsTabella("id")%>,'<%=rsTabella("CodiceAllievo")%>','0');" TYPE="RADIO" name="txtAttiva<%=i%>"  checked="true" value="0">No
				<% end if %>
				</td>
				<td class='hidden-480'><a onClick="cancella_iscrizione(<%=rsTabella("id")%>,'<%=rsTabella("CodiceAllievo")%>',<%=i%>);">
				<i class=" icon-trash" ></i></a>
				</td>
            </tr>
        <% 
			i=i+1
       	    rsTabella.movenext
        loop
%>
</table>
		
 
<script>

function cancella_iscrizione(id,codiceAllievo,riga) {
	if (window.confirm('Cancellare iscrizione?')) {
	  	 	var url="../cUtenti/gestisci_iscrizione_push_ajax.asp?api=1&id="+id+"&codiceallievo="+codiceAllievo;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								if (risposta=="Cancellazione avvenuta!")
									// ricarica la pagina per aggiornare la lista delle iscrizioni
									location.reload();
								else
									alert(risposta);
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();
	 }
 }

function modifica_iscrizione(id,codiceAllievo,stato) {
	  	 	var url="../cUtenti/gestisci_iscrizione_push_ajax.asp?api=2&id="+id+"&codiceallievo="+codiceAllievo+"&stato="+stato;
				 var xhttp = new XMLHttpRequest();
			   xhttp.onreadystatechange = function() {
			   	if (xhttp.readyState == 4 && xhttp.status == 200) {
						    var risposta=xhttp.responseText;
								alert(risposta);
					}
			   };
			   xhttp.open("GET", url, true);
			   xhttp.send();

}
</script>