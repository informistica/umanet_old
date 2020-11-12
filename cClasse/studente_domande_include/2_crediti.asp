<% QuerySQL="SELECT Allievi.CodiceAllievo, [2ESERCITAZIONI_SINGOLI].Descrizione, [2ESERCITAZIONI_SINGOLI].Data, [2CREDITI].Crediti, [2CREDITI].ID_Credito" &_
" FROM (Allievi INNER JOIN 2CREDITI ON Allievi.CodiceAllievo = [2CREDITI].Id_Stud) INNER JOIN 2ESERCITAZIONI_SINGOLI ON "&_ 
" [2CREDITI].Id_Esercitazione = [2ESERCITAZIONI_SINGOLI].ID_Esercitazione "  &_
" WHERE Allievi.CodiceAllievo='" & cod & "' And Descrizione<> 'Iscrizione' " &_
" and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
" ORDER By Data asc;"

 

'url="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logAllieviRisultati.txt"
	'			Set objCreatedFile = objFSO.CreateTextFile(url, True)
	'			objCreatedFile.WriteLine(QuerySQL)
	'			objCreatedFile.Close
	
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp

<!-- Div tendina per i risultatinei quiz -->
<a name="ancora_crediti" href="#" onClick="Effect.toggle('crediti','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">CREDITI</a> </span>
<div id="crediti" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 

<table id="zebra_stud" align=center border=1 bordercolor=pink style="table-layout:fixed; width:100%;border:1px solid #f00;word-wrap:break-word;">
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Non ci sono crediti !</th></tr>
			  
<% Else%>
		<tr><th colspan=5><center>Crediti nelle attività </center> </th></tr>
			<%if (session("Admin")=true) then %>
				<tr><td><b>Attività</b></td><td><b>Data</b></td><td><b>Crediti</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b>Attività</b></td><td><b>Data</b></td><td><b>Crediti</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (session("Admin")=true) then %>
				<tr><td><%=rsTabella(1)%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&amp;IdR=<%=rsTabella(4)%>" title="Cancella"><img src="../../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=rsTabella(1)%></td><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td>  
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>
 



 
</p> 
</div></div>
