<%	
QuerySQL="SELECT * FROM RISULTATI_ALLIEVI WHERE Allievi.CodiceAllievo='" & cod & "' "&_
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
<a name="ancora_quiz" href="#" onClick="Effect.toggle('quiz','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">QUIZ</a> </span>
<div id="quiz" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 

<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Non ci sono quiz svolti nei paragrafi!</th></tr>
			  
<% Else%>
		<tr><th colspan=6><center>Risultati nei quiz sui Paragrafi di <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%></center> </th></tr>
			<%if (session("Admin")=true) then %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td><b>Data</b><td><b>Ora</b></td><td><b>Risultato</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td><b>Data</b></td><td><b>Ora</b><td><b>Risultato</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (session("Admin")=true) then %>
				<tr><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td><td><%=rsTabella(5)%></td><td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&amp;IdR=<%=rsTabella(7)%>&amp;id_classe=<%=id_classe%>"title="Cancella"><img src="../../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=rsTabella(2)%></td><td><%=rsTabella(3)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td> <td><%=rsTabella(5)%></td>
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>
 



<% ' logica per mostrare i risultati nei quiz dello studente relativi ai singoli MODULI
     
QuerySQL="SELECT * FROM RISULTATI_ALLIEVI1 WHERE CodiceAllievo='" & cod & "'" &_
" and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
" ORDER By Data asc;"
  
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Non ci sono quiz svolti nei moduli!</th></tr>
			  
<% Else%>
		<tr><th colspan=5 align="center"><center>Risultati nei quiz sui Moduli di <%=rsTabella("Cognome")%>       <%=rsTabella("Nome")%></center> </th></tr>
			<%if (session("Admin")=true) then %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Data</b><td><b>Ora</b></td><td><b>Risultato</b></td><td><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Data</b></td><td><b>Ora</b><td><b>Risultato</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
			<%if (session("Admin")=true) then %>
				<tr><td><%=rsTabella(2)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td><td><%=rsTabella(5)%></td><td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href="../../cDomande/cancella_risultato.asp?cod=<%=rsTabella("CodiceAllievo")%>&amp;IdR=<%=rsTabella(7)%>&amp;id_classe=<%=id_classe%>" title="Cancella"><img src="../../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=rsTabella(2)%></td><td><%=rsTabella(4)%></td><td><%=rsTabella(8)%></td> <td><%=rsTabella(5)%></td>
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>

</p> 
</div></div>