<%' faccio una query per prelevare l'elenco dei moduli e paragrafi
' per ogni modulo e paragrafo faccio una query per contare le visualizzazioni e per ogni record aggregato metto il link al dettaglio delle visualizzazioni 
	   '                     0               1            2                 3                4                  5                   6                    7            8

QuerySQL="SELECT DISTINCT (Moduli.Titolo) AS Modulo, Paragrafi.Titolo as Paragrafo, Moduli.ID_Mod, Paragrafi.ID_Paragrafo " &_
" FROM Paragrafi INNER JOIN (Moduli INNER JOIN Visualizzazioni ON Moduli.ID_Mod = Visualizzazioni.ID_Mod) ON " &_
" Paragrafi.ID_Paragrafo = Visualizzazioni.ID_Paragrafo " &_
" WHERE Visualizzazioni.CodiceAllievo='" & cod & "' ORDER BY Moduli.ID_Mod, Paragrafi.ID_Paragrafo;"
'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni 
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
%>
<br>&nbsp
<!-- Div tendina per i le visualizzazioni di video -->
<a name="ancora_video" href="#" onClick="Effect.toggle('video','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">VIDEO</span></a> 
<div id="video" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p>


<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Non ci sono visualizzazioni</th></tr>
			  
<% Else%>
	
			 <% 'conto le visualizzazioni totali 
			  QuerySQL1="SELECT Count(*) AS Numero_visualizzazioni "&_
" FROM Allievi INNER JOIN Visualizzazioni ON Allievi.CodiceAllievo = Visualizzazioni.CodiceAllievo" &_
" WHERE Allievi.CodiceAllievo='"&cod &"';"
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
	num_visualizzazioni_totali=rsTabella1(0) %>
	          	<tr><th colspan=3><center>Visualizzazioni procedure in <b>Totale = <%=num_visualizzazioni_totali%></b></center> </th></tr>
				<tr><th><b><center>Modulo</center></b></th><th><b>Paragrafo</b></th><th><b>Totale</b></th></tr>
		 
		<% 'adesso per ogni recordset conto le visualizzazioni e aggiungo link per il dettaglio
		   do while not rsTabella.EOF 
		     QuerySQL1="SELECT Count(*) AS Numero_visualizzazioni "&_
" FROM Allievi INNER JOIN Visualizzazioni ON Allievi.CodiceAllievo = Visualizzazioni.CodiceAllievo" &_
" WHERE (Allievi.CodiceAllievo='"&cod &"' and Visualizzazioni.ID_Mod='" &rsTabella(2) &"' and Visualizzazioni.ID_Paragrafo='" &rsTabella(3) &"')"&_
" GROUP BY Visualizzazioni.ID_Mod, Visualizzazioni.ID_Paragrafo;"
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
	num_visualizzazioni=rsTabella1(0) 
			 %>
								<tr><td><a href="../../cDomande/dettagli_visualizzazioni.asp?tipo=0&amp;ID_Mod=<%=rsTabella(2)%>&amp;ID_Paragrafo=<%=rsTabella(3)%>&amp;cod=<%=cod%>"><%=rsTabella(0)%></a></td><td><a href="../../cDomande/dettagli_visualizzazioni.asp?tipo=1&amp;ID_Mod=<%=rsTabella(2)%>&amp;ID_Paragrafo=<%=rsTabella(3)%>&amp;cod=<%=cod%>"><%=rsTabella(1)%></a></td><td><%=num_visualizzazioni%></td> 
			 
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>

</p> 
</div></div>