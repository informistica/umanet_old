<%
' prelevo l'elenco delle metafore Topolino dello studente

'seleziono la metafora topolino
QuerySQL="SELECT Allievi.Cognome, Allievi.Nome, M_Topolino.Id_Arg, M_Topolino.Topolino, M_Topolino.Formaggio, M_Topolino.Fame, M_Topolino.Labirinto, M_Topolino.Strada, M_Topolino.Strada_OK, M_Topolino.Strada_KO,  Moduli.Titolo, Paragrafi.Titolo as [Tit], M_Topolino.CodiceMetafora," &_ 
" Moduli.ID_Mod, M_Topolino.Voto, M_Topolino.Data, Allievi.CodiceAllievo, M_Topolino.URL_Teoria, M_Topolino.Cartella,M_Topolino.Ora,M_Topolino.Testata,Paragrafi.ID_Paragrafo,M_Topolino.Distanza,M_Topolino.Segnalata " &_
" FROM Allievi INNER JOIN (Paragrafi INNER JOIN (Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod) ON Paragrafi.ID_Paragrafo = M_Topolino.Id_Arg) ON Allievi.CodiceAllievo = M_Topolino.Id_Stud " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &" order by CodiceMetafora asc;"
 
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 	
 ' seleziono le metafore navigazione
 QuerySQL="SELECT * " &_
" FROM Elenco_Metafore_Navigazione " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &" order by CodiceMetafora asc;"
	
 Set rsTabella0 = ConnessioneDB.Execute(QuerySQL)
 'rsTabella0.movefirst

QuerySQL="SELECT * " &_
" FROM Elenco_Metafore_Desideri " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &" order by CodiceMetafora asc;"
	'response.write("<br>"&QuerySQL)
 Set rsTabellaD = ConnessioneDB.Execute(QuerySQL)
 
    
 	 		
 ' seleziono le metafore ......
 
 i=0 ' serve per decidere quando aggiungere la riga con il modulo
%>
<br>&nbsp
<!-- Apro il div per l'effetto a tendina sulle metafore 
iN REALTà NON C'è PIù PERCHè NON POSSO VISUALIZZARE TUTTE LE METAFORE CON STRUTTURA DIVERSA CON LA STESSA PAGINA-->
<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND style="width:auto;padding:5px;"><span style="font-style:normal" class="sottotitoloquaderno"><a name="ancora_metafore">METAFORE</a></span>  </legend>
<div id="metafore" style="display:none;">
	<div style="background-color:#ffffff;width:auto;border:1px solid red;padding:10px;"> 
<p> 

<table id="zebra_stud" align=center border=1 width="95%"  align=center border=1 bordercolor=pink>
<% 
   If (rsTabella.BOF = True And rsTabella.EOF = True)   Then %>
			<br> <table id="zebra_stud" align=center border=1 width="95%"  align="center"> <tr><th>Metafore non ancora inserite! Inserisci per primo una metafora topolino</th></tr></table>		
<% Else%>
 
	<%' per riepilogare tutte le metafore topolino, navigazione,..., e  relativi punti da indicare nel titolo Interfaccia UWWW N(..)
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Topolino.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 'conta il numero di metafore topolino 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Topolino.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 'conta il numero di punti ottenuti nelle metafore topolino
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	  numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
	  'per riepilogare tutti le metafore navigazione e relativi punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Navigazione.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 'conta il numero di metafore navigazione
	 Set rsTabella3 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Navigazione.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	  'conta il numero di punti ottenuti nelle metafore navigazione
	
	 Set rsTabella4 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabella4=rsTabella4(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella4(0)&"" =""  then
	   numrsTabella4=0
	 end if  
	 
	   'per riepilogare tutti le metafore desideri e relativi punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Desideri.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 'conta il numero di metafore navigazione
	 response.write(QuerySQL1&"<br>")
	 Set rsTabella5 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and M_Desideri.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	  'conta il numero di punti ottenuti nelle metafore navigazione
	 response.write(QuerySQL2&"<br>")
	 Set rsTabella6 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabella6=rsTabella6(0)
	 ' se non restituisce nulla serve per dargli un valore
	if rsTabella6(0)&"" =""  then
	   numrsTabella6=0
	 end if  
 
	 
		%>
  

 </table>
   </div>
  </div><!--Questo chiude div Id="Metafore"-->  

<!-- QUA INIZIA LA LOGICA PER LA METAFORA TOPOLINO -->

		<%do while not rsTabella.EOF 
		if (rsTabella.fields("Tit")= "Esercitazioni varie") then
' non mostro niente
else
		if (i=0) then ' aggiungo la riga con il modulo e con il numero di nodi in quel modulo e la somma dei punti
     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and M_Topolino.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Topolino ON Moduli.ID_Mod = M_Topolino.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and M_Topolino.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	 divid=divid+1 
	 numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
   %>
  
   <br>
   <!--Questo è Titolo del MODULO :es Interfaccia UWWW N() Pt() clickando sul quale apre il div che mostra il menu metafore diviso in tabelle -->
   <a class="sottotitoloquaderno2"  href="#" onClick="Effect.toggle('sottometafore<%=divid%>','slide'); return false;"><%=rsTabella.fields("Titolo") & " N(" & rsTabella1(0)+rsTabella3(0)+rsTabella5(0) &") Pt(" & numrsTabella2+numrsTabella4+numrsTabella6 & ")"%></a> 
<div id="sottometafore<%=divid%>" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 

<table id="zebra_stud" align=center border=1 width="95%"  align="center">

<!--Questo è Titolo del paragrafo :es Topolino ed Obiettivi Nn() Pt() clickando sul quale si visualizano tutte le metafore del topolino-->
   <tr><th colspan="2" align="center"><b><a  target="_new" href="../../studente_domande_include/1inserisci_valutazioni_metafore.asp?id_classe=<%=id_classe%>&amp;Topolino=1&amp;DATA=<%=rsTabella.fields("Data")%>&amp;ID_MOD=<%=rsTabella.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella.fields("Tit")%>&amp;Paragrafo=<%=rsTabella.fields("Tit")%>&amp;CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>">
   <%=rsTabella.fields("Tit") & " Nn(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ")"%></b></a></th><th>Data</th><th>Ora</th><th>Punti</th><th>Elimina</th></tr> 
   <%end if      ' questo è l'elenco delle singole metafore, se sono admin visualizzo la x per cancellare altrimenti no.
			   %>
									<tr><td><%=rsTabella.fields("Tit")%></td><td><a target="_new" href="../../studente_domande_include/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&amp;DATA=<%=rsTabella.fields("Data")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>&amp;CodiceMetafora=<%=rsTabella.fields("CodiceMetafora")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Capitolo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella0.fields("Tit")%>&amp;Paragrafo=<%=rsTabella.fields("Tit")%>&amp;Topolino=<%=rsTabella(3)%>&amp;Formaggio=<%=rsTabella(4)%> &amp;Fame=<%=rsTabella(5)%>&amp;Labirinto=<%=rsTabella(6)%>&amp;Strada=<%=rsTabella(7)%>&amp;Strada_OK=<%=rsTabella(8)%>&amp;Strada_KO=<%=rsTabella(9)%>&amp;Distanza=<%=rsTabella.fields("Distanza")%>&amp;Testata=<%=rsTabella.fields("Testata")%>&amp;MO=<%=rsTabella.fields("ID_Mod")%>&amp;VAL=<%=rsTabella.fields("Voto")%>&amp;URL=<%=rsTabella.fields("URL_Teoria")%>&amp;Segnalata=<%=rsTabella.fields("Segnalata")%>&amp;Pippo=1 ">
									
									<%	if rsTabella.fields("Segnalata")=1 then%>
                         <font color="#FF0000">
						<%=rsTabella.fields("Topolino")%></a></td><td><%=rsTabella.fields("Data")%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella.fields("Voto")%></td> 
                        </font>
						 <%	else %>
                         
                            <%=rsTabella.fields("Topolino")%></a></td><td><%=rsTabella.fields("Data")%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella.fields("Voto")%></td>  
					 <% end if %>	 
						
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');" target="_new" href="../../studente_domande_include/cancella_metafora.asp?cla=<%=d%>&amp;cod=<%=rsTabella("CodiceAllievo")%>&amp;Cartella=<%=rsTabella(18)%>&amp;Modulo=<%=rsTabella(13)%>&amp;CodiceTest=<%=rsTabella(2)%>&amp;CodiceMetafora=<%=rsTabella(12)%>&amp;Capitolo=<%=rsTabella(10)%>&amp;Paragrafo=<%=rsTabella(11)%>&amp;id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>
</tr>			 
			 
		<% 
		end if
		i=i+1
	Modulo=rsTabella.fields("Titolo")
	rsTabella.movenext
      if not rsTabella.eof then     
		   Modu=rsTabella.fields("Titolo")
			if StrComp(Modulo, Modu) = 0 then
                  ' Response.Write("Le due stringhe sono uguali") quindi non aggiungo riga
             else 
                    i=0
					 
			end if
		end if
loop%>
</tr>
</table>		

<!-- INIZIO METAFORA NAVIGAZIONE******************************-->

<%
If (rsTabella0.BOF=True And rsTabella0.EOF=True) Then %>
			<br><table id="zebra_stud" align=center border=1 width="95%"  align="center"> <tr><th>Metafore  "Navigazione" non ancora inserite!</th></tr></table>
			  
<% Else 
	rsTabella0.movefirst
	i=0
	do while not rsTabella0.EOF 

 if (rsTabella0.fields("Tit")= "Esercitazioni varie") then
' non mostro niente
		else
		if (i=0) then ' aggiungo la riga con il modulo e con il numero di nodi in quel modulo e la somma dei punti
     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella0.fields("Titolo") & "' and M_Navigazione.Id_Stud='"& rsTabella0.fields("CodiceAllievo")& "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Navigazione ON Moduli.ID_Mod = M_Navigazione.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella0.fields("Titolo") & "' and M_Navigazione.Id_Stud='"& rsTabella0.fields("CodiceAllievo")& "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
	' divid=divid+1 
   %>
  
   <br>
   <!--Questo è Titolo del MODULO :es Interfaccia UWWW N() Pt() clickando sul quale apre il div che mostra il menu metafore diviso in tabelle  lo commento perchè lo crea la prima metafora 
   <a href="#" onClick="Effect.toggle('sottometafore<%=divid%>','slide'); return false;"><%=rsTabella0.fields("Titolo") & " N(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ")"%></a> 
<div id="sottometafore<%=divid%>" style="display:none;"><div style="background-color:#ffFFFF;width:auto;border:1px solid red;padding:10px;"> -->

<table id="zebra_stud" align=center border=1 width="95%"  align="center">

<!--Questo è Titolo del paragrafo :es Navigazione Nn() Pt() clickando sul quale si visualizano tutte le metafore della navigazione-->
   <tr><th colspan="2" align="center"><b><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_metafore.asp?id_classe=<%=id_classe%>&amp;Navigazione=1&amp;DATA=<%=rsTabella0.fields("Data")%>&amp;ID_MOD=<%=rsTabella0.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella0.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella0.fields("Cartella")%>&amp;Modulo=<%=rsTabella0.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella0.fields("Tit")%>&amp;Paragrafo=<%=rsTabella0.fields("Tit")%>&amp;CodiceTest=<%=rsTabella0.fields("ID_Paragrafo")%>">
   <%=rsTabella0.fields("Tit") & " Nn(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ")"%></b></a></th><th>Data</th><th>Ora</th><th>Punti</th><th>Elimina</th></tr> 
   <%end if%>  
   
       <%' questo è l'elenco delle singole metafore, se sono admin visualizzo la x per cancellare altrimenti no.
			    %>
										<tr><td><%=rsTabella0.fields("Tit")%></td>
										
										<td><a target="_new" href="../../studente_domande_include/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&amp;DATA=<%=rsTabella0.fields("Data")%>&amp;Cartella=<%=rsTabella0.fields("Cartella")%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabella0.fields("ID_Paragrafo")%>&amp;CodiceMetafora=<%=rsTabella0.fields("CodiceMetafora")%>&amp;CodiceAllievo=<%=rsTabella0.fields("CodiceAllievo")%>&amp;Capitolo=<%=rsTabella0.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella0.fields("Tit")%>&amp;Paragrafo=<%=rsTabella0.fields("Tit")%>&amp;Autista=<%=rsTabella0(4)%>&amp;Destinazione=<%=rsTabella0(5)%> &amp;Carburante=<%=rsTabella0(6)%>&amp;Luogo=<%=rsTabella0(7)%>&amp;Strada=<%=rsTabella0(8)%>&amp;Strada_OK=<%=rsTabella0(9)%>&amp;Strada_KO=<%=rsTabella0(10)%>&amp;Cespugli=<%=rsTabella0.fields("Cespugli")%>&amp;Cestino=<%=rsTabella0.fields("Cestino")%>&amp;Lupo=<%=rsTabella0.fields("Lupo")%>&amp;Distanza=<%=rsTabella0.fields("Distanza")%>&amp;MO=<%=rsTabella0.fields("ID_Mod")%>&amp;VAL=<%=rsTabella0.fields("Voto")%>&amp;URL=<%=rsTabella0.fields("URL_Teoria")%>&amp;Segnalata=<%=rsTabella0.fields("Segnalata")%>&amp;Pippo=1 ">
	<%	if rsTabella0.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabella0.fields("Autista")%></a></td><td><%=rsTabella0.fields("Data")%></td><td><%=rsTabella0.fields("Ora")%></td><td><%=rsTabella0.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabella0.fields("Autista")%></a></td><td><%=rsTabella0.fields("Data")%></td><td><%=rsTabella0.fields("Ora")%></td><td><%=rsTabella0.fields("Voto")%></td>
    <%end if%>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');" target="_new" href="../../studente_domande_include/cancella_metafora.asp?cla=<%=d%>&amp;cod=<%=rsTabella0.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella0.fields("Cartella")%>&amp;Modulo=<%=rsTabella0.fields("ID_Mod")%>&amp;CodiceTest=<%=rsTabella0.fields("ID_Paragrafo")%>&amp;CodiceMetafora=<%=rsTabella0.fields("CodiceMetafora")%>&amp;Capitolo=<%=rsTabella0.fields("Titolo")%>&amp;Paragrafo=<%=rsTabella0.fields("Tit")%>&amp;id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
	
		<% end if
		i=i+1
		
	Modulo=rsTabella0.fields("Titolo")
	rsTabella0.movenext
	
      if not rsTabella0.eof then     
		   Modu=rsTabella0.fields("Titolo")
			if StrComp(Modulo, Modu) = 0 then
                  ' Response.Write("Le due stringhe sono uguali") quindi non aggiungo riga
             else 
                    i=0
					 
			end if
		end if
   loop
end if ' end if di if rsTabella0.bof ....
%>
</table>	




<!-- INIZIO METAFORA DESIDERI******************************-->

<%
If (rsTabellaD.BOF=True And rsTabellaD.EOF=True) Then %>
			<br><table id="zebra_stud" align=center border=1 width="95%"  align="center"> <tr><th>Metafore  "Desideri" non ancora inserite!</th></tr></table>
			  
<% Else 

%>

	<%
    rsTabellaD.movefirst
	i=0
	
	do while not rsTabellaD.EOF 

 if (rsTabellaD.fields("Tit")= "Esercitazioni varie") then
' non mostro niente
		
		'response.write(rsTabellaD.fields("MotivazioneC"))
		else
		'response.write("ciao2")
		if (i=0) then ' aggiungo la riga con il modulo e con il numero di nodi in quel modulo e la somma dei punti
     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabellaD.fields("Titolo") & "' and M_Desideri.Id_Stud='"& rsTabellaD.fields("CodiceAllievo")& "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN M_Desideri ON Moduli.ID_Mod = M_Desideri.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabellaD.fields("Titolo") & "' and M_Desideri.Id_Stud='"& rsTabellaD.fields("CodiceAllievo")& "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	 numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
	' divid=divid+1 
   %>
  
   <br>
   <!--Questo è Titolo del MODULO :es Interfaccia UWWW N() Pt() clickando sul quale apre il div che mostra il menu metafore diviso in tabelle  lo commento perchè lo crea la prima metafora 
   <a href="#" onClick="Effect.toggle('sottometafore<%=divid%>','slide'); return false;"><%=rsTabellaD.fields("Titolo") & " N(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ")"%></a> 
<div id="sottometafore<%=divid%>" style="display:none;"><div style="background-color:#ffFFFF;width:auto;border:1px solid red;padding:10px;"> -->

<table id="zebra_stud" align=center border=1 width="95%"  align="center">

<!--Questo è Titolo del paragrafo :es Navigazione Nn() Pt() clickando sul quale si visualizano tutte le metafore della navigazione-->
   <tr><th colspan="2" align="center"><b><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_metafore.asp?id_classe=<%=id_classe%>&amp;Desideri=1&amp;DATA=<%=rsTabellaD.fields("Data")%>&amp;ID_MOD=<%=rsTabellaD.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabellaD.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabellaD.fields("Cartella")%>&amp;Modulo=<%=rsTabellaD.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabellaD.fields("Tit")%>&amp;Paragrafo=<%=rsTabellaD.fields("Tit")%>&amp;CodiceTest=<%=rsTabellaD.fields("ID_Paragrafo")%>">
   <%=rsTabellaD.fields("Tit") & " Nn(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ")"%></b></a></th><th>Data</th><th>Ora</th><th>Punti</th><th>Elimina</th></tr> 
   <%end if%>  
   
       <%' questo è l'elenco delle singole metafore, se sono admin visualizzo la x per cancellare altrimenti no.
			    %>
										<tr><td><%=rsTabellaD.fields("Tit")%></td>
										
										<td><a target="_new" href="../../studente_domande_include/inserisci_valutazione_metafore.asp?Cartella=<%=rsTabellaD.fields("Cartella")%>&amp;id_classe=<%=id_classe%>&amp;DATA=<%=rsTabellaD.fields("Data")%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabellaD.fields("ID_Paragrafo")%>&amp;CodiceMetafora=<%=rsTabellaD.fields("CodiceMetafora")%>&amp;CodiceAllievo=<%=rsTabellaD.fields("CodiceAllievo")%>&amp;Capitolo=<%=rsTabellaD.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabellaD.fields("Tit")%>&amp;Paragrafo=<%=rsTabellaD.fields("Tit")%>&amp;SoggettoC=<%=rsTabellaD("SoggettoC")%>&amp;DomandaC=<%=rsTabellaD("DomandaC")%>&amp;MotivazioneC=<%=rsTabellaD("MotivazioneC")%>&amp;DesiderioC=<%=rsTabellaD("DesiderioC")%>&amp;BisognoC=<%=rsTabellaD("BisognoC")%>&amp;SoggettoS=<%=rsTabellaD("SoggettoS")%>&amp;RispostaS=<%=rsTabellaD("RispostaS")%>&amp;MotivazioneS=<%=rsTabellaD("MotivazioneS")%>&amp;DesiderioS=<%=rsTabellaD.fields("DesiderioS")%>&amp;BisognoS=<%=rsTabellaD("BisognoS")%>&amp;TipoEvento=<%=rsTabellaD.fields("TipoEvento")%>&amp;TolleranzaC=<%=rsTabellaD.fields("TolleranzaC")%>&amp;MO=<%=rsTabellaD.fields("ID_Mod")%>&amp;VAL=<%=rsTabellaD.fields("Voto")%>&amp;URL=<%=rsTabellaD.fields("URL_Teoria")%>&amp;Segnalata=<%=rsTabellaD.fields("Segnalata")%>&amp;Pippo=1 ">
	<%	if rsTabellaD.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabellaD.fields("SoggettoC")%></a></td><td><%=rsTabellaD.fields("Data")%></td><td><%=rsTabellaD.fields("Ora")%></td><td><%=rsTabellaD.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabellaD.fields("SoggettoC")%></a></td><td><%=rsTabellaD.fields("Data")%></td><td><%=rsTabellaD.fields("Ora")%></td><td><%=rsTabellaD.fields("Voto")%></td>
    <%end if%>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../../studente_domande_include/cancella_metafora.asp?cla=<%=d%>&amp;cod=<%=rsTabellaD.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabellaD.fields("Cartella")%>&amp;Modulo=<%=rsTabellaD.fields("ID_Mod")%>&amp;CodiceTest=<%=rsTabellaD.fields("ID_Paragrafo")%>&amp;CodiceMetafora=<%=rsTabellaD.fields("CodiceMetafora")%>&amp;Capitolo=<%=rsTabellaD.fields("Titolo")%>&amp;Paragrafo=<%=rsTabellaD.fields("Tit")%>&amp;id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
	
		<% end if
		i=i+1
		
	Modulo=rsTabellaD.fields("Titolo")
	rsTabellaD.movenext
	
      if not rsTabellaD.eof then     
		   Modu=rsTabellaD.fields("Titolo")
			if StrComp(Modulo, Modu) = 0 then
                  ' Response.Write("Le due stringhe sono uguali") quindi non aggiungo riga
             else 
                    i=0
					 
			end if
		end if
   loop
end if ' end if di if rsTabellaD.bof ....
%>
</table>	




	      
  <%


end if%>       
        

<!--Chiudo il div che contiene l'effetto per le metefaore -->


 </p> 
</div></div> 


</fieldset>
