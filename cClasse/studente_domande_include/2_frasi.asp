<%
QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 " &_
" WHERE CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_
 " order by Moduli.In_Umanet,Moduli.Posizione,Paragrafi.Posizione,CodiceFrase asc;"
  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
  i=0 ' serve per decidere quando aggiungere la riga con il modulo

%>
<br>&nbsp
<!-- Effetto per  la tendiana sulla frasi-->
<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND style="width:auto;padding:5px;"><a name="ancora_frasi" href="#" onClick="Effect.toggle('frasi','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">FRASI</span></a>
</legend>
<div id="frasi" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;">
<p>
<table id="zebra_stud" align=center border=1 bordercolor=pink style="table-layout:fixed; width:100%;border:1px solid #f00;word-wrap:break-word;" >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Frasi non ancora inserite!</th></tr>

<% Else%>
	<%' per riepilogare il totale di tutte le frasi
	 QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and Frasi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and Frasi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if

	%>
		<tr><th colspan=4><center><a target="_new" href="../../studente_domande_include/2inserisci_valutazioni_frasi.asp?Tutte=1&amp;ID_MOD=<%=rsTabella.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella.fields("TitPar")%>&amp;id_classe=<%=id_classe%>">Dettaglio di tutte le frasi : Nf(<%= rsTabella1(0) &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a></center> </th></tr>
<!--

-->
		<%
		separatoUmanet=0 ' serve per aggiungere una riga di separazione tra moduli normali e moduli umanet, lo pongo ad 1 dopo la prima hr
do while not rsTabella.EOF
	if (rsTabella.fields(0)= "Esercitazioni varie") then
' non mostro niente
else
		if (i=0) then ' aggiungo la riga con il modulo e con il numero di frasi in quel modulo

		' vedo se devo mettere il separatore per i moduli umanet


     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and Frasi.Id_Stud='"& rsTabella.fields("CodiceAllievo")& "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Frasi ON Moduli.ID_Mod = Frasi.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and Frasi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if
	  divid=divid+1
   %>

   </table>
    </div></div>

   <br>
   <% if (separatoUmanet=0) and (rsTabella("In_Umanet")=1) and i=0 then
   			   separatoUmanet=1
				   %>
       		<!-- <hr class="hr" style="width:25%; background-color:transparent; color:#FC3; " title="Linea di confine tra Informatica ed Informistica"> -->
            <br><img width="200px" height="20px" class="imground_shadow" src="../../img/umanet_separatore.jpg" title="Linea di confine tra Informatica ed Informistica" ><br><br>
         <!--  <p align="center"> <img src="../img/umanet2.png" width="10%" height="10%" ></p>
            <hr class="hr" style="width:25%; background-color:transparent; color:#FC3; " title="Linea di confine tra Informatica ed Informistica"> <br>-->
  		<% end if%>



   <a class="sottotitoloquaderno2" href="#" onClick="Effect.toggle('sottofrasi<%=divid%>','slide'); return false;"><%=rsTabella.fields("Titolo") & " N(" & rsTabella1(0) &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a>
<div id="sottofrasi<%=divid%>" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;">


   <table id="zebra_stud" align=center border=1 width="95%"  align="center">

   <!-- Metto il link alla pagina inserisci_valutazioni per vedere le frasi dello studente per quel modulo -->

   <tr><th colspan="2" align="center"><b><a target="_new" href="../../studente_domande_include/2inserisci_valutazioni_frasi.asp?ID_MOD=<%=rsTabella.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella.fields("TitPar")%>&amp;id_classe=<%=id_classe%>"><%=rsTabella.fields("Titolo") & " Nf(" & rsTabella1(0) &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%></b> </a></th><th>Data</th><th>Ora</th><th>Punti</th><th>Elimina</th><th>Esposto</th></tr>
   <%end if
			    %>
                   <tr><td><%=rsTabella(0)%></td><td><a target="_new" href="../../studente_domande_include/2inserisci_valutazione_frase.asp?Cartella=<%=rsTabella.fields("Cartella")%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>&amp;CodiceFrase=<%=rsTabella.fields("CodiceFrase")%>&amp;Capitolo=<%=rsTabella(9)%>&amp;Paragrafo=<%=rsTabella(0)%>&amp;MO=<%=rsTabella.fields("ID_Mod")%>&amp;VAL=<%=rsTabella.fields("Voto")%>&amp;id_classe=<%=id_classe%>">
                 <%	if rsTabella.fields("Segnalata")=1 then%>

               			 <font color="#FF0000"> <%=rsTabella.fields("Chi")%></font></a></td><td><%=rsTabella.fields("Data")%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella.fields("Voto")%></td>
                <%else%>
										<%	if rsTabella.fields("Segnalata")=2 then%>
									      	<font color="#00FF00"> <%=rsTabella.fields("Chi")%></font></a></td><td><%=rsTabella.fields("Data")%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella.fields("Voto")%></td> 
										  <%else%>
					                 	<%=rsTabella.fields("Chi")%></a></td><td><%=rsTabella.fields("Data")%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella.fields("Voto")%></td>
                     <%end if%>
								 <%end if%>

<td><a onClick="return window.confirm('Vuoi veramente cancellare la frase?');" target="_new" href="../../studente_domande_include/cancella_frase.asp?cla=<%=d%>&amp;cod=<%=rsTabella("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("ID_Mod")%>&amp;CodiceTest=<%=rsTabella.fields("ID_Paragrafo")%>&amp;CodiceFrase=<%=rsTabella.fields("CodiceFrase")%>&amp;Capitolo=<%=rsTabella(9)%>&amp;Paragrafo=<%=rsTabella(0)%>&amp;id_classe=<%=id_classe%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>"><img src="../../img/elimina_small.jpg"></a>
</td><td><input type="checkbox"></td>


		<%
		end if
		i=i+1
Modulo=rsTabella.fields("Titolo") ' serve per vederer quando cambia il titolo del modulo, metto i=0 così che sopra venga aggiunta la riga
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

		<%end if%>
<!--Chiudo il div che contiene le frasi-->
</table>
 </p>
</div></div>
<!--</div></div>-->

</FIELDSET>
