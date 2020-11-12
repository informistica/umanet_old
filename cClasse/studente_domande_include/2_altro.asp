<%
if (session("Admin")=true) and (id_classe <> "") then %>

<fieldset style="margin: 0 auto 0 auto;">
<a style="font-size:12px;"href="#" onClick="Effect.toggle('admin','slide'); return false;">Altro</a>
</B>
<div id="admin" style="display:none;">
    <div class="style1" style="background-color:#ffffff;border:none;padding:10px;"> 



<!-- Per fare in modo che il quiz svolto in una certa data compaia nell'elenco delle attività PROBABILMENTE NON E? ATTIVO VOSTO CHE NON CE IL BOTTONE MA SOLO HREF-->
<form method="POST" form action="../../studente_domande_include/aggiorna_punteggio.asp?classe=<%=classe%>&amp;xQuiz=1&amp;id_classe=<%=id_classe%>&amp;DataCla=<%=DataCla%>&amp;DataCla2=<%=DataCla2%>" > 	
<FIELDSET style="margin-left:16px;" ><LEGEND class="sottotitoloquaderno2"><B> Convalida QUIZ</B></LEGEND>
 <p> 
<!--<select name="txtVerifica">-->
<%' adesso seleziono i quiz svolti dalla classe in modo da configurare il campo option select per la scelta del quiz da convalidar

 
QuerySQL1="SELECT DISTINCT Allievi.Id_Classe, Risultati1.CodiceTest, Risultati1.Data, Moduli.Titolo " &_
" FROM (Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo) INNER JOIN Moduli " &_ 
"ON Risultati1.CodiceTest = Moduli.ID_Mod WHERE Allievi.Id_Classe='"&id_classe &"'"

'
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logQuiz.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close
'	 
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
    i=0
	do while not rsTabella1.EOF %>	
		  <a href="../../studente_domande_include/aggiorna_punteggio_pulisci_test.asp?tipoTest=1&amp;classe=<%=classe%>&amp;xQuiz=1&amp;id_classe=<%=id_classe%>&amp;DataCla=<%=DataCla%>&amp;DataCla2=<%=DataCla2%>&amp;CodiceTest=<%=rsTabella1.fields("CodiceTest")%>&amp;DataTest=<%=rsTabella1.fields("Data")%>&amp;TitoloTest=<%=rsTabella1.fields("Titolo")%>"> <%=rsTabella1("Titolo") &" - "& rsTabella1("Data")%></a><br>
          
	    <%rsTabella1.movenext 
	i=i+1
	loop%>
 
<!--</select>
<input type="submit" value="Aggiorna" name="B1"><input type="reset" value="Azzera" name="B2"></p> <!--Definisce i due bottoni del form -->
 
 </p>

 </FIELDSET>
</form>

<FIELDSET style="margin-left:16px;" ><LEGEND class="sottotitoloquaderno2"><B><a name="cronologia"> Salva cronologia</a></B></LEGEND>
 <p> 

 <a target="_new" href="../../studente_domande_include/genera_grafico.asp?PS=<%=PS%>&amp;id_classe=<%=id_classe%>&amp;DataCla=<%=DataCla%>&amp;DataCla2=<%=DataCla2%>&amp;indice_periodo=<%=indice_periodo%>&amp;indice_periodo2=<%=indice_periodo2%>">Genera </a><br></p>
 
</FIELDSET>
<br>
<!-- Per estrarre uno studente a caso per l'interrogazione-->

<form method="POST" action="../../studente_domande_include/studente_domande.asp?divid=<%=divid%>&amp;classe=<%=classe%>&amp;DataCla=<%=DataCla%>&amp;xEstrazione=1&amp;id_classe=<%=id_classe%>"> 	
<FIELDSET style="margin-left:16px;" ><LEGEND class="sottotitoloquaderno2"><B><a name="sorteggia"> Sorteggia Studente</a></B></LEGEND>
 <p> 

<%' devo generare un numero casuale che servirà per accedere alla tabella studenti
 if xEstrazione<>"" then
			  randomize()
			  NumeroCasuale = Int(NumStud * Rnd + 1)
end if

%>
 &nbsp;&nbsp;
 <input type="text" name="txtSTUD" value="<%=NumeroCasuale&")"&vetstud(NumeroCasuale)%>" size="50">
 <p><input type="submit" value="Estrai" name="B1"></p> <!--Definisce i due bottoni del form -->
 </p></FIELDSET>
</form>

<%end if%></i>
</p> 

</FIELDSET>
<br>



</div></div>
