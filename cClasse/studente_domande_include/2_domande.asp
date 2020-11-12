<% QuerySQL="SELECT * FROM 2_PRELEVA_DOMANDE" &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#;"  

 'QuerySQLDomande=QuerySQL
 Set rsTabella3 = ConnessioneDB.Execute(QuerySQL)
 %>
 
 
 
 <!--Apro il div che contiene le domande per l'effetto -->
<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND style="width:auto; padding:5px;"><a name="ancora_domande" href="#" onClick="Effect.toggle('domande','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">    DOMANDE</span></a> </legend>






<div id="domande" style="display:none;">
<!--<div style="background-color:#ffffff;width:auto;border:1px solid red;padding:10px;"> -->
<div>
<p> 

<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabella3.BOF=True And rsTabella3.EOF=True Then %>
			  <tr><th align="center"> Domande non ancora inserite!</th></tr>
<% else%>
<!-- Metterò il link alla pagina inserisci_valutazioni per vedere tutte le domande dello studente-->

<% ' per riepilogare il totale delle domande e dei punti dello studente
 QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where  Domande.ID_Mod<>'6C'  and Domande.Id_Stud='"& rsTabella3.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	 
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where Domande.ID_Mod<>'6C'  and  Domande.Id_Stud='"& rsTabella3.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	  numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
	 
%>


<tr><th colspan=4><center><a target="_new" href="../../studente_domande_include/inserisci_valutazioni.asp?id_classe=<%=id_classe%>&amp;Tutte=1&amp;ID_MOD=<%=rsTabella3.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella3.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella3.fields("Cartella")%>&amp;Modulo=<%=rsTabella3.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella3.fields("Titolo")%>&amp;id_classe=<%=id_classe%>"> Dettaglio di tutte le domande : N   (<%=rsTabella1(0) &") Pt(" & numrsTabella2 & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a></center> </th></tr>


<%

i=0 ' AGGIUNTO PER PROVA

do while not rsTabella3.EOF 
 

 
if (rsTabella3.fields("Tit")= "Esercitazioni varie") then
' non mostro niente
else
 if (i=0) then ' aggiungo la riga con il modulo e con il numero di domande in quel modulo e il totale dei punti
     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella3.fields("Titolo") & "' and Domande.Id_Stud='"& rsTabella3.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	 
	 QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Domande ON Moduli.ID_Mod = Domande.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella3.fields("Titolo") & "' and Domande.Id_Stud='"& rsTabella3.fields("CodiceAllievo") & "'" &_
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
   <a class="sottotitoloquaderno2" href="#" onClick="Effect.toggle('sottodomande<%=divid%>','slide'); return false;"><%=rsTabella3.fields("Titolo") & " N(" & rsTabella1("Num") &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1("Num"),2) &")"%> </a>
<div id="sottodomande<%=divid%>" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 

 
   
 <table id="zebra_stud" align=center border=1 width="95%"  >
   
   
   <!-- Metto il link alla pagina inserisci_valutazioni per vedere le domande dello studente per quel modulo -->
   <tr><th>Paragrafo</th><th colspan="1" align="center"><b><a target="_new" href="../../studente_domande_include/inserisci_valutazioni.asp?ID_MOD=<%=rsTabella3.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella3.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella3.fields("Cartella")%>&amp;Modulo=<%=rsTabella3.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella3.fields("Titolo")%>&amp;id_classe=<%=id_classe%>"><%=rsTabella3.fields("Titolo") & " N(" & rsTabella1("Num") &") Pt(" & numrsTabella2 & ") Pb("& round( numrsTabella2/rsTabella1("Num"),2) &")"%> </a></b></th><th>Data</th><th>Punti</th><th>Elimina</th><th>Sposta</th></tr>
   <%end if%>
 
 	<tr><td><%=rsTabella3(11)%></td><td><a target="_new" href="../../studente_domande_include/inserisci_valutazione.asp?VF=<%=rsTabella3.fields("VF")%>&amp;Multiple=<%=rsTabella3.fields("Multiple")%>&amp;Segnalata=<%=rsTabella3.fields("Segnalata")%>&amp;DataClaq=<%=DataCla%>&amp;DataClaq2=<%=DataCla2%>&amp;DATA=<%=rsTabella3.fields("Data")%>&amp;Tipodomanda=<%=rsTabella3(20)%>&amp;Cartella=<%=rsTabella3(19)%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabella3(14)%>&amp;CodiceDomanda=<%=rsTabella3(12)%>&amp;Capitolo=<%=rsTabella3(10)%>&amp;Paragrafo=<%=rsTabella3(11)%>&amp;Quesito=<%=rsTabella3(3)%>&amp;R1=<%=rsTabella3(5)%> &amp;R2=<%=rsTabella3(6)%>&amp;R3=<%=rsTabella3(7)%>&amp;R4=<%=rsTabella3(8)%>&amp;RE=<%=rsTabella3(9)%>&amp;MO=<%=rsTabella3(13)%>&amp;VAL=<%=rsTabella3(15)%>&amp;URL=<%=rsTabella3(18)%>&amp;INQUIZ=<%=rsTabella3("In_Quiz")%>&amp;VALINQUIZ=<%=rsTabella3(23)%>&amp;id_classe=<%=id_classe%>">

	<%if rsTabella3.fields("Segnalata")=1 then ' se la domanda è segnalata la scrivo in rosso%>
	
<font color="#FF0000"><%=rsTabella3(3)%></font></a></td><td><%=rsTabella3(4)%></td><td><%=rsTabella3(17)%></td>
	<%else%>
	<%=rsTabella3(3)%></a></td><td><%=rsTabella3(4)%></td><td><%=rsTabella3(17)%></td>
	<%end if%> 	
	
	
		<td><a onClick="return window.confirm('Vuoi veramente cancellare la domanda?');" target="_new" href="../../studente_domande_include/cancella_domanda.asp?Verifica=0&amp;classe=<%=classe%>&amp;cod=<%=rsTabella3("CodiceAllievo")%>&amp;Cartella=<%=rsTabella3(19)%>&amp;Modulo=<%=rsTabella3(13)%>&amp;CodiceTest=<%=rsTabella3(14)%>&amp;CodiceDomanda=<%=rsTabella3(12)%>&amp;Capitolo=<%=rsTabella3(10)%>&amp;Paragrafo=<%=rsTabella3(11)%>&amp;id_classe=<%=id_classe%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>" title="Cancella"><img src="../../img/elimina_small.jpg"></a></td>
        
        <td><a onClick="return window.confirm('Vuoi veramente spostare la domanda?');" href="../../studente_domande_include/8_sposta_domanda.asp?Verifica=0&amp;classe=<%=classe%>&amp;cod=<%=rsTabella3("CodiceAllievo")%>&amp;Cartella=<%=rsTabella3(19)%>&amp;Modulo=<%=rsTabella3(13)%>&amp;CodiceTest=<%=rsTabella3(14)%>&amp;CodiceDomanda=<%=rsTabella3(12)%>&amp;Capitolo=<%=rsTabella3(10)%>&amp;Paragrafo=<%=rsTabella3(11)%>&amp;id_classe=<%=id_classe%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>" title="Cancella">></a></td>
 

 
   
<%
end if ' fine if per saltare Esercitazioni varie 



i=i+1
Modulo=rsTabella3.fields("Titolo") ' serve per vederer quando cambia il titolo del modulo, metto i=0 così che sopra venga aggiunta la riga
rsTabella3.movenext
      if not rsTabella3.eof then     
		   Modu=rsTabella3.fields("Titolo")
	 
			if StrComp(Modulo, Modu) = 0 then
                 '  Response.Write("Le due stringhe sono uguali") 'quindi non aggiungo riga
             else 
                    i=0
					 
			end if
		end if
loop%>
<%end if%> 
<!--Chiudo il div che contiene le domande per l'effetto -->
</table>
 </p> 
 
</div></div> 









</FIELDSET>
