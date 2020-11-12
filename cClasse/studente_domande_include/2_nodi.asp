<% '                      
QuerySQL="SELECT Allievi.Cognome, Allievi.Nome, Nodi.Id_Arg, Nodi.Chi, Nodi.Cosa, Nodi.Dove, Nodi.Quando, Nodi.Come, Nodi.Perche, Nodi.Quindi, Moduli.Titolo, Paragrafi.Titolo as [Tit], Nodi.CodiceNodo," &_ 
" Moduli.ID_Mod, Nodi.Voto, Nodi.Data, Allievi.CodiceAllievo, Nodi.URL_Teoria, Nodi.Cartella,Nodi.Ora,Nodi.Segnalata " &_
" FROM Allievi INNER JOIN (Paragrafi INNER JOIN (Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod) ON Paragrafi.ID_Paragrafo = Nodi.Id_Arg) ON Allievi.CodiceAllievo = Nodi.Id_Stud " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &" order by CodiceNodo asc;"

'response.Write(QuerySQL)


 'url="C:\Inetpub\umanetroot\anno_2012-2013\logClass.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
' 
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 i=0 ' serve per decidere quando aggiungere la riga con il modulo
%>
<br>&nbsp
<!-- Apro il div per l'effetto a tendina sui nodi-->
<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND style="width:auto; padding:5px;"><a name="ancora_nodi" href="#" onClick="Effect.toggle('nodi','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">NODI</span></a> </legend>
<div id="nodi" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 

<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Nodi della rete non ancora inseriti!</th></tr>
			  
<% Else%>
	<%' per riepilogare tutti i nodi e punti 
	  QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and Nodi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.ID_Mod<>'6C' and Nodi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 
	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
	   numrsTabella2=rsTabella2(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella2(0)&"" =""  then
	   numrsTabella2=0
	 end if 
	
	 
	 
		%>
   
   <!--<table align="center">-->
   
   
   
		<tr><th colspan=5>
        <center><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&amp;DATA=<%=rsTabella.fields("Data")%>&amp;Tutte=1&amp;ID_MOD=<%=rsTabella.fields("ID_Mod")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella.fields("Tit")%>">Dettaglio di tutti i nodi : Nn(<%=rsTabella1(0) &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a>
        </center> </th></tr>
		  
		
<!--		
			<%'if (session("Admin")=true) then %>
				<tr><td><b><center>Paragrafo</center></b></td><td><b>Nodo</b></td><td><b>Data</b></td><td><b>Ora</b></td><td><b>Punti</b></td><td><b>Cancella</b></td></tr>
			<%'else %>
				<tr><td><b><center>Paragrafo</center></b></td><td><b>Nodo</b></td><td><b>Data</b></td><td><b>Ora</b></td><td><b>Punti</b></td></tr>
			<%'end if %>
-->

<%do while not rsTabella.EOF 
		if (rsTabella.fields("Tit")= "Esercitazioni varie") then
' non mostro niente
else
	 if (i=0) then ' aggiungo la riga con il modulo e con il numero di nodi in quel modulo e la somma dei punti
     QuerySQL1="SELECT Count(*) AS Num FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and Nodi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
	 
	 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
	  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and Nodi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
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
    </div> </div>  
   
   <br>
   <a class="sottotitoloquaderno2" href="#" onClick="Effect.toggle('sottonodi<%=divid%>','slide'); return false;"><%=rsTabella.fields("Titolo") & " N(" & rsTabella1(0) &") Pt(" & numrsTabella2  & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a> 
<div id="sottonodi<%=divid%>" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 

<table id="zebra_stud" align=center border=1 width="95%"  >


   <tr><th>Paragrafo</th><th colspan="" align="center"><b><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&amp;DATA=<%=rsTabella.fields("Data")%>&amp;ID_MOD=<%=rsTabella.fields("ID_MOD")%>&amp;CodiceAllievo=<%=rsTabella.fields("CodiceAllievo")%>&amp;Cartella=<%=rsTabella.fields("Cartella")%>&amp;Modulo=<%=rsTabella.fields("ID_Mod")%>&amp;Capitolo=<%=rsTabella.fields("Titolo")%>&amp;TitoloParagrafo=<%=rsTabella.fields("Tit")%>">
   <%=rsTabella.fields("Titolo") & " Nn(" & rsTabella1(0) &") Pt(" & numrsTabella2 & ") Pb("& round( numrsTabella2/rsTabella1(0),2) &")"%> </a></th><th>Data</th><th>Ora</th><th>Punti</th><th>Elimina</th></tr>
   <%end if
			    %>
                   <tr><td><%=rsTabella(11)%></td><td><a target="_new" href="../../studente_domande_include/inserisci_valutazione_nodi.asp?Segnalata=<%=rsTabella.fields("Segnalata")%>&amp;Cognome=<%=rsTabella("Cognome")%>&amp;Nome=<%=rsTabella("Nome")%>&amp;id_classe=<%=id_classe%>&amp;DATA=<%=rsTabella.fields("Data")%>&amp;Cartella=<%=rsTabella(18)%>&amp;classe=<%=classe%>&amp;cod=<%=cod%>&amp;CodiceTest=<%=rsTabella(2)%>&amp;CodiceDomanda=<%=rsTabella(12)%>&amp;Capitolo=<%=rsTabella(10)%>&amp;Paragrafo=<%=rsTabella(11)%>&amp;Chi=<%=rsTabella(3)%>&amp;Cosa=<%=rsTabella(4)%> &amp;Dove=<%=rsTabella(5)%>&amp;Quando=<%=rsTabella(6)%>&amp;Come=<%=rsTabella(7)%>&amp;Perche=<%=rsTabella(8)%>&amp;Quindi=<%=rsTabella(9)%>&amp;MO=<%=rsTabella(13)%>&amp;VAL=<%=rsTabella(14)%>&amp;URL=<%=rsTabella(17)%> ">
                
                <%	if rsTabella.fields("Segnalata")=1 then%>
             <font color="#FF0000"><%=rsTabella(3)%></font>
             </a></td><td><%=rsTabella(15)%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella(14)%></td> 
				<%else%>
					<%=rsTabella(3)%>
                    </a></td><td><%=rsTabella(15)%></td><td><%=rsTabella.fields("Ora")%></td><td><%=rsTabella(14)%></td> 
              <%end if%>
<td><a onClick="return window.confirm('Vuoi veramente cancellare il nodo?');" target="_new" href="../../studente_domande_include/cancella_nodo.asp?cla=<%=d%>&amp;cod=<%=rsTabella("CodiceAllievo")%>&amp;Cartella=<%=rsTabella(18)%>&amp;Modulo=<%=rsTabella(13)%>&amp;CodiceTest=<%=rsTabella(2)%>&amp;CodiceDomanda=<%=rsTabella(12)%>&amp;Capitolo=<%=rsTabella(10)%>&amp;Paragrafo=<%=rsTabella(11)%>&amp;id_classe=<%=id_classe%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>"><img src="../../img/elimina_small.jpg"> </a>
</td>
			 
			  
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
		
		
<%end if%> 
<!--Chiudo il div che contiene l'effetto per i nodi -->
    
</table>

 </p> 
</div></div> 


</fieldset>