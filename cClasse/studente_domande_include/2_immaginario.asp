<%' prelevo l'elenco dei nodi dello studente
 
QuerySQL="SELECT * from TUTTESMILES2 " &_
" WHERE CodiceAllievo='" & cod & "';" 

'response.Write(QuerySQL)


 
 
 Set rsTabella = ConnessioneDB1.Execute(QuerySQL)
 i=0 ' serve per decidere quando aggiungere la riga con il modulo
%>
<br>&nbsp
<!-- Apro il div per l'effetto a tendina sui nodi-->
<fieldset style="margin: 0 auto 0 auto; border:none;"><LEGEND style="width:auto; padding:5px; text-align:center"><a name="ancora_immaginario" href="#" onClick="Effect.toggle('immaginario','slide'); return false;"><span style="font-style:normal" class="sottotitoloquaderno">  IMMAGINARIO</span></a> </legend>
<div id="immaginario" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 

<table id="zebra" align=center border=1 bordercolor=pink >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th align="center">Immaginario non presente !</th></tr>
              <tr><th align="center"><a target="_blank" href="../../studente_domande_include/upload_resize/ex2_imgsocial.asp">Crea immaginario</a></th></tr>
			  
<% Else%>
	<%' per riepilogare tutti i nodi e punti 
	  QuerySQL1="SELECT count(*) from TUTTESMILES2 " &_
" WHERE CodiceAllievo='" & cod & "';" 
	
 
	 
	 Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1)
	   numrsTabella1=rsTabella1(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabella1(0)&"" =""  then
	   numrsTabella1=0
	 end if 
		%>
   
   <!--<table align="center">-->
 
		<tr><th colspan=5>
        <center><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_immaginario.asp?CodiceAllievo=<%=cod%>">Dettaglio di tutte le immagini in tutte le categorie : Ntot(<%=numrsTabella1 &")"%></a>
        <tr><th align="center"><a target="_blank" href="../../studente_domande_include/upload_resize/ex2_imgsocial.asp">Aggiorna immaginario</a></th></tr>
        </center> </th></tr>
		  
 
		<%do while not rsTabella.EOF 
		'if (rsTabella.fields("Tit")= "Esercitazioni varie") then
		if (1=2) then ' per  nn togliere if 
' non mostro niente
else
	 if (i=0) then ' aggiungo la riga con illa categora e il num di immagini per quella cat.
     QuerySQL1="SELECT Count(*) from TUTTESMILES2 " &_
	 " where  ID_Categoria=" & rsTabella.fields("ID_Categoria") & ";"
	 
	 Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1)
	' response.write(Session("cartella"))
	'  QuerySQL2="SELECT SUM(Voto) AS Pt FROM Moduli INNER JOIN Nodi ON Moduli.ID_Mod = Nodi.Id_Mod " &_
'	 " where  Moduli.Titolo='" & rsTabella.fields("Titolo") & "' and Nodi.Id_Stud='"& rsTabella.fields("CodiceAllievo") & "'" &_
'	 " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"
'	 
'	 Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
'	   numrsTabella2=rsTabella2(0)
'	 ' se non restituisce nulla serve per dargli un valore
'	 if rsTabella2(0)&"" =""  then
'	   numrsTabella2=0
'	 end if 
	 divid=divid+1 
   %>
   
   
   </table>
   </div></div>  
   
   <br>
   <a class="sottotitoloquaderno2"  href="#" onClick="Effect.toggle('immaginario<%=divid%>','slide'); return false;"><%=rsTabella.fields("Testo") & " N(" & rsTabella1(0) &")"%></a> 
<div id="immaginario<%=divid%>" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 

<table id="zebra_stud" align=center border=1 width="95%"  align="center">

<tr><th colspan=5>
        <center><a target="_new" href="../../studente_domande_include/1inserisci_valutazioni_immaginario.asp?CodiceAllievo=<%=cod%>">Dettaglio di tutte le immagini di questa categoria : N(<%=numrsTabella1 &")"%></a>
         </th></tr>
   <tr><th>Nome</th><th>Descrizione</th><th>Img</th><th>Codice</th><th>Elimina</th></tr>
   
   <%end if
			    %>
                   <tr><td><a target="_new" href="../../studente_domande_include/2inserisci_valutazione_immaginario.asp?ID_Smile=<%=rsTabella("ID_Smile")%>&amp;CodiceAllievo=<%=rsTabella("CodiceAllievo")%>" title="Visualizza dettagli di questa immagine">
                
                
					<%=rsTabella("Nome")%>
                    </a></td>
                    <td>Descrizione</td>
                    <td>
                    <a target="_new" href="2inserisci_valutazione_immaginario.asp?ID_Smile=<%=rsTabella("ID_Smile")%>&CodiceAllievo=<%=rsTabella("CodiceAllievo")%>" title="Visualizza dettagli di questa immagine"> <img src="../<%=Session("cartella")%>/img_social/thumb/<%=rsTabella("Url")%>"></td><td><%=rsTabella.fields("Codice")%></a></td> 
              
<td><a onClick="return window.confirm('La vuoi veramente cancellare ?');" target="_new" href="../../studente_domande_include/cancella_immaginario.asp?cartella=<%=Session("cartella")%>&amp;ID_Smile=<%=rsTabella("ID_Smile")%>&amp;url=<%=rsTabella("Url")%>"><img src="../../img/elimina_small.jpg"> </a>
</td>
			 
			  
		<% 
		end if
		i=i+1
	Testo=rsTabella.fields("Testo")
	rsTabella.movenext
      if not rsTabella.eof then     
		   Testo2=rsTabella.fields("Testo")
			if StrComp(Testo, Testo2) = 0 then
                  ' Response.Write("Le due stringhe sono uguali") quindi non aggiungo riga
             else 
                    i=0
					 
			end if
		end if
loop
Set rsTabella=nothing
Set ConnessionDB1=nothing
%>
 
		
		
<%end if%> 
<!--Chiudo il div che contiene l'effetto per le frasi -->
</table>
 </p> 
</div></div> 


</fieldset>

