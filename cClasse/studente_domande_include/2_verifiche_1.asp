<%	

QuerySQL="select * from [2ESERCITAZIONI_SINGOLI] where Id_Classe='"&id_classe&"' and TipoVoto='V'"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)

'response.write(QuerySQL)

'QuerySQL="SELECT * FROM RISULTATI_ALLIEVI WHERE CodiceAllievo='" & cod & "' "&_
'" and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
'	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq"Descrizione")) &"', 104))"&_  
'" ORDER By Data asc;"
 
'url="C:\inetpub\umanetroot\expo"Descrizione"015Server\logAllieviRisultati.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
 'Set rsTabella = ConnessioneDB.Execute(QuerySQL)
' response.write(QuerySQL)
%>
 
<!-- Div tendina per i risultatinei quiz -->
 <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Risultati nelle verifiche
                                                </h3>
                                            </div> 
                                          <div class="box-content nopadding">
 <table class="table table-hover table-nomargin"> 
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			 <thead>
              <tr><th align="center"><center>Nessuna verifica svolta nei Paragrafi</center></th></tr>
			  </thead>
<% Else%>
         <thead>
		<tr><th colspan=7><center>Verifiche svolti sui Paragrafi </center> </th></tr>
		 </thead>	
         <tbody>
			<%if (session("admin")=true) then %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td class='hidden-480'><b>Data</b></td><td ><b>Risultato</b></td><td class='hidden-480'><b>Cancella</b></td></tr>
			<%else %>
				<tr><td><b><center>Modulo</center></b></td><td><b>Paragrafo</b></td><td class='hidden-480'><b>Data</b></td><td><b>Risultato</b></td></tr>
			<%end if %>
		<%do while not rsTabella.EOF %>
		<%
		CodiceTest=rsTabella("CodiceTest")
		QuerySQL="select Id_Modulo from Classi_Moduli_Paragrafi where Id_Paragrafo='"&CodiceTest&"'"
		Set rsMod = ConnessioneDB.Execute(QuerySQL)
		Id_Mod=rsMod(0)
		QuerySQL="select Titolo from Moduli where ID_Mod='"&Id_Mod&"'"
		Set rsMod = ConnessioneDB.Execute(QuerySQL)
		Titolo=rsMod(0)

		urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
	'ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&".xml"
	ulrRisorsa1=rsTabella("Descrizione")&"_correzione_"&cod&".xml"
	ulrRisorsa=urlRis&ulrRisorsa1
	ulrRisorsa=Replace(ulrRisorsa,"\","/")	
	shortUrl="Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
	shortUrl=shortUrl&ulrRisorsa1
	shortUrl=Replace(shortUrl,"\","/")	

	'response.write("<br>"&ulrRisorsa&"<br>")
    Set fso = CreateObject("Scripting.FileSystemObject") 
	If not(fso.FileExists(ulrRisorsa)) Then
	  voto="Non svolta"
	else
		Set objXMLDoc = Server.CreateObject("Microsoft.XMLDOM") ' per il file modello
		objXMLDoc.async = False 
		objXMLDoc.load ulrRisorsa
		Set Root = objXMLDoc.documentElement
		Set NodeList = Root.getElementsByTagName("Sentiment")
		Set Risposta = objXMLDoc.getElementsByTagName("Sentiment")(0)
		voto=Risposta.text
		'voto=NodeList.length
	end if

		%>
			<%if (session("admin")=true) then%>
				<tr><td><%=Titolo%></td><td><%=rsTabella("Descrizione")%></td><td><%=rsTabella("Data")%></td> <td><a target="_blank" href="../cFrasi/3visualizza_risultati_verifiche.asp?Titolo=<%=Titolo%>&cartella=<%=cartella%>&cod=<%=cod%>&paragrafo=<%=rsTabella("Descrizione")%>&shorturl=<%=shortUrl%>"><%=voto%>%</a></td>
  				<td><a onClick="return window.confirm('Vuoi veramente cancellare il risultato ?');" target="_new" href=""><img src="../../img/elimina_small.jpg"></a></td>
			<%else %>
		  <tr><td><%=Titolo%></td><td><%=rsTabella("Descrizione")%></td><td><%=rsTabella("Data")%></td><td><a target="_blank" href="../cFrasi/3visualizza_risultati_verifiche.asp?cartella=<%=cartella%>&cod=<%=cod%>&paragrafo=<%=rsTabella("Descrizione")%>&shorturl=<%=shortUrl%>"><%=voto%>%</a></td>
               
			<%end if %>
		<%rsTabella.movenext
		loop%>
	<%end if%> 
</table>
 



 

 