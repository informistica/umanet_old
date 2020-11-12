<%


QuerySQL="select * from [4PERIODI_CLASSIFICA] where CodiceAllievo='" &cod &"';"

'url="C:\Inetpub\umanetroot\Anno_2012-2013\logAllieviRisultati.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
	
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)%>
 
<a name="ancora_cronologia" href="#" onClick="Effect.toggle('cronologia','slide'); return false;"> <span style="font-style:normal" class="sottotitoloquaderno">CRONOLOGIA</a> </span>
<div id="cronologia" style="display:none;"><div  class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 

<table id="zebra_stud" align=center border=1 bordercolor=pink style="table-layout:fixed; width:100%;border:1px solid #f00;word-wrap:break-word;" >
<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
			  <tr><th width="100%" align="center">Non ci sono classifiche !</th></tr>
			  
<% Else%>
		<tr><th width="100%" colspan=8><center>Cronologia delle classifiche </center> </th></tr>
			 
				<tr><th>N</th><th><b>Dal</b></th><th><b>Al</b></th><th><b>Voto</b></th><th><b>Posizione</b></th><th><b>Trend</b></th><th><b>Punti</b></th><th>Dettagli</th></tr>
			 
		<% i=1
		   pos1=0 ' prima posizione in classifica serve per calcolare il trend
			do while not rsTabella.EOF 
			trend= pos1 - rsTabella("Posizione")%>
			 
		  <tr><td><%=i%></td><td><%=rsTabella("Dal")%></td><td><%=rsTabella("Al")%></td><td><%=rsTabella("Vv")%></td>  
              <td><%=rsTabella("Posizione")%></td>
              <td>
			  <% if pos1<>0 then 
			     
					 if trend >0 then %>
						<img src="../../../img/pollice_su.jpg" width="18" height="18">
					 <%else%>  
                       <% if trend<0 then%>              
						<img src="../../../img/pollice_giu.jpg"  width="18" height="18"> 
				      <%end if%>
				   <%end if%>
                  <% response.write(abs(trend))%>
               <%end if%>
              </td>
              <td><%=rsTabella("TOT")%></td>
              <td><a href="#" onClick="Effect.toggle('dDet<%=i%>','appear'); return false;">
              <img src="../../../img/Next.gif" width="14" height="13"></a> 
        </td>
        <tr><td colspan="8">
        <div id="dDet<%=i%>" style="display:none;"><div style="width:auto;border:1px solid pink;padding:3px; background-color:#FFC;"> 
          <table style="table-layout:fixed;width:100%;border:1px solid #f00;word-wrap:break-word;"}>
             <tr><th>PD</th><th>PN</th><th>PF</th><th>PM</th><th>PC</th><th>PS</th></tr>
			 <tr><td><%=rsTabella("Pd")%></td><td><%=rsTabella("Pn")%></td><td><%=rsTabella("Pf")%></td><td><%=rsTabella("Pm")%></td>
            <td><%=rsTabella("Pc")%></td><td><%=rsTabella("Ps")%></td></tr>
           
          </table>
		</div></div>
		 
		<%  i=i+1
		   pos1=rsTabella("Posizione")
		   rsTabella.movenext
		loop%>
	<%end if%> 
     <tr><th colspan=8><center><a target="_blank" href="../../cGrafici/genera_grafico_studente.asp?CodiceAllievo=<%=cod%>">Visualizza Grafico</a></center> </th></tr>
</table>
 



 
</p> 
</div></div>
