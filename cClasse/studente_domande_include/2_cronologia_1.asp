<%

QuerySQL="select Cognome,Nome from [Allievi] where CodiceAllievo='" &cod &"';"
 Set rsTabellaAllievo = ConnessioneDB.Execute(QuerySQL)
cognome=trim(rsTabellaAllievo("Cognome"))
nome=trim(rsTabellaAllievo("Nome"))

QuerySQL="select * from [4PERIODI_CLASSIFICA] where CodiceAllievo='" &cod &"';"
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
 'response.Write(QuerySQL)

 	classe=request.querystring("classe")
    anno="as_1920"

   ' urlreport=protocollo&dominio&homesito& "/grafici/"&anno&"/report&"& classe &".asp"
	'urlreport="../../../grafici/"&anno&"/report&"& classe &".json"  'json risolve problema caratteri speciali
    'urlreport=Replace(urlreport,"\","/")
 %>
 							<div class="box-content">
								<div class="accordion" id="accordion2">
								<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapseOne">
											 Elenco
											</a>
										</div>
										<div id="collapseOne" class="accordion-body collapse in">
											<div class="accordion-inner">
											 <a href="../cGrafici/terzaVersione/index.php?nome=<%=cognome%><%=nome%>&classe=<%=classe%>&anno=<%=anno_scolastico%>" target="blank">Apri</a>
											   <iframe src="../cGrafici/terzaVersione/index.php?nome=<%=cognome%><%=nome%>&classe=<%=classe%>&anno=<%=anno_scolastico%>" name="grafico" id="grafico" width="100%" height="100%" frameborder="0" SCROLLING="yes" border="0" class="iframe"></iframe>

											</div>
										</div>
									</div>
                                  <% Else%>
                                   <% i=1
		   pos1=0 ' prima posizione in classifica serve per calcolare il trend
			do while not rsTabella.EOF 
			trend= pos1 - rsTabella("Posizione")%>
            
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion<%=i%>" href="#collapse<%=i%>">
Dal <%=rsTabella("Dal")%> &nbsp;al &nbsp; <%=rsTabella("Al")%>&nbsp;; Punti = <%=rsTabella("TOT")%> ; Voto =&nbsp;<%=rsTabella("Vv")%>&nbsp;; Pos =&nbsp; <%=rsTabella("Posizione")%> 
  <% if pos1<>0 then %>
			&nbsp;; Trend =     
			<%		 if trend >0 then %>
						<img src="../../img/pollice_su.jpg" width="18" height="18"> +
					 <%else%>  
                       <% if trend<0 then%>              
						<img src="../../img/pollice_giu.jpg"  width="18" height="20">  -
				      <%end if%>
				   <%end if%>
                  <% response.write(abs(trend))%>
               <%end if%>

  </a>  
										</div>
                                        
										<div id="collapse<%=i%>" class="accordion-body collapse ">
											<div class="accordion-inner">
												<table class="table table-hover table-nomargin"> 
                                                  	
             <tr><th>PD</th><th>PN</th><th>PF</th><th>PM</th><th>PC</th><th>PS</th></tr>
			 <tr><td><%=rsTabella("Pd")%></td><td><%=rsTabella("Pn")%></td><td><%=rsTabella("Pf")%></td><td><%=rsTabella("Pm")%></td>
            <td><%=rsTabella("Pc")%></td><td><%=rsTabella("Ps")%></td></tr>
                                                </table>
											</div>
										</div>
									</div>
            
								
                                 <%  i=i+1
		   pos1=rsTabella("Posizione")
		   rsTabella.movenext
		loop%>
	<%end if%>
    </div>
</div>

 



			  
			  
 
		 
	 
