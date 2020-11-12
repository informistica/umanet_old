<div class="box box-color box-bordered <%=Session("stile")%>">
								<div class="box-title">
									<h3>
										<i class="icon-table"></i>
										Contatti
									</h3>
								</div>
								<div class="box-content nopadding">
									 <table class="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped"> 
                                   
                                   
										<thead>
											<tr>
                                                <th title="CognomeNome"><b>Cognome Nome</b></th>
                                                <th title="CodiceAllievo"><b>Codice Allievo</b></th>
                                                <th title="Classe"><b>Classe</b></th>
                                                <th title="ApriChat"><b>Azione</b></th>
											</tr>
										</thead>
										<tbody>
										
										<%
										
										
										
										QuerySQL = "SELECT * FROM Allievi WHERE CodiceAllievo <> '"&Session("CodiceAllievo")&"' AND Id_Classe IN (SELECT Id_Classe FROM Allievi WHERE Id_Classe IN (SELECT DISTINCT Id_Classe FROM stud_as_classe WHERE Id_As = (SELECT MAX(Id_As) FROM anni_classi) ) GROUP BY Id_Classe HAVING COUNT(*) >= 7) order by Classe,Cognome;"
										'response.write QuerySQL
										set rsTabellaAllievi = ConnessioneDB.Execute(QuerySQL)
										
										do while not rsTabellaAllievi.EOF
										
										%>
										
										<tr>
                                                <td title="CognomeNome"><%=rsTabellaAllievi("Cognome")&" "&rsTabellaAllievi("Nome")%></td>
                                                <td title="CodiceAllievo"><%=rsTabellaAllievi("CodiceAllievo")%></td>
                                                <td title="Classe"><% if rsTabellaAllievi("Classe") <> "Expo" and rsTabellaAllievi("Classe") <> "Admin" and rsTabellaAllievi("Classe") <> "DOC" then %><%=rsTabellaAllievi("Classe")%><% else %><%=rsTabellaAllievi("Classe")%><%end if%></td>
                                                <td title="ApriChat"><a target="_blank" style="text-decoration:none" href="leggichat.asp?cod=<%=session("CodiceAllievo")%>&contatto=<%=rsTabellaAllievi("CodiceAllievo")%>">Apri Chat</a></td>
											</tr>
										
										<% ' non consento apertura chat con docenti -> andrebbe fatto controllo anche nella pagina della chat!
										' bisogna implementare anche il cerca altrimenti l'elenco è troppo lungo 
										%>
										
										<%
										rsTabellaAllievi.movenext
										loop
										
										%>
										
										</tbody>
										