 
  <form method="POST" action="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&amp;Id_Stud=<%=cod%>" name="frmDocument1"  class="form-horizontal form-bordered form-validate"> 
						 <%
						 QuerySQL="SELECT count(*)  FROM Eccezioni_Frasi  WHERE  Id_Stud='"&cod&"';"
						 Set rsTabella = ConnessioneDB.Execute(QuerySQL)
						 eccFrasi = rsTabella(0)
						 QuerySQL="SELECT count(*)  FROM Eccezioni_Nodi  WHERE  Id_Stud='"&cod&"';"
						  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
						 eccNodi=rsTabella(0)
						 QuerySQL="SELECT count(*)  FROM Eccezioni_Domande  WHERE  Id_Stud='"&cod&"';"
						  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
						 eccDomande=rsTabella(0)
						 
						 %>
                        			 
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Frasi </label>
										<div class="controls">
											<input type="text" style="height: auto;" disabled class="input-mini" value="<%=eccFrasi%>" rel="tooltip"  title="Numero di proroghe concesse" >
										 &nbsp; &nbsp;
										<a href="2_rimuovi_eccezioni.asp?frasi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta Eccezioni" rel="tooltip"  title="Consulta" ></a>
										</div>
									</div>
                                    
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Domande </label>
										<div class="controls">
											<input type="text" style="height: auto;" disabled class="input-mini"  value="<%=eccDomande%>" rel="tooltip"  title="Numero di proroghe concesse">
                                             &nbsp; &nbsp;
										<a href="2_rimuovi_eccezioni.asp?domande=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta Eccezioni" rel="tooltip"  title="Consulta" ></a>
										</div>
									</div>
                                    
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nodi</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-mini" disabled  value="<%=eccNodi%>" rel="tooltip"  title="Numero di proroghe concesse">
                                             &nbsp; &nbsp;
										<a href="2_rimuovi_eccezioni.asp?nodi=1&amp;cod=<%=cod%>">	<input type="button" clas="btn" value="Consulta Eccezioni" rel="tooltip"  title="Consulta" ></a>	 
										</div>
                                        
                                        
									</div>  
                                  <br />
                                     
                                 
								</form>