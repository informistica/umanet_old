 
     <form method="POST" name="frmDocument3" class="form-horizontal form-bordered form-validate" action="../cUtenti/modifica_contatti.asp"  > 
								 
									 
                                    
                                    
                                  
									<div class="control-group">
										<label for="textfield"  class="control-label">Email</label>
										<div class="controls">
                                         <input type="hidden" name="txtCodiceAllievo" value="<%=rsTabella("CodiceAllievo")%>">
                                            
                                       
                                             <% if strcomp(rsTabella("Email")&"","")=0 then %>
                                                <input type="text" style="height: auto;"  class="input-xlarge" placeholder="Nessuna" name="txtEm" id="emailfield" data-rule-required="true"  data-rule-email="true">
										     <%else%>
                                                <input type="text" style="height: auto;"  id="emailfield" class="input-xlarge" data-rule-required="true"  data-rule-email="true"  name="txtEm" value="<%=rsTabella("Email")%>" >
                                             <%end if%>
                                        </div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nuova email</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"      name="txtNewEm" placeholder="Inserisci la nuova email">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield"  class="control-label">Conferma email</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtNewEm1" placeholder="Conferma la nuova email" data-rule-equalTo="#emailfield" data-rule-required="true"  data-rule-email="true">
										</div>
                                     </div>
                                     
                                     
                                  <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then %>
                                       
                                  	 <div class="form-actions">
										<button type="submit"  class="btn btn-primary" name="B2">Salva modifiche</button>	 
									</div>
                                    <%end if%>
									
								</form>