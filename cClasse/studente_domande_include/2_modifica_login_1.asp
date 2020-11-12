
              <form  class='form-horizontal form-bordered form-validate' name="frm0"  METHOD = "POST">
   
   
   							  <div class="control-group">
										<label for="textfield"  class="control-label">Username</label>
										<div class="controls">
											<input type="text"   style="height: auto;"  class="input-xlarge"  name="txtCodiceAllievo" value="<%=rsTabella("CodiceAllievo")%>">
										</div>
                                     </div>
                                     
                                     
                                      <%' tolgo visualizzazione della password perché con crittografia non serve più a niente : bisognerebbe implementare un bel recupera password!!
									  %>
                                     
                                        <!--<div class="control-group">
                                        <label for="textfield"  style="height:auto" class="control-label">Password</label>
										<div class="controls">
                                          <% if (session("Admin")=true) and (strcomp(ucase(pwdAdmin),ucase(rsTabella("Password")))<>0) then%>
											<input  type="text" style="height: auto;" class="input-xlarge"  name="txtPwdAllievo" value="<%=rsTabella("Password")%>">
                                            <%else%>
                                            <input  type="password" style="height:auto"   class="input-xlarge"  name="txtPwdAllievo" value="">
                                            <%end if%>
										</div>
								      	</div>-->
                                    
                                   
                        
                                    <div class="control-group">
										<label for="textfield"  class="control-label">Password</label>
										<div class="controls">
											<input type="password" style="height: auto;" data-rule-required="true" data-rule-minlength="2"  class="input-xlarge"  name="txtPwdOld"  placeholder="Inserisci password attuale">
										</div>
                                     </div>
                                     
                                     
                                     <div class="control-group">
										<label for="textfield"  class="control-label">Nuova Password</label>
										<div class="controls">
										<input type="password" style="height:auto"   name="txtNewPwd"   id="pwfield2" class="input-xlarge" data-rule-required="true" placeholder="Nuova password">
										</div>
									</div>
									<div class="control-group">
                                        <label for="textfield"  style="height: auto;" class="control-label">Conferma Password</label>
										<div class="controls">
											<input type="password" style="height:auto"  name="txtNewPwd1"   id="confirmfield" class="input-xlarge" data-rule-equalTo="#pwfield2" data-rule-required="true" placeholder="Conferma nuova password">
										</div>
                                     </div>
     
                                    <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then %>
                                       
                                  	 <div class="form-actions">
										<button type="button"  class="btn btn-primary" name="B2" onclick="crittapwd();">Salva modifiche</button>	 
									</div>
                                    <%end if%>
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>