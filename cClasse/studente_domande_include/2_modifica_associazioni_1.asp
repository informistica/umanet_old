 
     <% QuerySQL = "SELECT * FROM AssociazioniAllievi WHERE CodiceAllievo = '"&rsTabella("CodiceAllievo")&"' OR UtenteAssociato = '"&rsTabella("CodiceAllievo")&"'"
	 set rsTabellaA = ConnessioneDB.Execute(QuerySQL)
	 i = 1
	 %>
	 
	 <form class="form-horizontal form-bordered form-validate"> 
								 
									 
                                    
                                    <% do while not rsTabellaA.EOF %>
                                  
									<div class="control-group">
										<label for="textfield"  class="control-label">Associazione <%=i%></label>
										<div class="controls">
                                            
										<input type="text" style="height: auto;"  id="ca<%=i%>" class="input-xlarge" disabled  value="<%=rsTabellaA("CodiceAllievo")%>" >
										<input type="text" style="height: auto;"  id="ua<%=i%>" class="input-xlarge" disabled value="<%=rsTabellaA("UtenteAssociato")%>" > &nbsp; &nbsp;
										<a style="text-decoration:none" href="javascript:void(0)" onclick="eliminaassociazione('<%=rsTabellaA("CodiceAllievo")%>','<%=rsTabellaA("UtenteAssociato")%>')"><i style="color:red"  class="icon-remove"></i></a>
                                             
                                        </div>
                                     </div>
                                     
                                     <% rsTabellaA.movenext
									 i=i+1
									 loop %>
                                     
									 <div class="control-group">
										<label for="textfield"  class="control-label">Nuova Associazione</label>
										<div class="controls">
                                            
										<input type="text" style="height: auto;"  id="canew" class="input-xlarge" readonly disabled  value="<%=rsTabella("CodiceAllievo")%>" ><br><br>
										<input type="text" style="height: auto;"  id="uanew" class="input-xlarge"  value="" placeholder="Inserisci il Codice Allievo da associare" >
										<input type="password" style="height: auto;"  id="pwdnew" class="input-xxlarge"  value="" placeholder="Inserisci la password del Codice Allievo da associare" >
                                             
                                        </div>
                                     </div>
									 
                                  <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then %>
                                       
                                  	 <div class="form-actions">
										<button type="button" onclick="aggiungiassociazione()" class="btn btn-primary" name="B2">Aggiorna</button>	 
									</div>
                                    <%end if%>
									
								</form>
								
								<script src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>

								
								<script>
								
								function aggiungiassociazione(){
								
									var user = $("#uanew").val();
									var pass = $("#pwdnew").val();
									
									if(user.trim() == "" || pass.trim() == ""){
										alert("Non puoi lasciare campi non compilati");
									}else{

										$.ajax({
										  method: "POST",
										  url: "aggiungiassociazione.asp",
										  dataType: "html",
										  data: { user: user.trim(), pass:  Sha256.hash(pass.trim()) }
										}) /* .ajax */
										 .done(function( ans ) {
										 //alert(ans);
											if(ans.trim() == "associato"){
												alert("Associazione effettuata correttamente");
												window.location.reload();
											}else if(ans.trim() == "presente"){
												alert("Associazione già presente");
											}else{
												alert("Username e/o Password errati");
											}
										 }); /* .done */
									
									}
									
								}
								
								function eliminaassociazione(ca, ua){
								
									var conferma = confirm("Vuoi eliminare l'associazione?");
									
									/*alert(ca);
									 alert(ua);*/
									
									if(conferma){
									
									$.ajax({
										  method: "POST",
										  url: "eliminaassociazione.asp",
										  dataType: "html",
										  data: { all1: ca, all2: ua }
										}) /* .ajax */
										 .done(function( ans ) {
										 //alert(ans);
											if(ans.trim() == "eliminato"){
												alert("Associazione eliminata correttamente");
												window.location.reload();
											}else{
											//alert(ans.trim());
												alert("Errore");
											}
										 }); /* .done */
									
									
									}
								
								}
								
								</script>