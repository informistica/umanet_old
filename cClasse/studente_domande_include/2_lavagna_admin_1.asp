<%
' a dispetto del nome vado sul diario scegli=2

 QuerySQL="SELECT * FROM [FORUM_MESSAGES]  Where ID_Classe='"& ID_Classe &"' and ParentMessage=0 and Comments<>'InizializzaDB' and Id_Social=2 order by DatePosted desc;"


	' response.Write(QuerySQL)
 Set rsTabella = ConnessioneDB3.Execute(QuerySQL)

 %>
 <div class="box-content">
								<div class="accordion" id="accordion2">
								<%If rsTabella.BOF=True And rsTabella.EOF=True Then %>
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapseOne">
												Non ci sono avvisi !
											</a>
										</div>
										<div id="collapseOne" class="accordion-body collapse in">
											<div class="accordion-inner">
											..
											</div>
										</div>
									</div>
                                  <% Else%>
                                   <% i=1
		   i=0 ' prima posizione in classifica serve per calcolare il trend%>
		    <form method="POST" action="modifica_avviso.asp?i=<%=i%>&tipoAvviso=1&classe=<%=classe%>" name="Aggiorna" class='form-horizontal form-striped' >
		    <input type="hidden" name="txtClasse" value="<%=cartella%>">
		   
			<%do while not rsTabella.EOF 
			 %>
            
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse"  href="#collapseL<%=i%>">
											<%=rsTabella ("Topic") %> </a>  
										</div>
                                        
										<div id="collapseL<%=i%>" class="accordion-body collapse ">
											<div class="accordion-inner">
                                             
                                            
                                            <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Modifica post</h3>
							</div>
							<div class="box-content nopadding">
								<form action="#" method="POST" class='form-horizontal form-striped'>
									<div class="control-group">
										<label for="textfield" class="control-label">Post</label>
										<div class="controls">
											<input type="text" name="textPost" id="textfield" placeholder="Text input" class="input-xlarge" value="<%=ReplaceCar( rsTabella("Topic"))%>">
										</div>
									</div>
									<div class="control-group">
										<label for="password" class="control-label">Data e Ora</label>
										<div class="controls">
											<input type="text" name="txtData" id="data" placeholder="Data e ora" class="input-xlarge" value="<%= rsTabella("DatePosted")%>">
										</div>
									</div>
									<div class="control-group">
									<%
   
				sReadAll=rsTabella.fields("Comments")
				if len(sReadAll)=0 then
				     sReadAll="Nessuna spiegazione"
				end if
			  %>
                                    
                                    
									<div class="control-group">
										<label for="textarea" class="control-label">Spiegazione</label>
										<div class="controls">
											<textarea  name="txtSpiegazione<%=i%>"  id="textarea" rows="5" class="input-block-level">
                                            <%=ReplaceCar(sReadAll)%>
                                            </textarea>
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textarea" class="control-label">Azione </label>
										<div class="controls">
											<textarea name="txtAzione<%=i%>" id="textarea" rows="5" class="input-block-level">
                                            <%=rsTabella.fields("Azione")%>
                                            </textarea>
                                            
										</div>
									</div>
                                    
                                    
                                    
                                    	<label class="control-label">Visibile<small>Mostra sulla lavagna</small></label>
										<div class="controls">
											<label class='checkbox'>
												<input type="checkbox" name="checkbox"> Mostra
											</label>
											<label class='checkbox'>
												<input type="checkbox" name="checkbox"> Nascondi
											</label>
										</div>
									</div>
									<div class="form-actions">
										<button type="button" class="btn" onClick="javascript:modifica(<%=i%>);">Aggiorna</button>
                                        <input type="hidden" value ="<%=rsTabella.fields("ID")%>" name="txtIdA<%=i%>">
										 
									</div>
								</form>
							</div>
						</div>
                                            
                                            
												 
											</div>
										</div>
									</div>
            
								
                                 <%  i=i+1
		   
		   rsTabella.movenext
		loop%>
	<%end if%>
    </div>
</div>

 



			  
			  
 
		 
	 
