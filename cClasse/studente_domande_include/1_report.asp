 <%%>	
 </b><center>
 
										 
										<div class="span12"> 
										
									
     <div class="bs-docs-example">
                       
                          <div class="accordion-group">                
                          <div class="accordion-heading"><br />
                           <center>      <p> <a target="_new" href="../cGrafici/genera_grafico.asp?byGrafico=1&amp;PS=<%=PS%>&amp;id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>&indice_periodo=<%=indice_periodo%>&indice_periodo2=<%=indice_periodo2%>">Mostra Grafico</a><br></p>
<p><a target="_blank" href="../cAdmin/consulta_profili_new.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=ClasseProfili%>">Profili Classe </a></p>
<!--
<p><a target="_blank" href="../cAdmin/consulta_profili_new_regno.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=ClasseProfili%>">Profili Regno </a></p>
-->
<p><a target="_blank" href="../cAdmin/consulta_profili_quiz.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=ClasseProfili%>">Profili Quiz </a></p>

<p><a target="_blank" href="../cGrafici/grafico.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=ClasseProfili%>&periodoInizio=<%=replace(DataCla,"/","_")%>&periodoFine=06_06_2020&primachiamata=1">Profili Andamenti </a></p>
<p><a target="_blank" href="../cGrafici/secondaVersione/index.php?id_classe=<%=id_classe%>&periodoInizio=<%=replace(DataCla,"/","_")%>&periodoFine=<%=replace(DataCla2,"/","_")%>&classe=<%=ClasseProfili%>&anno=<%=anno_scolastico%>&primachiamata=1">Profili Andamenti (2) </a></p>
</center>  
                            <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3.1" href="#collapseReport">
                            <span style="text-align:center"> Report attivit&agrave;</span>
                            </a>
                            <br />
                          </div>
                          <div id="collapseReport" class="accordion-body collapse">
       
                              <ul id="myTab3" class="nav nav-tabs">
                                  <li class="active"><a href="#profileVerifiche" data-toggle="tab">Verifiche</a></li>
                                  <li><a href="#profileCrediti" data-toggle="tab">Crediti</a></li>
                                  
    
                                  
                            </ul>
                            <div id="myTabContent2" class="tab-content">
                              <div class="tab-pane fade in active" id="profileVerifiche">
                               
                             <%    ' PREPARO REPORT
		QuerySQL="SELECT Url, Data, Descrizione FROM VERIFICHE Where Id_Classe='"& id_classe &"'"
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)%> 
                               
                               
                               
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Risultati nelle verifiche
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                            <th>Argomento</th>
                                                            <th>Data</th>
                                                           
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                      <%
													  if rsTabella.eof then%>
                                                      <tr><td colspan="2">Nessuna verifica svolta</td></tr>
													 <% else
													    do while not rsTabella.eof%> 
														<tr><td><a href="<%=rsTabella(0)%>"><%=rsTabella(2)%></a></td><td><%=rsTabella(1)%></td></tr>
														<%
															rsTabella.movenext
														loop
													   end if
													   %>
																									   
                                                        
                                                     
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>

                              </div>
                              <div class="tab-pane fade" id="profileCrediti">
                            <%
							QuerySQL= "SELECT distinct [2CREDITI].Id_Esercitazione, Descrizione,Data  FROM [2ESERCITAZIONI_SINGOLI] INNER JOIN [2CREDITI] ON " &_
		" [2ESERCITAZIONI_SINGOLI].ID_Esercitazione = [2CREDITI].Id_Esercitazione " &_
		"  Where Id_Classe='"& id_classe &"' and Descrizione<>'Iscrizione';"		 
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)
							%>
                                                 
                            
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Punti Extra
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                            
                                            
                                            
                                               	<table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                          <th>
                                                            <b>Attivit&agrave;
                                                            <% if session("Admin")=true then %> 
                                                            <a target=blank href="../cReport/report_verifiche.asp?id_classe=<%=id_classe%>&amp;AggiungiReport=1">(+)
                                                            </a><%end if%></b>
                                                            </th> 
                                                            <th> Data
                                                            </th>
                                                             <th> Elimina
                                                            </th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                         <%
		 
													do while not rsTabella.eof%> 
														<tr>
                                                        <td><a target="_new" href="../cReport/report_verifiche.asp?ID_ESER=<%=rsTabella(0)%>"><%=rsTabella(1)%></a></td><td><%=rsTabella(2)%></a></td><td><a target="_new" href="../cReport/report_verifiche_aggiorna.asp?Cancella=1&amp;ID_ESER=<%=rsTabella(0)%>"><i class="icon-trash"></i></a><td></tr> 
														
															<%
														rsTabella.movenext
													loop
													%>
                                                        
                                                            
                                                        
                                                        
                                                     
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>
                            
                           
                            
                              </div>
                                                       
                            </div>
                        </div>
                        </div> 
              </div>          
           </div> 
                         