 <div class="bs-docs-example">

                          <div class="accordion-group">
                          <div class="accordion-heading">
                            <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3.1" href="#collapseAltro">
                          <center> </b> <span style="text-align:center">  Admin</span></center>
                            </a>

                          </div>
                          <div id="collapseAltro" class="accordion-body collapse">

                              <ul id="myTab4" class="nav nav-tabs">
                                  <li class="active"><a href="#profileSorteggia" data-toggle="tab">Sorteggia</a></li>
                                  <li><a href="#profileConvalida" data-toggle="tab">Convalida</a></li>
                                  <li><a href="#profileGenera" data-toggle="tab">Genera</a></li>
                                  <li><a href="#profileAggiorna" data-toggle="tab">Service</a></li>



                            </ul>
                            <div id="myTabContent3" class="tab-content">
                              <div class="tab-pane fade in active" id="profileSorteggia">

           <div class="box box-color box-bordered">

                           <form method="POST" action="classifica.asp?divid=<%=divid%>&classe=<%=classe%>&DataCla=<%=DataCla%>&xEstrazione=1&id_classe=<%=id_classe%>">
<FIELDSET style="margin-left:16px;" ><LEGEND class="sottotitoloquaderno2"><B><a name="sorteggia" style="text-decoration:none"> Sorteggia Studente</a></B></LEGEND>
 <p>

<%' devo generare un numero casuale che servirà per accedere alla tabella studenti
 if xEstrazione<>"" then
			  randomize()
			  NumeroCasuale = Int(NumStud * Rnd + 1)
end if

%>
 &nbsp;&nbsp;
 <input type="text" name="txtSTUD" id="txtSTUD" value="" size="50"> &nbsp;&nbsp; <input class="btn" type="button" onclick="estrai()" style="vertical-align:top" value="Estrai" name="B1"> <!--Definisce i due bottoni del form -->

 </p></FIELDSET>
</form>


<script>

function getRandomArbitrary(min, max) {
  return Math.random() * (max - min) + min;
}

function estrai(){

	var vettstud = new Array();
	var strstud = "null,<%=Left(strstud,Len(strstud)-1)%>";

	vettstud = strstud.split(",");

	var num = Math.round(getRandomArbitrary(1,<%=NumStud%>));

	$("#txtSTUD").fadeOut(1000);
	var t = setTimeout(function(){ document.getElementById("txtSTUD").value = vettstud[num]; clearTimeout(t); }, 1250);
	$("#txtSTUD").fadeIn(2000);

}




</script>



   </div>



                              </div>
                              <div class="tab-pane fade" id="profileConvalida">




                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Quiz svolti
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">

                                            <form method="POST" form action="aggiorna_punteggio.asp?classe=<%=classe%>&xQuiz=1&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>" >

                                               	<table class="table table-hover table-nomargin">
                                                   <thead>
                                                        <tr>
                                                            <th>Argomento</th>
                                                            <th>Data</th>
                                                            <th>Elimina</th>

                                                        </tr>
                                                    </thead>
                                                     <tbody>
                                                             <%
							'seleziono le sessioni per la classe e quindi i qui corrispondenti
					QuerySQL1="SELECT * FROM [dbo].[2SESSIONI_QUIZ] WHERE Id_Classe='"&id_classe &"'"
							'response.write(QuerySQL1)

							'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\78.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close
'

					Set rsTabellaSess = ConnessioneDB.Execute(QuerySQL1)

	   			    do while not rsTabellaSess.eof

							QuerySQL1="SELECT DISTINCT Id_Classe, CodiceTest, Data, Titolo " &_
							" FROM (Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo) INNER JOIN Moduli " &_
							"ON Risultati1.CodiceTest = Moduli.ID_Mod WHERE Id_Classe='"&id_classe &"' and Sessione="&rsTabellaSess("ID_Sessione")&";"
							'response.write(QuerySQL1)


						'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\90.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

							Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) %>



                                                         <%
														 i=0
														 if rsTabella1.eof then
														 %><tr><td colspan="2" >Sessione (<%=rsTabellaSess("ID_Sessione")%>):&nbsp;<%=rsTabellaSess("Titolo")%>, non ci sono quiz svolti </td></tr><%
														 end if
													do while not rsTabella1.eof%>
														<tr>
                                                        <td>
                                                        <a href="../cDomande/aggiorna_punteggio_pulisci_test.asp?SessioneQuiz=<%=rsTabellaSess("ID_Sessione")%>&tipoTest=<%=rsTabellaSess("TipoQuiz")%>&classe=<%=classe%>&xQuiz=1&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>&CodiceTest=<%=rsTabella1.fields("CodiceTest")%>&DataTest=<%=rsTabella1.fields("Data")%>&TitoloTest=<%=rsTabella1.fields("Titolo")%>"> (<%=rsTabellaSess("ID_Sessione")%>):&nbsp;<%=rsTabellaSess("Titolo")%> &nbsp;(Risposta&nbsp;
                                                        <% select case(rsTabellaSess.fields("TipoQuiz"))
				case 0:
				  response.write("Vero/Falso")
				  case 1:
				   response.write("Singola")
				   case 2:
				    response.write("Multipla")
				end select
				%>)
                                                        </a></td>
                                                        <td><%=rsTabella1("Data")%></td>
                                                         <td><i class="icon-remove"></i></td>
                                                        </tr>

															<%
														rsTabella1.movenext
													loop
													%>

                             <%
									rsTabellaSess.movenext
									loop
								%>



                                                    </tbody>
                                                </table>
                                             </div>
                                        </div>
                                  </div>


                              <div class="tab-pane fade" id="profileGenera">

                             <%    ' PREPARO REPORT
		QuerySQL="SELECT Url, Data, Descrizione FROM VERIFICHE Where Id_Classe='"& id_classe &"'"
		Set rsTabella = ConnessioneDB.Execute(QuerySQL)%>



                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Cronologia
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                            <th>Salva la Classifica </th>


                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                       <td>
                                                        <a target="_new" href="../cGrafici/genera_grafico.asp?PS=<%=PS%>&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>&indice_periodo=<%=indice_periodo%>&indice_periodo2=<%=indice_periodo2%>">Salva per grafico </a><br></p>
                                                       </td>
                                                     </tr>
                                                     <tr>
                                                    <td>
                                                     <a target="_new" href="classifica_report.asp?PS=<%=PS%>&id_classe=<%=id_classe%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>&indice_periodo=<%=indice_periodo%>&indice_periodo2=<%=indice_periodo2%>">Salva per report</a><br></p>
                                                    </td>
                                                  </tr>


                                                    </tbody>
                                                </table>
                                             </div>
                                        </div>

                              </div>



                               <div class="tab-pane fade" id="profileAggiorna">



                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Funzioni extra
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                            <th>Progressione </th>


                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                       <tr>
                                                       <td>
                                                        <a target="_new" href="../cAdmin/aggiorna_anni_scolastici.asp?id_classe=<%=id_classe%>">Aggiorna a/s </a><br></p>
                                                       </td>
                                                     </tr>
                                                     <tr>
                                                     <td>
                                                      <a target="_new" onclick="caricaeccezioni('<%=id_classe%>');">Eccezioni </a><br></p>

                                                     </td>
                                                   </tr>



                                                    </tbody>
                                                </table>
                                                <div id="compitispec">

                                                </div>
                                             </div>
                                        </div>

                              </div>


                            </div>
                        </div>
                        </div>
              </div>
