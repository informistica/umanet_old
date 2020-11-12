<!-- #include file = "studente_domande_include/4_quaderno.asp" -->

		<% DataClaN = DataCla
		DataCla2N = DataCla2 %>


		<!-- #include file = "../var_globali.inc" -->


 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

		<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->
        <!-- #include file = "../cClasse/studente_domande_include/1_periodi_date.asp" -->

        <!-- #include file = "../extra/test_server.asp" -->

 <!-- #include file = "../cUtenti/adovbs.inc" -->

 <% DataCla = DataClaN
 DataClaq = DataClaN
		DataCla2 = DataCla2N
		DataClaq2 = DataCla2N
		%>

 		<!-- #include file = "../include/formattaDataCla.inc" -->

	
<% 'response.write DataCla
%>

 <%

 CodiceAllievo = Request.querystring("cod")
 idmod=Request.querystring("idmod")

 'per le store procedure
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set cmd1 = Server.CreateObject("ADODB.Command")
set cmd2 = Server.CreateObject("ADODB.Command")
set cmd3 = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd1.activeconnection = conn
set cmd2.activeconnection = conn
set cmd3.activeconnection = conn
%>

<%
umanet=""
umanet=request.querystring("umanet")

'if strcomp(umanet,"1")=0 then
'QuerySQL="SELECT * FROM MODULI_CLASSE_UMANET " &_
' " WHERE Id_Classe='" & id_classe  &"' and ID_Mod='"&idmod&"'"
' else
' QuerySQL="SELECT * FROM MODULI_CLASSE " &_
' " WHERE Id_Classe='" & id_classe  &"' and ID_Mod='"&idmod&"'"
 
'end If 


 if umanet="" then
 QuerySQL="SELECT * FROM MODULI_CLASSE " &_
 " WHERE Id_Classe='" & id_classe  &"' and ID_Mod='"&idmod&"'"
else
QuerySQL="SELECT * FROM MODULI_CLASSE_UMANET " &_
 " WHERE Id_Classe='" & id_classe  &"' and ID_Mod='"&idmod&"'"
end If 
 

  Set rsTabellaModuli = ConnessioneDB.Execute(QuerySQL)
   '  response.write(QuerySQL)
 %>

 <% 'k=0
 'p=0
   'compiti=0 ' serve per mettere il box se non ci sono compiti inseriti


			
		     %>
               <!-- #include file = "studente_domande_include/3_statistica_frasi.asp" -->
              <!-- #include file = "studente_domande_include/3_statistica_nodi.asp" -->
              <!-- #include file = "studente_domande_include/3_statistica_domande.asp" -->
              <%if umanet<>"" then%>
              <!-- #include file = "studente_domande_include/3_statistica_metafore.asp" -->
              <%end if%>
                    <table class="table table-hover table-nomargin table-condensed">
                        <thead>
                        <tr align="center">
                         <th>
                        
                    <%
                        'on error resume next
						if numrsPreFrasi<>0 then
						percFrasi=fix((numrsFrasi/numrsPreFrasi)*10)/10*100
						else
						percFrasi=0
						end if
						if numrsPreDomande<>0 then
						percDomande=fix((numrsDomande/numrsPreDomande)*10)/10*100
						else
						percDomande=0
						end if
						if numrsPreNodi<>0 then
						percNodi=fix((numrsNodi/numrsPreNodi)*10)/10*100
						else
						percNodi=0
						end if
                        if numrsPreMetafore<>0 then
						percMetafore=fix((numrsMetafore/numrsPreMetafore)*10)/10*100
						else
						percMetafore=0
						end if

						numrsDomandeBack=numrsDomande%>
						 
                         <%
						 QuerySQL="SELECT Cartella FROM Classi  WHERE ID_Classe='"&id_classe&"' " ' order aggiunto by 27/09 per ordinare i paragrafi nel quaderno dei compiti svolti
						Set rsTabellaCartella = ConnessioneDB.Execute(QuerySQL)
						cartella=rsTabellaCartella(0)

if umanet="" then
QuerySQL="SELECT * FROM MODULI_PARAGRAFI_CLASSE " &_
" WHERE ID_Mod='" & rsTabellaModuli("ID_Mod") & "' and Id_Classe='"&id_classe&"'  order By Posizione, Expr1;" ' order aggiunto by 27/09 per ordinare i paragrafi nel quaderno dei compiti svolti
   else
QuerySQL="SELECT * FROM MODULI_PARAGRAFI_CLASSE_UMANET " &_
" WHERE ID_Mod='" & rsTabellaModuli("ID_Mod") & "' and Id_Classe='"&id_classe&"'  order By Posizione, Expr1;" ' order aggiunto by 27/09 per ordinare i paragrafi nel quaderno dei compiti svolti
end if
QueryTuttoCap="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_MOD='"&  rsTabellaModuli("ID_Mod") &"' and CodiceAllievo='"&CodiceAllievo&"'order by Posizione, CodiceFrase;"

  Set rsTabellaParagrafi = ConnessioneDB.Execute(QuerySQL)%>

       <%

				' servono solo per i parametri per aprire tutti i compiti del cap, forse si può anche fare a meno usando i parametri di rsTabellaModuli
				%>
                <!-- #include file = "studente_domande_include/2_nodi_0.asp" -->

                <!-- #include file = "studente_domande_include/2_domande_0.asp" -->
                <!-- #include file = "studente_domande_include/2_frasi_0.asp" -->
                <% if umanet<>"" then %>
                <!-- #include file = "studente_domande_include/2_metafore_0.asp" -->
                <% end if%>


                       <ul class="pagestats style-3">
											<li>
                                                <div class="spark">
													<div title="% di Frasi svolte" class="chart" data-percent="<%=percFrasi%>" data-color="#368ee0" data-trackcolor="#d5e7f7">
													<%=percFrasi%> %
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaFrasi.eof then %>
                                                 <span style="color:#000" title="Apri tutte le frasi del capitolo" href="../cFrasi/2inserisci_valutazioni_frasi.asp?TutteCap=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>">
                                                 <%end if%>
													<span class="name"><%=numrsFrasi%> su <%=numrsPreFrasi%></span>
                                                    <span class="name">PF.<%=numrsFrasi2%> </span>
                                                      </span>
												</div>
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Domande svolte" class="chart" data-percent="<%=percDomande%>" data-color="#56af45" data-trackcolor="#dcf8d7">
													<%=percDomande%> %
                                                    </div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaDomande.eof then %>
                                                  <a style="color:#000" title="Apri tutte le domande del capitolo" href="../cDomande/inserisci_valutazioni.asp?Tutte=1&ID_MOD=<%=rsTabellaDomande("ID_MOD")%>&ID_PAR=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_MOD")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&TitoloParagrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">
                                                   <%end if%>
													<span class="name"><%=numrsDomandeBack%> su <%=numrsPreDomande%></span>
                                                    <span class="name">PD.<%=numrsDomande2%> </span>
                                                    </a>
												</div>
											</li>
                                            <li>
												<div class="spark">
													<div title="% di Nodi svolti" class="chart" data-percent="<%=percNodi%>" data-color="#f96d6d" data-trackcolor="#fae2e2"><%=percNodi%>%</div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaNodi.eof then %>
                                                 <a style="color:#000" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>">
													<%end if%>
                                                    <span class="name"><%=numrsNodi%> su <%=numrsPreNodi%></span>
                                                    <span class="name">PN.<%=numrsNodi2%> </span>
                                                    </a>
												</div>
											</li>
                                               <li>
												<div class="spark">
													<div title="% di Metafore svolte" class="chart" data-percent="<%=percMetafore%>" data-color="#FC0" data-trackcolor="#F90"><%=percMetafore%>%</div>
												</div>
												<div class="bottom">
                                                <%if not rsTabellaMetafore.eof then %>
                                                 <a style="color:#000" title="Apri tutte le Metafore del paragrafo"  href="../cMetafore/2inserisci_valutazioni_Metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaMetafore("ID_Mod")%>&CodiceAllievo=<%=rsTabellaMetafore("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore("Cartella")%>&Modulo=<%=rsTabellaMetafore("ID_Mod")%>&Capitolo=<%=rsTabellaMetafore("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore("TitoloParagrafo")%>"> 
													<%end if%>
                                                    <span class="name"><%=numrsMetafore%> su <%=numrsPreMetafore%></span>
                                                    <span class="name">PM.<%=numrsMetafore2%> </span>
                                                    </a>
                                                   
												</div>
                                               
											</li>

										</ul>
                      </th>
                                                        </tr>
                                                    </thead>
                     </table>
                 				 <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("Admin")=true) and (numrsFrasi<>0) then%>
								<!--<form name="dati" method="POST" target="_blank" action="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">
                           -->
                         <!--  <i class="icon-print"></i>-->
						      <img src="../../img/printer.jpg" title="Stampa frasi, domande, nodi">
                               <!--  <input type="submit" class="btn" value="Stampa Frasi Capitolo" >  -->
								 <a href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=<%=umanet%>&tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>"" target="_blank">
								 <input type="button" class="btn" value="Stampa Frasi Capitolo" >
								 </a>
								 <% if session("admin")=true then%>
                                     <a title="Stampa paragrafi e domande" target="_blank" href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=<%=umanet%>&sint=1&tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">(Sint)</a>

                                       <a title="Stampa solo paragrafi" target="_blank" href="../cFrasi/7_stampa_schede_frasi_elenco_sint.asp?umanet=<%=umanet%>&supersint=1&tutto=1&CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&Paragrafo=<%=rsTabellaFrasi("TitPar")%>&Cartella=<%=cartella%>&DataCla=<%=DataCla%>&DataCla2=<%=DataCla2%>">(Super)</a>
                                   <a href="#"><input type="button" class="btn" value="Stampa Domande Capitolo" >
                                  <%end if%>

								  
								 </a>
								 <% if session("admin")=true then%>
                                  <%end if%>

                               <!-- </form>	 -->

								<% end if%>
                     <% p=0
										 riga=0
		     do while not rsTabellaParagrafi.EOF
                %>

 							    <!-- #include file = "studente_domande_include/2_frasi_1.asp" -->
                                <!-- #include file = "studente_domande_include/2_domande_1.asp" -->
                                <!-- #include file = "studente_domande_include/2_nodi_1.asp" -->

					 <!--Qua il controllo per vedere se ci sono compiti svolti per quel paragrafo-->
                     <% 'Response.write(rsTabellaParagrafi("ID_Paragrafo") & numrsFrasi &" " & " " & numrsNodi & " " &numrsDomande & "<br>")
					 %>


                            <%numrsMetafore=0
							   select case rsTabellaParagrafi("Paragrafo")
							 
							 case "Topolino ed Obiettivi" 
							
							 %>
							 <!--#include file = "studente_domande_include/2_metaforeT_1.asp"-->
							 <%
							 numrsMetafore=numrsTabellaT
							 numrsMetafore2=numrsTabellaPT
							   case "Navigazione nella Rete della Vita" 
								 
							 %> 
								 <!--#include file = "studente_domande_include/2_metaforeN_1.asp"-->
							 <%
								 numrsMetafore=numrsTabellaN
								 numrsMetafore2=numrsTabellaPN	 
							  case "Relazione Cliente Servitore" 
							 %>
							 	 <!--#include file = "studente_domande_include/2_metaforeD_1.asp"-->
							 <%
							   numrsMetafore=numrsTabellaD
							   numrsMetafore2=numrsTabellaPD
							   
							   case else 
							    numrsMetafore=0
							   numrsMetafore2=0
							   
							 end select  %>


					<% if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0) or (numrsMetafore<>0) then %>

                          <div class="accordion-group">

                          <div class="accordion-heading">

                            <a style="text-decoration:none" id="toggleSottoPar<%=k%><%=p%>" title="<%=k%><%=p%>" class="accordion-toggle" data-toggle="collapse" data-parent="#accordionnew<%=k%><%=p%>" href="#collapseTrenew<%=k%><%=p%>">
                            <%=rsTabellaParagrafi("Paragrafo") %> <small> (<% Response.write(numrsFrasi2+numrsNodi2+numrsDomande2+numrsMetafore2)%>)</small>&nbsp;&nbsp;
						<% if numrsFrasi2 > 0 then %><small><i class="icon-reply"></i>(<% Response.write(numrsFrasi2)%>)</small><%end if%>&nbsp;&nbsp;
						<% if numrsNodi2 > 0 then %><small><i class="glyphicon-snowflake"></i>(<% Response.write(numrsNodi2)%>)</small><%end if%>&nbsp;&nbsp;
						<% if numrsDomande2 > 0 then %><small><i class="icon-question-sign"></i>(<% Response.write(numrsDomande2)%>)</small><%end if%>&nbsp;&nbsp;
                        <% if numrsMetafore2 > 0 then %><small><i class="icon-picture"></i>(<% Response.write(numrsMetafore2)%>)</small><%end if%>
                    </a>

                          </div>




                          <div id="collapseTrenew<%=k%><%=p%>" class="accordion-body collapse">
                              <ul id="myTab3" class="nav nav-tabs">
                                <% if numrsFrasi<>0 then %>
                                  <li  class="active">
								  <%else%>
                                  <li>
								  <%end if%>
                                 <a id="toggleFrasi<%=k%><%=p%>" href="#profileFrasi<%=k%><%=p%>" data-toggle="tab">Frasi (<%=numrsFrasi2%>)</a></li>


                                    <% if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                  <a id="toggleDomande<%=k%><%=p%>" href="#profileDomande<%=k%><%=p%>" data-toggle="tab">Domande (<%=numrsDomande2%>)</a></li>



                                       <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>

                                  <a id="toggleNodi<%=k%><%=p%>" href="#profileNodi<%=k%><%=p%>" data-toggle="tab">Nodi (<%=numrsNodi2%>)</a></li>
                                   <% if (numrsNodi=0 ) and (numrsFrasi=0) and (numrsDomande=0) and (numrsMetafore<>0) then %>
                                         <li class="active">
                                     <%else%>
                                         <li>
                                     <%end if%>
                                    <a id="toggleMetafore<%=k%><%=p%>" href="#profileMetafore<%=k%><%=p%>" data-toggle="tab">Metafore (<%=numrsMetafore2%>)</a></li> 

                            </ul>
                            <div id="myTabContent2<%=k%><%=p%>" class="tab-content">

                              <% if numrsFrasi<>0 then %>
                                  <div class="tab-pane fade in active" id="profileFrasi<%=k%><%=p%>">

								  <%else%>
                                   <div class="tab-pane fade" id="profileFrasi<%=k%><%=p%>">

								  <%end if%>

                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                    <a title="Apri tutte le frasi del paragrafo" style="color:#FFF"  href="../cFrasi/2inserisci_valutazioni_frasi.asp?TuttePar=1&ID_MOD=<%=rsTabellaFrasi("ID_MOD")%>&ID_PAR=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaFrasi("CodiceAllievo")%>&Cartella=<%=rsTabellaFrasi("Cartella")%>&Modulo=<%=rsTabellaFrasi("ID_Mod")%>&Capitolo=<%=rsTabellaFrasi("Titolo")%>&TitoloParagrafo=<%=rsTabellaFrasi("TitPar")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>">
                                                 Apri tutte le frasi: N(<%= numrsFrasi &") Pt(" & numrsFrasi2  & ") Pb("& round( numrsFrasi2/numrsFrasi,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                    <% if not rsTabellaFrasi.eof then %>
                                                        <tr>
                                                            <th>Frase</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Risposto</th>
															 <%if strcomp(cod,Session("CodiceAllievo"))=0 then%>
                                                            <th class='hidden-480'>Elimina</th></tr>
															<%end if%>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">nessun compito inserito</th>

                                                        </tr>
                                                    <%end if%>
                                                    </thead>
                                                    <tbody>



                     <% Sottoparagrafo=""
					' p=0

		     do while not rsTabellaFrasi.EOF
			   if StrComp(Sottoparagrafo, rsTabellaFrasi("SotPar")) <> 0 then
			  ' response.write(p&")<br>strcomp="&Sottoparagrafo&"="&rsTabellaFrasi("SotPar")&" "&StrComp(Sottoparagrafo, (rsTabellaFrasi("SotPar"))))
			   Sottoparagrafo=rsTabellaFrasi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaFrasi("SotPar")%></b></center></td></th>
			 <%end if%>
                        <tr id="riga_<%=riga%>">
															  <%if rsTabellaFrasi("Img")=1 then
															      image="  <i class='icon-picture' title='richiede immagine'></i>"
																  else
																  image=""
															   end if

															 	if rsTabellaFrasi("Segnalata")=1 then
															 		 colore="#F00" 'rosso'
															 	else
															 			if rsTabellaFrasi("Segnalata")=2 then
															 				 colore="#228b22" ' verde foresta'
															 			else
															 				 colore=""
															 			end if
															 	end if
															 	%>

                                                            <td > <a style="color:<%=colore%>"  href="../cFrasi/2inserisci_valutazione_frase.asp?Cartella=<%=Cartella%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaFrasi("ID_Paragrafo")%>&CodiceFrase=<%=rsTabellaFrasi("CodiceFrase")%>&Capitolo=<%=rsTabellaFrasi(9)%>&Paragrafo=<%=rsTabellaFrasi(0)%>&MO=<%=rsTabellaFrasi("ID_Mod")%>&VAL=<%=rsTabellaFrasi("Voto")%>&id_classe=<%=id_classe%>&tCap=<%=k-1%>&tSot=<%=k-1%><%=p%>&tFra=<%=k%><%=p%>"><%=Server.HTMLEncode(rsTabellaFrasi("Chi"))%>&nbsp;<%=image%></a></td>
                                                             <td style="color:<%=colore%>"><%=rsTabellaFrasi("Voto")%></td>


                                                            <td><%=rsTabellaFrasi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaFrasi("Ora"),5)%> </td>
                                                            <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("admin")=true) then%>
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> &nbsp; <i class="icon-reply"></i></td>
                                                            <td class='hidden-480'>

                                                            <a onClick="cancella_frase(<%=rsTabellaFrasi("CodiceFrase")%>,<%=riga%>,'<%=rsTabellaFrasi("ID_Mod")%>','<%=rsTabellaFrasi(0)%>','<%=cartella%>','<%=rsTabellaFrasi("CodiceAllievo")%>');">

                                                            <i class=" icon-trash" ></i></a>
                                                            </td>
															<%end if%>
                                                        </tr>

                 <% f=f+1
				    riga=riga+1
				    rsTabellaFrasi.movenext()
				 loop%>




                                                    </tbody>
                                                </table>
                                             </div>
                                        </div>
                              </div>


                              <%
							'  p=0
							  if (numrsDomande<>0 ) and (numrsFrasi=0) then %>
                                         <div class="tab-pane fade in active" id="profileDomande<%=k%><%=p%>">

                                     <%else%>
                                          <div class="tab-pane fade" id="profileDomande<%=k%><%=p%>">

                                     <%end if%>




                                   <!-- inizio blocco frasi che diventa domande-->




                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaDomande.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte le domande"  href="../cDomande/inserisci_valutazioni.asp?ID_MOD=<%=rsTabellaDomande("ID_Mod")%>&Paragrafo=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceAllievo=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&Capitolo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>">  Apri tutte le domande:&nbsp;
                                                    N(<%= numrsDomande &") Pt(" & numrsDomande2  & ") Pb("& round( numrsDomande2/numrsDomande,2) &")"%> </a>
                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                         <% if not rsTabellaDomande.eof then %>
                                                        <tr>
                                                            <th>Domanda</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
															 <th class='hidden-480'>Esposta</th>

															 <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("admin")=true) then%>
                                                            <th class='hidden-480'>Elimina</th></tr>
															<%end if%>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessuna compito inserito</th>

                                                       </tr>
                                                    <%end if%>


                                                    </thead>
                                                    <tbody>

                      <% Sottoparagrafo=""
					' p=0
					n=0

		     do while not rsTabellaDomande.EOF


			   if ((StrComp(Sottoparagrafo, rsTabellaDomande("SotPar")) <> 0) ) then
			   Sottoparagrafo=rsTabellaDomande("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaDomande("SotPar")%></b></center></td></th>

            <%end if%>



                                                        <tr>



                                                             <%if rsTabellaDomande("Segnalata")=1 then%>
                                                            <td > <a style="color:<%=color%>"  href="../cDomande/inserisci_valutazione.asp?Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&URL=<%=rsTabellaDomande("URL_Teoria")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>"><%=rsTabellaDomande("Quesito")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaDomande("Voto")%></td>
                                                             <%else%>
                                                              <td> <a   href="../cDomande/inserisci_valutazione.asp?Multiple=<%=rsTabellaDomande("Multiple")%>&ORA=<%=left(rsTabellaDomande("Ora"),5)%>&DATA=<%=rsTabellaDomande("Data")%>&Tipodomanda=<%=rsTabellaDomande("Tipo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&Quesito=<%=rsTabellaDomande("Quesito")%>&R1=<%=rsTabellaDomande("Risposta1")%> &R2=<%=rsTabellaDomande("Risposta2")%>&R3=<%=rsTabellaDomande("Risposta3")%>&R4=<%=rsTabellaDomande("Risposta4")%>&RE=<%=rsTabellaDomande("RispostaEsatta")%>&MO=<%=rsTabellaDomande("ID_Mod")%>&VAL=<%=rsTabellaDomande("Voto")%>&VF=<%=rsTabellaDomande("VF")%>&INQUIZ=<%=rsTabellaDomande("In_Quiz")%>&VALINQUIZ=<%=rsTabellaDomande("In_QuizStud")%>&Segnalata=<%=rsTabellaDomande("Segnalata")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>">  <%=rsTabellaDomande("Quesito")%></a></td>
                                                              <td><%=rsTabellaDomande("Voto")%></td>
                                                              <%end if%>





                                                            <td><%=rsTabellaDomande("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaDomande("Ora"),5)%> </td>
                                                            <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("admin")=true) then%>
                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare la domanda?');"  href="../cDomande/cancella_domanda.asp?Verifica=0&classe=<%=classe%>&cod=<%=rsTabellaDomande("CodiceAllievo")%>&Cartella=<%=rsTabellaDomande("Cartella")%>&Modulo=<%=rsTabellaDomande("ID_Mod")%>&CodiceTest=<%=rsTabellaDomande("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaDomande("CodiceDomanda")%>&Capitolo=<%=rsTabellaDomande("Tit")%>&Paragrafo=<%=rsTabellaDomande("Titolo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tDom=<%=k%><%=p%>" title="Cancella">
                                                            <i class="icon-trash" ></i></a>
                                                            </td>
															<%end if%>
                                                        </tr>


                 <% f=f+1
				  '  p=p+1
				  n=n+1
				    rsTabellaDomande.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div>
                                        </div>

                                  <!-- fine blocco frasi che diventa domande-->






                              </div>

                                <% if (numrsNodi<>0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                        <div class="tab-pane fade in active" id="profileNodi<%=k%><%=p%>">

                                     <%else%>
                                          <div class="tab-pane fade" id="profileNodi<%=k%><%=p%>">

                                     <%end if%>

                                  <!-- inizio blocco nodi -->



                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                   <i class="icon-table"></i>

                                                    <% if not rsTabellaNodi.eof then %>
                                                    <a style="color:#FFF" title="Apri tutte i nodi del paragrafo"  href="../cNodi/2inserisci_valutazioni_nodi.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaNodi("Data")%>&Tutte=1&ID_MOD=<%=rsTabellaNodi("ID_Mod")%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&TitoloParagrafo=<%=rsTabellaNodi("TitoloParagrafo")%>">
                                               Apri tutti i nodi: N(<%= numrsNodi2 &") Pt(" & numrsNodi2  & ") Pb("& round( numrsNodi2/numrsNodi,2) &")"%> </a>
											   - <a style="color:#FFF" title="Apri la mappa concettuale del paragrafo" href="../cNodi/spiegazione_nodi.asp?Cartella=<%=rsTabellaNodi("Cartella")%>&Stato=0&Stato0=0&CodiceTest=<%=rsTabellaNodi("Id_Paragrafo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>&CodiceAllievo=<%=rsTabellaNodi("CodiceAllievo")%>&daQuaderno=1">Apri Mappa</a>

                                                    <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div>
                                            <div class="box-content nopadding">
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                           <% if not rsTabellaNodi.eof then %>
                                                        <tr>
                                                            <th>Nodi</th>
                                                            <th>Punti</th>
                                                            <th>Data</th>
                                                            <th class='hidden-480'>Ora</th>
                                                            <th class='hidden-480'>Risposto</th>
															  <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("admin")=true) then%>
                                                            <th class='hidden-480'>Elimina</th></tr>
															<%end if%>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessun compito inserito</th>

                                                        </tr>
                                                    <%end if%>

                                                        </tr>
                                                    </thead>
                                                    <tbody>



                     <% Sottoparagrafo=""
					' p=0



		     do while not rsTabellaNodi.EOF
			   if StrComp(Sottoparagrafo, rsTabellaNodi("SotPar")) <> 0 then
			   Sottoparagrafo=rsTabellaNodi("SotPar")
                %>
                <th><td colspan="6"><center><b><%=rsTabellaNodi("SotPar")%></b></center></td></th>
			 <%end if%>

                                                        <tr>


                                                             <%if rsTabellaNodi("Segnalata")=1 then%>
                                                   <td><a  style="color:red" title="Apri il nodo"  href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>
                                                             <td style="color:#F00"><%=rsTabellaNodi("Voto")%></td>

                                                             <%else%>


                                                             <td><a title="Apri il nodo"   href="../cNodi/inserisci_valutazione_nodi.asp?DATA=<%=rsTabellaNodi("Data")%>&Ora=<%=left(rsTabellaNodi("Ora"),5)%>&Cartella=<%=rsTabellaNodi("Cartella")%>&cla=<%=d%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaNodi("ID_paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&Chi=<%=rsTabellaNodi("Chi")%>&Cosa=<%=rsTabellaNodi("Cosa")%> &Dove=<%=rsTabellaNodi("Dove")%>&Quando=<%=rsTabellaNodi("Quando")%>&Come=<%=rsTabellaNodi("Come")%>&Perche=<%=rsTabellaNodi("Perche")%>&Quindi=<%=rsTabellaNodi("Quindi")%>&MO=<%=rsTabellaNodi("ID_Mod")%>&VAL=<%=rsTabellaNodi("Voto")%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>"><%=rsTabellaNodi("Chi")%></a></td>

                                                             <td><%=rsTabellaNodi("Voto")%></td>

                                                              <%end if%>


                                                            <td><%=rsTabellaNodi("Data")%> </td>
                                                             <td  class='hidden-480'><%=left(rsTabellaNodi("Ora"),5)%> </td>

                                                            <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
												 <%if (strcomp(cod,Session("CodiceAllievo"))=0) or (session("admin")= true) then%>
                                                            <td class='hidden-480'>
                                                            <a onClick="return window.confirm('Vuoi veramente cancellare il nodo?');"  href="../cNodi/cancella_nodo.asp?cla=<%=d%>&cod=<%=rsTabellaNodi("CodiceAllievo")%>&Cartella=<%=rsTabellaNodi("Cartella")%>&Modulo=<%=rsTabellaNodi("ID_Mod")%>&CodiceTest=<%=rsTabellaNodi("ID_Paragrafo")%>&CodiceDomanda=<%=rsTabellaNodi("CodiceNodo")%>&Capitolo=<%=rsTabellaNodi("Titolo")%>&Paragrafo=<%=rsTabellaNodi("TitoloParagrafo")%>&id_classe=<%=id_classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&tCap=<%=k%>&tSot=<%=k%><%=p%>&tNod=<%=k%><%=p%>">
                                                            <i class=" icon-trash" ></i></a>
                                                            </td><%end if%>
                                                        </tr>

                 <% f=f+1
				    rsTabellaNodi.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div>
                                        </div>

                                  <!-- fine blocco frasi che diventa domande-->  
                                  </div> <!-- fine profile nodi-->







                                  
                                  
                                  
                                  
                                  
                                   
                                <% if (numrsMetafore<>0 ) and  (numrsNodi=0 ) and (numrsFrasi=0) and (numrsDomande=0) then %>
                                        <div class="tab-pane fade in active" id="profileMetafore<%=k%><%=p%>">
                              
                                     <%else%>
                                          <div class="tab-pane fade" id="profileMetafore<%=k%><%=p%>">
                              
                                     <%end if%>
                              
                                  <!-- inizio blocco nodi -->  
                                  
                               
                                  
                                   <div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    <% if not rsTabellaMetafore.eof then 
														rsTabellaMetafore.movefirst
													%>
                                                    
                                                     <%=" Nn(" & numrsMetafore &") Pt(" & numrsMetafore2 &  ") ; Pb("& round( numrsMetafore2/numrsMetafore,2) &")"%>
                                                       <%else%>
                                                    Punti (0)
                                                    <%end if%>
                                                </h3>
                                            </div> 
                                            <div class="box-content nopadding"> 
                                              <table class="table table-hover table-nomargin">
                                                    <thead>
                                                        <tr>
                                                           <% if not rsTabellaMetafore.eof then %>
                                                        <tr>
                                                             <th>Percorso</th>
                                                            <th>Metafora</th>
                                                          
                                                            <th>Data</th>
                                                            <th>Ora</th>
                                                            <th class='hidden-480'>Punti</th>
                                                            <th class='hidden-480'>Esposto</th>
                                                            <th class='hidden-480'>Elimina</th                                                          
                                                        ></tr>
                                                         <%else%>
                                                     <tr>
                                                            <th colspan="6">Nessun compito inserito</th>
                                                                                                                
                                                        </tr>
                                                    <%end if%>
                                                           
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                  
                     <% Sottoparagrafo=""
					' p=0
					
					
				
								
					
		     do while not rsTabellaMetafore.EOF  %>
             
             
              <% select case rsTabellaParagrafi("Paragrafo")
							 
							 case "Topolino ed Obiettivi" %>
							 
  <%
     QuerySQL2="SELECT SUM(Voto) AS Pt FROM Elenco_Metafore_Topolino" &_
	 " where  ThreadParent ="& rsTabellaMetafore.fields("CodiceMetafora") & ";"  
	 'conta il numero di punti ottenuti nelle metafore topolino
	' response.write(QuerySQL2)
	 Set rsTabellaPuntiPercorso = ConnessioneDB.Execute(QuerySQL2)
	 ' numrsTabella2=rsTabella2(0)
	    puntiPercorso=rsTabellaPuntiPercorso(0)
	 ' se non restituisce nulla serve per dargli un valore
	 if rsTabellaPuntiPercorso(0)&"" =""  then
	   puntiPercorso=0
	 end if 
  %>                             
 
									<tr>
                                    <td><a target="_blank"  href="../cMetafore/sintesi_metafore.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&CodiceAllievo=<%=cod%>"><i class="icon-play-circle"></i></a></td> 
                                    
                                    <td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&Topolino=<%=rsTabellaMetafore("Topolino")%>&Formaggio=<%=rsTabellaMetafore("Formaggio")%> &Fame=<%=rsTabellaMetafore("Fame")%>&Labirinto=<%=rsTabellaMetafore("Labirinto")%>&Strada=<%=rsTabellaMetafore("Strada")%>&Strada_OK=<%=rsTabellaMetafore("Strada_OK")%>&Strada_KO=<%=rsTabellaMetafore("Strada_KO")%>&Distanza=<%=rsTabellaMetafore.fields("Distanza")%>&Testata=<%=rsTabellaMetafore.fields("Testata")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
									
									<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">
						<%=rsTabellaMetafore.fields("Topolino")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=puntiPercorso%></td> 
                        </font>
						 <%	else %>
                         
                            <%=rsTabellaMetafore.fields("Topolino")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=puntiPercorso%></td>  
					 <% end if %>	 
						
  <td class='hidden-480'><input name="checkbox" type="checkbox"> </td><td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore("Cartella")%>&Modulo=<%=rsTabellaMetafore("Id_Mod")%>&CodiceTest=<%=rsTabellaMetafore("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore("Titolo")%>&Paragrafo=<%=rsTabellaMetafore("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>
</tr>	
							 
							 
							 
							<% case "Navigazione nella Rete della Vita" %>
							<tr>
                              <td><a  target="_blank"   href="../cMetafore/sintesi_metafore.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&classe=<%=classe%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&CodiceAllievo=<%=cod%>"><i class="icon-play-circle"></i></a></td> 
                                    	 
							
							 
										
										<td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?id_classe=<%=id_classe%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&Autista=<%=rsTabellaMetafore(4)%>&Destinazione=<%=rsTabellaMetafore(5)%> &Carburante=<%=rsTabellaMetafore(6)%>&Luogo=<%=rsTabellaMetafore(7)%>&Strada=<%=rsTabellaMetafore(8)%>&Strada_OK=<%=rsTabellaMetafore(9)%>&Strada_KO=<%=rsTabellaMetafore(10)%>&Cespugli=<%=rsTabellaMetafore.fields("Cespugli")%>&Cestino=<%=rsTabellaMetafore.fields("Cestino")%>&Lupo=<%=rsTabellaMetafore.fields("Lupo")%>&Distanza=<%=rsTabellaMetafore.fields("Distanza")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
	<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabellaMetafore.fields("Autista")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabellaMetafore.fields("Autista")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td>
    <%end if%>
      <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
							
							
							<% case "Relazione Cliente Servitore" %>
							<td><i class="icon-play-circle"></i></td>
							
                            
                            
										
										<td><a  href="../cMetafore/inserisci_valutazione_metafore.asp?Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&id_classe=<%=id_classe%>&ThreadParent=<%=rsTabellaMetafore.fields("ThreadParent")%>&DATA=<%=rsTabellaMetafore.fields("Data")%>&classe=<%=classe%>&cod=<%=cod%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&CodiceAllievo=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&TitoloParagrafo=<%=rsTabellaMetafore.fields("Tit")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&SoggettoC=<%=rsTabellaMetafore("SoggettoC")%>&DomandaC=<%=rsTabellaMetafore("DomandaC")%>&MotivazioneC=<%=rsTabellaMetafore("MotivazioneC")%>&DesiderioC=<%=rsTabellaMetafore("DesiderioC")%>&BisognoC=<%=rsTabellaMetafore("BisognoC")%>&SoggettoS=<%=rsTabellaMetafore("SoggettoS")%>&RispostaS=<%=rsTabellaMetafore("RispostaS")%>&MotivazioneS=<%=rsTabellaMetafore("MotivazioneS")%>&DesiderioS=<%=rsTabellaMetafore.fields("DesiderioS")%>&BisognoS=<%=rsTabellaMetafore("BisognoS")%>&TipoEvento=<%=rsTabellaMetafore(14)%>&TolleranzaC=<%=rsTabellaMetafore.fields("TolleranzaC")%>&MO=<%=rsTabellaMetafore.fields("ID_Mod")%>&VAL=<%=rsTabellaMetafore.fields("Voto")%>&Segnalata=<%=rsTabellaMetafore.fields("Segnalata")%>&Pippo=1 ">
	<%	if rsTabellaMetafore.fields("Segnalata")=1 then%>
                         <font color="#FF0000">									
<%=rsTabellaMetafore.fields("SoggettoC")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td> </font>
    <%else%>
        <%=rsTabellaMetafore.fields("SoggettoC")%></a></td><td><%=rsTabellaMetafore.fields("Data")%></td><td><%=left(rsTabellaMetafore.fields("Ora"),5)%></td><td><%=rsTabellaMetafore.fields("Voto")%></td>
    <%end if%>
      <td class='hidden-480'>
												<input name="checkbox" type="checkbox"> </td>
<td><a onClick="return window.confirm('Vuoi veramente cancellare la metafora?');"  href="../cMetafore/cancella_metafora.asp?cla=<%=d%>&cod=<%=rsTabellaMetafore.fields("CodiceAllievo")%>&Cartella=<%=rsTabellaMetafore.fields("Cartella")%>&Modulo=<%=rsTabellaMetafore.fields("ID_Mod")%>&CodiceTest=<%=rsTabellaMetafore.fields("ID_Paragrafo")%>&CodiceMetafora=<%=rsTabellaMetafore.fields("CodiceMetafora")%>&Capitolo=<%=rsTabellaMetafore.fields("Titolo")%>&Paragrafo=<%=rsTabellaMetafore.fields("Tit")%>&id_classe=<%=id_classe%>"><img src="../../img/elimina_small.jpg"></a>
</td>

</tr>			  
							
							   
							<% end select  %>
							 
			 
                                                    
                                                       
                                                     
                 <% f=f+1
				    rsTabellaMetafore.movenext()
				 loop%>
                                                    </tbody>
                                                </table>
                                             </div> 
                                        </div>  
                                        
                                  <!-- fine blocco frasi che diventa domande-->                               
                                  </div> 














                            </div><!-- fine MyTabContent2-->

                          </div><!-- fine collapse(treuno)-->
                        </div> <!-- fine accordino group-- da Descrizione capitolo in giù >-->
                         <%end if %> <!--if (numrsFrasi<>0) or (numrsDomande<>0) or (numrsNodi<>0) then-->





                         <% p=p+1
						   rsTabellaParagrafi.movenext()
						   Loop
						%>




                  

        <!--lo tolgo e lo aggiungo nel chiamante        </div>--> <!--  fine accordion group uno per ogni capitolo-->
       <%'compiti=compiti+1  %>