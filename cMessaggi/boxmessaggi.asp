<div class="box box-bordered box-color">
		  
		  <!-- #include file = "../cClasse/studente_domande_include/2_chat_2.asp" -->
  <!-- #include file = "../cClasse/studente_domande_include/2_messaggi_2.asp" -->
		  
            <div class="box-title">
              <h3> <i class="icon-envelope"></i> Centro Messaggi </h3>
            </div>
            <div class="box-content nopadding">
              <ul class="tabs tabs-inline tabs-left" id="menumessaggi">
                <li class='write hidden-480'><a href="nuovomessaggio.asp" rel="tooltip" data-placement="bottom" title="Inizia nuova chat">(+) Chat</a></li>
                <%if session("admin")=true then%>
                <li class='write hidden-480'><a href="nuovomessaggio_mail.asp" rel="tooltip" data-placement="bottom" title="Invia email">(+) Messaggio</a></li>
                <%end if%>
                <li class='active'> <a href="#notifichemessaggi" data-toggle="tab"><i class="icon-inbox"></i> Notifiche <strong>(<%=numMessaggi%>)</strong></a> </li>
                <li> <a href="#lavagnamessaggi" data-toggle="tab"><i class="icon-bullhorn"></i> Lavagna</a> </li>
                <li> <a href="#forummessaggi" data-toggle="tab"><i class="icon-group"></i> Forum</a> </li>
                <li> <a href="#diariomessaggi" data-toggle="tab"><i class="icon-book"></i> Diario</a> </li>
                <li> <a href="#sentmessaggi" data-toggle="tab"><i class="icon-comments-alt"></i> Messaggi (<%=numMessaggiChat%>)</a> </li>
                <li> <a href="#archiviomessaggi" data-toggle="tab"><i class="icon-inbox"></i> Lette (<%=numMessaggiArchivio%>)</a> </li>
              </ul>
			  
			   <% if session("fraseinesistente") = true then %>
					<script>alert("Errore interno: frase inesistente")</script>
					<% session("fraseinesistente") = false %>
				<% end if %>	
			  
              <div class="tab-content tab-content-inline" >
                <div class="tab-pane active" id="notifichemessaggi" style="min-height:307px"> <!-- metto min-height per non fare l'effetto di riadeguazione dei px -->
                  
				  <% If rsTabellaAvvisiP.BOF=True And rsTabellaAvvisiP.EOF=True then %>
				  
				  <% else %>
				  
				  
				  <div class="highlight-toolbar">
                    <div class="pull-left">
                      <div class="btn-toolbar"> 
                    </div>
					</div>
                    <div class="pull-right">
                      <div class="btn-toolbar"> </div>
                      <div class="pull-right">
                        <div class="btn-toolbar">
                          <div class="btn-group text hidden-768"> <span> <strong> Notifiche</strong></span> </div>
                          <div class="btn-group hidden-768">
                            <div class="dropdown"> <a href="#" class="btn" data-toggle="dropdown"><i class="icon-cog"></i><span class="caret"></span></a>
                              <ul class="dropdown-menu pull-right">
                                <li><a href="cambiastato.asp?Tutte=1&Leggi=1&cod=<%=session("CodiceAllievo")%>">Leggi tutte</a></li>
                              </ul>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
				  
				  
                  <% end if %>
                  
				  <% If rsTabellaAvvisiP.BOF=True And rsTabellaAvvisiP.EOF=True then %>
                      <div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono notifiche da leggere</span></center>
				  
				  <% else %>
				  
                  <table class="table table-striped table-nomargin table-mail">
                    <thead>
                      <tr>
                        <th class='table-checkbox hidden-480'>  </th>
                        <th>Mittente</th>
                        <th>Testo</th>
                        <th>Oggetto</th>
                        <th class='table-date hidden-480'> Data <i class="icon-calendar"></i></th>
                      </tr>
                    </thead>
                    <tbody>
                      
                      
                      <% 
						 k=0
						 do while not rsTabellaAvvisiP.EOF
                     QuerySQL2="SELECT Cognome,Nome,Url_img,Classe FROM Allievi WHERE CodiceAllievo='"&rsTabellaAvvisiP("CodiceAllievo2")&"'"
							'response.write(QuerySQL2)
							Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
							Cognome2=rsTabella2("Cognome")
							Nome2=rsTabella2("Nome")
							'Url_img2=rsTabella2("Url_img")
							   ' if strcomp(Url_img2&"","")=0 then 
								' urlimmagine="../img/no-avatar.jpg"
									
								 ' else
								   ' if rsTabellaAvvisiP("CodiceAllievo2") = Session("CodAdmin") then
									  ' url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/Admin/Profili/thumb"
								   ' else
									   ' url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&rsTabella2("Classe")&"/Profili/thumb" ' vuole il percorso relativo della cartella
								   ' end if
								  ' url=Replace(url,"\","/")
								  ' urlimmagine=url&"/"& Url_img2 
			 
							  ' end if
							  %>
                      <tr class='unread'>
                        <td class='table-checkbox hidden-480'><!--<input type="checkbox" name="cbArchivia<%=k%>"  >--> 
                          <a href="cambiastato.asp?IdNotifica=<%=rsTabellaAvvisiP("ID_Avviso")%>&Leggi=1" class='btn' rel="tooltip" data-placement="bottom" title="Archivia"><i class="icon-inbox"></i></a></td>
                        <td class='table-fixed-medium'>&nbsp;<%=trim(Cognome2)%>&nbsp; <%=left(Nome2,1)&"."%></td>
                        <td class='table-fixed-medium'><% if isNull(rsTabellaAvvisiP("Testo")) then%><i>Per leggere il contenuto del messaggio, aprire la notifica</i><%else%><%=server.htmlencode(rsTabellaAvvisiP("Testo"))%><%end if%></td>
						
						<% action = rsTabellaAvvisiP("Azione")
						vett = split(action, ">")
						vett(1) = Replace(vett(1), " !", "!")
						
						%>
						
                        <td class='table-fixed-medium'><a href="legginotifica.asp?IdNotifica=<%=rsTabellaAvvisiP("ID_Avviso")%>&action=<%=action%>"><%=vett(1)&"<>"%></a></td>
                        <td class='table-date hidden-480'><%=left(rsTabellaAvvisiP("Data"), 10)%></td>
                      </tr>
                      <%
							  k=k+1
							  rsTabellaAvvisiP.movenext%>
                      <% loop %>
                      <%end if%>
                    </tbody>
                  </table>
                  <!--     </form>--> 
                </div>
                
                <!-- #include file = "../cClasse/studente_domande_include/2_lavagna_2.asp" -->
                
                <div class="tab-pane" id="lavagnamessaggi" style="min-height:307px">
                  
                  
                    <%If rsTabellaLavagna.BOF=True And rsTabellaLavagna.EOF=True Then %>
                     <div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono attività nella Bacheca</span></center>
                    <% Else%>
                    <!--<table class="table table-hover table-nomargin dataTable table-bordered">-->
                    <table class="table table-striped table-nomargin table-mail">
                    <thead>
                      <tr>
                        <th colspan="5" style="background-color:white"> Discussioni aperte (<%=num_post_totali%>) + Commenti (<%=num_messaggi%>) = Punti (<%=num_post_totali_punti+num_messaggi_punti%>) </th>
                      <tr>
                        <th class='table-checkbox hidden-480'> <!--	<input type="checkbox" class='sel-all' rel="tooltip" title="Seleziona tutti">--> </th>
                        <th>Post</th>
                        <th>Messaggio</th>
                        <th class='hidden-480'>Data</th>
                        <th class='hidden-480'>Punti</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   i=0
		   do while not rsTabellaLavagna.EOF  'and i<10
		   i=i+1%>
                      <tr class='unread'>
                        <td class='table-checkbox hidden-480'><center><% if rsTabellaLavagna("ParentMessage") <> 0 then response.write "C" else response.write "D" end if %></center><!--<input type="checkbox" class='selectable'>--></td>
                        <%
		     QuerySQL1="SELECT * "&_
" FROM FORUM_MESSAGES " &_
" WHERE ID=" & rsTabellaLavagna("ThreadParent") &" and comments<>'InizializzaDB'" &_
 " ORDER BY ID desc;"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

Set rsTabella1 = ConnessioneDB2.Execute(QuerySQL1) 
	'num_visualizzazioni=rsTabella1(0) 
	
	'discussione principale=categoria=Programma&id_categoria=101&nome=&cognome=&scegli=1&bacheca=&ID=3094&Zip=0&RCount=6&TParent=3094&divid=&id_classe=14COM&visibile=1&privato=0
	'messaggio discussione=scegli=1&bacheca=&Rispondi=1&ID=4321&Zip=0&CodiceAllievo=federico.ballarini&divid=&id_classe=14COM&RCount=6&categoria=Programma&id_categoria=101
	
	'aggiunto id e nome categoria!! così non perde più parametri!!
	
	QueryCat = "SELECT Descrizione FROM CAT_CAT WHERE Id_Categoria = '"&rsTabellaLavagna("Id_Categoria")&"';"
		Set rsTabellaCat = ConnessioneDB.Execute(QueryCat)
		
		categorianome = rsTabellaCat(0)
	
			 %>
                        <% if not rsTabella1.eof then%>
                        <td><a title="Visualizza Post di apertura discussione" href="../cSocial/ShowMessage.asp?scegli=1&ID=<%=rsTabellaLavagna("ThreadParent")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaLavagna("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabella1("Topic")%></a></td>
                        <td><a title="Visualizza il messaggio nella discussione"    href="../cSocial/ShowMessage.asp?scegli=1&ID=<%=rsTabellaLavagna("ID")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaLavagna("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabellaLavagna("Topic")%></a></td>
                        <td class='hidden-480'><%=left(rsTabellaLavagna("DatePosted"), 10)%></td>
                        <td class='hidden-480'><center><%=rsTabellaLavagna("Punti")%></center></td>
                        <!--   <td class='hidden-480'><a onClick="return window.confirm('Vuoi veramente cancellare il messaggio ?');" target="_new" href="../cancella_messaggio.asp?ID=<%=rsTabellaLavagna("ID")%>" title="Cancella"><i class=" icon-trash" ></i></a></td>--> 
                      </tr>
                      <%
		  end if
		 rsTabellaLavagna.movenext
		loop%>
                    </tbody>
                    <%end if%>
                  </table>
                </div>
                
                <!-- #include file = "../cClasse/studente_domande_include/2_forum_2.asp" -->
                
				<div class="tab-pane" id="forummessaggi" style="min-height:307px">
				
				<%If rsTabellaForum.BOF=True And rsTabellaForum.EOF=True Then %>
                    <div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono attività nel Forum</span></center>
				
				<% Else%>
				
                
                  <table class="table table-striped table-nomargin table-mail">
                    
                    
                    <!--<table class="table table-hover table-nomargin dataTable table-bordered">-->
                    
                    <thead>
                      <tr>
                        <th colspan="5" style="background-color:white"> Discussioni aperte (<%=num_post_totali%>) + Commenti (<%=num_messaggi%>) = Punti (<%=num_post_totali_punti+num_messaggi_punti%>) </th>
                      <tr>
                        <th class='table-checkbox hidden-480'> <!--	<input type="checkbox" class='sel-all' rel="tooltip" title="Seleziona tutti">--> </th>
                        <th>Post</th>
                        <th>Messaggio</th>
                        <th class='hidden-480'>Data</th>
                        <th class='hidden-480'>Punti</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   i=0
		   do while not rsTabellaForum.EOF  'and i<10
		   i=i+1%>
                      <tr class='unread'>
                        <td class='table-checkbox hidden-480'><center><% if rsTabellaForum("ParentMessage") <> 0 then response.write("C") else response.write("D") end if %></center><!--<input type="checkbox" class='selectable'>--></td>
                        <%
		     QuerySQL1="SELECT * "&_
" FROM FORUM_MESSAGES " &_
" WHERE ID=" & rsTabellaForum("ThreadParent") &" and comments<>'InizializzaDB'" &_
 
" ORDER BY ID desc;"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1) 
	'num_visualizzazioni=rsTabella1(0) 
	
		
		QueryCat = "SELECT Descrizione FROM CAT_CAT WHERE Id_Categoria = '"&rsTabellaForum("Id_Categoria")&"';"
		Set rsTabellaCat = ConnessioneDB.Execute(QueryCat)
		
		categorianome = rsTabellaCat(0)
	
			 %>
                        <% if not rsTabella1.eof then%>
                        <td><a title="Visualizza Post di apertura discussione" href="../cSocial/ShowMessage.asp?scegli=0&ID=<%=rsTabellaForum("ThreadParent")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaForum("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabella1("Topic")%></a></td>
                        <td><a title="Visualizza il messaggio nella discussione"    href="../cSocial/ShowMessage.asp?scegli=0&ID=<%=rsTabellaForum("ID")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaForum("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabellaForum("Topic")%></a></td>
                        <td class='hidden-480'><%=left(rsTabellaForum("DatePosted"), 10)%></td>
                        <td class='hidden-480'><center><%=rsTabellaForum("Punti")%></center></td>
                        <!--   <td class='hidden-480'><a onClick="return window.confirm('Vuoi veramente cancellare il messaggio ?');" target="_new" href="../cancella_messaggio.asp?ID=<%=rsTabellaForum("ID")%>" title="Cancella"><i class=" icon-trash" ></i></a></td>--> 
                      </tr>
                      <%
		  end if
		 rsTabellaForum.movenext
		loop%>
                    </tbody>
                    <%end if%>
                  </table>
                </div>
                
                <!-- #include file = "../cClasse/studente_domande_include/2_diario_2.asp" -->
                
                <div class="tab-pane" id="diariomessaggi" style="min-height:307px">
                 
                 
                    <%If rsTabellaDiario.BOF=True And rsTabellaDiario.EOF=True Then %>
                    
					<div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono attività nel Diario</span></center>
					
					
                    <% Else%>
					
					 <table class="table table-striped table-nomargin table-mail">
                    <!--<table class="table table-hover table-nomargin dataTable table-bordered">-->
                    
                    <thead>
                      <tr>
                        <th colspan="5" style="background-color:white"> Discussioni aperte (<%=num_post_totali%>) + Commenti (<%=num_messaggi%>) = Punti (<%=num_post_totali_punti+num_messaggi_punti%>) </th>
                      <tr>
                        <th class='table-checkbox hidden-480'> <!--<input type="checkbox" class='sel-all' rel="tooltip" title="Seleziona tutti">--> </th>
                        <th>Post</th>
                        <th>Messaggio</th>
                        <th class='hidden-480'>Data</th>
                        <th class='hidden-480'>Punti</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   i=0
		   do while not rsTabellaDiario.EOF  'and i<10
		   i=i+1%>
                      <tr >
                        <td class='table-checkbox hidden-480'><center><% if rsTabellaDiario("ParentMessage") <> 0 then response.write("C") else response.write("D") end if %></center><!-- <input type="checkbox" class='selectable'>--></td>
                        <%
		     QuerySQL1="SELECT * "&_
" FROM FORUM_MESSAGES " &_
" WHERE ID=" & rsTabellaDiario("ThreadParent") &" and comments<>'InizializzaDB'" &_
 
" ORDER BY ID desc;"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	'num_visualizzazioni=rsTabella1(0) 
	
	QueryCat = "SELECT Descrizione FROM CAT_CAT WHERE Id_Categoria = '"&rsTabellaDiario("Id_Categoria")&"';"
		Set rsTabellaCat = ConnessioneDB.Execute(QueryCat)
		
		categorianome = rsTabellaCat(0)
	
			 %>
                        <% if not rsTabella1.eof then%>
                        <td><a title="Visualizza Post di apertura discussione" href="../cSocial/ShowMessage.asp?scegli=2&ID=<%=rsTabellaDiario("ThreadParent")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaDiario("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabella1("Topic")%></a></td>
                        <td><a title="Visualizza il messaggio nella discussione"    href="../cSocial/ShowMessage.asp?scegli=2&ID=<%=rsTabellaDiario("ID")%>&id_classe=<%=id_classe%>&divid=<%=divid2%>&id_categoria=<%=rsTabellaDiario("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabellaDiario("Topic")%></a></td>
                        <td class='hidden-480'><%=left(rsTabellaDiario("DatePosted"), 10)%></td>
                        <td class='hidden-480'><center><%=rsTabellaDiario("Punti")%></center></td>
                        <!--   <td class='hidden-480'><a onClick="return window.confirm('Vuoi veramente cancellare il messaggio ?');" target="_new" href="../cancella_messaggio.asp?ID=<%=rsTabellaDiario("ID")%>" title="Cancella"><i class=" icon-trash" ></i></a></td>--> 
                      </tr>
                      <%
		  end if
		 rsTabellaDiario.movenext
		loop%>
                    </tbody>
                    <%end if%>
                  </table>
                </div>
                 



				 
				 <div class="tab-pane" id="sentmessaggi" style="min-height:307px">
                  
				  <% If rsTabellaContattiChat.BOF=True And rsTabellaContattiChat.EOF=True then %>
				  
				  <% else %>
				  
				  
				  <!--<div class="highlight-toolbar">
                    <div class="pull-left">
                      <div class="btn-toolbar"> 
                    </div>
					</div>
                    <div class="pull-right">
                      <div class="btn-toolbar"> </div>
                      <div class="pull-right">
                        <div class="btn-toolbar">
                          <div class="btn-group text hidden-768"> <span> <strong> Chat</strong></span> </div>
                          <div class="btn-group hidden-768">
                            <div class="dropdown"> <a href="#" class="btn" data-toggle="dropdown"><i class="icon-cog"></i><span class="caret"></span></a>
                              <ul class="dropdown-menu pull-right">
                                <li><a href="cambiastatochat.asp?Tutte=1&Leggi=1&cod=<%=session("CodiceAllievo")%>">Segna tutte come già lette</a></li>
								<li><a href="cambiastatochat.asp?Tutte=1&Rimuovi=1&cod=<%=session("CodiceAllievo")%>">Cancella tutte le chat</a></li>
                              </ul>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>-->
				  
				  
                  <% end if %>
                  
				  <% If rsTabellaContattiChat.BOF=True And rsTabellaContattiChat.EOF=True then %>
                      <div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono chat da leggere</span></center>
				  
				  <% else %>
				  
                  <table class="table table-striped table-nomargin table-mail">
                    <thead>
                      <tr>
                        <th class='table-checkbox hidden-480'>  </th>
                        <th>Contatto</th>
                        <th>Ultimo Messaggio</th>
                        <th class='table-date hidden-480'> Data <i class="icon-calendar"></i></th>
                      </tr>
                    </thead>
                    <tbody>
                      
                      
                      <% 
						 do while not rsTabellaContattiChat.EOF
							
							if LCase(rsTabellaContattiChat("CodiceAllievo")) = LCase(cod) then
								colonna = "CodiceAllievo2"
								tipo = "Ricevuto"
							else
								colonna = "CodiceAllievo"
								tipo = "Inviato"
							end if
							
							'response.write("Stringa elenco: "&stringaelenco&"<br>")
							elenco = split(stringaelenco, "$")
							trovato = false
							'response.write(rsTabellaContattiChat("CodiceAllievo2")&"<br>")
							
							for i = 0 to Ubound(elenco) and trovato = false
								
								'response.write("<br>"&elenco(i)&"<br>"&rsTabellaContattiChat(tabella))
								'response.write(StrComp(elenco(i), rsTabellaContattiChat(tabella)))
								
								if StrComp(LCase(elenco(i)), LCase(rsTabellaContattiChat(colonna))) = 0 then
									trovato = true
								end if
																
							next
							
							'response.write(trovato&"<br>")
							
							if trovato = false then
							
							if stringaelenco <> "" then 
								stringaelenco = stringaelenco&"$"&rsTabellaContattiChat(colonna)
							else
								stringaelenco = rsTabellaContattiChat(colonna)
							end if
							
							QuerySQL2="SELECT Cognome,Nome,Url_img,Classe FROM Allievi WHERE CodiceAllievo='"&rsTabellaContattiChat(colonna)&"'"
							
							Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
							Cognome2=rsTabella2("Cognome")
							Nome2=rsTabella2("Nome")
							
							if LCase(rsTabellaContattiChat("CodiceAllievo")) = LCase(cod) then
							'ricevuto
								QuerySQLMessaggi = "SELECT * FROM AVVISI WHERE CodiceAllievo='"&LCase(cod)&"' AND CodiceAllievo2='"&LCase(rsTabellaContattiChat("CodiceAllievo2"))&"' AND CAST(Testo AS ntext) like Azione order by Data desc;"
							else
							'inviato
								QuerySQLMessaggi = "SELECT * FROM AVVISI WHERE CodiceAllievo='"&LCase(rsTabellaContattiChat("CodiceAllievo"))&"' AND CodiceAllievo2='"&LCase(cod)&"' AND CAST(Testo AS ntext) like Azione order by Data desc;"
							end if
							'response.write(QuerySQLMessaggi)
							
							Set rsTabellaMessaggi = ConnessioneDB.Execute(QuerySQLMessaggi)
							TestoUltimo = rsTabellaMessaggi("Testo")
							DataUltimo = rsTabellaMessaggi("Data")
							
							
						
							  %>
                      <tr class='unread'>
                        <td class='table-checkbox hidden-480'><!--<input type="checkbox" name="cbArchivia<%=k%>"  >--> 
                          <a target="_blank" href="leggichat.asp?cod=<%=cod%>&contatto=<% if LCase(rsTabellaContattiChat("CodiceAllievo")) = LCase(cod) then%> <%=rsTabellaContattiChat("CodiceAllievo2")%><% else %><%=rsTabellaContattiChat("CodiceAllievo")%><% end if %>" class='btn' rel="tooltip" data-placement="bottom" title="Apri Chat"><i class="icon-book"></i></a></td>
						<td class='table-fixed-medium'>&nbsp;<%=trim(Cognome2)%>&nbsp; <%=left(Nome2,1)&"."%></td>
                        <td class='table-fixed-medium'><a style="text-decoration:none"><i>
						<%if tipo = "Inviato" then 
							if rsTabellaMessaggi("Visto") = 0 then %>
								<%=tipo%>
							<% else response.write("Letto")
							end if 
						else%>
							<%=tipo%>
						<%end if%></i></a>&nbsp;&nbsp;
						<% if len(TestoUltimo) > 35 then
						response.write(left(TestoUltimo, 35)&"...")
						else
						%><%=TestoUltimo%>
						<%end if%></td>
						<td class='table-date hidden-480'><%=DataUltimo%></td>
                      </tr>
                      <%
							  end if
							  rsTabellaContattiChat.movenext%>
                      <% loop %>
                      <%end if%>
                    </tbody>
                  </table>
                  <!--     </form>--> 
                </div>
				 
				 
				 
				 




				 
                <div class="tab-pane" id="archiviomessaggi" style="min-height:307px">
				
				<% If rsTabellaAvvisiLetti.BOF=True And rsTabellaAvvisiLetti.EOF=True then %>
				  
				  <% else %>
				
                  <div class="highlight-toolbar">
                    <div class="pull-left">
                      <div class="btn-toolbar"> 
                    </div>
					</div>
                    <div class="pull-right">
                      <div class="btn-toolbar"> </div>
                      <div class="pull-right">
                        <div class="btn-toolbar">
                          <div class="btn-group text hidden-768"> <span> <strong> Notifiche</strong></span> </div>
                          <div class="btn-group hidden-768">
                            <div class="dropdown"> <a href="#" class="btn" data-toggle="dropdown"><i class="icon-cog"></i><span class="caret"></span></a>
                              <ul class="dropdown-menu pull-right">
								<li><a href="cambiastato.asp?Lette=1&Ripristina=1&cod=<%=session("CodiceAllievo")%>">Ripristina Notifiche Lette</a></li>
                                <li><a href="cambiastato.asp?Lette=1&Rimuovi=1&cod=<%=session("CodiceAllievo")%>">Cancella Notifiche Lette</a></li>
								<li><a href="cambiastato.asp?Tutte=1&Rimuovi=1&cod=<%=session("CodiceAllievo")%>">Cancella Tutte</a></li>
                              </ul>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
				  
				  
				  <% end if %>
                  
				  <% If rsTabellaAvvisiLetti.BOF=True And rsTabellaAvvisiLetti.EOF=True then %>
                      <div style="height:153px"></div>
					  <center><span class="alert-error">Non ci sono notifiche lette</span></center>
				  
				  <% else %>
				  
				  
                  <table class="table table-striped table-nomargin table-mail">
                    <thead>
                      <tr>
                        <th class='table-checkbox hidden-480'> </th>
						<th class='table-checkbox hidden-480'> </th>
                        <th>Mittente</th>
                        <th>Testo</th>
                        <th>Oggetto</th>
                        <th class='table-date hidden-480'> Data <i class="icon-calendar"></i></th>
                      </tr>
                    </thead>
					
					
					
                    <tbody>
                      
					  <% 
						 k=0
						 do while not rsTabellaAvvisiLetti.EOF %>
                      <%QuerySQL2="SELECT Cognome,Nome,Url_img FROM Allievi WHERE CodiceAllievo='"&rsTabellaAvvisiLetti("CodiceAllievo2")&"'"
							Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
							Cognome2=rsTabella2("Cognome")
							Nome2=rsTabella2("Nome")
							' Url_img2=rsTabella2("Url_img")
							   ' if strcomp(Url_img2&"","")=0 then 
								' urlimmagine="../img/no-avatar.jpg"
									
								 ' else
								   ' if rsTabellaAvvisiP("CodiceAllievo2") = Session("CodAdmin") then
									  ' url= "../Materie/"&Session("ID_Materia") &"/Admin/Profili/thumb"
								   ' else
									   ' url= "../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
								   ' end if
								  ' url=Replace(url,"\","/")
								  ' urlimmagine=url&"/"& Url_img2 
			 
							   ' end if
							   %>
                      <tr class='unread'>
                        <td class='table-checkbox hidden-480'><a href="cambiastato.asp?IdNotifica=<%=rsTabellaAvvisiLetti("ID_Avviso")%>&Rimuovi=1" class='btn' rel="tooltip" data-placement="bottom" title="Elimina dall'archivio"><i class="icon-trash"></i></a></td>
						<td class='table-checkbox hidden-480'><a href="cambiastato.asp?IdNotifica=<%=rsTabellaAvvisiLetti("ID_Avviso")%>&Ripristina=1" class='btn' rel="tooltip" data-placement="bottom" title="Segna come Da Leggere"><i class="icon-repeat"></i></a></td>
                        <td class='table-fixed-medium'>&nbsp;<%=trim(Cognome2)%>&nbsp; <%=left(Nome2,1)&"."%></td>
                        <td class='table-fixed-medium'><% if isNull(rsTabellaAvvisiLetti("Testo")) then%><i>Per leggere il contenuto del messaggio, aprire la notifica</i><%else%><%=rsTabellaAvvisiLetti("Testo")%><%end if%></td>
						
						<% action = rsTabellaAvvisiLetti("Azione")
						vett = split(action, ">")
						vett(1) = Replace(vett(1), " !", "!")
						
						%>
						
                        <td class='table-fixed-medium'><a href="legginotifica.asp?IdNotifica=<%=rsTabellaAvvisiLetti("ID_Avviso")%>&action=<%=action%>"><%=vett(1)&"<>"%></a></td>
                        <td class='table-date hidden-480'><%=left(rsTabellaAvvisiLetti("Data"),10)%></td>
                      </tr>
                      <%
							  k=k+1
							  rsTabellaAvvisiLetti.movenext%>
                      <% loop %>
                      <%end if%>
                  </table>
                </div>
              </div>
            </div>
          </div>