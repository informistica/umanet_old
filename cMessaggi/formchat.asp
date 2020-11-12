<div class="box">

<a href="#testo"><button id="pertesto" style="display:none"></button></a>

							<div class="box-title">
								<h3>
									<i class="icon-comments"></i>
									Conversazione
								</h3>
								<div class="actions">
								</div>
							</div>
							<div class="slimScrollDiv" style="position: relative; overflow: hidden; width: auto;"><div class="box-content nopadding" style="overflow: hidden; width: auto;">
								<ul id="conversazione" class="messages">
										
									
									<% contatto = Request.QueryString("contatto")
									cod = Request.QueryString("cod")
									
									QuerySQL2="SELECT Cognome,Nome FROM Allievi WHERE CodiceAllievo='"&cod&"'"
							
									Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
									Cognome_cod=rsTabella2("Cognome")
									Nome_cod=rsTabella2("Nome")
									
									
									QuerySQL2="SELECT Cognome,Nome FROM Allievi WHERE CodiceAllievo='"&contatto&"'"
							
									Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
									Cognome_contatto = rsTabella2("Cognome")
									Nome_contatto=rsTabella2("Nome")
									
									
									QuerySQL = "SELECT * FROM AVVISI WHERE ((CodiceAllievo='"&LCase(cod)&"' AND CodiceAllievo2 = '"&LCase(contatto)&"') OR (CodiceAllievo2='"&LCase(cod)&"' AND CodiceAllievo = '"&LCase(contatto)&"')) AND CAST(Testo AS ntext) like Azione order by Data asc;"
									Set rsTabellaMessaggi = ConnessioneDB.Execute(QuerySQL)
									
									%>
									
									<% do while not rsTabellaMessaggi.EOF 
									
									if LCase(rsTabellaMessaggi("CodiceAllievo")) = LCase(cod) then %>
									
									<li class="left">
										<!--<div class="image">
											<img src="img/demo/user-1.jpg" alt="">
										</div>-->
										<div class="message">
											<span class="caret"></span>
											<span class="name"><%=Cognome_contatto&" "&Nome_contatto%></span>
											<p><%=rsTabellaMessaggi("Testo")%> </p>
											<span class="time">
												<%=rsTabellaMessaggi("Data")%>
											</span>
										</div>
									</li>
									
									<% else %>
									<li class="right">
										<!--<div class="image">
											<img src="img/demo/user-2.jpg" alt="">
										</div>-->
										<div class="message">
											<span class="caret"></span>
											<span class="name"><%=Cognome_cod&" "&Nome_cod%></span>
											<p><%=rsTabellaMessaggi("Testo")%></p>
											<span class="time">
												<%=rsTabellaMessaggi("Data")%>
											</span>
										</div>
									</li>
									
									<% end if
									
									rsTabellaMessaggi.movenext
									loop
									%>
									
									<%
									
									'elimino notifiche
									
									QuerySQL = "UPDATE AVVISI SET Visto = 1 WHERE (CodiceAllievo='"&LCase(cod)&"' AND CodiceAllievo2 = '"&LCase(contatto)&"') AND CAST(Testo AS ntext) like Azione;"
									ConnessioneDB.Execute(QuerySQL)
									
									
									%>
									
									<!--<li class="typing">
										<span class="name">John Doe</span> is typing <img src="img/loading.gif" alt="">
									</li>--><br>
								
								</ul>
								<ul class="messages">
									<li class="insert">
										<form id="message-form" method="POST" action="inserisci_messaggio_personale.asp?CodiceAllievo=<%=contatto%>">
											<div class="text">
												<a name="testo"><input type="text" name="txtMessaggio" placeholder="Inserisci qui il tuo messaggio (max 250 caratteri)..." class="input-block-level"></a>
											</div>
											<div class="submit">
												<button type="submit"><i class="icon-share-alt"></i></button>
											</div>
										</form>
									</li>
								</ul>
							</div><div class="slimScrollBar ui-draggable" style="background: rgb(102, 102, 102); width: 7px; position: absolute; top: 0px; opacity: 0.4; display: none; border-radius: 7px; z-index: 99; right: 1px; height: 405px;"></div><div class="slimScrollRail" style="width: 7px; height: 100%; position: absolute; top: 0px; display: none; border-radius: 7px; background: rgb(51, 51, 51); opacity: 0.2; z-index: 90; right: 1px;"></div></div>
						</div>