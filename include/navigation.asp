<%



	Dim ConnessioneDB , rsTabella,QuerySQL, CodiceTest, CodiceAllievo,PwdAllievo, CodiceCorso, i,StringaConnessione,capitolo
	function formattaData

	  					if day(date()) < 10 then
						giorno="0" & day(date())
						else
						giorno=day(date())
						end if
						anno=year(date())

						if month(date()) < 10 then
						mese="0" & month(date())
						else
						mese=month(date())
						end if

						formattaData = giorno & "/" & mese& "/" & anno


	end function
				   'Apertura della connessione al database



		'			id_classe=Session("Id_Classe")
'					' per chiamare direttamente il quaderndo
'					QuerySQL="SELECT Data FROM [dbo].[3PERIODI] WHERE Id_Classe='"&id_classe&"' and Iniziale=1"
'					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'					DataClaq=rsTabella("Data")
'
'					DataClaq2=formattaData()
'					cod=Session("CodiceAllievo")
'
'
'
'					QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
'
'
'					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'					divid=request.querystring("divid")
'					cartella=rsTabella.fields("Cartella")
'                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze
'
<!-- #include file = "../service/formatta_data_LO.asp" -->
if Session("CodiceAllievo")="" then%>
	<!-- #include file = "../stringhe_connessione/stringa_connessione_refresh.asp" -->

	      <%
		  if (Session("CodiceAllievo")="") or (Session("Id_Classe")="")  then
		     response.redirect "../service/redirect.asp"
	      end if

	 end if



' per chiamare direttamente il quaderndo
					id_classe=request.QueryString("Id_Classe")
				'	end if
				if id_classe="" then
				id_classe=Session("Id_Classe")
				end if

					'QuerySQL="SELECT Data FROM [dbo].[3PERIODI] WHERE Id_Classe='"&id_classe&"' and Iniziale=1"
'					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'					if daStud="" then
'					 if rsTabella.eof then
'					     DataClaq=cdate(inizio_anno)
'					 else
'					    DataClaq=rsTabella("Data")
'					  end if
'					DataClaq2=formattaData()
'					end if


					cod=Session("CodiceAllievo")


					QuerySQL="SELECT Cognome,Nome,Url_img FROM Allievi WHERE CodiceAllievo='"&cod&"'"
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					Cognome=rsTabella("Cognome")
					Nome=rsTabella("Nome")
					Url_img=rsTabella("Url_img")
					
					'response.write(QuerySQL&" --- " & rsTabella("Cognome"))
					QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					'response.write("jjj"&rsTabella(0))
					divid=request.querystring("divid")
					cartella=rsTabella.fields("Cartella")
					classe=rsTabella.fields("Classe")
					urlfeedback=rsTabella.fields("Url_feedback")
					response.cookies("Dati")("Cartella")=cartella
					response.cookies("Dati")("Classe")=classe
                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze
					'response.cookies("Dati")("Cartella")=cartella

 QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
 Set rsTabellaSetting = ConnessioneDB.Execute(QuerySQL)
 ValidaQuiz=rsTabellaSetting("ValidaQuiz")
 set rsTabellaSetting = nothing




					%>

<div class="container-fluid">
			<a href="../cClasse/quaderno.asp?umanet=0&id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>&cod=<%=cod%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("Datacla2")%>&daMenu=1" id="brand">Umanet <%=replace(left(classe,1+len(classe)-(instr(classe,"$")-1)),"$","")%> </a>
			<a href="#" class="toggle-nav" rel="tooltip" data-placement="bottom" title="Toggle navigation"><i class="icon-reorder"></i></a>
			<ul class='main-nav'>
            <li>
            <%if (strcomp(Session("DB"),"1")=0) then%>
					<a href="../../home.asp">
             <%else%>
                <%if (strcomp(Session("DOC"),"1")=0) then%>
            	 <a href="../../home.asp">
                  <%else%>
                      <%if (strcomp(Session("AS"),"1")=0) then%>
                          <a href="../../home4.asp">
                     <%else%>
                     <a href="../../home.asp">
                     <% end if%>
				 <%end if%>

			 <%end if%>
						<i class="icon-home"></i>
						<span>Home </span>
					</a>
				</li>


                <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Classe</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">
                     <li><a href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-log_book"></i>&nbsp;Diario</a></li>
						 <li><a href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-desktop"></i>&nbsp;Bacheca</a></li>
							<li><a href="../cClasse/home_app.asp?divid=<%=divid%>&amp;id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>Libro</a></li>
<% 'if ValidaQuiz=1 then %>
 <li> <a href="../cClasse/quaderno.asp?umanet=<%=umanet%>&id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>&daMenu=1"><i class="icon-paste"></i>&nbsp;Quaderno</a></li>
 <%' end if%>
  <li > <a   href="../cClasse/quaderno_mappe.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>"><i class="glyphicon-snowflake"></i>&nbsp;Mappe</a></li>
  <% 'if Session("Admin")=true then
  %>
	<li><a href="../cSocial/default0.asp?scegli=3&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-question-sign"></i>&nbsp;Interrogazioni</a></li>
   <li> <a href="../cClasse/classifica.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>"><span></span><i class="glyphicon-charts"></i>&nbsp;Classifica</a> </li>
  <%'end if
  %>
  <% if urlfeedback<>"" then %>
 <li id="feedback"> <a href="<%=urlfeedback%>"><span></span><i class="icon-exchange"></i>&nbsp;Feedback</a> </li>
 <% end if %> 
 <li id="recupero"> <a  href="../cFrasi/2compilaprefrase_recupero.asp?CodiceAllievo=<%=cod%>&id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("Cartella")%>"><span></span><i class="icon-signal"></i>&nbsp;Recupero</a> </li>
 
   <li> <a href="../cClasse/calendario.asp"><span></span><i class="glyphicon-calendar"></i>&nbsp;Calendario</a> </li>



					</ul>
				</li>


                 <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Umanet</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">
                    	<li><a href="../cClasse/home_app.asp?umanet=1&divid=<%=divid%>&amp;id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>Libro U</a></li>     <% if ValidaQuiz=1 then %>
                     <li> <a href="../cClasse/quaderno_metafore.asp?umanet=1&id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>&daMenu=1"><i class="icon-paste"></i>&nbsp;Quaderno U</a></li>
						<%end if%>
  								<li><a href="../cSocial/default0.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="icon-comments"></i>&nbsp;Forum</a></li>
								<li> <a href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="icon-comments-alt"></i>&nbsp;Chat</a></li>


					</ul>
				</li>

				<% if (strcomp(Session("DB"),"1") <> 0) and (strcomp(Session("DOC"),"1") <> 0) and (strcomp(Session("AS"),"1") <> 0) then %>

                <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-refresh"></i>
						<span>Cambia Classe</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">

				<%
					QuerySQL="SELECT * FROM anni_scolastici where Attivo=1"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					id_as=rsTabella("ID_AS")



					'QuerySQL="SELECT * FROM anni_classi where Id_Classe <> '8COM' and Id_As="&id_as
					'QuerySQL="SELECT *  FROM Classi where ID_Classe not in (SELECT Id_Classe  FROM anni_classi where  Id_As<>"&id_as&") and Classe<>'2B'"

					'dopo il cambio di strategia per la gestione degli a/s la nuova query è la seguente
					QuerySQL="SELECT *  FROM Classi where Visibile=1 order by Classe"

					Set rsTabella = ConnessioneDB.Execute(QuerySQL)

					randomize()
					i=0 ' indice per cambiare il colore e distinguera la classe
					if rsTabella.eof then
					' non ci sono classi rimando ad Admin per l'inserimento della prima classe
					response.Redirect "script/cAdmin/inserisci_classe.asp"
					end if%>

					<%
					do while not rsTabella.eof
					 %>

								<li><a href="../cUtenti/form_login2.asp?app=1&id_classe=<%=rsTabella.fields("Id_Classe")%>&id_materia=<%=id_materia%>&cartella=<%=rsTabella.fields("Classe")%>&id_as=<%=id_as%>&id_scuola=<%=id_scuola%>">
								Classe <%=replace(left(rsTabella("Classe"),1+len(rsTabella("Classe"))-(instr(rsTabella("Classe"),"$")-1)),"$","") %>

                                </a></li>


						 <%  i=i+1
						     if i=max_stile_classi then
							    i=0
							 end if
						    rsTabella.movenext
					loop
				%>
				<%
					QuerySQL="SELECT *  FROM Classi where Visibile=0 and (Classe like '5%' )  order by Contatore"

					Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

					randomize()
					i=0 ' indice per cambiare il colore e distinguera la classe
					 %>
				<li class="dropdown-submenu">
							<a href="#" data-toggle="dropdown" class="dropdown-toggle">Anni scorsi</a>
							<ul class="dropdown-menu">
							<%
								do while not rsTabella1.eof
							%>
								 	<li><a href="../cClasse/home_app.asp?divid=<%=divid%>&amp;id_classe=<%=rsTabella1.fields("ID_Classe")%>&cartella=<%=rsTabella1.fields("cartella")%>"><span></span><i class="glyphicon-book"></i><%=rsTabella1.fields("Nome")%></a></li>

						 <% 
						    rsTabella1.movenext
					loop
%>
							</ul>
						</li>

				<li>
					<ul class="dropdown-menu">
						<li>
						<a href="../../home.asp?classi=1&anniscorsi=1">Anni Scorsi</a>
						</li>
					</ul>
				</li>

				</ul>
				</li>







				<% end if %>

               <% if (session("Admin") = true) then %>
                  <!--<li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Gestione</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">

                           <li><a href="../cClasse/studente_domande_gruppi.asp"><span></span>Gruppi</a></li>
						   <li> <a href="../cAdmin/admin.asp?id_Classe=<%=id_classe%>&amp;divid=<%=divid%>"><span></span>Admin</a></li>

                         <li class="sub-menu"><a href="../service/logout.asp">Logout</a>
					</ul>
				</li>-->

				<li> <a href="../cAdmin/admin.asp?id_Classe=<%=id_classe%>&amp;divid=<%=divid%>"><i class="icon-key"></i><span> Admin</span></a></li>

                  <% else %>

					<!--	 <li class="sub-menu"><a href="logout.asp">Logout</a>-->
					</a>
				</li>

                <%end if%>


			</ul>



            <div class="user" id="notifiche">

		   <!-- #include file = "../cClasse/studente_domande_include/2_messaggi_3.asp" -->

				<ul class="icon-nav" >
					<li class='dropdown' >



						<a href="#" class='dropdown-toggle' data-toggle="dropdown"><i class="icon-envelope-alt"></i>
                        <% if numMessaggi>0 then %>
                        <span class="label label-lightred" ><%=numMessaggi%></span>
                         <%end if%>
                        </a>

                         <% If rsTabellaAvvisiP.BOF=True And rsTabellaAvvisiP.EOF=True then %>
              				<ul class="dropdown-menu pull-right message-ul">
                            <li>
                  <!--<a href="../cMessaggi/centro_messaggi.asp?daMenu=1&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>-->
						<a href="../cMessaggi/centro_messaggi.asp" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>
							</li>
						</ul>

                       <% else %>
						<ul class="dropdown-menu pull-right message-ul">
                         <%
						 k=0
						 do while not rsTabellaAvvisiP.EOF  and k<5%>
                           		<%QuerySQL2="SELECT Cognome,Nome,Url_img FROM Allievi WHERE CodiceAllievo='"&rsTabellaAvvisiP("CodiceAllievo2")&"'"
							'response.write(QuerySQL2)
							Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
							Cognome2=rsTabella2("Cognome")
							Nome2=rsTabella2("Nome")
							Url_img2=rsTabella2("Url_img")

							   ' if strcomp(Url_img2&"","")=0 then
								' urlimmagine="../../img/no-avatar.jpg"

								 ' else
								     ' if  (strcomp(rsTabellaAvvisiP("CodiceAllievo2"),Session("CodAdmin"))=0) then
									  ' url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/Admin/Profili/thumb"
								   ' else
									   ' url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&session("id_classe_img")&"/Profili/thumb" ' vuole il percorso relativo della cartella
								   ' end if
								   ' response.write(""&rsTabellaAvvisiP("CodiceAllievo2") &" - " &rsTabellaAvvisiP("Azione")  &" - " &rsTabellaAvvisiP("ID_Avviso"))
								  ' url=Replace(url,"\","/")
								  ' urlimmagine=url&"/"& Url_img2

							   ' end if
							   %>
						<li>

									<!--<img src="<%=urlimmagine%>" title="<%=trim(Cognome2)%>&nbsp; <%=left(Nome2,1)&"."%>" width="84px" height="84px">-->

									<% if rsTabellaAvvisiP("Testo") = rsTabellaAvvisiP("Azione") then %>

									<a href="../cMessaggi/leggichat.asp?cod=<%=cod%>&contatto=<%=rsTabellaAvvisiP("CodiceAllievo2")%>" <%if (k Mod 2) <> 0 then %> style="background-color:#EEEEEE" onMouseOver="this.style='background-color:#F9F9F9'" onMouseOut="this.style='background-color:#EEEEEE'" <% end if %>>
									Nuovo Messaggio da <b><%=Cognome2&" "%><%=Nome2%></b>

									<% else %>

									<a href="../cMessaggi/legginotifica.asp?IdNotifica=<%=rsTabellaAvvisiP("ID_Avviso")%>&action=<%=rsTabellaAvvisiP("Azione")%>" <%if (k Mod 2) <> 0 then %> style="background-color:#EEEEEE" onMouseOver="this.style='background-color:#F9F9F9'" onMouseOut="this.style='background-color:#EEEEEE'" <% end if %>>
									Nuova Notifica da <b><%=Cognome2&" "%><%=Nome2%></b>

									<% end if %>

									<!--<div class="details">

									  <div class="name">&nbsp;&nbsp;&nbsp;
											<%=rsTabellaAvvisiP("Testo")%><br />
										</div>
									</div>-->
								</a>
							</li>
                         <!--
                            <%


							%>
                            <li>
								<a href="#">
									<img src="../img/demo/user-1.jpg" alt="">
									<div class="details">
										<div class="name">Jane Doe</div>
										<div class="message">
											Lorem ipsum Commodo quis nisi ...
										</div>
									</div>
								</a>
							</li>
                          -->


                              <%
							  k=k+1
							  rsTabellaAvvisiP.movenext%>
                         <% loop %>
							<li>


							<!--	<a href="components-messages.html" class='more-messages'>Go to Message center <i class="icon-arrow-right"></i></a>-->
                                <a href="../cMessaggi/centro_messaggi.asp" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>
							</li>
						</ul>
                         <%end if%>
					</li>

					<li class="dropdown sett" >
						<a href="#" class='dropdown-toggle' data-toggle="dropdown"><i class="icon-cog"></i></a>
					  <%' if Session("Admin")=true then
					  %>
                        <ul class="dropdown-menu pull-right theme-settings">
							<!--<li>
								<span>Layout-width</span>
								<div class="version-toggle">
									<a href="#" class='set-fixed'>Fixed</a>
									<a href="#" class="active set-fluid">Fluid</a>
								</div>
							</li>-->
							<li>
								<span>Menu Superiore </span>
								<div class="topbar-toggle">
									<a href="#" class='active set-topbar-fixed' id="FissaTopBar">Fisso</a>
									<a href="#" class="set-topbar-default">Scorrevole</a>
								</div>
							</li>
							<li>
								<span>Menu Laterale </span>
								<div class="sidebar-toggle">
									<a href="#" class='active set-sidebar-fixed' id="FissaSideBar">Fisso</a>
									<a href="#" class="set-sidebar-default" id="sidebarScorr">Scorrevole</a>
								</div>
							</li>
						</ul>
                        <% 'end if
						%>
					</li>
					<li class='dropdown colo' >
						<a href="#" class='dropdown-toggle' data-toggle="dropdown"><i class="icon-tint"></i></a>
						<ul class="dropdown-menu pull-right theme-colors">
							<li class="subtitle" >
								Imposta il tuo colore
							</li>

							<li>
								<span class='red' ></span>
								<span class='green'></span>
								<span class="brown"></span>
								<span class='lime' ></span>
								<span class="purple"></span>
								<span class="pink" ></span>
                                <span class="teal"></span>
								<span class="magenta" ></span>
								<span class="grey" ></span>
								<span class="darkblue"></span>
								<span class="lightred" ></span>
								<span class="lightgrey"></span>
								<span class="satblue"></span>
                                <span class='orange' ></span>
                                <span class="blue"></span>
                                <span class="satgreen" ></span>

							</li>
						</ul>
					</li>

					<li>
					<!--	<a href="more-locked.html" class='lock-screen' rel='tooltip' title="Lock screen" data-placement="bottom"><i class="icon-lock"></i></a>-->
                        <a href="#" class='lock-screen' rel='tooltip' title="Lock screen" data-placement="bottom"><i class="icon-lock"></i></a>
					</li>
				</ul>
				<div class="dropdown" >
					<a href="#" class='dropdown-toggle' data-toggle="dropdown"><%=Cognome%>&nbsp; <% response.write(left(Nome,1)&".")%>
                    <% if strcomp(Url_img&"","")=0 then%>
                       <img src="../../img/no-avatar.jpg" alt=""  style="width:28px; height:28px" class="imground">
                    <%else
					    'if(session("admin") = true) or Session("Admin2")<>""  then
					     if session("Admin") = true then
						 ' Session("Admin")=true ' per il problema della foto admi che spariva
					      url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&Session("CartellaAdmin")&"/Profili/thumb"
					   else
					       url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&session("id_classe_img")&"/Profili/thumb" ' vuole il percorso relativo della cartella
    				   end if
					  url=Replace(url,"\","/")
					  urlimg=url&"/"& Url_img

					 %>
                      <img src="<%=urlimg%>" title="<%=trim(Cognome)%>&nbsp; <%=trim(Nome)%>&nbsp; " style="width:28px; height:28px" class="imground">
                    <%end if%>
                    </a>
					<ul class="dropdown-menu pull-right">
						<% if Session("identita")=true then %>

						<li>
							<a href="../cUtenti/login256.asp?identita=1&CodiceAllievo=informistica&Cartella=Expo&id_classe=6COM">Torna ad Admin</a>
						</li>

						<% end if %>

						<li>
							<a href="../cUtenti/form_cambia_pwd_new.asp?dividApro=0&cartella=<%=cartella%>">Modifica Profilo</a>
						</li>

						<li>
							<a href="../service/logout.asp">Log out  </a>
						</li>
					</ul>
				</div>
			</div>
            </div>
			

        <script type="text/javascript">

		$('.red').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=red"
			 event.stopPropagation();
		});
			$('.green').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=green"
			 event.stopPropagation();
		});
			$('.brown').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=brown"
			 event.stopPropagation();
		});
			$('.lime').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=lime"
			 event.stopPropagation();
		});
			$('.purple').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=purple"
			 event.stopPropagation();
		});
		$('.pink').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=pink"
			 event.stopPropagation();
		});

		$('.magenta').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=magenta"
			 event.stopPropagation();
		});

			$('.grey').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=grey"
			 event.stopPropagation();
		});
			$('.darkblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=darkblue"
			 event.stopPropagation();
		});
			$('.lightgrey').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=lightgrey"
			 event.stopPropagation();
		});


			$('.satblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.orange').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=orange"
			 event.stopPropagation();
		});
			$('.blue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=blue"
			 event.stopPropagation();
		});
		$('.satblue').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.satgreen').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=satgreen"
			 event.stopPropagation();
		});
			 $('.teal').bind('click', function() {
            document.location = "../service/aggiorna_stile.asp?stile=teal"
			 event.stopPropagation();
		});

		</script>
