<%







					id_classe=request.QueryString("id_classe")
					if  id_classe="" then
					id_classe=Session("Id_Classe")
					end if

						QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"



					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					divid=request.querystring("divid")
					cartella=rsTabella.fields("Cartella")
					urlfeedback=rsTabella.fields("Url_feedback")
                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze

					 QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
 Set rsTabellaSetting = ConnessioneDB.Execute(QuerySQL)
 ValidaQuiz=rsTabellaSetting("ValidaQuiz")
 set rsTabellaSetting = nothing


					%>

<div id="left">
		<!--	<form action="#search-results.html" method="GET" class='search-form'>
				<div class="search-pane">
					<input type="text" name="search" placeholder="Search here...">
					<button type="submit"><i class="icon-search"></i></button>
				</div>
			</form>-->
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i id="classe" class="icon-angle-down"></i><span>Classe</span></a>
				</div>
				<ul class="subnav-menu">
					 <li id="diario"><a   href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-log_book"></i>&nbsp;Diario</a></li>
                   <li id="lavagna"><a   href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-desktop"></i>&nbsp;Bacheca</a></li>
								<li id="libro"><a   href="../cClasse/home_app.asp?divid=<%=divid%>&id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>&nbsp;Libro</a></li>

<% 'if ValidaQuiz=1 then %>
 <li id="compiti"> <a   href="../cClasse/quaderno.asp?umanet=0&id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>&cod=<%=cod%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("Datacla2")%>&daMenu=1"><i class="icon-paste"></i>&nbsp;Quaderno</a></li>
 <%' end if%>
 <li id="mappe" > <a   href="../cClasse/quaderno_mappe.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>"><i class="glyphicon-snowflake"></i>&nbsp;Mappe</a></li>
<%'if session("Admin")=true then
%>
 <li id="interrogazioni"><a   href="../cSocial/default0.asp?scegli=3&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-question-sign"></i>&nbsp;Interrogazioni</a></li>
 <li id="classifica"> <a   href="../cClasse/classifica.asp?daMenu=1&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("Datacla2")%>&id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>"><span></span><i class="glyphicon-charts"></i>&nbsp;Classifica</a> </li>
 <% if urlfeedback<>"" then %>
 <li id="feedback"> <a href="<%=urlfeedback%>"><span></span><i class="icon-exchange"></i>&nbsp;Feedback</a> </li>
 
 <% end if %> 
  <li id="recupero"> <a   href="../cFrasi/2compilaprefrase_recupero.asp?CodiceAllievo=<%=cod%>&id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("Cartella")%>"><span></span><i class="icon-signal"></i>&nbsp;Recupero</a> </li>
 

 
  <%'end if
  %>



   <li id="calendario" > <a href="../cClasse/calendario.asp"><span></span><i class="glyphicon-calendar"></i>&nbsp;Calendario</a> </li>


				</ul>
			</div>
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-up"></i><span>Umanet</span></a>
				</div>
				<ul class="subnav-menu">
				  <li id="librou"><a   href="../cClasse/home_app.asp?umanet=1&divid=<%=divid%>&id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>&nbsp;Libro U</a></li>
                <% if ValidaQuiz=1 then %>
                  <li id="quadernou"> <a   href="../cClasse/quaderno_metafore.asp?umanet=1&id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>&cod=<%=cod%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("Datacla2")%>&daMenu=1"><i class="icon-paste"></i>&nbsp;Quaderno U</a></li>
                  <% end if%>
                 	<li id="forum"><a   href="../cSocial/default0.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="icon-comments"></i>&nbsp;Forum</a></li>
								<li id="chat"> <a   href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="icon-comments-alt"></i>&nbsp;Chat</a></li>



				</ul>
			</div>

			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Gestione </span></a>
				</div>
				<ul class="subnav-menu">

				<% if Session("identita")=true then %>

						<li>
							<a href="../cUtenti/login256.asp?identita=1&CodiceAllievo=informistica&Cartella=Expo&id_classe=6COM">Torna ad Admin</a>
						</li>

						<% else %>

                      <%'if (session("admin") = true) then
						 if session("admin") = true then %>
                           <!--<li><a href="../cClasse/studente_domande_gruppi.asp"><span></span>Gruppi</a></li>-->

						   <li> <a href="../cAdmin/admin.asp?id_Classe=<%=id_classe%>&amp;divid=<%=divid%>"> Admin</a></li>
						   <li> <a href="../cClasse/home_app.asp?id_classe=8COM" >DOC</li>
						<% end if %>
					   <%end if%>
                         <li ><a href="../service/logout.asp">Logout</a></li>

				</ul>
			</div>
		</div>
