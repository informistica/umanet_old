<%
	 
 		   
				  
				 
					
					
					
					    id_classe=Session("Id_Classe")
						QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
						
 
				
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					divid=request.querystring("divid")
					cartella=rsTabella.fields("Cartella")
                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze
					
					
					
					
					%>

<div id="left">
			<form action="#search-results.html" method="GET" class='search-form'>
				<div class="search-pane">
					<input type="text" name="search" placeholder="Search here...">
					<button type="submit"><i class="icon-search"></i></button>
				</div>
			</form>
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Classe</span></a>
				</div>
				<ul class="subnav-menu">
					
                   <li><a href="default.asp?scegli=1&id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Lavagna</a></li>
								<li><a href="../../home_app.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>"><span></span>Libro</a></li>

 <li> <a href="../quaderno.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&daMenu=1">Quaderno</a></li>
  <li> <a href="../classifica.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>"><span></span>Classifica</a> </li>

								<li><a href="default.asp?scegli=0&id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span>Forum</a></li>
								<li> <a href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span>Chat</a></li>
                                <li><a href="default.asp?scegli=2&id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span>Diario</a></li>
				</ul>
			</div>
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>U-ecdl</span></a>
				</div>
				<ul class="subnav-menu">
				 
                 	 <li><a href="../../U-ECDL/home_uecdl_app.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>"><span></span>Apprendimento</a></li>
								<li><a href="../../U-ECDL/home_uecdl_ver.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>"><span></span>Verifica</a></li>
                 
                 
				</ul>
			</div>
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Gestione</span></a>
				</div>
				<ul class="subnav-menu"> 
                      <%if (session("Admin")=true) then %>
                           <li><a href="../studente_domande_gruppi.asp"><span></span>Gruppi</a></li>
						   <li> <a href="../admin.asp?id_Classe=<%=id_classe%>&divid=<%=divid%>"><span></span>Admin</a></li> 
                       <%end if%>
                         <li class="sub-menu"><a href="../logout.asp">Logout</a>
                     				 
				</ul>
			</div>
		</div>