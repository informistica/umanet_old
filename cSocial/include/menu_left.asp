<%
	 
 		   
				  
				 
					
					
					
					    id_classe=Session("Id_Classe")
						QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
						
 
				
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					divid=request.querystring("divid")
					cartella=rsTabella.fields("Cartella")
                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze
					
					
					
					
					%>

<div id="left">
			<!--<form action="#search-results.html" method="GET" class='search-form'>
				<div class="search-pane">
					<input type="text" name="search" placeholder="Search here...">
					<button type="submit"><i class="icon-search"></i></button>
				</div>
			</form>-->
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Classe</span></a>
				</div>
				<ul class="subnav-menu">
					 <li><a href="../../social/include/default.asp?scegli=2&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>Diario</a></li>
                 
                   <li><a href="../../social/include/default.asp?scegli=1&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>">Lavagna</a></li>
								<li><a href="../../home_app.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>"><span></span>Libro</a></li>

 <li> <a href="../../social/quaderno.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>&amp;daMenu=1">Quaderno</a></li>
 <%'if session("Admin")=true then %>
  <li> <a href="../../social/classifica.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>"><span></span>Classifica</a> </li>
  <% 'end if%> 
   <li> <a href="../../social/calendario.asp"><span></span>Calendario</a> </li>


								 
				</ul>
			</div>
            
            	<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Umanet</span></a>
				</div>
				 
                
                <ul class="subnav-menu">
				  <li><a href="../../U-ECDL/home_uecdl_app.asp?uecdl=1&stato=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&divid=<%=divid%>"><span></span>Libro U</a></li>
                  <li> <a href="../../social/quaderno_metafore.asp?umanet=1&id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>&amp;daMenu=1">Quaderno U</a></li>
                 	<li><a href="../../social/social/default.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>Forum</a></li>
								<li> <a href="../../social/ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>Chat</a></li>
                               
                 
                 
				</ul>
			</div>
			 
			<div class="subnav">
				<div class="subnav-title">
					<a href="#" class='toggle-subnav'><i class="icon-angle-down"></i><span>Gestione</span></a>
				</div>
				<ul class="subnav-menu"> 
                      <%if (session("Admin")=true) or (session("Admin2")<>"") then %>
                           <li><a href="../../social/studente_domande_gruppi.asp"><span></span>Gruppi</a></li>
						   <li> <a href="../../social/admin.asp?id_Classe=<%=id_classe%>&amp;divid=<%=divid%>"><span></span>Admin</a></li> 
                       <%end if%>
                         <li class="sub-menu"><a href="../../social/logout.asp">Logout</a>
                     		 
				</ul>
			</div>
		</div>