            <div class="box">
					<div class="visible-tablet">
						<div class="navbar">
							<div class="navbar-inner">
								<ul class="nav">
                                <li><a href="../cClasse/home_app.asp?divid=<%=divid%>&amp;id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>Libro</a></li>
								<li><a href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-log_book"></i>Diario</a></li>
						         <li><a href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-desktop"></i>Bacheca</a></li>
                                 <li> <a href="../cClasse/quaderno.asp?umanet=<%=umanet%>&id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>&daMenu=1"><i class="icon-paste"></i>Quaderno</a></li>
								<li> <a href="../cClasse/classifica.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>"><span></span><i class="glyphicon-charts"></i>Classifica</a> </li>
								<li> <a href="../cClasse/calendario.asp"><span></span><i class="glyphicon-calendar"></i>Calendario</a> </li>
								<li><a href="../cSocial/default0.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="icon-comments"></i>Forum</a></li>
                                </ul>
							</div>
						</div>
					</div>

					<div class="visible-phone">
						<div class="navbar">
							<div class="navbar-inner">
								<ul class="nav">
                                <li><a href="../cClasse/home_app.asp?divid=<%=divid%>&amp;id_classe=<%=id_classe%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-book"></i>Libro</a></li>
								<li><a href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span><i class="glyphicon-log_book"></i>Diario</a></li>
						         <li><a href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><i class="icon-desktop"></i>Bacheca</a></li>
                                 <li> <a href="../cClasse/quaderno.asp?umanet=<%=umanet%>&id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>&daMenu=1"><i class="icon-paste"></i>Quaderno</a></li>
								<li> <a href="../cClasse/classifica.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&DataClaq=<%=Session("DataCla")%>&DataClaq2=<%=Session("DataCla2")%>"><span></span><i class="glyphicon-charts"></i>Classifica</a> </li>
								 </ul>
							</div>
						</div>
					</div>
				</div>
 
		 
