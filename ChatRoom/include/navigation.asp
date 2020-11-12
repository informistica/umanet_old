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



' per chiamare direttamente il quaderndo
					id_classe=Session("Id_Classe")
					QuerySQL="SELECT Data FROM [dbo].[3PERIODI] WHERE Id_Classe='"&id_classe&"' and Iniziale=1"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					if daStud="" then
					 if rsTabella.eof then
					     DataClaq=cdate(inizio_anno)
					 else
					    DataClaq=rsTabella("Data")
					  end if
					DataClaq2=formattaData()
					end if
					cod=Session("CodiceAllievo")	
					
					
					QuerySQL="SELECT Cognome,Nome,Url_img FROM Allievi WHERE CodiceAllievo='"&cod&"'"
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					Cognome=rsTabella("Cognome")
					Nome=rsTabella("Nome")
					Url_img=rsTabella("Url_img")
					
					QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"
					Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					divid=request.querystring("divid")
					cartella=rsTabella.fields("Cartella")
                    Id_Stud=request.QueryString("Id_Stud") ' è settato se lo devo inoltrare a modifica scadenze






					
					%>

<div class="container-fluid">
			<a href="#" id="brand">Umanet <%=session("Cartella")%></a>
			<a href="#" class="toggle-nav" rel="tooltip" data-placement="bottom" title="Toggle navigation"><i class="icon-reorder"></i></a>
			<ul class='main-nav'>     
            <li>
					<a href="../../home.asp">
						<i class="icon-home"></i>
						<span>Home</span>
					</a>
				</li>
				
                
                <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Classe</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">
                     <li><a href="../social/default.asp?scegli=2&id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>"><span></span>Diario</a></li>
						 <li><a href="default.asp?scegli=1&id_classe=<%=rsTabella.fields("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella.fields("cartella")%>">Lavagna</a></li>
							<li><a href="../../home_app.asp?divid=<%=divid%>&id_classe=<%=rsTabella.fields("Id_Classe")%>"><span></span>Libro</a></li>

 <li> <a href="../quaderno.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>&daMenu=1">Quaderno</a></li>
     <%'if session("Admin")=true then %>
   <li> <a href="../classifica.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&classe=<%=rsTabella.fields("Classe")%>"><span></span>Classifica</a> </li> 
  <%' end if%> 
   <li> <a href="../calendario.asp"><span></span>Calendario</a> </li>



								
					</ul>
				</li>
                
                
                 <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Umanet</span>
						<span class="caret"></span>
					</a>
                    
                    
				<ul class="dropdown-menu">
                   <li><a href="../../U-ECDL/home_uecdl_app.asp?uecdl=1&amp;stato=1&amp;id_classe=<%=id_classe%>&amp;cartella=<%=cartella%>&amp;divid=<%=divid%>"><span></span>Libro U</a></li>
                     <li> <a href="../quaderno_metafore.asp?id_classe=<%=id_classe%>&amp;divid=<%=divid%>&amp;classe=<%=rsTabella.fields("Classe")%>&amp;cod=<%=cod%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq2%>&amp;daMenu=1">Quaderno U</a></li>
								 
  								<li><a href="../social/default.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>Forum</a></li>
								<li> <a href="../ChatRoom/showChat.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>Chat</a></li>      
					
					</ul>
			
                    
                    
				</li>
            
                
                
                 
            
            
            
            
               <%if (session("Admin")=true) then %>
                  <li>
					<a href="#" data-toggle="dropdown" class='dropdown-toggle'>
						<i class="icon-edit"></i>
						<span>Gestione</span>
						<span class="caret"></span>
					</a>
					<ul class="dropdown-menu">
						
                           <li><a href="../studente_domande_gruppi.asp"><span></span>Gruppi</a></li>
						   <li> <a href="../admin.asp?id_Classe=<%=id_classe%>&divid=<%=divid%>"><span></span>Admin</a></li>                        
                         
                         <li class="sub-menu"><a href="../logout.asp">Logout</a>
					</ul>
				</li>
                  <% else %>
          
					<!--	 <li class="sub-menu"><a href="logout.asp">Logout</a>-->
					</a>
				</li>
                
                <%end if%>


			</ul>
            
               <!-- #include file = "../../cClasse/studente_domande_include/2_messaggi_3.asp" --> 
          
            
            <div class="user">
				
				<ul class="icon-nav">
					<li class='dropdown'>
                    
						<a href="#" class='dropdown-toggle' data-toggle="dropdown"><i class="icon-envelope-alt"></i>
                        <% if numMessaggi>0 then %>
                        <span class="label label-lightred"><%=numMessaggi%></span>
                         <%end if%>
                        </a>
                        
                         <% If rsTabellaAvvisiP.BOF=True And rsTabellaAvvisiP.EOF=True then %>
              				<ul class="dropdown-menu pull-right message-ul">
                            <li>
                 <a href="../centro_messaggi.asp?DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>
							</li>
						</ul>
                            
                       <% else %>
						<ul class="dropdown-menu pull-right message-ul">
                         <% 
						 k=0
						 do while not rsTabellaAvvisiP.EOF  and k<5%>
                           		<%QuerySQL2="SELECT Cognome,Nome,Url_img FROM Allievi WHERE CodiceAllievo='"&rsTabellaAvvisiP("CodiceAllievo2")&"'"
							Set rsTabella2 = ConnessioneDB.Execute(QuerySQL2)
							Cognome2=rsTabella2("Cognome")
							Nome2=rsTabella2("Nome")
							Url_img2=rsTabella2("Url_img")
							   if strcomp(Url_img2&"","")=0 then 
								urlimmagine="../../img/no-avatar.jpg"
									
								 else
								   if rsTabellaAvvisiP("CodiceAllievo2") = Session("CodAdmin") then
									  url= "../../Materie/"&Session("ID_Materia") &"/Admin/Profili/thumb"
								   else
									   url= "../../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
								   end if
								  url=Replace(url,"\","/")
								  urlimmagine=url&"/"& Url_img2
			 
							   end if%>
						<li> 
								<a href="#">
									<img src="<%=urlimmagine%>" title="<%=trim(Cognome2)%>&nbsp; <%=left(Nome2,1)&"."%>" width="84px" height="84px">
									<div class="details">
                                  
									  <div class="name">&nbsp;&nbsp;&nbsp;
											<%=rsTabellaAvvisiP("Testo")%>
										</div>
									</div>
								</a>
							</li>
                         <!--   
                            
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
                                <a href="../centro_messaggi.asp?DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>
							</li>
						</ul>
                         <%end if%>
					</li>
					
					<li class="dropdown sett">
						<a href="#" class='dropdown-toggle' data-toggle="dropdown"><i class="icon-cog"></i></a>
					  <%' if session("Admin")=true then %>
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
									<a href="#" class="set-sidebar-default">Scorrevole</a>
								</div>
							</li>
						</ul>
                        <% 'end if%>
					</li>
					<li class='dropdown colo'>
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
				<div class="dropdown">
					<a href="#" class='dropdown-toggle' data-toggle="dropdown"><%=Cognome%> <% response.write(left(Nome,1)&".")%> 
                    <% if strcomp(Url_img&"","")=0 then%>
                       <img src="../../img/no-avatar.jpg" alt=""  width="28px" height="28px" class="imground">
                    <%else
					    if session("Admin")=true or session("Admin2")<>""  then
					      url= "../../Materie/"&Session("ID_Materia") &"/"&Session("CartellaAdmin")&"/Profili/thumb"
					   else
					       url= "../../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
    				   end if
					  url=Replace(url,"\","/")
					  urlimg=url&"/"& Url_img 
					  
					 %>
                      <img src="<%=urlimg%>" title="<%=trim(Cognome)%>&nbsp; <%=trim(Nome)%>&nbsp; " width="28px" height="28px" class="imground">
                    <%end if%>
                    </a>
					<ul class="dropdown-menu pull-right">
						<li>
							<a href="../form_cambia_pwd_new.asp?dividApro=0">Modifica Profilo</a>
						</li>
						 
						<li>
							<a href="../logout.asp">Log out</a>
						</li>
					</ul>
				</div>
			</div>
            
            </div>
            
     
			
        
        <script type="text/javascript">
		$('.red').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=red"
			 event.stopPropagation();
		});
			$('.green').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=green"
			 event.stopPropagation();
		});
			$('.brown').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=brown"
			 event.stopPropagation();
		});
			$('.lime').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=lime"
			 event.stopPropagation();
		});
			$('.purple').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=purple"
			 event.stopPropagation();
		});
		$('.pink').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=pink"
			 event.stopPropagation();
		});
		
		$('.magenta').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=magenta"
			 event.stopPropagation();
		});
		
			$('.grey').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=grey"
			 event.stopPropagation();
		});
			$('.darkblue').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=darkblue"
			 event.stopPropagation();
		});
			$('.lightgrey').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=lightgrey"
			 event.stopPropagation();
		});
		
		 
			$('.satblue').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.orange').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=orange"
			 event.stopPropagation();
		});
			$('.blue').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=blue"
			 event.stopPropagation();
		});
		$('.satblue').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=satblue"
			 event.stopPropagation();
		});
			$('.satgreen').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=satgreen"
			 event.stopPropagation();
		});
			 $('.teal').bind('click', function() {
            document.location = "aggiorna_stile.asp?stile=teal"
			 event.stopPropagation();
		});
	 
		
		
		
		
		
		
		</script>