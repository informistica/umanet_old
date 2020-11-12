<%@ Language=VBScript %>
<!doctype html>
 
<html>
<head>
   
   <title>Crea Quiz</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">
    
    


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery UI -->
	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.draggable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body>
	<div id="navigation">
     
        <% 
		  id_classe=request.QueryString("id_classe")
  Cartella=Request.QueryString("Cartella") 
  TitoloCapitolo=Request.QueryString("Capitolo") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")  
  dividA=request.QueryString("dividApro")
     Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
       
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Crea Quiz </h1> 
                    
					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->	 
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>
						<li>
							<a href="#more-login.html">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
                        <% if by_UECDL=1 then %>
                        <a href="#more-files.html">Libro UWWW</a>
                        <%else%>
                        <a href="#more-files.html">Libro</a>
                        <%end if%>
							
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html"><%=response.write(TitoloCapitolo)%></a>
						</li>
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				 
                 
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  <%=response.write(Paragrafo)%> </h3>
			          </div>
				      <div class="box-content">
                      
                       <%
					   if CodiceSottopar<>"" then
		QuerySQL="SELECT * " &_
"FROM [dbo].[PREDOMANDE1] WHERE Id_Paragrafo='" & CodiceTest & "' and  Id_Sottoparagrafo='" & CodiceSottopar & "' order by Posizione;"
						else
						QuerySQL="SELECT * " &_
"FROM [dbo].[PREDOMANDE1] WHERE PREDOMANDE1.Id_Paragrafo='" & CodiceTest & "' order by Posizione;" 
						end if
'response.write(QuerySQL & "<br>" & Modulo)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)	%>	  
          
                      
                      
<%if rsTabella.eof and rsTabella.bof then%>

   <div class="alert alert-error">
                     	<b>
                        </b>
                        <a rel="popover" data-trigger="hover" data-content="Crea o svolgi Quiz" title="Domande" target="_blank" href="../cClasse/scegli_azione_test.asp?id_classe=<%=id_classe%>&Cartella=<%=Cartella%>&Stato=0&Stato0=0&CodiceTest=<%=CodiceTest%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">
						<%=response.write("Non ci sono compiti assegnati!<br>Se volevi inserire un quiz clicca qui  ")%>
						
						
						<i class="icon-edit"></i>

                     </div>
   
<%else 											 
	 
		 
		 %>
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   <%
				    QuerySQL="SELECT * " &_
	"FROM preDomande WHERE preDomande.Id_Paragrafo='" & CodiceTest & "' and ID_Predomanda not in (Select Id_Predomanda from Domande WHERE Domande.Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "' order by Posizione;" 

				
	Set rsTabella = ConnessioneDB.Execute(QuerySQL)
				   
				   
				    if rsTabella.eof and rsTabella.bof then%>
                     <div class="alert alert-success">
                     	<b><%=response.write("Hai gia' svolto tutti i compiti assegnati")%></b>
                     </div>
       			
					<%end if
                    'response.write(QuerySql & " " &Paragrafo)
                    end if
                i=0
                'paragrafo=rsTabella(2)


    predomanda=1
	do while not rsTabella.eof%>
		    <div class="box-content"> 
                     <% if (i=0) then %>
 					 <ol>
					 <%end if %>
                    <% QuerySQL="SELECT count(*) FROM Domande WHERE Quesito='" & rsTabella.fields("Quesito")&"'"
    Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) %>
                      
                      <li style="line-height:10px;">
					<a title="Scade il " href="inserisci_test.asp?by_UECDL=<%=by_UECDL%>&predomanda=<%=predomanda%>&ID_Predomanda=<%=rsTabella("ID_Predomanda")%>&Tipo=0&Quesito=<%=rsTabella("Quesito")%>&Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Scadenza=<%=rsTabella("Scadenza")%>&CodiceSottopar=<%=rsTabella("Id_Sottoparagrafo")%>"><%=rsTabella("Quesito")& " (" & rsTabella1(0) &")"%>
					</a> 
							
				<%
	i=i+1
	rsTabella.movenext 
loop%>
               </ol>     <br>
               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> 
                      </div>         
			        </div>
			      </div>
			    </div>
	
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina.asp" -->
         

			 
	</body>

 </html>

