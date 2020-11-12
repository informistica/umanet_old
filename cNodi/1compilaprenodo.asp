<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<script src="../js/google.js"></script><title>Crea Nodo</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">

    


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	
    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>   
	
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
	 <!--<script src="../js/plugins/plupload/plupload.full.js"></script>
	<script src="../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../js/plugins/mockjax/jquery.mockjax.js"></script>-->
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
		
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
        <%
		
		if CodiceSottopar<>"" then
		 QuerySQL="SELECT * " &_
"FROM preNodi WHERE Id_Paragrafo='" & CodiceTest & "' and Id_Sottoparagrafo='" & CodiceSottopar & "' order by Posizione" 
		else
		 QuerySQL="SELECT * " &_
"FROM preNodi WHERE Id_Paragrafo='" & CodiceTest & "' order by Posizione" 

end if
'response.write(QuerySQL & "<br>" & Modulo)
Set rsTabellaNodi = ConnessioneDB.Execute(QuerySQL)	


%>	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="glyphicon-snowflake"></i>&nbsp;Crea Nodo</h1> 
                    
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
							<a href="#more-files.html">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html"><%=response.write(Capitolo)%></a>
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
					 ' response.write("rsTabellaNodi="&rsTabellaNodi.eof)
					  
					  if rsTabellaNodi.eof and rsTabellaNodi.bof then%>
 
   <div class="alert alert-error">
                     	<b><%=response.write("Non ci sono compiti assegnati. Se vuoi inserire un nodo libero clicca")%></b>
						  <a rel="popover" data-trigger="hover" data-content="Crea o svolgi Nodo" title="Domande" target="_blank" href="../cNodi/inserisci_nodo.asp?Tipo=0&id_classe=<%=id_classe%>&Cartella=<%=Cartella%>&Stato=0&Stato0=0&CodiceTest=<%=CodiceTest%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">&nbsp;qui&nbsp;<i class="icon-edit"></i>

                     </div>
   
<%else 	
										 
											 
	'QuerySQL="SELECT * " &_
'	"FROM PRENODI1 WHERE Id_Paragrafo='" & CodiceTest & "' and ID_Prenodo not in (Select Id_Prenodo from Nodi WHERE Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "' order by Posizione;" 



 if CodiceSottopar<>"" then	
    QuerySQL="SELECT * " &_
	"FROM preNodi WHERE preNodi1.Id_Paragrafo='" & CodiceTest & "' and preNodi.Id_Sottoparagrafo='" & CodiceSottopar & "' and ID_Prenodo not in (Select ID_Prenodo from Nodi WHERE Nodi.Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "' order by Posizione;" 
   
   else									    
QuerySQL="SELECT  ID_Prenodo, Id_Mod, Id_Paragrafo, CodiceMetafora, Quesito, Eseguita, Posizione, Scadenza, Img, Files, Id_Sottoparagrafo " &_
	"FROM dbo.preNodi WHERE  Id_Paragrafo='" & CodiceTest & "' and ID_Prenodo not in (Select ID_Prenodo from dbo.Nodi WHERE Id_Arg='" & CodiceTest & "' and Id_Stud='"&Session("CodiceAllievo")&"') and Id_Mod='" &Modulo  & "' order by Posizione;" 
end if
 


	'response.write(QuerySQL)			
	Set rsTabellaNodi = ConnessioneDB.Execute(QuerySQL)
	 
		 
		 %>
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   <% if rsTabellaNodi.eof and rsTabellaNodi.bof then%>
                     <div class="alert alert-success">
                     	
						  <a rel="popover" data-trigger="hover" data-content="Crea Nodo" title="Nodi" target="_blank" href="../cNodi/inserisci_nodo.asp?Tipo=0&id_classe=<%=id_classe%>&Cartella=<%=Cartella%>&Stato=0&Stato0=0&CodiceTest=<%=CodiceTest%>&Capitolo=<%=TitoloCapitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">
						  <b><%=response.write("Hai gia' svolto tutti i compiti assegnati. Se vuoi inserire un nodo libero clicca")%></b>
						  &nbsp;qui&nbsp;<i class="icon-edit"></i>

                     </div>
       			
					<%end if
                    'response.write(QuerySql & " " &Paragrafo)
                    end if
                i=0
                'paragrafo=rsTabellaNodi(2)



	do while not rsTabellaNodi.eof%>
		    <div class="box-content"> 
                     <% if (i=0) then %>
 					 <ul>
					 <%end if %>
                      
                      <li style="line-height:10px;">
					<a title="Scade il <%=rsTabellaNodi("Scadenza")%>" href="inserisci_nodo.asp?by_UECDL=<%=by_UECDL%>&Tipo=0&Quesito=<%=rsTabellaNodi(4)%>&Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&prenodo=1&ID_Prenodo=<%=rsTabellaNodi("ID_Prenodo")%>&Scadenza=<%=rsTabellaNodi("Scadenza")%>&CodiceSottopar=<%=rsTabellaNodi("Id_Sottoparagrafo")%>"><%=rsTabellaNodi(4)%>
					</a> 
							
				<%
	i=i+1
	rsTabellaNodi.movenext 
loop%>
               </ul>     <br>
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

