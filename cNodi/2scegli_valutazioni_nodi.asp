<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Valuta nodi</title>   
   
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
       
       
      
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
    </script>
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
    
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<%
  Response.Buffer = true
 ' On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>'>
  <% end if %>


	<div id="navigation">
     
   
	
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
 <%
  
  Cartella=Request.QueryString("Cartella") 
  Capitolo=Request.QueryString("Capitolo") 
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest") 
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  
   BoxApro=Request.QueryString("BoxApro")
  
'--------------
  'Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("Modulo")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  if left(Cartella,1)<>"" then
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if
'---------

 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="glyphicon-snowflake"></i> Valuta nodi</h1> 
                        
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
							<a href="#">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Valuta nodi</a>
                            
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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%> : <%=TitoloParagrafo%></h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	<%  QuerySQL="SELECT * " &_
"FROM preNodi WHERE preNodi.Id_Paragrafo='" & CodiceTest & "'" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySql & " " &Paragrafo)

if rsTabella.eof then%>
<span class="alert-error">
Non ci sono compiti da valutare<br>
  
</span>
<a href="javascript:history.back()">	Indietro </a> 
<%else%>


<%i=0
'paragrafo=rsTabella(2)

do while not rsTabella.eof
	'if (i=0) or (StrComp(capitolo, rsTabella(0)) <> 0) then'
	 %>
					
								
					<li>
					<p style="text-align: left">
					<a href="2inserisci_valutazioni_nodi.asp?BoxApro=<%=BoxApro%>&Tipo=0&ID_Prenodo=<%=rsTabella("ID_Prenodo")%>&NodoScelto=<%=rsTabella.fields("Quesito")%>&Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>&prenodo=1"><%=rsTabella.fields("Quesito")%>
					</a> </li>
							
				<%
	
	 
	i=i+1
	
	cap=rsTabella(1)
	'response.write(capitolo)
	rsTabella.movenext
	if not rsTabella.eof then
		c=rsTabella(1)		 
	  '  response.write(capitolo & " " & c)
			    if StrComp(cap, c) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")
                   
                   else 
                    i=0 
                   ' Response.Write("Le due stringhe sono diverse")
			       %>
			       </ol>
  </div>
				  <%
                end if   	
         end if 
		loop%>
<% end if ' if rsTabella.eof then%>
														 
						

 <br>
				 
				 
                   
                   
 
		  			   
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        

			 
	</body>

 </html>

