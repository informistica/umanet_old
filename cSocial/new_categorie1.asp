<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci categorie</title>   
   
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
    <!-- Notify -->
	<script src="js/plugins/gritter/jquery.gritter.min.js"></script>
<!-- Theme framework -->
	<script src="js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="js/application.min.js"></script>
	<!-- Just for demonstration -->
	<script src="js/demonstration.min.js"></script>

	 
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />
       
 
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
  On Error Resume Next  
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
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserisci categoria</h1> 
                    
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
							<a href="../home_app.asp?id_classe=<%=session("id_classe")%>">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Inserisci categoria</a>
                            <i class="icon-angle-right"></i>
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
				      
				      <div class="box-content">
                      
 <%
 									 
	 
	 
				 
				 



' se inserisco in più sessioni devo accodare le ultime frasi dietro alle prime, quindi mi serve la posizione raggiunta per proseguire da li

function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"a"&Chr(96))
  sAns = Replace(sAns,chr(225),"a"&Chr(96))
  sAns = Replace(sAns,chr(232),"e"&Chr(96))
  sAns = Replace(sAns,chr(233),"e"&Chr(96))
  sAns = Replace(sAns,chr(236),"i"&Chr(96))
  sAns = Replace(sAns,chr(237),"i"&Chr(96))
  sAns = Replace(sAns,chr(242),"o"&Chr(96))
  sAns = Replace(sAns,chr(243),"o"&Chr(96))
  sAns = Replace(sAns,chr(249),"u"&Chr(96))
  sAns = Replace(sAns,chr(250),"u"&Chr(96)) 
  sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
  sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
  sAns=  Replace(sAns,"&","e") 
  sAns=  Replace(sAns,"/","-") 
  sAns=  Replace(sAns,"\","-") 
  sAns=  Replace(sAns,"?",".") 
  sAns=  Replace(sAns,"*","x") 
  sAns=  Replace(sAns,"<","_")
  sAns=  Replace(sAns,">","_") 
  
ReplaceCar = sAns
end function

scegli = Request.QueryString("scegli") 
txtCategorie = Request.Form("MyTextArea") 
		strText = txtCategorie
        arrLines = Split(strText, vbCrLf)
    k=1
	For Each strLine in arrLines
	     Categoria=strLine
		    
			Categoria =  ReplaceCar(Categoria) 
		       
				    QuerySQL="  INSERT INTO CAT_CAT (Id_Classe, Descrizione,Id_Social)  SELECT '" & Session("Id_Classe") & "','" & Categoria & "', " & scegli & ";"
				 
				  '  response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL 
				  
		response.write Domanda & "<br>"
       k=k+1
	Next
	
	 
%>
<%

	On Error Resume Next
		If Err.Number = 0 Then%>
	<span class="alert-success">
	<%

		Response.Write "Inserimento avvenuto! "
	Else
	%>
	<span class="alert-error">
	<%

		Response.Write Err.Description 
		Err.Number = 0
	End If


   %>
   </span>
	</font>   
	 
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

