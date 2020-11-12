<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci nodo</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
    
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
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	
	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	 <script src="../../js/eak_app_dem.min.js"></script>
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	 <!--<script src="../js/plugins/plupload/plupload.full.js"></script>
	<script src="../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../js/plugins/mockjax/jquery.mockjax.js"></script>
    -->
    
    
    
    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
<script language="javascript" type="text/javascript"> 
function showText4() {window.alert("Non adesso grazie! Troppo tardi o troppo presto !")
location.href="../home.asp"
 
 }
 
  function showText5(proroga) {
	 window.alert("Attenzione il compito era scaduto! Il prof. ti ha concesso una proroga"); 
	 getElement1();
 }
 </script>
    
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
</head>
<% Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
  
  'Request.Cookies("Dati")("CodiceTest")= Codice_Test
  
  Codice_Test=Request.QueryString("CodiceTest")
   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito") ' prenodo
   Id_Stud=Session("CodiceAllievo")
  prenodo=Request.QueryString("prenodo")
  Cartella=Request.QueryString("Cartella") 
   Scadenza=Request.QueryString("Scadenza")  
  'Response.Cookies("Dati")("StrConn")="../database/Copiaditestonline.mdb"
  Num = Request.QueryString("Num")
  Num=Num+1
  ID_Prenodo=Request.QueryString("ID_Prenodo") 
  ' ID_Prenodo=18
  by_UECDL=Request.QueryString("by_UECDL")%>
  
 <% Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function
if not(strcomp(Scadenza,"gg/mm/aaaa")=0 or (Scadenza="")) then
' se non è impostata la scadenza la pongo uguale ad oggi per evitare errori
      Scadenza=Cdate(Request.QueryString("Scadenza"))
   else
      Scadenza=gira_data()
end if
  
   
 Function verifica_eccezione(id_prenodo,id_stud,data)
   QuerySql="Select * from Eccezioni_Nodi where Id_Prenodo="&id_prenodo&" and Id_Stud='"&id_stud&"';"
  '  response.write(QuerySql&"<br>")
   set rsTabella=ConnessioneDB.execute(QuerySql)
   if not rsTabella.eof then ' se è presente l'eccezione verifico se è ancora valida
  
	  if Datediff("d",data,rsTabella("Scadenza"))>=0 then
            verifica_eccezione=1
			Proroga=rsTabella("Scadenza")
		else
         verifica_eccezione=0
		end if
   else
       verifica_eccezione=0
   end if
end function

Function data_eccezione(id_prenodo,id_stud,data)
   QuerySql="Select * from Eccezioni_Nodi where Id_Prenodo="&id_prenodo&" and Id_Stud='"&id_stud&"';"
   
   set rsTabella=ConnessioneDB.execute(QuerySql)
   if not rsTabella.eof then ' se è presente l'eccezione verifico se è ancora valida
        ' response.write("InFunct, Scadenza="& rsTabella("Scadenza") &"Data="&data&"Datedif="& Datediff("d",rsTabella("Scadenza"),data))
	data_eccezione=rsTabella("Scadenza")
	  
   end if
end function

   
   
   
     Data = gira_data()
	  if ID_Prenodo<>"" then 
     eccezione=verifica_eccezione(ID_Prenodo,Id_Stud,Data)
	  end if
   

  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %> 
  
     <%
	 
	  if ID_Prenodo<>"" then 
	    if Datediff("d",Scadenza,Data)>0 and eccezione=0 then 
		%> 
        <BODY onLoad="showText4();">%>
        </BODY>
	    <%else%>    
         
   		     <% if eccezione=1 then
			 Proroga=data_eccezione(ID_Prenodo,Id_Stud,Data)
			 %>
			     <body onLoad="showText5(<%=Proroga%>);" >
      		    </body>
			  <% end if%>
	    <% end if %>
       <% else ' se sono chaiamato da scegli_azione_test per inserire un nodo a piacere(libero) allora non faccio controllo sulla scadenz
	   %>
	    <body class='theme-<%=session("stile")%>'>
        
        <%end if%>
  <% end if%>

  <%
  'Response.write "Sessione: "&Session("CartellaIniz")
 ' Response.write "<br>Cartella: "&Cartella
  if Session("CartellaIniz") <> Cartella and Session("Admin") <> true then
	Response.write "<script>alert('Con questo utente non puoi inserire compiti in questa classe');window.location.href='"&Request.ServerVariables("HTTP_REFERER")&"'</script>"
	'Response.write "ifasdfsdfsd"
  
  end if
%>
  

	<div id="navigation">
     
   
	
		
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="glyphicon-snowflake"></i> Inserisci nodo della rete  </h1> 
                    
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
							<a href="#">Verifica</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							<a href="#">Inserisci nodo</a>
                             
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
				        <h3> <i class="icon-reorder"></i>  "<%=Capitolo%> : <%=Paragrafo%>"</h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				 <form method="POST" class='form-vertical' action="inserisci_nodo1.asp?by_UECDL=<%=by_UECDL%>&prenodo=<%=prenodo%>&ID_Prenodo=<%=ID_prenodo%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=Codice_Test%>&StringaConnessione=../database/Copiaditestonline.mdb&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>" > <!-- Alla pressione del bottone INVIA il form chiama la pagina che verifica il login accedendo al data base specificato dalla stringa di connessione-->
 
  
  	<div class="control-group">
		<label for="textfield" class="control-label"><b>Chi</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtChi"  placeholder="Nome del nodo concettuale" class="input-xxlarge" maxlength="149" value="<%=response.write(left (Quesito,len(Quesito)))%>">
	    	</div>
    </div>
  
  <div class="control-group">
		<label for="textfield" class="control-label"><b>Cosa</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR1Cosa"  placeholder="Azione caratterizzante" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Dove</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR2Dove"  placeholder="Luogo dell'azione" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Quando</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR3Quando"  placeholder="Tempo dell'azione" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Come</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR4Come"  placeholder="Modo dell'azione" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Perch&egrave;</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtR5Perche"  placeholder="Motivo dell'azione" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
     <div class="control-group">
		<label for="textfield" class="control-label"><b>Quindi</b></label>
	     	<div class="controls">
		     	<input type="text" name="txtREQuindi"  placeholder="Senso dell'azione" maxlength="149" class="input-xxlarge">
	    	</div>
    </div>
    
    <div class="control-group">
		<label for="textfield" class="control-label"><b>Sintesi</b></label>
	     	<div class="controls">
		     <p><textarea class="input-block-level" rows="6" name="S1"  placeholder="Unisci i vari livelli di significato in una spiegazione di riepilogo"></textarea></p>
 
	    	</div>
    </div>
    
	 <div class="form-actions">
									 <button type="submit" class="btn btn-primary" name="B1">Invia</button>
	</div>
   
 
</form> 
                   
                   
 
		  			 
                             
			       
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

