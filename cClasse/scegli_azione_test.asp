<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Scegli Verifica </title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	  <meta charset="UTF-8">

	<!-- Bootstrap -->
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

	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
   
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>
<% Response.Buffer=True 
   on error resume next   
 
  Cartella=Request.QueryString("Cartella")
  DataTest = Request.Cookies("Dati")("DataTest")
  CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Nome=Request.Cookies("Dati")("Nome")
  Cognome=Request.Cookies("Dati")("Cognome")
  
  CodiceTest = Request.QueryString("CodiceTest") 
  if instr(CodiceTest,",")>0 then ' se devo tagliare prima della virgola per bug misterioso del doppio codice
  CodiceTest=left(CodiceTest,instr(CodiceTest,",")-1)
  end if
  Response.Cookies("Dati")("CodiceTest")=CodiceTest
  Capitolo=Request.QueryString("Capitolo") 
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  id_classe=request.querystring("id_classe")
  Tutti=request.querystring("Tutti")
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  Stato = Request.QueryString("Stato") 
 ' if Stato="" then
'    Stato=0
' end if	
  
  
 

%>
 <body class='theme-<%=session("stile")%>'>
      
	<div id="navigation">
     
        <% 
		
 
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
						<h1> <i class="icon-signal"></i> Verifica il tuo apprendimento... </h1> 
                    
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
							<a href="#"><b>Verifica</b></a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							<a href="#"><b><%=Capitolo%></b></a>
                            <i class="icon-angle-right"></i>
						</li>
                        <%if Tutti="" then%>
                         <li>
							<a href="#"><b><%=Paragrafo%></b></a> 
                             <i class="icon-angle-right"></i>   
						</li>
                          <li>
							<a href="#"><b><%=Sottoparagrafo%></b></a>    
						</li>
                        <%end if%>
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
				        <h3> <i class="icon-reorder"></i>Scegli il tipo di attivit&agrave;</h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     
                     <div class="accordion" id="accordion2">
					
                    
                    
                    <% 
					 
					 
					 					 
 QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
 Set rsTabellaSetting = ConnessioneDB.Execute(QuerySQL)
 ValidaTest=rsTabellaSetting("ValidaQuiz")
					 
					if strcomp(Tutti,"1")=0 then%>
                    
                    
      <%  
	  
	  
	   QuerySql="SELECT  Moduli.ID_Mod,Moduli.Titolo, Paragrafi.ID_Paragrafo, Paragrafi.Titolo as [Tit],URL_O,URL_OL, Moduli.Posizione as [posMod],Paragrafi.Posizione as [posPar] " &_
" FROM Paragrafi, Moduli, Classi_Moduli_Paragrafi " &_
" WHERE  Classi_Moduli_Paragrafi.Id_Modulo=Moduli.ID_Mod and Classi_Moduli_Paragrafi.Id_Paragrafo=Paragrafi.ID_Paragrafo " &_
" And Moduli.ID_Mod='" & request.QueryString("Modulo") &"' order by Moduli.Posizione, Paragrafi.Posizione ;"

'response.write("<br>"&QuerySql & " " &Paragrafo)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)

	  
	  
	  
	  
	  Dim vf(),rs(),rm()

					 QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
		   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & request.QueryString("Modulo") & "' "
		 ' response.write("<br>"&QuerySQL)
		   set rsTabellaNQ=ConnessioneDB.Execute(QuerySQL)
		 if not isnull(rsTabellaNQ(0)) then
		   NumQuiz=rsTabellaNQ(0)
		 else
		  NumQuiz=0
		 end if  
		   redim vf(NumQuiz+1),rs(NumQuiz+1),rm(NumQuiz+1) %>           
                    
                    <% 
					 
				 %>
				 <%	if ValidaTest=1 or session("admin")= true then%>
                    <div class="accordion-group">
										<div class="accordion-heading">
                                        
                                        
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse0">
												<center><b>Tutti i Quiz</b></center>
											</a>
										</div>
										<div id="collapse0" class="accordion-body collapse">
											<div class="accordion-inner">
                                            <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Quadro generale :  <%=request.QueryString("Capitolo")%> </h3>
							</div>
							<div class="box-content nopadding">
                            
                            <table class="table-bordered table-condensed">
                            <thead>
                          
							<th>Paragrafo</th>
							<th><a href="../cDomande/esegui_test_vf.asp?verifica=1&testnodo=0&Stato=1&Tutti=1&Id_Classe=<%=Id_Classe%>&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&vf=1&NUMTEST=-1">
                            Vero/Falso</a></th>
						   <th> <a href="../cDomande/esegui_test.asp?verifica=1&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&NUMTEST=-1">
						   Singola</a></th>
						   <th>  <a href="../cDomande/5_esegui_test_multiple.asp?verifica=1&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&rm=1&NUMTEST=-1">
						   Multipla</a></th>
                            
                            
                            
                            </thead>
                            <tbody>
                            <tr>
                              <tr><td><b>Riepilogo per batterie e tipologie</b></td>
                         <td><b>
                           <%for i=1 to NumQuiz 'Riga di riepilogo di tutto il modulo
						   ' comincio con il VF						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Mod='" & request.QueryString("Modulo") & "' and Domande.Segnalata=0 and Domande.VF=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   
							   set rsTabellaNVF=ConnessioneDB.Execute(QuerySQL)
							   NumVF=rsTabellaNVF(0)
							   vf(i)=NumVF %>
                               <a href="../cDomande/esegui_test_vf.asp?verifica=1&testnodo=0&Stato=1&Tutti=1&Id_Classe=<%=Id_Classe%>&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&vf=1&NUMTEST=<%=i%>">
                               <%
							   if vf(i)<9 then
							   response.write(vf(i)&"&nbsp;")
							   else
							   response.write(vf(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if				   
						   next %>
                          </b></a> </td>
                            <td><b>
                           <%for i=1 to NumQuiz ' singole						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Mod='" & request.QueryString("Modulo") & "' and Domande.Segnalata=0 and Domande.VF=0 and Domande.Multiple=0 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRS=ConnessioneDB.Execute(QuerySQL)
							   NumRS=rsTabellaNRS(0)
							   rs(i)=NumRS %>
                                <a href="../cDomande/esegui_test.asp?verifica=1&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&NUMTEST=<%=i%>">
							   <%  if rs(i)<9 then
							   response.write(rs(i)&"&nbsp;")
							   else
							   response.write(rs(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if					   
						   next %>
                           </a></td></b>
                           
                            <td><b>
                           <%for i=1 to NumQuiz ' multiple						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Mod='" & request.QueryString("Modulo") & "' and Domande.Segnalata=0 and Domande.Multiple=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRM=ConnessioneDB.Execute(QuerySQL)
							   NumRM=rsTabellaNRM(0)
							   rm(i)=NumRM %>
							     <a href="../cDomande/5_esegui_test_multiple.asp?verifica=1&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&rm=1&NUMTEST=<%=i%>">
								<%if rm(i)<9 then
							   response.write(rm(i)&"&nbsp;")
							   else
							   response.write(rm(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if					   
						   next %>
                          </a> </td></b>
                         </tr>                        
                            </tr>
                           <%do while not rsTabella.eof%> 
                           <tr><td><%=rsTabella.fields("Tit")%></td>
                           
                           <td>
                           <%for i=1 to NumQuiz ' comincio con il VF						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Arg='" & rsTabella("ID_Paragrafo") & "' and Domande.Segnalata=0 and Domande.VF=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNVF=ConnessioneDB.Execute(QuerySQL)
							   NumVF=rsTabellaNVF(0)
							   vf(i)=NumVF  %>
                               <a href="../cDomande/esegui_test_vf.asp?verifica=1&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&vf=1&NUMTEST=<%=i%>">
                               <%
							    if vf(i)<9 then
							   response.write(vf(i)&"&nbsp;")
							   else
							   response.write(vf(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if					   
						   next %>
                          </a> </td>
                            <td>
                           <%for i=1 to NumQuiz ' singole						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Arg='" & rsTabella("ID_Paragrafo") & "' and Domande.Segnalata=0 and Domande.VF=0 and Domande.Multiple=0 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRS=ConnessioneDB.Execute(QuerySQL)
							   NumRS=rsTabellaNRS(0)
							   rs(i)=NumRS  %>
                               <a href="../cDomande/esegui_test.asp?verifica=1&testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&NUMTEST=<%=i%>">
                               <%
							    if rs(i)<9 then
							   response.write(rs(i)&"&nbsp;")
							   else
							   response.write(rs(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if				   
						   next %>
                          </a> </td>
                           
                            <td>
                           <%for i=1 to NumQuiz ' multiple						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							    " WHERE Domande.Id_Arg='" & rsTabella("ID_Paragrafo") & "' and Domande.Segnalata=0 and Domande.Multiple=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRM=ConnessioneDB.Execute(QuerySQL)
							   NumRM=rsTabellaNRM(0)
							   rm(i)=NumRM  %>
                               <a href="../cDomande/5_esegui_test_multiple.asp?verifica=1&testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&rm=1&NUMTEST=<%=i%>">
                               <%
							    if rm(i)<9 then
							   response.write(rm(i)&"&nbsp;")
							   else
							   response.write(rm(i))
							   end if
							   if i<>NumQuiz then							   
							   response.write(" - ")		
							   end if				   
						   next %>
                           </a></td>
                           
                           
		 				</tr>
						   <% rsTabella.movenext%>
                         <%loop

						 
						 %>
                         
                       
   		 					 
						 
                          </tbody>
                          </table>
        
							</div>
						</div>
                                             
                         <%end if ' if ValidaTest=1%>                    
											</div>
										</div>
									</div>	
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    
                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse1">
												<center>
												<b>Quiz Vero/Falso..</b>
												</center>
											</a>
										</div>
										<div id="collapse1" class="accordion-body collapse">
											<div class="accordion-inner">
                                             <p> <ul> 
 
 	  <li><a href="../cDomande/esegui_test_vf.asp?verifica=0&testnodo=0&Stato=1&Tutti=1&Id_Classe=<%=Id_Classe%>&Cartella=<%=Cartella%>&CodiceTest=1_0&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&vf=1&NUMTEST=-1">
                            Verifica tutti</a></li>
	  <li><a href="../cDomande/esegui_test_vf.asp?Lingua=it&CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Verifica</a> 
	  (<a href="../cDomande/esegui_test_vf.asp?Lingua=en&CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">in English</a>) </li> 
      <% if ValidaTest=1 then%>
         <li><a href="../cDomande/esegui_test_vf.asp?Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo qualità</a>
		 <a href="../cDomande/esegui_test_vf.asp?Lingua=en&Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
		 </li>
        <%end if%> 
	 
 
    
  <% if Session("Admin")=True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin (Beta)</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a>
      </li>
	  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&vf=1">Modifica</a>
      <i class="glyphicon-warning_sign"></i></li>
       
      
      <% end if%>
</ul>
</p> 
                                             
                                             
											</div>
										</div>
									</div>	
                                    
                                    
                                    
                                    
                                    
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse2">
												<center><b>Quiz a risposta singola</b></center>
											</a>
										</div>
										<div id="collapse2" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p><ul> 
										
			 
											
	   <li><a href="../cDomande/esegui_test.asp?verifica=0&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=1_0&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&NUMTEST=-1">
						   Verifica tutti</a></li>
	  <li><a href="../cDomande/esegui_test.asp?Lingua=it&CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Esegui</a>
	  <a href="../cDomande/esegui_test.asp?Lingua=en&CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">(in English)</a></li> 
	 <% if ValidaTest=1 then%>
	    <li><a href="../cDomande/esegui_test.asp?Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo qualità</a>
		<a href="../cDomande/esegui_test.asp?Lingua=en&Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
		</li> 
        <%end if%>
        
    
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin (Beta)</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a></li>
	  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modifica </a></li>
	  </FIELDSET>
      <% end if%>
</ul>
</p> 
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse3">
												<center><b>Quiz a risposta multipla</b></center>
											</a>
										</div>
										<div id="collapse3" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p><ul>  
 
		<li>
		    <a href="../cDomande/5_esegui_test_multiple.asp?verifica=0&Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=1_0&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&rm=1&NUMTEST=-1">
						   Verifica tutti</a> 	
		</li>
	  <li><a href="../cDomande/5_esegui_test_multiple.asp?CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>">Verifica</a>
	  <a href="../cDomande/5_esegui_test_multiple.asp?Lingua=en&CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>">(in English)</a></li> 
      <% if ValidaTest=1 then%>
        <li><a href="../cDomande/5_esegui_test_multiple.asp?Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo qualità</a>
		<a href="../cDomande/5_esegui_test_multiple.asp?Lingua=en&Tutti=1&Verifica=1&CodiceTest=1_0&Stato=1&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
		</li> 
        <%end if%>
	 
	  
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin (Beta)</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a></li>  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&rm=1">Modifica </a></li>
   </FIELDSET>
     <% end if
  %> 
</p> </ul>
                                         
                                              
											</div>
										</div>
									</div>
                                    
     <% if Session("Admin")= True then %>                                 
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse4">
												<center><b>Quiz immaginario (Solo Admin)</b></center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p> <ul>

	  <li><a href="../cDomande/esegui_test_immagini.asp?CodiceTest=1_0&Stato=1&Tutti=<%=Tutti%>&Modulo=<%=Modulo%>">Verifica</a></li> 
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin(Beta)</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test_img.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Mescola</a></li>
      <% end if
  %>
</ul>
</p> 
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse5">
												<center><b>Nodi e Link</b></center>
											</a>
										</div>
										<div id="collapse5" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p> <ul>
  <li><a href="../cNodi/inserisci_nodo.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Inserisci un nodo della rete</a></li> 
 <!--  <li><a href="studente_quiz.asp?testnodo=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modifica i nodi della rete</a></li>-->
 <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
  <li><a href="../cNodi/inserisci_collegamento.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Inserisci un collegamento nella rete</a></li> <%'response.write("Paragrafo="&Paragrafo)%>
 </FIELDSET>
  <% end if%>
</p> </ul>
                                         
                                              
											</div>
										</div>
									</div>
                                    
  <%end if 'if session(admin)=true%>                             
                                    
                                    
 <%else ' if Tutti=1%>
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse1">
												<center><b>Quiz Vero/Falso</b></center>
											</a>
										</div>
										<div id="collapse1" class="accordion-body collapse">
											<div class="accordion-inner">
                                             <p> <ul><li>
 <a href="../cDomande/inserisci_test_vf.asp?Tipo=0&Multiple=0&VF=1&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Inserisci</a>
 </li>
 
 	  <li><a href="../cDomande/esegui_test_vf.asp?Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Verifica</a>
	  <a href="../cDomande/esegui_test_vf.asp?Lingua=en&Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">(in English)</a></li> 
	 
	 
	   <li><a href="../cDomande/esegui_test_vf.asp?Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo qualità</a>
	   <a href="../cDomande/esegui_test_vf.asp?Lingua=en&Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
	   </li> 
	  
 <!-- <li><a href="studente_quiz.asp?VF=1&testnodo=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>">Modifica</a></li>-->
    
  <% if Session("Admin")=True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Mescola</a></li>
	  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&vf=1">Modifica </a></li>
	 
      
       <li>Crea file.csv per esportare il test    <label for="select" class="control-label"> <small> (url=https://elexpo.net/app/csv/Nome del file.csv)</small></label>
      <i class="glyphicon-warning_sign"></i></li>
	   
      
      <form method="POST" action="../cDomande/esegui_test_csv_vf.asp?testnodo=0&Stato=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&vf=1" class="form-horizontal" >

								<div class="control-group">
										 
                                       
										<div class="controls">
											<input type="text" name="txtCSV" id="textfield" class="input-xxlarge" placeholder ="Inserisci nome del file" >
                                            <input type="text" name="NUMTEST"  class="input-small" placeholder ="Num da 1 a 4" >
                                            
											<input type="submit" value="Genera" class="btn">
										</div>
									</div>
								</form>
 				 </FIELDSET>
        <% end if%>     
      
      
      
      
</ul>
</p> 
                                             
                                             
											</div>
										</div>
									</div>	
                                    
                                    
                                    
                                    
                                    
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse2">
												<center><b>Quiz a risposta singola</b></center>
											</a>
										</div>
										<div id="collapse2" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p><ul> 
<li><a href="../cDomande/inserisci_test.asp?Multiple=0&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Inserisci</a></li> <!-- se il login è corretto richima la pagina per inserire le domande del test -->
	  <li><a href="../cDomande/esegui_test.asp?Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Verifica</a>
	  <a href="../cDomande/esegui_test.asp?Lingua=en&Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a></li> 
	 
	   <li><a href="../cDomande/esegui_test.asp?Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo Qualità</a>
	   <a href="../cDomande/esegui_test.asp?Lingua=en&Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
	   </li> 
	  
 <!-- <li><a href="studente_quiz.asp?testnodo=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&Cartella=<%=Cartella%>">Modifica</a></li>-->
    
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Mescola</a></li>
	  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Modifica (beta)</a></li> 
      
      
       <li>Crea file.csv per esportare il test    <label for="select" class="control-label"> <small> (url=https://elexpo.net/app/csv/Nome del file.csv)</small></label>
      <i class="glyphicon-warning_sign"></i></li>
	   
      
      <form method="POST" action="../cDomande/esegui_test_csv.asp?testnodo=0&Stato=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&vf=1" class="form-horizontal" >

								<div class="control-group">
										<div class="controls">
											<input type="text" name="txtCSV" id="textfield" class="input-xxlarge" placeholder ="Inserisci nome del file" >
                                            <input type="text" name="NUMTEST"  class="input-small" placeholder ="Num da 1 a 4" >
                                            
											<input type="submit" value="Genera" class="btn">
										</div>
									</div>
								</form>
      

	  </FIELDSET>
      <% end if%>
</ul>
</p> 
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse3">
												<center><b>Quiz a risposta multipla</b></center>
											</a>
										</div>
										<div id="collapse3" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p><ul>  
 
<li><a href="../cDomande/inserisci_test.asp?Multiple=1&Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&id_classe=<%=id_classe%>">Inserisci</a></li> <!-- se il login è corretto richima la pagina per inserire le domande del test -->
	  <li><a href="../cDomande/5_esegui_test_multiple.asp?Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Verifica</a>
	  <a href="../cDomande/5_esegui_test_multiple.asp?Lingua=en&Verifica=0&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a></li> 
	 
	   <li><a href="../cDomande/5_esegui_test_multiple.asp?Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>">Controllo Qualità</a>
	   <a href="../cDomande/5_esegui_test_multiple.asp?Lingua=en&Verifica=1&CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>"> (in English)</a>
	   </li> 
	  
 <!-- <li><a href="studente_quiz.asp?Multiple=1&testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Modifica</a></li>-->
    
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Mescola</a></li>  
	  <li><a href="../cDomande/3correggi_test.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>&rm=1">Modifica </a></li>
   </FIELDSET>
     <% end if
  %> 
</p> </ul>
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse4">
												<center><b>Quiz immaginario</b></center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p> <ul>

	  <li><a href="../cDomande/esegui_test_immagini.asp?CodiceTest=<%=CodiceTest%>&Stato=0&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Verifica</a></li> 
  <% if Session("Admin")= True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
      <li><a href="../cAdmin/mescola_test_img.asp?testnodo=0&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>&TitoloCapitolo=<%=Capitolo%>&TitoloParagrafo=<%=Paragrafo%>">Mescola</a></li>
      <% end if
  %>
</ul>
</p> 
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse5">
												<center><b>Nodi e Link</b></center>
											</a>
										</div>
										<div id="collapse5" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                                            <p> <ul>
  <li><a href="../cNodi/inserisci_nodo.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci un nodo della rete</a></li> 
  <!-- <li><a href="studente_quiz.asp?testnodo=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Modifica i nodi della rete</a></li>-->
 <% 
   link="true" ' andrebbe aggiunto campo nella tabella setting 
  if (Session("Admin")= True) or (link="true") then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><!--<h5>Admin</h5>--></font></B></LEGEND>
  <li><a href="../cNodi/inserisci_collegamento.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci un collegamento nella rete</a></li> <%'response.write("Paragrafo="&Paragrafo)%>
 </FIELDSET>
  <% end if%>
</p> </ul>
                                         
                                              
											</div>
										</div>
									</div>
                                    
                                    
                                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse6">
												<center><b>Metafore</b></center>
											</a>
										</div>
										<div id="collapse6" class="accordion-body collapse">
											<div class="accordion-inner">
                                            
                              <ul>               
                                            
                                            <% if (strcomp(Paragrafo,"Topolino ed Obiettivi")=0) then
 %>
   <!----Inizio metafora Topolino -->
        <br>
            
          <li> 
        <a href="../cMetafore/inserisci_metafore.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci una metafora</a></li> 
           
         <% if Session("Admin")= True then %>
              <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
             
           <!--   <li><a href="studente_quiz.asp?testmetafora=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Modifica una metafora</a></li>-->
              <li><a href="../cNodi/inserisci_collegamento.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci un collegamento tra metafore</a></li> <%'response.write("Paragrafo="&Paragrafo)%>
         <% end if%>
           </div></div> 
             <!----Fine metafora Topolino -->

 <% else%>
       <%' if CodiceTest = Cartella&"_U_3_5" then 
	   if (strcomp(Paragrafo,"Navigazione nella Rete della Vita")=0) then 
	   %>
   <!----Inizio metafora Navigazione -->

  <li><a href="../cMetafore/inserisci_metafore.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Inserisci una metafora</a></li> 
   
			 <% if Session("Admin")= True then %>
              <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
        <!--      <li><a href="studente_quiz.asp?testmetafora=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Modifica una metafora</a></li>-->
              <li><a href="../cNodi/inserisci_collegamento.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>">Inserisci un collegamento tra metafore</a></li> <%'response.write("Paragrafo="&Paragrafo)%>
      
 			<% end if%>
   
     <!----Fine metafora Navigazione -->
 
 	<% else
	
		'if CodiceTest = Cartella&"_U_2_7" then
		if (strcomp(Paragrafo,"Relazione Cliente Servitore")=0) then 
		 %>
	   <!----Inizio metafora Database dei Desideri -->
	 
	  
	  
	  <li><a href="../cMetafore/inserisci_metafore.asp?Tipo=0&Cartella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci una metafora</a></li> 
	   
				 <% if Session("Admin")= True then %>
				  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
				<!--  <li><a href="studente_quiz.asp?testmetafora=1&Cartella=<%=Cartella%>&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Modifica una metafora</a></li>-->
				  <li><a href="../cNodi/inserisci_collegamento.asp?Tipo=0&Cassrtella=<%=Cartella%>&Num=0&Cognome=<%=Cognome%>&Nome=<%=Nome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottoPar=<%=CodiceSottopar%>&Modulo=<%=Modulo%>">Inserisci un collegamento tra metafore</a></li> <%'response.write("Paragrafo="&Paragrafo)%>
		  
				<% end if%>
       <% else%>
    
         <div class="alert alert-error">
                       Non ci sono metafore
                     </div>
      
       <% end if%>
    
	 
    
 	<% end if%>
	
    
    <!-- Per aggiunger altre metafore copio ed incollo tutto questo codice da if Codice_TEst=... 
	   fino all'end if qui sotto, e lo incollo in questa area (tra else ed end if)-->
 	
    
 <% end if%>
 
</p> 
 <!----Fine metafora -->
 
 
</FIELDSET>
      </ul>                                      
                                            
                                         
                                              
											</div>
										</div>
									</div>
 
 <% end if%>
                                    		  
                     </div>
                     
                     
                      
                     
                    <!--
                      <div class="alert alert-error">
                     KO..
                     </div>
                     
                     <div class="alert alert-success">
                     OK
                     </div>
                     
                      -->
                     
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

