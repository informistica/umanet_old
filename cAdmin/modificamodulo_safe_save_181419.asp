<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Modifica Moduli</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	  <meta charset="UTF-8">

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
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />-->

  


   
</head>
<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
   
	</div>
    
 <%
 on error resume next
 ID_Mod=Request.QueryString("ID_Mod")
Classe=Request.QueryString("Classe")
divid=Request.QueryString("divid")
Id_Classe=Request.QueryString("Id_Classe")
Caricato=Request.QueryString("Caricato")
Conta=Request.QueryString("Conta")
byUmanet=Request.QueryString("byUmanet")
If Conta="" then
  Conta=0
end if  
URL_OL=Request.QueryString("URL_OL")
 %>    
  
      
    
	<div class="container-fluid" id="content">
       <!-- #include file = "../include/menu_left.asp" -->
       
          <% if (URL_OL<>"") and (Request.Form("txtURL_OL")<>"") then  
 QuerySQL ="UPDATE Moduli SET URL_OL = '" & Request.Form("txtURL_OL")  & "'  WHERE ID_Mod ='" &ID_Mod&"';"
ConnessioneDB.Execute(QuerySQL)	 

 end if%>
 
 <% QuerySQL="SELECT [ID_Mod],[Titolo],[ID_Paragrafo],[Tit],[URL_O],[URL_OL],[posMod],[posPar],[Visibile] " &_
" FROM MODULI_PARAGRAFI_CLASSE1 " &_
" WHERE [ID_Mod]='" & Id_Mod&"' Order by posPar ;"


'dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'					url="C:\Inetpub\umanetroot\expo2015Server\logModuli.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.WriteLine("<br>byUmanet="&byUmanet&"<br>"& QuerySQL & " " &Paragrafo)
'				
				



'response.write( "                        "&QuerySql )
Set rsTabella = ConnessioneDB.Execute(QuerySQL)



if byUmanet="" then
 QuerySQL="SELECT Titolo, Posizione, Id_Classe,Visibile " &_
" FROM MODULI_CLASSE WHERE (((MODULI_CLASSE.Id_Classe)='"&Id_Classe&"')) ORDER BY MODULI_CLASSE.Posizione asc;"
else
 QuerySQL="SELECT Titolo, Posizione, Id_Classe,Visibile " &_
" FROM MODULI_CLASSE_UMANET WHERE (((MODULI_CLASSE_UMANET.Id_Classe)='"&Id_Classe&"')) ORDER BY MODULI_CLASSE_UMANET.Posizione asc;"

end if
'response.write("byUmanet="&byUmanet&"<br>"& QuerySQL & " " &Paragrafo)

'objCreatedFile.WriteLine(QuerySQL)

	'objCreatedFile.Close			


Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)%>
  
         
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Gestisci moduli </h1> 
                    
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
							<a href="#">Admin</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">....</a>
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
				      <h3> <i class="icon-reorder"></i>  <%response.write("Modulo : " & rsTabella("Titolo")&"<br>") %> </h3>
			          </div>
				      <div class="box-content">
                      
   <%
TitoloModulo=rsTabella("Titolo")
'paragrafo=rsTabella(2)
if rsTabella.eof and rsTabella.bof then%>
   <span class="alert-error"><b><%=response.write("Non ci sono paragrafi nel modulo")%></b></span>
<%end if%>       
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    
                    
                    
                    
                    
                    
                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapse1">
												<center>Riordina posizione e modifica visibilità moduli</center>
											</a>
										</div>
										<div id="collapse1" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                               
                                                <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Riordina i moduli nel libro</h3>
							</div>
							<div class="box-content nopadding">
                            
                            
                              <%i=0%>
<form method="POST" class='form-vertical form-bordered' action="modificamodulo1.asp?byUmanet=<%=byUmanet%>&Id_Mod=<%=Id_Mod%>&Id_Classe=<%=Id_Classe%>&URL_OL=1&divid=<%=divid%>&Conta=<%=Conta%>&Classe=<%=Classe%>&posMod=1">
<div class="control-group">
									
										<label for="textfield" class="control-label"> Ordine di visualizzazione nel libro</label>
										<div class="controls">
							
<%do while not rsTabella1.eof%> 
		<b><%=i+1%>) Modulo </b><input class="input-xxlarge" type="text" disabled="true" value="<%=rsTabella1.fields("Titolo")%>" name="txtModulo<%=i%>" >
          
      <b>  Posizione</b> <input class='input-mini'  type="text" value="<%=rsTabella1.fields("Posizione")%>" name="txtPosMod<%=i%>" > <b>Visibile (1=si/0=no)</b> <input class='input-mini'  type="text" name="txtVisibile<%=i%>" value="<%=rsTabella1.fields("Visibile")%>" >  <br>			 
<%  i=i+1 
    indice=indice+1 
	rsTabella1.movenext%>
 <%loop%>
  <br>
  <input type="submit" value=" Aggiorna" class="btn">
		 
  </div>
						</div>
 </form>

        
							</div>
						</div>  
                                       
                                                
                                                
											</div>
										</div>
									</div>
                                    
                                    
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapse2">
												<center>Gestisci Risorse modulo</center>
											</a>
										</div>
										<div id="collapse2" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                               
                                                 <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Collega risorsa per il modulo</h3>
							</div>
							<div class="box-content nopadding">
                            
                            
                            <form method="POST" class='form-vertical form-bordered' action="modificamodulo1.asp?Id_Mod=<%=Id_Mod%>&Id_Classe=<%=Id_Classe%>&URL_OL=1&divid=<%=divid%>&Conta=<%=Conta%>&Classe=<%=Classe%>&URLRISORSA=1&byUmanet=<%=byUmanet%>">
 									<div class="control-group">
										<label for="textfield" class="control-label"> Aggiungi URL di una Risorsa :</label>
										<div class="controls">
											<input type="text"  name="txtURL_OL" placeholder="Incolla url " class="input-xxlarge">
                                            <input type="submit" value="Aggiungi" class="btn">	 
										</div>
									</div>

						</form>
        
							</div>
						</div>
                    
                    
                    
                    
                    <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Carica risorsa per il modulo</h3>
							</div>
							<div class="box-content nopadding">
                            
		<form  method="POST" class='form-vertical form-bordered' name="frmDocument" ENCTYPE="multipart/form-data" ACTION="Upload/confirm_update.asp?AggRisMod=1&Classe=<%=Classe%>&Id_Mod=<%=ID_Mod%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>&byUmanet=<%=byUmanet%>">
                                
  
									<div class="control-group">
										 
										<div class="controls">
											
<b>Classe :</b> <input type="text" value="<%=Classe%>" disabled="disabled"><br>
<b>Modulo :</b> <input type="text" name="txtId_Mod "value="<%=ID_Mod%>" disabled="disabled"><br>
<b>Risorsa del Modulo :</b> 
<%
	 if rsTabella("URL_OL")&"" = ""  then
%>
<input type="text" name="txtRis" placeholder="Nessuna risorsa caricata"><br><br>
<%   else  %>
<input type="text" name="txtRis" value="<%=rsTabella("URL_OL")%>" class="input-xxlarge"><br><br>
<%   end if  %>
Aggiungi una Risorsa : <INPUT TYPE="file" name="flname" class="btn"> 
 <input  class="btn" type="Submit" name="btnUpload" value="Upload" onClick="mostra()"> 
  <br><img src="Upload/nulla.jpg" width="35" height="35" name="loading">     
                                            
                                            
										</div>
									</div>
									 
									 
								</form>
                                
							</div>
						</div>
                    
                                       
                                                
                                                
											</div>
										</div>
									</div>
                    
                    
                    
                    
                    
                    
                    
                      <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapse3">
												<center>Gestisci Risorse paragrafi</center>
											</a>
										</div>
										<div id="collapse3" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                       
                                       
                                       <%i=0
indice=0
%>

                    
                    
                      <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Collega risorse per i paragrafi </h3>
							</div>
							<div class="box-content nopadding">
                            
                            
                            <form method="POST" class='form-vertical form-bordered' action="inserisci_paragrafo.asp?Id_Mod=<%=Id_Mod%>&Id_Classe=<%=Id_Classe%>&byUmanet=<%=byUmanet%>">	
 									<div class="control-group">
										<label for="textfield" class="control-label"> Aggiungi URL di una Risorsa :</label>
										<div class="controls">
											<%do while not rsTabella.eof%> 
		<b><%=i+1%>)</b> <input type="text"  value="<%=rsTabella.fields("Id_Paragrafo")%>" name="txtIdPar<%=i%>" class="input-xxsmall">  
		
		<input type="text" disabled="true" value="<%=rsTabella.fields("Tit")%>" name="txtParagrafo<%=i%>" class="input-xlarge">
       <b> Risorsa</b> <input type="text" value="<%=rsTabella.fields("URL_O")%>" name="txtURL<%=i%>"  class="input-xlarge">  
        <b>Posizione</b> <input type="text" value="<%=rsTabella.fields("posPar")%>" name="txtPosPar<%=i%>" size="3" class="input-mini">&nbsp;
		<a title="Gestisci sottoparagrafi" href="#"><img src="../../img/Next.gif" width="14" height="13"></a> &nbsp;&nbsp;
		<a onClick="return window.confirm('Vuoi veramente cancellare il paragrafo?');"
		href="cancella_paragrafo.asp?cancella=1&Id_Par=<%=rsTabella.fields("Id_Paragrafo")%>&Id_Classe=<%=Id_Classe%>&Classe=<%=Classe%>">
		<i class="icon-trash"></i></a>
		<%
		qsl="SELECT * FROM  ParagrafiSottoparagrafi2 where Id_Paragrafo='"&rsTabella.fields("Id_Paragrafo")&"'"
		set rsTabSottoPar= ConnessioneDB.execute(qsl)
		j=0
		do while not rsTabSottoPar.eof%>
		<b><%=j+1%>)</b> <input type="text"  value="<%=rsTabSottoPar("ID_Sottoparagrafo")%>" name="txtIdSotPar<%=i%>" class="input-xxsmall">  
		
		<input type="text" disabled="true" value="<%=rsTabSottoPar("Titolo")%>" name="txtSotParagrafo<%=j%>" class="input-xlarge">
       <b> Risorsa</b> <input type="text" value="<%=rsTabSottoPar("URL")%>" name="txtSotURL<%=j%>"  class="input-xlarge">  
        <b>Posizione</b> <input type="text" value="<%=rsTabSottoPar("Posizione")%>" name="txtSotPosPar<%=j%>" size="3" class="input-mini">&nbsp;
		<a title="Gestisci sottoparagrafi" href="#"><img src="../../img/Next.gif" width="14" height="13"></a> &nbsp;&nbsp;
		<a onClick="return window.confirm('Vuoi veramente cancellare il paragrafo?');"
		href="cancella_sottoparagrafo.asp?cancella=1&Id_Par=<%=rsTabSottoPar("Id_Paragrafo")%>&Id_Classe=<%=Id_Classe%>&Classe=<%=Classe%>">
		<i class="icon-trash"></i></a>
		<%rsTabSottoPar.movenext
		loop
		%>
		  <br>			 
<%  i=i+1 
    indice=indice+1 
	rsTabella.movenext%>
 <%loop%>
  <%=i+1%>) Paragrafo <input type="text"  value="" name="txtParagrafoNuovo" size="50">Url <input type="text"  value="" name="txtUrlParagrafoNuovo" size="50">
  <input type="submit" class="btn" value="Aggiungi/Aggiorna">
		<%rsTabella.movefirst
		'i=1 %>
   </ol></p> 
										</div>
									</div>

						</form>
        
							</div>
						</div>
                    
                    
                    
                    
                    
                    
                    
                    
                    
                      <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Carica risorsa per il paragrafo</h3>
							</div>
							<div class="box-content nopadding">
                            
                            
                            <form  class='form-vertical form-bordered' name="frmDocument" METHOD="Post" ENCTYPE="multipart/form-data" action="Upload/confirm_update.asp?Classe=<%=Classe%>&Id_Mod=<%=Id_Mod%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>&AggRisPar=1&byUmanet=<%=byUmanet%>">							
 									<div class="control-group">
                                    	<%if Caricato<>"" then%>
										<label for="textfield" class="control-label"> 
                                        <%response.write("Risorsa aggiunta Aggiungi altre risorse ?<br>")%>
                                        </label>
                                        <%end if%>
										<div class="controls">
											 Paragrafo <select name="txtId_Par">
   
									   <% i=1
                                          do while not rsTabella.eof %>					
                                        <% 'session("Id_Par")=rsTabella.fields("ID_Paragrafo")%>
                                           <%if i= (cint(conta)+1) then %>
                                               <option selected value="<%=rsTabella.fields("ID_Paragrafo")%>"><%=i &") "&rsTabella.fields("Tit")%> </option>
                                         
                                           <% else%>
                                           <option value="<%=rsTabella.fields("ID_Paragrafo")%>"><%=i &") "&rsTabella.fields("Tit")%> </option>
                                            <% end if%>   
                                         <% i=i+1
                                               rsTabella.movenext
                                           loop %>
                                           </select>
                                           File : <INPUT TYPE="file"  name="flname"  ><BR><br>
                                           <input type="Submit" name="btnUpload" value="Upload" class="btn" onClick="return validate();">
                                                                        </div>
                                                                    </div>
              				</form>
        
							</div>
						</div>
                    
                                       
                                       
                                                
                                                
											</div>
										</div>
									</div>
                                    
                                    
                                    
                       
                       
                       <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapseSegnalazioniFrasi">
												<center>Gestisci segnalazioni frasi</center>
											</a>
										</div>
										<div id="collapseSegnalazioniFrasi" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                        
                     <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Seleziona paragrafo da controllare  </h3>
							</div>
							<div class="box-content nopadding">
                           
                               <div class="controls">
							      	 
				&nbsp; <i class="icon-reply"></i>
                                  <div class="controls">
                                  <ol>
											<% rsTabella.movefirst()
											i=0
											do while not rsTabella.eof%> 
                                            <li> <a href="../cFrasi/2scegli_valutazioni_frasi.asp?solosegnalate=1&Cartella=<%=Session("Cartella")%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=rsTabella("Titolo")%>&Paragrafo=<%=rsTabella("ID_Paragrafo")%>&TitoloParagrafo=<%=rsTabella("Tit")%>&Modulo=<%=rsTabella("ID_Mod")%>&tutto=1" ><%=rsTabella("Tit")%></a></li>
	 		 
										<%  i=i+1 
                                            
                                            rsTabella.movenext%>
                                         <%loop%>
                                         <%rsTabella.movefirst()%>
   									</ol>  
										</div>
                                  
					     		</div>
							</div>
                                              
						</div>
					</div>
				  </div>
                    
                 
			      </div>
                       
                                    
                                    
                    
                      <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapseQuiz">
												<center>Gestisci Quiz </center>
											</a>
										</div>
										<div id="collapseQuiz" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                       
                                       
                                       <%i=0
									    rsTabella.movefirst
									indice=0
									%>

                    <%' 
'	trovi numero quiz del modulo, all'interno del ciclo paragrafi, per ognuno conto numero di quiz vf,s,m
' while 
' for i=1 to NUM_QUIZ
'   select count (*) where vf=1 and  InQUiz=i
  ' vf(i)=rsTab(0)
  'next 
  'table
'Vettori per numeri di risposte vf(0)=numero di risposte vf del quiz n.1,....
Dim vf(),rs(),rm()

					 QuerySQL="SELECT MAX([In_Quiz]) AS [Num_Quiz] "&_
		   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
		   " WHERE Domande.Id_Mod='" & Id_Mod & "'"  ' and Domande.Multiple=0  and Domande.VF=0;"
		   set rsTabellaNQ=ConnessioneDB.Execute(QuerySQL)
		   
			 if not isnull(rsTabellaNQ(0)) then
			   NumQuiz=rsTabellaNQ(0)
			 else
			  NumQuiz=0
			 end if  
		  
		   redim vf(NumQuiz+1),rs(NumQuiz+1),rm(NumQuiz+1)
					%>
                     <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i><a href="sessioni_quiz.asp?Id_Classe=<%=Id_Classe%>&Id_Mod=<%=Id_Mod%>&byUmanet=<%=byUmanet%>"> Gestisci sessioni QUIZ</a> </h3>
							</div>
							<div class="box-content nopadding">
                            </div>
                      </div>
                    
                    
                     <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i><a href="../cDomande/4correggi_segnalazioni.asp?cod=<%=Session("CodiceAllievo")%>&Id_Classe=<%=Id_Classe%>&Id_Mod=<%=Id_Mod%>&byUmanet=<%=byUmanet%>"> Gestisci segnalazioni QUIZ</a> </h3>
							</div>
							<div class="box-content nopadding">
                            </div>
                      </div>
                    
                      <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Bilanciamento quiz:  <%=rsTabella("Titolo")%> </h3>
							</div>
							<div class="box-content nopadding">
                            
                            <table class="table-bordered table-condensed">
                            <thead>
                            <th>Paragrafo</th><th>Vero/Falso</th><th>Singola</th><th>Multipla</th>
                            
                            
                            
                            </thead>
                            <tbody>
                            <tr>
                              <tr><td><b>Riepilogo per batterie e tipologie</b></td>
                         <td><b>
                           <%for i=1 to NumQuiz 'Riga di riepilogo di tutto il modulo
						   ' comincio con il VF						   
						    QuerySQL="SELECT count(*) "&_
							   " FROM Paragrafi INNER JOIN Domande ON Paragrafi.ID_Paragrafo = Domande.Id_Arg" &_
							   " WHERE Domande.Id_Mod='" & Id_Mod & "' and Domande.Segnalata=0 and Domande.VF=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNVF=ConnessioneDB.Execute(QuerySQL)
							   NumVF=rsTabellaNVF(0)
							   vf(i)=NumVF %>
                               <a href="../cDomande/3correggi_test_new.asp?testnodo=0&Stato=1&Tutti=1&Id_Classe=<%=Id_Classe%>&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&vf=1&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
							   " WHERE Domande.Id_Mod='" & Id_Mod & "' and Domande.Segnalata=0 and  Domande.VF=0 and Domande.Multiple=0 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRS=ConnessioneDB.Execute(QuerySQL)
							   NumRS=rsTabellaNRS(0)
							   rs(i)=NumRS %>
                                <a href="../cDomande/3correggi_test_new.asp?Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
							   " WHERE Domande.Id_Mod='" & Id_Mod & "' and Domande.Segnalata=0 and Domande.Multiple=1 and (In_Quiz="&i &" or In_Quiz=-1);"
							   set rsTabellaNRM=ConnessioneDB.Execute(QuerySQL)
							   NumRM=rsTabellaNRM(0)
							   rm(i)=NumRM %>
							     <a href="../cDomande/3correggi_test_new.asp?Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=1&Tutti=1&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&rm=1&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
                               <a href="../cDomande/3correggi_test_new.asp?Id_Classe=<%=Id_Classe%>&testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&vf=1&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
                               <a href="../cDomande/3correggi_test_new.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
                               <a href="../cDomande/3correggi_test_new.asp?testnodo=0&Stato=0&Cartella=<%=Cartella%>&CodiceTest=<%=rsTabella("ID_Paragrafo")%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=rsTabella("Id_Mod")%>&rm=1&NUMTEST=<%=i%>&byUmanet=<%=byUmanet%>">
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
   
   
                    
                     
                                       
                                       
                                                
                                                
											</div>
										</div>
									</div>
                                    
                                    
                                    
                                    
                                    
                                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapse4">
												<center>Gestisci verifiche</center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
												
                          
						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Visualizza modelli </h3>
							</div>
							<div class="box-content">
								<div class="controls">

								<% QuerySQL="SELECT * from Paragrafi "&_
							   " WHERE ID_Paragrafo like '" & Id_Mod & "%' and Verifica=1 order by Posizione"
							   'response.write(QuerySQL)
							   set rsTabellaVerifiche=ConnessioneDB.Execute(QuerySQL)
								
								 %>
								<ul>
							   <%do while not rsTabellaVerifiche.eof %>
								<li><a target="_blank" href="../cFrasi/3visualizza_modello_verifiche.asp?TitoloModulo=<%=TitoloModulo%>&Modulo=<%=Id_Mod%>&CodiceTest=<%=rsTabellaVerifiche("ID_Paragrafo")%>&Cartella=<%=cartella%>&Classe=<%=Classe%>&Paragrafo=<%=rsTabellaVerifiche("Titolo")%>">
								<%=rsTabellaVerifiche("Titolo")%>
								</a>
								</li>
								<%rsTabellaVerifiche.movenext
							   loop
							   %>
								</ul>	 
								</div>
							</div>
						</div>   

						 <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Correzioni </h3>
							</div>
							<div class="box-content">
								<div class="controls">
								<ul>
								<% 'QuerySQL="SELECT * from Paragrafi "&_
							   '" WHERE ID_Paragrafo like '" & Id_Mod & "%' and Verifica=1 order by Posizione"
							   'response.write(QuerySQL)
							   'set rsTabellaVerifiche=ConnessioneDB.Execute(QuerySQL)
							 '  if not rsTabellaVerifiche.eof then
									rsTabellaVerifiche.movefirst
						    '	end if
									do while not rsTabellaVerifiche.eof %>
										<li><a target="_blank" href="../cFrasi/3correggi_verifica_paragrafo.asp?TitoloModulo=<%=TitoloModulo%>&Modulo=<%=Id_Mod%>&CodiceTest=<%=rsTabellaVerifiche("ID_Paragrafo")%>&Cartella=<%=cartella%>&Classe=<%=Classe%>&Paragrafo=<%=rsTabellaVerifiche("Titolo")%>">
										<%=rsTabellaVerifiche("Titolo")%>
										</a>
										</li>
										<%rsTabellaVerifiche.movenext
									loop
							
									%>
								</ul>	 
								</div>
							</div>
						</div>     

						<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Convalida </h3>
							</div>
							<div class="box-content">
								<div class="controls">
								<ul>
								<% 'QuerySQL="SELECT * from Paragrafi "&_
							  ' " WHERE ID_Paragrafo like '" & Id_Mod & "%' and Verifica=1 order by Posizione"
							   'response.write(QuerySQL)
							  ' set rsTabellaVerifiche=ConnessioneDB.Execute(QuerySQL)
							 ' if not rsTabellaVerifiche.eof then
							  rsTabellaVerifiche.movefirst
							 ' end if
							   do while not rsTabellaVerifiche.eof %>
								<li><a target="_blank" href="../cFrasi/3aggiorna_punteggio_verifiche.asp?TitoloModulo=<%=TitoloModulo%>&Modulo=<%=Id_Mod%>&CodiceTest=<%=rsTabellaVerifiche("ID_Paragrafo")%>&Cartella=<%=cartella%>&Classe=<%=Classe%>&Paragrafo=<%=rsTabellaVerifiche("Titolo")%>">
								<%=rsTabellaVerifiche("Titolo")%>
								</a>
								</li>
								<%rsTabellaVerifiche.movenext
							   loop
							   
							   %>
								</ul>	 
								</div>
							</div>
						</div>                        
                                       
                      <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Carica risultato verifica (Beta)</h3>
							</div>
							<div class="box-content nopadding">
                            
                            <form  class='form-vertical form-bordered' name="frmDocument" METHOD="Post" ENCTYPE="multipart/form-data" action="Upload/confirm_update.asp?Classe=<%=Classe%>&Id_Mod=<%=Id_Mod%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>&AggRisVer=1&byUmanet=<%=byUmanet%>">							
 									<div class="control-group">
                                    <%if Caricato<>"" then%>
										<label for="textfield" class="control-label"> Aggiungi risultato della verifica ?</label>
                                     <%end if%>
										<div class="controls">
				Argomento : <input type="text" size="50" name="txtVerifica" class="input-xlarge">  
        		Data : <input type="text" size="10" name="txtData"  class="input-small">  <br><br> 
          	    File del risultato : <INPUT TYPE="file"  name="flname"  ><BR><br>
           <input type="Submit" name="btnUpload" value="Upload" onClick="return validate();" class="btn"   title="Carica il file selezionato come risultato della verifica">
										</div>
									</div>

						</form>
        
							</div>
						</div>
                    
                                                
                                                
											</div>
										</div>
									</div>
                    
                    
                    
                    
                     <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapse5">
												<center>Importa compiti</center>
											</a>
										</div>
										<div id="collapse5" class="accordion-body collapse">
											<div class="accordion-inner">
												
                                               
                                        
                     <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Importa frasi (F)  </h3>
							</div>
							<div class="box-content nopadding">
                            
                               <div class="controls">
							      Visibile : <input type="text" size="50" name="txtVerifica" class="input-xlarge">  
        		
								   <input type="Submit" name="btnUpload" value="Upload" onClick="return validate();" class="btn"   title="Carica il file selezionato come risultato della verifica">
					     		</div>
							</div>
                                              
						</div>
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

