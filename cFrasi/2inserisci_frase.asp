<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Crea Frase</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 <meta charset="UTF-8">

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

<!--	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>-->
  	<script src="ckeditor/ckeditor.js"></script>


	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
	<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
	
<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>



 <script language="javascript" type="text/javascript">
function showText4() {window.alert("Non adesso grazie! Troppo tardi o troppo presto !");
<% if Request.ServerVariables("HTTP_REFERER") <> "" then %>
window.location.href = "<%=Request.ServerVariables("HTTP_REFERER")%>";
<% else %>
<% if session("DB")=1 then%>
location.href="../../home.asp"
<% else%>
location.href="../../home.asp"
<%end if%>
<% end if %>
 }
 function showText5(proroga, probabilita) {
	 window.alert("Attenzione il compito era scaduto! Il prof. ti ha concesso una proroga, hai il "+ probabilita+"% di probabilita' di avere 0 punti");
	// getElement1();
 }
 </script>


</head>


<%Function gira_data()

	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())

End Function



  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
  CodiceTest=Request.QueryString("CodiceTest")
  Capitolo=Request("Capitolo")
  Paragrafo=Request("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Quesito=Request.QueryString("Quesito")
  'response.write("Quesito="&Quesito)
  prefrase=Request.QueryString("prefrase")
  ID_Prefrase=Request.QueryString("ID_Prefrase")
  Img=Request.QueryString("Img")
  cFile=Request.QueryString("cFile")
  if cFile="" then
     cFile=0
  end if
  estesa=Request.QueryString("estesa")
  if strcomp(estesa,"True")=0 then
    ext=1
  else
	ext=0
  end if

  'cFile=Request.QueryString("cFile") ' 1 se devo caricare il file .cpp senza copia ed incolla
  Cartella=Request.QueryString("Cartella")
  Id_Stud=Session("CodiceAllievo") ' per la verifica dell'eccezione


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		<!-- #include file = "../service/controllo_sessione.asp" -->
  	 
        <%


QuerySql="Select Scadenza from preFrasi where Id_Prefrase="&id_prefrase&" and Quesito='"& Quesito&"' ;"
'response.write(QuerySql)
set rsTabellaScad=ConnessioneDB.execute(QuerySql)
Scadenza=rsTabellaScad(0)
      'Scadenza=formatta_data_LO(Cdate(Request.QueryString("Scadenza")))
	  ' Scadenza=Cdate(Request.QueryString("Scadenza"))
set rsTabellaScad=nothing
if  ((strcomp(Scadenza,"gg/mm/aaaa")=0) or (Scadenza="")) then
 ' se non ? impostata la scadenza la pongo uguale a ieri per bloccare la domanda ed  evitare errori
 		    yesterday = Year(Date)&Right("0" & Month(Date),2)&  Right("0" & Day(Date() -1),2) 
			giorno= Right("0" & Day(Date() -1),2)
			mese=Right("0" & Month(Date),2)
			if (giorno=31) and (mese>1) then
			 yesterday = Year(Date)&Right("0" & Month(Date)-1,2)&  Right("0" & Day(Date() -1),2)
			end if
     ' Scadenza=gira_data()
	 Scadenza=yesterday
end if

Function verifica_eccezione(id_prefrase,id_stud,data)
   QuerySql="Select * from Eccezioni_Frasi where Id_Prefrase="&id_prefrase&" and Id_Stud='"&id_stud&"'"
  '  response.write(QuerySql&"<br>")
   set rsTabella=ConnessioneDB.execute(QuerySql)
   if not rsTabella.eof then ' se ? presente l'eccezione verifico se ? ancora valida
  ' rsTabella.movelast ' se ci sono duplicati vado a quella pi? recente
   ' Set objFSO = CreateObject("Scripting.FileSystemObject")
	'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log.txt"
	'Set objCreatedFile = objFSO.CreateTextFile(url, True)
			'	objCreatedFile.WriteLine("ID_Prefrase,Id_Stud,Data,Scadenza<br>" & ID_Prefrase & " " & Id_Stud & " " & Data & " " & rsTabella("Scadenza"))
			'	objCreatedFile.WriteLine("Datediff <br>" & Datediff("d",Data, rsTabella("Scadenza")))
				'	objCreatedFile.WriteLine("Eccezzione<br>" &eccezione)
				'objCreatedFile.Close
	     'response.write("InFunct, Scadenza="& rsTabella("Scadenza") &"Data="&data&"Datedif="& Datediff("d",rsTabella("Scadenza"),data))

	  ' if Datediff("d",rsTabella("Scadenza"),data)>=0 then una ? in formato gg/mm altra mm/gg Porco d...
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
Function data_eccezione(id_prefrase,id_stud,data)
   QuerySql="Select * from Eccezioni_Frasi where Id_Prefrase="&id_prefrase&" and Id_Stud='"&id_stud&"';"
  '  response.write(QuerySql&"<br>")
   set rsTabella=ConnessioneDB.execute(QuerySql)
   if not rsTabella.eof then ' se ? presente l'eccezione verifico se ? ancora valida
        ' response.write("InFunct, Scadenza="& rsTabella("Scadenza") &"Data="&data&"Datedif="& Datediff("d",rsTabella("Scadenza"),data))
	data_eccezione=rsTabella("Scadenza")

   end if
end function


    

%>

    <%QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
    Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
    CIAbilitato=rsTabellaCI("CIAbilitato")
    JSAbilitato=rsTabellaCI("JSAbilitato")
	RecuperoAttivo=rsTabellaCI("RecuperoAttivo")
    rsTabellaCI.close

    Data =  gira_data()
	  eccezione=verifica_eccezione(ID_Prefrase,Id_Stud,Data)
    session("eccezione")=eccezione
	session("recupero")=RecuperoAttivo

    QuerySQL = "Select CIAbilitato,Probabilita from Allievi where CodiceAllievo = '"&Session("CodiceAllievo")&"'"
    'response.write QuerySQL
    Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
    CIAbilitato2=rsTabellaCI("CIAbilitato")
    Probabilita=rsTabellaCI("Probabilita")
    session("Probabilita")=Probabilita
  ''  response.write("perobabilita="&Probabilita)
  ''  Probabilita=Probabilita*10
    'Probabilita=10
    rsTabellaCI.close
  '  response.write("probabilita="&probabilita)
    if CIAbilitato2 = 1 then
    CIAbilitato = 1
    end if

      %>
<%
'RecuperoAttivo=1
if RecuperoAttivo=1 then
else
      'response.write(Scadenza & " (1) " & Data & " (2) " & Datediff("d",Scadenza,Data) )
	   if Datediff("d",Scadenza,Data)>0 and eccezione=0 and daUpload="" then
     	    Session("Scaduto") = true
		 	Response.Redirect Request.ServerVariables("HTTP_REFERER")
      %>
        <body onLoad="showText4();">
        </body>
    <%else %>
       <% if eccezione=1 then
			      Proroga=data_eccezione(ID_Prefrase,Id_Stud,Data)
              '' response.write("probabilita="&CIAbilitato2)
                'response.write("<script>alert('ciao"&probabilita&"');</script>")
       %>
			     <body onLoad="showText5(<%=Proroga%>,<%=Probabilita%>);" >
      		 </body>
			  <% end if%>

    <%end if%>
<%end if%>
<%

 
   
    datediffe=Datediff("d",Scadenza,Data)

%>
<!--log javascript
<script>alert('<%'=datediffe%>-<%'=Scadenza%>-<%'=Data%>-<%'=ID_Prefrase%>-<%'=eccezione%>');</script>
-->
<%
 
%>
 
       <body   class='theme-<%=session("stile")%>' >
       

			<%
  'Response.write "Sessione: "&Session("CartellaIniz")
 ' Response.write "<br>Cartella: "&Cartella
  if Session("CartellaIniz") <> Cartella and Session("Admin") <> true then
	'Response.write "<script>alert('Con questo utente ("&Session("CartellaIniz")&"<>"& Cartella&") non puoi inserire compiti in questa classe');window.location.href='"&Request.ServerVariables("HTTP_REFERER")&"'</script>"
 ' 03/10/2019 COMMENTO LA RIGA SOPRA per il problema : quelli di 3c eD EX 3A hanno NELLA TABELLA ALLIEVI DUE CLASSI DIVERSE, CON LO STESSO ID_CLASSE 
  end if
%>



	<div id="navigation">

         
    
 	 
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%

 
   
  


   

	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  	<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Crea Frase </h1>

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
							<a href="#"><%=Response.write (Capitolo)%></a>
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
				        <h3> <i class="icon-reorder"></i> <%=Quesito%>
                         </h3>
			          </div>

				      <div class="box-content">
					  <% if ext=1 then ' carico testo esteso 
					   Set objFSO = CreateObject("Scripting.FileSystemObject") 
					   Const ForReading = 1, ForWriting = 2, ForAppending = 8
					   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" &cartella &"/" &Modulo&"_Esercizi/"&CodiceTest&"_"&ID_Prefrase&".txt"
					   	url=Replace(url,"\","/")
						   'response.write(url)
					   if objFSO.FileExists(url) then
						Set objTextFile = objFSO.OpenTextFile(url, ForReading)
						sReadAll="" 'pulisco sReadAll -> altrimenti rimane la vecchia spiegazione
						sReadAll = objTextFile.ReadAll
						'sReadAll=url
						objTextFile.Close
						else
						' response.write("Il file non esiste:"&url)
						 sReadAll="Il file " &url &" non esiste"
						end if
					   end if %>
							<div class="row-fluid">
								<div class="span12">
									<div class="box">
									
										<div class="box-content">
										<% if ext=1 then%>
											<%response.write(sReadAll)%>
										<%end if%>
										
										<form name="frmDocument" id="frmDocument" method="POST" >
										<textarea name="editor1" id="editor1" rows="10" cols="80">
										
										</textarea><br>
										<textarea name="txtEncode"  id="txtEncode" style="display:none;" rows="10" cols="80">
										</textarea><br>
										<input type="text" value="<%=Quesito%>" style="display:none;" id="txtQuesito">
										<!--
										<input type="text" value="<%'=ID_Prefrase%>" style="display:none;" name="ID_Prefrase" id="ID_Prefrase">
								-->
								<% if strcomp(Img,"1")=0  or Img=1 then %>
								
                                 <div class="accordion" id="accordion3">
                                 <div class="accordion-group ">
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseLinkImg">
												<center>
                                               <i class="icon-link"  title="Incolla URL Immagini"></i>&nbsp;<i class="icon-picture"  title="Incolla URL Immagini"></i>
                                                </center>

											</a>
										</div>

											<div id="collapseLinkImg" class="accordion-body in collapse">
                                            <div class="accordion-inner">
 										 		Inserisci immagini tramite <a target="_blank" href="https://postimages.org/it/">questo servizio di hosting esterno</a>.&nbsp;<font color="#000000">
												<br><b>N.B.</b>&nbsp;</font>Ridimensiona l'immagine (es.640x480) ed incolla (ctrl+v) il <b>Link Diretto</b> (secondo della lista), inoltre <b>non inserire caratteri speciali nel nome del file</b>.
												<br><b>N.B.2 Puoi inserire anche url a pagine .html .php o googledoc</b>

                                                 <div class="control-group">
                                                
												<!--
												<div class="controls">
                                                    <input name="txtImg1" id="textfield" placeholder="Incolla collegamento diretto poi clicca su Aggiungi link" class="input-xxlarge" type="text" value=""> 					 
                                                </div>
												 <div class="controls">
													<button type="button" class="btn btn-secondary" onclick="aggiungiUrl()" onchange="aggiornaStato()" id="BtnUrl" >Aggiungi link</button>
                                                </div>-->

												 <div class="controls">
                                                    <input name="txtImg1" id="textfield" placeholder="Incolla collegamento diretto " class="input-xxlarge" type="text" value=""> 					 
                                                </div>
                                                 <div class="controls">
                                                    <input name="txtImg2" id="textfield2" placeholder="Incolla collegamento diretto" class="input-xxlarge" type="text">
                                                </div>
                                                 <div class="controls">
                                                    <input name="txtImg3" id="textfield3" placeholder="Incolla collegamento diretto" class="input-xxlarge" type="text">
                                                </div>
												 
                                                </div>
                     						 </div>
                                              
										</div>
                                     </div>
                                 </div>
								 <%end if%>

										  <button type="button" class="btn btn-primary" onclick="inviaDati()" id="Btn1" >Invia</button>
										<!--
										  <button type="button" class="btn" onclick="caricaDati()" id="Btn2" >Carica</button>
										  <button type="button" class="btn" onclick="stampaDati()" id="Btn3" >Stampa</button>
										 -->
										<div class="row-fluid">
													<!--
													<div class="span1" id="conta">
													</div>--><br>
													<div class="span11">
														<div class="progress">
														<!--	<div class="bar bar-danger" style="width: 10%;"></div>
															<div class="bar bar-warning" style="width: 20%;"></div>-->
															<div class="bar bar-success" id="progressbar" style="width: 0%;"></div>
														</div>
													</div> 
												</div>


										
									</form>

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
	    <!-- #include file = "../include/pull.asp" -->
    	 

	</body>

 </html>
<script>
 
var statoLink=0; // campo non modificato
function aggiornaStato(){
	var statoLink=1;
 }
	function aggiungiUrl(){
				var url=document.getElementById('textfield').value;
				var link='<a target="_blank" href="'+url+'">Link</a>';
				alert(link);
				CKEDITOR.instances.editor1.setData(link);
	   
		}




</script>