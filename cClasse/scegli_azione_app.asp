<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Scegli Apprendimento</title>   
   <meta charset="utf-8">
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
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />



   
</head>
<% Response.Buffer=True 
       
 
      
  Stato=Request.QueryString("Stato") 
  Stato0=Request.QueryString("Stato0")
  Codice_Test=Request.QueryString("CodiceTest") 
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome")
  Cartella=Request.QueryString("cartella")
  uecdl=request.querystring("uecdl")
 
 
  
  DataTest = Request.Cookies("Dati")("DataTest")
  CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
   Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  
  
   
  Response.Cookies("Dati")("CodiceTest")=CodiceTest
 

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
						<h1> <i class="icon-cloud-download"></i> Approfondisci  il tuo apprendimento </h1> 
                    
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
							<a href="#more-blank.html"><b>Approfondisci</b></a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							<a href="#more-blank.html"><b><%=Capitolo%></b></a>
                            <i class="icon-angle-right"></i>
						</li>
                         <li>
							<a href="#more-blank.html"><b><%=Paragrafo%></b></a>
                           
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
				        <h3> <i class="icon-reorder"></i>Scegli il tipo di approfondimento</h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
    <!--               
                    <ul id="myTab2" class="nav nav-tabs">
                                 
                                 	  <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Scegli azione sul quiz">Frasi <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                     
                                      <li><a href="#dropdownConFrasi" data-toggle="tab" title="I miei commenti">Consulta</a></li>
                                      
                                    </ul>
                                    </li>

                                     <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown" title="Scegli azione sul quiz">Quiz <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownInsQuiz" data-toggle="tab" title="I miei commenti">Inserisci</a></li>
                                      <li><a href="#dropdownEseQuiz" data-toggle="tab" title="I miei commenti">Esegui</a></li>
                                      <% if session("Admin")=true then%>
                                      <li><a href="#dropdownBilQuiz" data-toggle="tab" title="I miei commenti">Bilancia</a></li>  
                                      <li><a href="#dropdownMesQuiz" data-toggle="tab" title="I miei commenti">Mescola</a></li> 
                                      <% end if%>
                                    </ul>
                                    </li>
                                    
                                     <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Nodi <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownConNodo" data-toggle="tab">Consulta</a></li>
                                      <li><a href="#dropdownColNodo" data-toggle="tab">Collega</a></li>
                                               
                                              
                                    </ul>
                                    </li>
                                  
                                 
                            </ul>
  
    <div id="myTabContent2" class="tab-content">
  
   					 <div class="tab-pane fade" id="dropdownConFrasi">          
                         <div class="box-content nopadding">
                           <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Consulta Frasi
								</h4>
			 	    	</div>			  
			  		</div>     
                  </div>
           </div>
                   
                   
                   -->
                   
                   
                   
                   
                   
 
		    <div class="box-content"> 
                     
                     <div class="accordion" id="accordion2">
					
                    <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapse1">
												<center><b>Leggi</b></center>
											</a>
										</div>
										<div id="collapse1">
                                        <!-- se lo voglio collassato aggiungo sopra  class="accordion-body collapse" -->
											<div class="accordion-inner">
                                             <p> <ul><li>
										   <b>Domande del Quiz :</b><ol><li> <a href="../cDomande/spiegazione_test_1.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">Tutte</a> (<a href="../cDomande/spiegazione_test_1.asp?Lingua=en&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">In English</a>)</li>
										   <li><a href="../cDomande/spiegazione_test_1_vf.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">Vero/Falso</a> (<a href="../cDomande/spiegazione_test_1_vf.asp?Lingua=en&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">In English</a>)</li>
										   <li> <a href="../cDomande/spiegazione_test_1_rs.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">Singola</a> (<a href="../cDomande/spiegazione_test_1_rs.asp?Lingua=en&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">In English</a>)</li>
										   <li> <a href="../cDomande/spiegazione_test_1_rm.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">Multipla</a> (<a href="../cDomande/spiegazione_test_1_rm.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">In English</a>)</li></ol>
  </li>
 
 	  <li><b>Mappa Concettuale</b><ol><li>
      <a href="../cNodi/spiegazione_nodi.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">	Rete di Nodi</a> </li>
	  <li> <a target="_self" href="../cMap/spiegazione_mappa.asp?cod=<%=CodiceAllievo%>&Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>&idclasse=<%=Session("Id_Classe")%>">	Rete di Nodi interattiva</a> </li></ol>
	
      </li> <hr>
	  <% if Session("Admin")=True then %>
	   <li><b>Rete Domanda/Risposta (?)/(!)</b>(Beta test)<ol><li>
      <a href="../cFrasi/spiegazione_frasi.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>">	Frasi ...</a> </li></ol>
       </li> 
	  
   <li><b>Percorsi Operativi</b>(Beta test)<ol><li>
   <a href="../cDomande/spiegazione_test_img.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>">	Procedure...</a> </li></ol>
  </li>
    <%end if%>
      <%  
	      'if mid(Modulo,6,1)="U" then
		  if inStr(Modulo,"_U_")>0 then %>
	      <%' metafora topolino 
		    if Codice_Test =  Cartella&"_U_3_3" then
    
	  %>
	       <li><a href="../cMetafore/spiegazione_metafora_topolino.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>">	Metafore...</a>  
		   <% end if%>
		   
		     <%  if Codice_Test = Cartella&"_U_3_5" then %>
	       <li><a href="../cMetafore/spiegazione_metafora_navigazione.asp?Cartella=<%=Cartella%>&Stato=<%=Stato%>&Stato0=<%=Stato0%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>&Sottoparagrafo=<%=Sottoparagrafo%>">	Metafore...</a></li> 
		   <% end if%>
		   
		   
	  <% end if%>
    
  <% if Session("Admin")=True then %>
  <FIELDSET><LEGEND align="center"><B><font color="#000000"><h5>Admin</h5></font></B></LEGEND>
      <li>???</li>
	  
	 
	  </FIELDSET>
      <% end if%>
</ul>
</p> 
                                             
                                             
											</div>
										</div>
									</div>	
                                    
                                    
                                    
                                    
                                     
                                   
                                    		  
                     </div>
                     
                     
                      
                     
                    <!--
                      <div class="alert alert-error">
                     KO..
                     </div>
                     
                     <div class="alert alert-success">
                     OK
                     </div>
                     
                      -->
                      <br><center>
                      
                    
                      
               <h6 align="center"><i class="glyphicon-remove"></i><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6> 
               </center>
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
        
        
        	

			 
	</body>
<script type="text/javascript">

$(window).load(function () {
	   
	   $('#collapse1').click();
	    
	});
	

 	
</script>
        <!-- #include file = "../include/colora_pagina.asp" -->
         
 </html>

