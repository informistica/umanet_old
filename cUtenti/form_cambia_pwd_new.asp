<%@ Language=VBScript %>

<%
    
   CodiceAllievo=Session("CodiceAllievo")
   id_classe=Session("Id_Classe")
   cartella=request.QueryString("cartella")
    Session("cartella")=cartella
	 Session("Cartella")=cartella
   classe=Session("Cartella")
    
	
   dividA=request.QueryString("dividApro")
   
  
  'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    
' tolgo sotto il controllo sessione per quando ritorno dall'upload foto profilo che mi dice sessione scaduta
%>   
 
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../var_globali.inc" -->
 
 <%
' VERIFICHIAMO SE L'UTENTE E' IDENTIFICATO (LOGGATO)

IF Session("Loggato") = True then

QuerySQL="Select * from Allievi where CodiceAllievo='"&CodiceAllievo& "';" 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)

End IF


Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
 
 

%>

<!doctype html>
<html>
<head>
   
   <title>Modifica Dati Personali</title>   
   <link rel="shortcut icon" href="../favicon.ico" />

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
<meta charset="utf-8">
    
    


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
	<!-- 29/11/2019 commento per evitare l'errore validate not a function che impedisce la chiusura del menu laterale
	<script src="../../js/eakroko.min.js"></script>
	-->
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />
       
    <!-- Script gestione iscrizioni notifiche -->
    <script src="../../script/js/push/push-subscription-manager.js"></script>

    <script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta. Esegui di nuovo il login!")
location.href="../../../../"
//location.href=window.history.back();
 }
    </script>
    <script type="text/javascript" src="../js/selezionatutti.js"></script>
    
<script language="javascript" type="text/javascript"> 
function showText3() {window.alert("Il nodo è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"
 
 }
    </script>
     
  <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
     
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

   <script>
   
   
   
function PopUpWindow(w,h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

	window.open('../upload_resize/ex2_imgprofilo.asp','../upload_resize/ex2_imgprofilo.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl)
}
// -->
 

	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}

</script>

 <script language="javascript" type="text/javascript" >
 function validate2() {
	 
	 
 if (frmDocument.txtCodiceAllievo.value=="")
	{
	   alert("Non hai lo username ");
	   frmDocument.txtCodiceAllievo.setfocus();
	   return 0;
	}
	else
	if (frmDocument.txtPwdAllievo.value=="")
	{
	   alert("Non hai la vecchia password");
	   frmDocument.txtPwdAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewCodiceAllievo.value=="")
	{
	   alert("Non hai inserito il nuovo username ");
	   frmDocument.txtNewCodiceAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd.value=="")
	{
	   alert("Non hai inserito la nuova password");
	   frmDocument.txtNewPwd.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd1.value=="")
	{
	   alert("Non hai confermato la nuova password");
	   frmDocument.txtNewPwd1.setfocus();
	   return 0;
	}else
	{
		
		 document.frmDocument.action = "modifica_pwd_new.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>";  
		
	   
		document.frmDocument.submit();
		
	 
    }
	
}


function validate3() {
	 
	 if (frmDocument3.txtNewEm.value=="")
	{
	   alert("Non hai inserito la nuova email");
	   frmDocument3.txtNewEm.setfocus();
	   return 0;
	}
 else
 if (frmDocument3.txtNewEm1.value=="")
	{
	   alert("Non hai confermato la nuova email");
	   frmDocument3.txtNewEm1.setfocus();
	   return 0;
	}
 else
 if (frmDocument3.txtNewEm1.value !=  frmDocument3.txtNewEm.value)
	{
	   alert("Le due email non corrispondono ");
	   frmDocument3.txtEm1.setfocus();
	   return 0;
	}
	else
	{
		
		 document.frmDocument3.action = "modifica_contatti.asp";  
		
	   
		document.frmDocument3.submit();
		
	 
    }
	
}
 </script>
   
   <script src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>
      

 <script language="javascript" type="text/javascript"> 
  
  function crittapwd() {
 var PwdAllievoOld=frm0.txtPwdOld.value;
 var PwdAllievo=frm0.txtNewPwd.value;
 var PwdAllievo1=frm0.txtNewPwd1.value;
 var PwdAllievoSHA256 = Sha256.hash(PwdAllievo)
 var PwdAllievo1SHA256 = Sha256.hash(PwdAllievo1)
 var PwdAllievoSHA256Old = Sha256.hash(PwdAllievoOld)
 
 
 if (PwdAllievoOld=="")
	{
	   alert("Non hai inserito la password in uso");
	   return 0;
	}
 else if (PwdAllievo=="")
	{
	   alert("Non hai inserito la nuova password ");
	   return 0;
	}
 else if (PwdAllievo1=="")
	{
	   alert("Non hai confermato la nuova password ");
	   return 0;
	}
 else if (PwdAllievoSHA256=="")
	{
	   alert("Password non crittografata");
	   return 0;
	}
 else if (PwdAllievoSHA256 != PwdAllievo1SHA256)
	{
	   alert("Password non corrispondenti");
	   return 0;
	}
 else
 
	{
    document.frm0.action = "../cUtenti/modifica_pwd_new.asp?stato=<%=stato%>&cla=<%=cla%>&id_classe=<%=id_classe%>&divid=<%=divid%>&PwdAllievoSHA256="+PwdAllievoSHA256+"&PwdAllievoSHA256Old="+PwdAllievoSHA256Old;
	 
	document.frm0.submit();
		
	 
    }
	
}
  </script>

  


   
</head>

<%
  Response.Buffer = true
  'On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="")then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
     <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
  <% end if %>

	<input type="text" style="display: none;" id="CodiceAllievo" value="<%=session("CodiceAllievo")%>">

	<div id="navigation">
     
   
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
       
          
         
	</div>
	
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h3> <i class="icon-user"></i> &nbsp;Modifica Dati Personali</h3> 
                    
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
							<a href="#">Impostazioni</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Modifica Profilo</a>
                           
						</li>
                         
					</ul>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>
				
				<br>
		         
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				 
				      <div class="box-content">
                     		 	 
				 
						<!--<div class="box-content">-->
                      
						<% ' contenuto pagina
						%>
						
						<div class="row-fluid">
							<div class="bs-docs-example"> 
                 
								<ul id="myTab2" class="nav nav-tabs">
                                    <li class="active"><a href="#profileL" data-toggle="tab">Login</a></li> 
                                    <li><a href="#profileC" data-toggle="tab">Contatti</a></li>  
                                    <li><a href="#profileP" data-toggle="tab">Profilo</a></li>   
									<% if Session("DB") <> 1 then %><li><a href="#profileA" data-toggle="tab">Associazioni</a></li> <%end if%>
									<li><a href="#profileN" data-toggle="tab">Notifiche</a></li>   
								</ul>
                            
								<div id="myTabContent2" class="tab-content">
									
									
									<% 
										   CodiceAllievo=Session("CodiceAllievo")
											QuerySQL="SELECT * " &_
									" FROM Allievi " &_
									" WHERE Allievi.CodiceAllievo='" & CodiceAllievo & "'"

									Set rsTabella = ConnessioneDB.Execute(QuerySQL)

									cognome = rsTabella.fields("Cognome")
									nome = rsTabella.fields("Nome")
									
									
									if strcomp(CodiceAllievo,"ospite")<>0 then ' visualizzo solo se non sono ospite 
									
									%>              
                             
									<div class="tab-pane fade" id="profileP">
									
									<!----Inizio -->           
									<div class="row-fluid">
										<div class="span12">
											<div class="box box-color box-bordered">
												<div class="box-title">
													<h3><i class="icon-user"></i>Modifica Profilo</h3>
												</div>
												<div class="box-content nopadding">
													<!-- #include file = "../cClasse/studente_domande_include/2_modifica_profilo_1.asp" -->
												</div>
											</div>
										</div>
									</div>
									<!-- >fine form -->              
									</div>
                              
									<div class="tab-pane active" id="profileL">
									<!----Inizio -->           
									<div class="row-fluid">
										<div class="span12">
											<div class="box box-color box-bordered">
												<div class="box-title">
													<h3><i class="icon-user"></i>Modifica Login</h3>
												</div>
												<div class="box-content nopadding">
													<!-- #include file = "../cClasse/studente_domande_include/2_modifica_login_1.asp" -->
												</div>
											</div>
										</div>
									</div>
									<!-- >fine form -->                 
									</div>
                               
									<div class="tab-pane fade" id="profileC">
									<!----Inizio -->           
									<div class="row-fluid">
										<div class="span12">
											<div class="box box-color box-bordered">
												<div class="box-title">
													<h3><i class="icon-user"></i>Modifica Contatti</h3>
												</div>
												<div class="box-content nopadding">
													<!-- #include file = "../cClasse/studente_domande_include/2_modifica_contatti_1.asp" -->
												</div>
											</div>
										</div>
									</div>
									<!-- >fine form -->              
									</div>
									
									<% if Session("DB") <> 1 then %>
									<div class="tab-pane fade" id="profileA">
									<!----Inizio -->           
									<div class="row-fluid">
										<div class="span12">
											<div class="box box-color box-bordered">
												<div class="box-title">
													<h3><i class="icon-user"></i>Modifica Associazioni</h3>
												</div>
												<div class="box-content nopadding">
													<!-- #include file = "../cClasse/studente_domande_include/2_modifica_associazioni_1.asp" -->
												</div>
											</div>
										</div>
									</div>
									<!-- >fine form -->              
									</div>
									<%end if%>

									<div class="tab-pane fade" id="profileN">
									<!----Inizio -->           
									<div class="row-fluid">
										<div class="span12">
											<div class="box box-color box-bordered">
												<div class="box-title">
													<h3><i class="icon-bell"></i>Notifiche</h3>
												</div>
												<div class="box-content nopadding">
													<!-- uso <div> invece che <form> perche tanto non c'è nulla da inviare -->
													<div class="form-horizontal form-bordered form-validate">
														<div class="control-group">
															<label for="textfield"  class="control-label">Stato iscrizione</label>
															<div class="controls">
																<input type="text" style="height: auto;" class="input-xlarge" disabled id="push-subscription-status" value="Caricamento...">
															</div>
                                     					</div>
														<div class="form-actions">
															<button type="submit" class="btn btn-primary" disabled id="push-subscribe-button" onclick="subscribeButton()">Iscriviti</button>
														</div>
														
													</div>
													<!-- #include file = "../cClasse/studente_domande_include/5_elenco_subscription.asp" -->
												</div>
											</div>
										</div>
									</div>

									
									<!-- fine form -->              
									</div>
                            
									<%end if%>         
                            
								</div>

							</div>
						</div>
						
						
 
						<!--</div>-->
					
					
					
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

