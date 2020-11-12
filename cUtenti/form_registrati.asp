<%@ Language=VBScript %>
<%
  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu
  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
 
  


Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 

Dim stato
  Dim ConnessioneDB , rsTabella,rsTabella1
  Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
  
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
 session("registrati")=true ' serve per ex2inmgprofilo altrimenti controlla sessioni rimanda ad home
 ' lo metto a falso dopo l'invio
 
 ' stato=Request.QueryString("stato")
  id_classe=Request.QueryString("id_classe")
 ' app=Request.QueryString("app") ' vale 1 se sono stato chiamata da apprendimento
  divid=request.querystring("divid")
%>
 <!-- #include file = "../stringhe_connessione/stringa_connessione_registrati.inc" -->
    <!-- #include file = "../var_globali.inc" -->
   <%
	  QuerySQL="SELECT * FROM Classi where Id_Classe='"&id_classe&"';"
	  Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	  'response.write()
	  classe=rsTabella("Classe")	
	   
%>
<!doctype html>
<html>
<head>
<title>Registrazione utente</title>
	<meta charset="utf8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
	<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">
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
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
	<!-- Validation -->
	<script src="../../js/plugins/validation/jquery.validate.min.js"></script>
	<script src="../../js/plugins/validation/additional-methods.min.js"></script>
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
	<!-- Wizard -->
	<script src="../../js/plugins/wizard/jquery.form.wizard.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	<script src="../../js/demonstration.min.js"></script><script  src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
  
    <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    
	<script src="../js/sha256.js">/* SHA-256 JavaScript implementation */</script>
   
    
    
    <script type="text/javascript">
	

	
$(document).ready(function() {
	 
		/*$("#next").attr("value", "Ok!");*/
  $("#bottone").click(function(){
    var username = $("#username").val();
	var password = $("#password").val();
	var email = $("#emailfield").val();
	var nome = $("#nome").val();
    var cognome = $("#cognome").val();
	var sesso = $("#gend").val();
	var mipiace = $("#mipiace").val();
	var nonmipiace = $("#nonmipiace").val();
	var S1 =$("#S1").val();
	var id_classe =$("#id_classe").val();
	var classe =$("#classe").val();
	var tag =$("#gruppo").val();
	 
 var PwdAllievoSHA256 = Sha256.hash(password)
 
 

	

	
	 
	/*$("#nome1").val(nome);
	$("#cognome1").val(cognome);*/
	/*alert(username); 
	alert(password); 
	alert(id_classe); 
	 alert(S1); */ 
	
	 document.frmDocument.action = "registrati2_new.asp?passwordsha256="+PwdAllievoSHA256+"&nome=" + nome + "&cognome=" + cognome+ "&username=" + username+ "&password=" + password+ "&email=" + email+ "&mipiace=" + mipiace+ "&nonmipiace=" + nonmipiace+ "&S1=" + S1+ "&id_classe=" + id_classe+ "&classe=" + classe+ "&tag=" + tag;
		document.frmDocument.submit();
	 
	
	
	/*
	 $.ajax({
      type: "POST",
      url: "registrati2_new.asp",
      data: "nome=" + nome + "&cognome=" + cognome+ "&username=" + username+ "&password=" + password+ "&email=" + email+ "&mipiace=" + mipiace+ "&nonmipiace=" + nonmipiace+ "&S1=" + S1+ "&id_classe=" + id_classe,
      dataType: "html",
      success: function(msg)
      {
       //  $("#risultato").html(msg);
	    alert("Chiamata ok, e...");
      },
      error: function()
      {
        alert("Chiamata fallita, si prega di riprovare...");
      }
    });
	
	*/
	 
  });
  
 
});


/*
$("#ssss").submit(function(event) {

 var nome = $("#nome").val();
    var cognome = $("#cognome").val();
	$("#nome1").val(nome);
	$("#cognome1").val(cognome);
	alert($("#nome1").val()); 

 alert(cognome);


});
*/

 


</script>
 <script type="text/javascript">
 	
	$("#frmDocument").submit(function(e) {
    
	e.preventDefault();
	
	 
    var username = $("#username").val();
	var password = $("#password").val();
	
	var PwdAllievoSHA256 = Sha256.hash(password)	
	
	var email = $("#emailfield").val();
	var nome = $("#nome").val();
    var cognome = $("#cognome").val();
	var sesso = $("#gend").val();
	var mipiace = $("#mipiace").val();
	var nonmipiace = $("#nonmipiace").val();
	var S1 =$("#S1").val();
	var id_classe =$("#id_classe").val();
	var classe=$("#classe").val();



	
	 
	/*$("#nome1").val(nome);
	$("#cognome1").val(cognome);*/
	/*alert(username); 
	alert(password); 
	alert(id_classe); 
	    
	
	 document.frmDocument.action = "registrati2_new.asp?passwordsha256="+PwdAllievoSHA256+"&nome=" + nome + "&cognome=" + cognome+ "&username=" + username+ "&password=" + password+ "&email=" + email+ "&mipiace=" + mipiace+ "&nonmipiace=" + nonmipiace+ "&S1=" + S1+ "&id_classe=" + id_classe;
		document.frmDocument.submit();
	 




return false;*/

 });
   
  </script>  
    
    
    

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />

</head>

<body class='login'>
<!-- <div class="wrapper">
		<h1><a href="index.html"> <img class="img-rounded" src="../img/umanet2.jpg" alt=""  width="50%" height="20%"> 2.0</a></h1>
		 </div>
			 
             -->
            	<div class="row-fluid">
					<div class="span12"> 
             
             
             <div class="box box-color box-bordered blue">
							<div class="box-title">
								<h3>
									<i class="icon-magic"></i>
									Registrati nella classe  <%=left(Classe,1+len(Classe)-instr(Classe,"$"))%> 
                                    <% Session("Cartella")=classe
									   Session("Registrati")=true
									  
									%>
								</h3>
							</div>
							<div class="box-content nopadding">
								<form  method="POST" class='form-horizontal form-wizard wizard-vertical form-validate'  id="frmDocument" name="frmDocument" action="../aaaa.asp">
									<div class="step" id="firstStep">
										<ul class="wizard-steps steps-3">
											<li class='active'>
												<div class="single-step">
													<span class="title">1</span>
													<span class="circle">
														<span class="active"></span>
													</span>
													<span class="description">
														&nbsp;Login
													</span>
												</div>
											</li>
											<li>
												<div class="single-step">
													<span class="title">2</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Contatti 
													</span>
												</div>
											</li>
											<li>
												<div class="single-step">
													<span class="title">3</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Profilo
													</span>
												</div>
											</li>
										</ul>
										<div class="form-content">
                                      
                                        
                                        
											<div id="usr" class="control-group">
												<label for="firstname" id="firstname" class="control-label">Username</label>
												<div class="controls">
                               
													<input type="text" name="username" id="username" onblur="controllacaratteri()" class="input-xlarge" data-rule-required="true" placeholder="Non inserire spazi bianchi, caratteri speciali o accentati">
                                                                      <a href="#" class="btn" rel="popover" data-trigger="hover" title="Attenzione" data-content="Non inserire spazi bianchi,  caratteri speciali o accentati" ><i  class="icon-info"><b>(i)</b></i></a>
                                                   
												</div>
											</div>
                                            
                                            <div class="control-group">
										<label for="pwfield" class="control-label">Password</label>
										<div class="controls">
											<input type="password" name="password" id="password" class="input-xlarge" data-rule-required="true">
										</div>
									</div>
									<div class="control-group">
										<label for="confirmfield" class="control-label">Conferma password</label>
										<div class="controls">
											<input type="password" name="password_conferma" id="confirmfield" class="input-xlarge" data-rule-equalTo="#password" data-rule-required="true">
										</div>
									</div>
                                            
											 
                                            
                                            
                                            
                                            
										</div>
									</div>
									<div class="step" id="secondStep">
										<ul class="wizard-steps steps-3">
											<li>
												<div class="single-step">
													<span class="title">1</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Login
													</span>
												</div>
											</li>
											<li class='active'>
												<div class="single-step">
													<span class="title">
														2</span>
													<span class="circle">
														<span class="active"></span>
													</span>
													<span class="description">
														&nbsp;Contatti
													</span>
												</div>
											</li>
											<li>
												<div class="single-step">
													<span class="title">
														3</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Profilo
													</span>
												</div>
											</li>
										</ul>
										<div class="form-content">
										
                                        
										<div class="control-group">
										<label for="pwfield" class="control-label">Email</label>
										<div class="controls">
											<input type="text" name="email" id="emailfield" class="input-xlarge" data-rule-required="true"  data-rule-email="true">
										</div>
									</div>
									<div class="control-group">
										<label for="confirmfield1" class="control-label">Conferma email</label>
										<div class="controls">
											<input type="text" name="email_conferma" id="confirmfield1" class="input-xlarge" data-rule-equalTo="#emailfield" data-rule-required="true"  data-rule-email="true">
										</div>
									</div>
                                    
                                    
										</div>
									</div>
									<div class="step" id="thirdStep">
										<ul class="wizard-steps steps-3">
											<li>
												<div class="single-step">
													<span class="title">
														1</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Login
													</span>
												</div>
											</li>
											<li>
												<div class="single-step">
													<span class="title">
														2</span>
													<span class="circle">
													</span>
													<span class="description">
														&nbsp;Contatti
													</span>
												</div>
											</li>
											<li class='active'>
												<div class="single-step">
													<span class="title">
														3</span>
													<span class="circle">
														<span class="active"></span>
													</span>
													<span class="description">
														&nbsp;Profilo
													</span>
												</div>
											</li>
										</ul>
										<div class="form-content">
                                       
                                       
                                           <%   
											   url= "../Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
										   
										  url=Replace(url,"\","/")
										  
										   
										 urlimg=Server.MapPath(homesito)& "/img/no-image.jpg"
										  urlimg=url&"/"& Url_img 
										 
										 %>
                                        
                                      <!--
                                        <div class="control-group">
                                        
                                       <div class="accordion" id="accordion2">
  
   <div class="accordion-group">
    <div class="accordion-heading">
      <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapseOne">
        Foto (+)
      </a>
    </div>
   
    <div id="collapseOne" class="accordion-body collapse">
     
      <div class="accordion-inner">
        <iframe src="upload_resize/ex2_imgprofilo.asp" name="postmessage" id="postmessage" width="100%" height="60%" frameborder="0" SCROLLING="no" border="0" class="iframe">
      </iframe> 
      </div>
      
    </div>
    
  </div>
</div> 

                                        
								 
									</div>
                                        
                                        -->
                                        
									    <div class="control-group">
										<label for="textfield"  class="control-label">Cognome </label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="cognome" id="cognome"  data-rule-required="true">
                                            <input type="hidden"  id="id_classe" name="id_classe" value="<%=id_classe%>">
                                            <input type="hidden"  id="classe" name="id_classe" value="<%=classe%>">
                                             
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nome</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="nome"  id="nome"  data-rule-required="true" >
                                            
										</div>
									</div>
									 <div class="control-group">
                                        <label for="textfield" class="control-label">Gruppo</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-mini"  name="gruppo"  id="gruppo"  data-rule-required="true" >
                                            
										</div>
									</div>
                                            
                                       <!--     
                                          <div class="control-group">
											<label for="text" class="control-label">Sesso</label>
											<div class="controls">
												<select name="gend" id="gend">
													<option value="">-- Scegli --</option>
													<option value="1">Maschio</option>
													<option value="2">Femmina</option>
												</select>
											</div>
										</div>
                                            
                                         -->   
                                            <div class="control-group">
										<label for="textfield"  class="control-label" >Mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;"   class="input-xxlarge"  name="mipiace" id="mipiace"value="" >
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Non mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xxlarge"  name="nonmipiace" id="nonmipiace" value="">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label">Descriviti</label>
										<div class="controls">
											 <p><textarea  rows="6" name="descriviti"  id="S1"  class="input-block-level"></textarea></p>
   
										</div>
                                 <div class="control-group">
										<label for="textfield" class="control-label">
                                        <input type="button" class="btn btn-primary" value="Completa la registrazione" id="bottone">				 				
                                        </label>       
                                       
                                        
									</div>
                               </div>             
                                            
										</div>
									</div>
									<div class="form-actions">
										<input type="reset" class="btn" value="Indietro" id="back">
										<input type="submit" class="btn "  id="next">
                                   
									</div>
								</form>
							</div>
						</div>
             
             
             
             </div>
             </div>
             
         <% rsTabella.Close()
	Set rsTabella = nothing
 %>    
      
	<script type="text/javascript">
	
	function controllacaratteri(){
	
		var stringa = document.getElementById("username").value;
		var re = /^[a-z0-9]+$/i;
		
		if(!re.test(stringa)){
			alert("Non inserire caratteri speciali, né spazi bianchi");
						document.getElementById("next").disabled=true;
		}else{
			document.getElementById("next").disabled=false;
		}
	}
	

		 
//$(window).load(function () {
	   
	//   $('#avvisoUser').click();
	   
  
	 
	  // event.stopPropagation();
	    
	//});
	

/*$(".red").click(function(event){
   
   // alert("Hai cliccato sull'Elemento");
	document.location = "script/aggiorna_stile.asp?stile=red"
});
*/	
	
</script> 
	
</body>

</html>
