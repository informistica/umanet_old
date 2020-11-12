<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Inserisci prefrase</title>   
   
       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	 <meta charset="UTF-8">
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
    
    <meta charset="utf-8">


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

	<script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />

	  <!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>   
       
       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
<!--	<script src="ckeditor/ckeditor.js"></script>-->
<!--	<script src="//cdn.ckeditor.com/4.14.0/standard/ckeditor.js"></script>-->
	 <script src="//cdn.ckeditor.com/4.14.0/full/ckeditor.js"></script>
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
     <body class='theme-<%=session("stile")%>' data-layout-topbar="fixed">
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
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
 
TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Paragrafo=Request.QueryString("Paragrafo")
  CodiceTest = Request.QueryString("CodiceTest") 
  Modulo=Request.QueryString("Modulo")
  Nome=Request.QueryString("Nome")
  Cognome=Request.QueryString("Cognome") 
  Cartella=Request.QueryString("Cartella")
  'Scadenza=Request.Form("txtScadenza")
  Scadenza=Request.Form("date3")
  Num = Request.Form("TxtNum") ' numero di domande che si vogliono inserire
  
   by_UECDL=Request.QueryString("by_UECDL")
   Segnalibro=Request.QueryString("Segnalibro")
   BoxApro=Request.QueryString("BoxApro")
  
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar") 
  Segnalibro=Request.QueryString("Segnalibro")
iddiv=Request.QueryString("iddiv")
  
 %>   
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserisci frasi</h1> 
                    
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
							<a href="#">Crea compito</a>
                            
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
				        <h3> <i class="icon-reorder"></i><%=Capitolo%> : <%=TitoloParagrafo%>
                        <%if CodiceSottopar<>"" then%>
                              /<%=response.write(Sottoparagrafo)%>
                            <%end if%>
                        </h3>
			          </div>
				      <div class="box-content">
   
       <form method="POST" class="form-vertical" name="dati" action="2inserisci_prefrase1.asp?Segnalibro=<%=Segnalibro%>&BoxApro=<%=BoxApro%>&by_UECDL=<%=by_UECDL%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Scadenza=<%=Scadenza%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>" >
        
       <span class="alert-info"> 
        <b>Incolla l'elenco (una su ogni riga)</b>  
        </span>
        
        
        
        <textarea name="MyTextArea" rows=8 cols=70 class="input-block-level" placeholder="aggiungi di seguito alla domanda il $ se &egrave; previsto il caricamento di immagini, aggiungi il # se &egrave; previsto il caricamento di file " ></textarea> 
        
        
        <p align="center"><b>
		<% response.write("Scadenza ?") %><br> </b>
        <!--<input type="text" name="txtScadenza" size="10" value="gg/mm/aaaa">-->
        <i class="icon-calendar"></i>&nbsp;<b>Data:</b> 
        
        <input type="text" name="date3" id="datepicker" class="input-small datepick" /></p>
        </p>
        <p align="center"><input type="submit" value="Invia" name="B1" class="btn" rel="tooltip" title="inserisci in blocco"></p>
        </form>
        </p>
                      </div>
					  <hr>
					    <div class="box-content">
   
       <form method="POST" class="form-vertical" name="dati" action="2inserisci_prefrase1.asp?Estesa=1&Segnalibro=<%=Segnalibro%>&BoxApro=<%=BoxApro%>&by_UECDL=<%=by_UECDL%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Scadenza=<%=Scadenza%>&Sottoparagrafo=<%=Sottoparagrafo%>&CodiceSottopar=<%=CodiceSottopar%>" >
        
       <label> 
        <b>Domanda estesa</b>  
        </label>
        <input type="text"  placeholder="Inserisci nome della domanda"  id="txtQuesito"  class="input-xxlarge">
        <textarea name="editor1" id="editor1"rows=8 cols=70 class="input-block-level" placeholder=".......">
		</textarea><br>
		<textarea name="txtEncode"  id="txtEncode" style="display:none;" rows="10" cols="80">
		</textarea><br>
		  
         
        
        <p align="center"> 
		 <a target="_blank" href="https://postimages.org/it/">hosting esterno per immagini</a>.&nbsp;<font color="#000000"><br><br>
 		<b> <span title="Richiede caricamento immagine ?">Img</span> </b>                                 
		<INPUT TYPE="RADIO" id ="txtImg1" name="txtImg"  value="1">Si
        <INPUT TYPE="RADIO"  id ="txtImg0" name="txtImg" checked="true" value="0"> No
                           <br><br>       		
		<% response.write("Scadenza ?") %><br> </b>
        <!--<input type="text" name="txtScadenza" size="10" value="gg/mm/aaaa">-->
        <i class="icon-calendar"></i>&nbsp;<b>Data:</b> 
        
        <input type="text" name="date3" id="datepicker1" class="input-small datepick" /></p>
        </p>
        <p align="center"><input type="button" value="Invia"onclick="invia();" name="B1" class="btn btn-primary" rel="tooltip" title="inserisci in blocco"></p>
        </form>
	
        </p>
			<hr>
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->
       
            
		</div> <!--fine main-->
        </div>
        
        

			 
	</body>
<script>
  
  
CKEDITOR.replace('editor1' );
 function rimpiazza(testo){
	var pulito = new String(testo);
		pulito = pulito.replace(/&agrave;/g,"à");
		pulito = pulito.replace(/&ograve;/g,"ò");
		pulito = pulito.replace(/&ugrave;/g,"ù");
		pulito = pulito.replace(/&egrave;/g,"è");
		pulito = pulito.replace(/&igrave;/g,"ì");
		pulito = pulito.replace(/&nbsp;/g," ");
		//pulito = pulito.replace(/&/g,"e");
		pulito = pulito.replace(/&#39;/g,"`");
		
		pulito = pulito.replace("'","`");

		return pulito;
}


	function invia(){

	   var domanda=document.getElementById("txtQuesito").value;	
	   var Modulo ="<%=Modulo%>";
	   var Paragrafo="<%=Paragrafo%>"; 
	   var CodiceSottopar="<%=CodiceSottopar%>" ;
	   var Cartella="<%=cartella%>" ;
	   var scadenza=document.getElementById("datepicker1").value;
	   var testo= rimpiazza(CKEDITOR.instances.editor1.getData());
	   var img=1;
	   if (document.getElementById("txtImg0").checked===true) img=0;
	    

	   var xhttp = new XMLHttpRequest();
	   xhttp.onreadystatechange = function() {
	   if (xhttp.readyState == 4 && xhttp.status == 200) {
		  var testoJSON=JSON.parse(xhttp.responseText);
					stato=testoJSON["stato"];
					messaggio=testoJSON["messaggio"];
					if (stato==0) 
                        alert('Errore: '+messaggio);
                    else {
						alert(messaggio);
					CKEDITOR.instances.editor1.setData('');
					document.getElementById("txtQuesito").value='';
                   
                     }
				//	document.getElementById("titolo"+globalpostid).innerHTML="<b>"+titolo+"</b>";
				//	document.getElementById(globalpostid).innerHTML=decodeURIComponent(testo);
					//document.getElementById(globalpostid).innerHTML=decodeURI(testo);
				//	 $('#chiudi').click();
	  }
	};
 
       
		var url="2inserisci_prefrase1_estesa_ajax.asp?domanda="+domanda+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&CodiceSottopar="+CodiceSottopar+"&scadenza="+scadenza+"&Img="+img+"&cartella="+Cartella;
		testo=encodeURIComponent(testo);
		params="testo="+testo;
		// alert(testo);
		// alert(params);
		xhttp.open('POST', url) 
		xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')
		xhttp.send(params);
		 
		}




</script>
 </html>

