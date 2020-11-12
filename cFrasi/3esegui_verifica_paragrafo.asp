<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Verifica</title>

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

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />


       <!-- PLUpload -->
	 <!--<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
 <!--	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>-->



   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />


 <%'if session("Admin")=false then
%>
    <script type="text/javascript" src="../js/utility.js"></script>
 <% 'end if
 %>



</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%
 ' 	lettura dei parametri passati alla pagina
  Cartella=Request.QueryString("Cartella")
  TitoloCapitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")

  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")


  

		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo

 urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Modulo,len(Modulo)-instr(Modulo,"_"))&"/"
	'ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&".xml"
	ulrRisorsa1=paragrafo&".xml"
	ulrRisorsa=urlRis&ulrRisorsa1
	ulrRisorsa=Replace(ulrRisorsa,"\","/")	
	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Esegui verifica </h1>

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
							<a href="#"><%=response.write(TitoloCapitolo)%></a>
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
				        <h3> <i class="icon-reorder"></i>  <%=response.write(Paragrafo)%>
						 
                         </h3>
			          </div>
				      <div class="box-content">

   <%' se la query che preleva i compiti non restituisce risultati
   Set objXMLDocM = Server.CreateObject("Microsoft.XMLDOM") ' per il file modello
objXMLDocM.async = False 
objXMLDocM.load ulrRisorsa
Set RootM = objXMLDocM.documentElement
Set NodeListM = RootM.getElementsByTagName("Domanda")

		 %>

		<div class="row-fluid">
		 <div class="span12">
		   <div class="box">
              <div class="box-content">
                <form method="POST" name="frmDocument" id="frmDocument" class="form-vertical" action="3consegna_verifica_paragrafo.asp?Id_Classe=<%=Id_Classe%>&ID_Mod=<%=ID_Mod%>&Titolo=<%=Titolo%>&classe=<%=classe%>&cartella=<%=cartella%>&Num=<%=NodeListM.length%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceTest=<%=CodiceTest%>">
              <%
			  
			  

'response.write("File modello:"& NodeListM.length -1 &"<br>")
For i = 0 to NodeListM.length -1
	Set IdPrefraseM = objXMLDocM.getElementsByTagName("IdPrefrase")(i)
    Set TestoM = objXMLDocM.getElementsByTagName("Testo")(i)
    Set RispostaM = objXMLDocM.getElementsByTagName("Risposta")(i)
  'Response.Write IdPrefraseM.text & "<br> " & TestoM.text & "<br>"& RispostaM.text & "<br><br>"
%>
	 <b><%=i+1%>) Domanda </b>  
	 	   <input type="text" class="hidden" value="<%=IdPrefraseM.text%>" name="txtIdFrase<%=i%>" size="3" >
		   <input type="text"  value="<%=TestoM.text%>" class="input-block-level" name="txtFrase<%=i%>">
		    
		   	<script>

	  		 document.write("<p><textarea  oninput=\"aumenta(<%=i%>,<%=len(RispostaM.text)%>)\" onFocus=\"azzera(<%=i%>)\"  onkeydown = \"onKeyDown()\" id=\"txtRisposta<%=i%>\"  rows=\"3\" name=\"txtRisposta<%=i%>\" cols=\"116\" class=\"input-block-level\"></textarea></p>");
			</script>
		<noscript>Abilita JavaScript per visualizzare correttamente il sito.</noscript>
		<div class="row-fluid">
			<div class="span2" id="cont<%=i%>"></div>
			<div class="span10"></div> 
		</div>
		   
		<!-- 
		   <textarea class="input-block-level" rows="3"  name="txtRisposta<%=i%>" id="txtRisposta<%=i%>" onFocus="this.value=''">
		   </textarea></p>-->
<% 
Next
			  %>            
               </div>
               <br>
			     </div>
	          </div>
    	    </div>
<button type="submit" class="btn btn-primary"  id="Btn1" >Invia</button>
   </form>
                      </div>
			        </div>
			      </div>
			    </div>
			</div>


		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->

	 <script>
	  function aumenta(i,lung) {
		 // alert(document.getElementById("txtRisposta"+i).value.length);
        let l= document.getElementById("txtRisposta"+i).value.length;
        document.getElementById("cont"+i).innerHTML = `${l}/${lung}`;
	}
	  function azzera(i) {
		 // alert(document.getElementById("txtRisposta"+i).value.length);
        document.getElementById("txtRisposta"+i).value="";
        document.getElementById("cont"+i).innerHTML = "";
    }
	 </script>

	</body>

 </html>
