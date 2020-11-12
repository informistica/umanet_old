<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Modifica prefrasi estesa</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<meta charset="utf-8">

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
	<!--
      <script type="text/javascript" src="../js/utility.js"></script>

	  -->
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

    <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
 <script src="../../js/datapicker_it.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />



<!--Controllo accesso quaderno e sessione scaduta con redirect ad index.html-->
       <script src="../js/privacy.js"></script>

	    

	<!--	<script src="ckeditor/ckeditor.js"></script>-->
     <!--   <script src="//cdn.ckeditor.com/4.14.0/standard/ckeditor.js"></script>-->
       <script src="//cdn.ckeditor.com/4.14.0/full/ckeditor.js"></script>

<script language="javascript" type="text/javascript">
function showText3() {window.alert("Il compito è già stato inserito, lo puoi modificare dal tuo quaderno!")
location.href="../home.asp"

 }
    </script>

	<% x = Request.ServerVariables("HTTP_REFERER")
if x = "" then %>
<script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

 //alert("<%=x%>");
 location.href="../../../../index.html";

//location.href=window.history.back();
 }
 </script>

 <% else %>
 <script>
function showText() {window.alert("Non puoi visualizzare i dati degli altri studenti!")

 location.href="<%=x%>";

//location.href=window.history.back();
 }
 </script><% end if%>




<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<%
  Response.Buffer = true
  'On Error Resume Next









  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu

  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>

    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<%
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
 
  'cla=Request.QueryString("cla")
 
  
  
  Domanda=Request.QueryString("Domanda")
  Modulo=Request.QueryString("Modulo")
  Paragrafo=Request.QueryString("CodiceTest")
  Codice_Test=Request.QueryString("CodiceTest")
  Id_Prefrase=Request.QueryString("Id_Prefrase")
  Capitolo=Request.QueryString("Capitolo")
TitoloParagrafo=Request.QueryString("Paragrafo")
Sottoparagrafo=Request.QueryString("Sottoparagrafo")
cartella=Request.QueryString("cartella")
 'Capitolo=Request.QueryString("Capitolo")
 'Paragrafo=Request.QueryString("Paragrafo")
 


  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" &cartella &"/" &Modulo&"_Esercizi/"&Paragrafo&"_"&Id_Prefrase&".txt"
  url=Replace(url,"\","/")
 Set objFSO = CreateObject("Scripting.FileSystemObject") 
 'response.write(url)
 Set objTextFile = objFSO.OpenTextFile(url, ForReading)

	sReadAll=""

	' Use different methods to read contents of file.
	sReadAll = objTextFile.ReadAll

	if sReadAll = "" then
		sReadAll = "File spiegazione mancante. Elimina e reinserisci la frase nel tuo quaderno."
		dis = true
	end if
	 

  

%>
        <body   class='theme-<%=session("stile")%>' >	 

	<div id="navigation">
        <!-- #include file = "../var_globali.inc" -->

  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
	</div>

 <%

 %>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Modifica prefrase estesa</h1>

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
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%>: <%=TitoloParagrafo%><% if Sottoparagrafo<>"" then response.write("/"&Sottoparagrafo) end if%></h3>
			          </div>
				      <div class="box-content">

 <div class="immagini" style="height:auto; width:auto; border:none;" >
  <form id="frmDocument" name="dati" class='form-vertical form-bordered form-striped' method="POST"  action="2inserisci_valutazione_frase1.asp?CodiceAllievo=<%=CodiceAllievo%>&CodiceTest=<%=Codice_Test%>&Cartella=<%=Cartella%>&Modulo=<%=ID_MOD%>&Paragrafo=<%=Paragrafo%>&tCap=<%=tCap%>&tSot=<%=tSot%><%=p%>&tFra=<%=tFra%>&Capitolo=<%=Capitolo%>" >
 <legend> <input type="text" value="<%=Domanda%>" name="txtQuesito" id="txtQuesito" class="input-xxlarge"></legend> &nbsp;&nbsp; 
										
             <div class="control-group">
				  <div class="controls">
<div id="editor1">
    <%=sReadAll%>
</div>
<textarea name="txtSpiegazione"  id="txtSpiegazione" style="display:none;">
</textarea>	 
	  <br>
	</div>
</div>

 <br>

<!-- *** Ripristinare ?
<img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa();">
&nbsp;
-->

 
<input type="button" onclick="invia()" id="btnImg" value="Aggiorna" name="B1" class="btn"> </p>
 <br><br>




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
   

 <script language="javascript" type="text/javascript">
/*
 $("#btnImg").click(function(){

	getParametri(1);

 });
*/

  

CKEDITOR.replace('editor1');
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
       var Id_Prefrase="<%=Id_Prefrase%>"; 
	   var CodiceSottopar="<%=CodiceSottopar%>" ;
	   var Cartella="<%=cartella%>" ;
	   var testo= rimpiazza(CKEDITOR.instances.editor1.getData());
	 
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
                     }
				//	document.getElementById("titolo"+globalpostid).innerHTML="<b>"+titolo+"</b>";
				//	document.getElementById(globalpostid).innerHTML=decodeURIComponent(testo);
					//document.getElementById(globalpostid).innerHTML=decodeURI(testo);
				//	 $('#chiudi').click();
	  }
	};
 
       
		var url="2modifica_prefrase1_estesa_ajax.asp?Id_Prefrase="+Id_Prefrase+"&domanda="+domanda+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo+"&CodiceSottopar="+CodiceSottopar+"&cartella="+Cartella;
		testo=encodeURIComponent(testo);
		params="testo="+testo;
		//alert(testo);
        //alert(params);
        //alert(url);
		xhttp.open('POST', url) 
		xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')
		xhttp.send(params);
		 
		}
 </script>


 </html>
