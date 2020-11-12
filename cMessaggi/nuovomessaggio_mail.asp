<%@ Language=VBScript %>
<!doctype html>
<html>
 
<head>
<title>Nuovo Invio Mail</title>
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
<!-- Notify -->
<link rel="stylesheet" href="../../css/plugins/gritter/jquery.gritter.css">
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
<!-- Notify -->
<script src="../../js/plugins/gritter/jquery.gritter.min.js"></script>

<!-- Theme framework -->
<script src="../../js/eakroko.min.js"></script>
<!-- Theme scripts -->
<script src="../../js/application.min.js"></script>
<!-- Just for demonstration -->

<!-- Favicon -->
<link rel="shortcut icon" href="../../img/favicon.ico" />
<!-- Apple devices Homescreen icon -->
<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />

<script language="javascript" type="text/javascript"> 
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
}

 $(window).ready(function () {	   
	   $('#msg').click();
	   
	  // event.stopPropagation();
	    
	});
</script>

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
</head>

<%
  Response.Buffer = true
  On Error Resume Next  
    ' per il controllo della validità della sessione, se è scaduta -> nuovo login
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
<BODY onLoad="showText2();">
</BODY>
<% else %>
<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
<% end if %>
<div id="navigation">
  <%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")   
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario
  
		%>
  <!-- #include file = "../var_globali.inc" --> 
  <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" --> 
  <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" --> 
  <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" --> 
  <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" --> 
  <!-- #include file = "../service/controllo_sessione.asp" --> 
  
  <!-- #include file = "../include/navigation.asp" --> 
  <!-- #include file = "../include/formattaDataCla.inc" --> 
  
</div>
<div class="container-fluid" id="content"> 
  
  <!-- #include file = "../include/menu_left.asp" -->
  
  <div id="main">
    <div class="container-fluid">
      <div class="page-header">
        <div class="pull-left">
          <h3> <i class="icon-comment"></i> Nuova Chat</h3>
        </div>
        <div class="pull-right"> 
          <!-- se mi interessa devo includere
                         include pull_right.asp--> 
        </div>
      </div>
      <div class="breadcrumbs">
        <ul>
          <li> <a href="#">Home</a> <i class="icon-angle-right"></i> </li>
          <li> <a href="#">Messaggi</a> <i class="icon-angle-right"></i> </li>
          <li> <a href="#">Nuova Email</a> </li>
        </ul>
        </ul>
        <div class="close-bread"> <a href="#"><i class="icon-remove"></i></a> </div>
      </div>
      <br>
      <div class="row-fluid">
        <div class="span12">
          <!-- #include file = "newemail.asp" -->
        </div>
      </div>
    </div>
    <!-- #include file = "../include/colora_pagina.asp" --> 
    
  </div>
  <!--fine main--> 
</div>
</body>
<script language="javascript" type="text/javascript">

/*h = "307" //altezza del box messaggi -> imposto l'altezza minima per grafica
h2 = "153"
document.getElementById("notifichemessaggi").style="min-height:"+h+"px";
document.getElementById("lavagnamessaggi").style="min-height:"+h+"px";
document.getElementById("forummessaggi").style="min-height:"+h+"px";
document.getElementById("diariomessaggi").style="min-height:"+h+"px";
document.getElementById("archiviomessaggi").style="min-height:"+h+"px";
document.getElementById("sentmessaggi").style="min-height:"+h+"px";
*/
//se necessario va aggiornato anche il px degli spazi di divisione nei vari box


function cancella_avviso() {
	
	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {  
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();	
	 }
}
   
   
 function aggiornaStud() {
	 // alert (DataClaq);
	 var DataCla=document.dati.txtData.value;
	 var DataCla2=document.dati.txtData2.value;
	// alert (DataClaq);
	 // alert (DataClaq2);
		 
		 
		 
		    document.dati.action = "?daForm=1&centromsg=1&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		 
	
	   
		document.dati.submit();		
}
 

 

 

</script>
<script type="text/javascript">
function checkTutti() {
	 
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		if (elements[i].type == 'checkbox')
		    {
		     elements[i].checked = true;
			 
			}
		}
	}	 
}
</script>
<script language="javascript" type="text/javascript">
function archivia_notifica() {
	
	  if (confirm("Vuoi archiviare tutte le notifiche selezionate ?")) {  
    document.dati.action = "cancella_notifica.asp?CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.dati.submit();	
	 }
}
 
 </script>
</html>
