<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<meta charset="utf-8">
   <title>Convalida verifica</title>

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
		  <!-- #include file = "../service/replacecar.asp" -->
<link rel="stylesheet" href="../../jquery-ui.css" />





</head>
<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">

	<div id="navigation">

        <%
Paragrafo=Request.QueryString("Paragrafo")
CodiceTest=Request.QueryString("CodiceTest")
Classe=Request.QueryString("classe")
cartella= Request.QueryString("cartella")
  
ID_Mod=Request.QueryString("Modulo") 
TitoloModulo=Request.QueryString("TitoloModulo") 
Titolo = Replace(Titolo, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
Titolo=  Replace(Titolo,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql	    
Titolo=ReplaceCar(Titolo)
  

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
						<h1> <i class="icon-comments"></i> Convalida verifica </h1>

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
							<a href="#">Verifica</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
						  <a href="#"><%=TitoloModulo%></a>
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
				        <h3> <i class="icon-reorder"></i>  Risultati :  <%=Paragrafo%> </h3>
			          </div>
				      <div class="box-content">






				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">



		    <div class="box-content">

 <% 
   
    urlRisModello=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
    urlRisorsaModello=paragrafo&".xml"
    urlModel=urlRisModello&urlRisorsaModello
	urlModello=Replace(urlModel,"\","/")
 
Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0") ' 
objXMLDoc.async = False 



    urlRisRisposte=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"	
    	
    querySQL="Select * from Allievi where Id_Classe='"&Session("Id_Classe")&"' and Attivo=1 order by Cognome;"
    set rsTabella=ConnessioneDB.Execute(QuerySQL)
    Set fso = CreateObject("Scripting.FileSystemObject") %>
 <form method="POST" class="form-horizontal" name="Aggiorna" action="3aggiorna_punteggio_verifiche1.asp?id_classe=<%=id_classe%>&CodiceTest=<%=CodiceTest%>">
 <table class="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped">
 <tr><th>Cognome</th><th>Nome</th><th width="10%"><span ocClick="normalizza();">Punteggio</th> </tr>
	<%i=0
    do while not rsTabella.eof
            
                urlRisorsaRisposteCorrezione=paragrafo&"_correzione_"&rsTabella("CodiceAllievo")&".xml"
                urlCorrezione=urlRisRisposte&urlRisorsaRisposteCorrezione
                urlCorrezione=Replace(urlCorrezione,"\","/")
   
           ' response.write("<br><b>"&rsTabella("Cognome") & " " &rsTabella("Nome")&" : </b>")
            If (fso.FileExists(urlCorrezione)) Then    ' se esiste e quindi lo stud ha consegnato
                ' creo file per salvare la correzione
               
			objXMLDoc.load urlCorrezione
			Set Root = objXMLDoc.documentElement
			Set NodeList = Root.getElementsByTagName("Sentiment")
			Set Risposta = objXMLDoc.getElementsByTagName("Sentiment")(0)
			voto=Risposta.text
			else
			' non ha consegnato
			voto=0
			end if%>
            
 <tr><td><input type="hidden" value="<%=rsTabella.fields("CodiceAllievo")%>" name="txtStud<%=i%>" />
 <input type="hidden" value="<%=rsTabella.fields("Cognome") & " " & left(rsTabella("Nome"),1) &"."%>" name="txtStudNome<%=i%>" />
 <%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td><td><input type="text" class="input-mini" value="<%=voto%>" id="txtPunti<%=i%>" name="txtPunti<%=i%>" /></td> </tr>
  

        <%rsTabella.movenext
       i=i+1
    loop

%>
</table>
 
 
 
<br />
 <input type="hidden" value="<%=i-1%>" name="txtnumRec" />
  <input class="btn " type="button" value="Assegna punti" onClick="javascript:normalizza();"/><br>

 <br>
<input class="btn btn-primary" disabled type="submit" id="btnInvia" value="Convalida verifica" /><br>
<br>
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



	</body>
    <script language="javascript" type="text/javascript">

 
 function normalizza() { 
	with (document.Aggiorna) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'text')
		elements[i].value = Math.round((elements[i].value)/10);
		document.getElementById("btnInvia").disabled=false;
		}
	}
}


  </script>

 </html>
