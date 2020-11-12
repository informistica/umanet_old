<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Consegna e correzione</title>

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
s
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />

 

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>s
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
		<!-- #include file = "../service/replacecar.asp" -->
        <%
		 ' esecuzione della query per prelevare le i dati di un dato paragrafo di un dato modulo

 
Num=Request.QueryString("Num")
Titolo=Request.QueryString("Capitolo")
Paragrafo=Request.QueryString("Paragrafo")
CodiceTest=Request.QueryString("CodiceTest")
Classe=Request.QueryString("classe")
cartella= Request.QueryString("cartella")
  
ID_Mod=Request.QueryString("Modulo") 
	   
Titolo = Replace(Titolo, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
Titolo=  Replace(Titolo,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql	    
Titolo=ReplaceCar(Titolo)
  
 
  
    urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
	'ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&".xml"
	ulrRisorsa1=paragrafo&"_"&session("CodiceAllievo")&".xml"
	ulrRisorsa=urlRis&ulrRisorsa1
	ulrRisposta=Replace(ulrRisorsa,"\","/")	
	'response.write("<br>"&ulrRisorsa&"<br>")

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
							<div class="row-fluid">
								<div class="span12">
									<div class="box">
										<div class="box-content">
									
									
 <%

    Set fso = CreateObject("Scripting.FileSystemObject") 

    If not(fso.FileExists(ulrRisposta)) Then
        Set objCreatedFile = fso.CreateTextFile(ulrRisposta, True)
        objCreatedFile.WriteLine("<Domande>")
        cont=0
        for k=0 to Num-1
            Domanda = Request.Form("txtFrase"&k) 
            Domanda = ReplaceCar(Domanda) 
            Risp = Request.Form("txtRisposta"&k) 
            Risp = ReplaceCar(Risp)  
            IdPrefrase = Request.Form("txtIdFrase"&k)  
			if Risp="" then
			Risp="In_bianco"
			end if
	        Risp=LTrim(Rtrim(Risp))
             ' response.write(k&"---"&Risposta&"<br>")
               'if strcomp(Trim(Domanda),"")<>0 then ' controllo per le righe vuote 
                cont=cont+1
                'response.write(Domanda&"<br>")
                objCreatedFile.WriteLine("<Domanda>")
                objCreatedFile.WriteLine("	<IdPrefrase>"&IdPrefrase&"</IdPrefrase>")
                objCreatedFile.WriteLine("	<Numero>"&cont&"</Numero>")
                objCreatedFile.WriteLine("	<Testo>"&Domanda&"</Testo>")
                objCreatedFile.WriteLine("		<Risposta>")
                objCreatedFile.WriteLine(Risp)
                objCreatedFile.WriteLine("		</Risposta>")
                objCreatedFile.WriteLine("</Domanda>") 
        next 
        objCreatedFile.WriteLine("</Domande>")
        objCreatedFile.Close
    Set objCreatedFile=nothing
    'On Error Resume Next
	If Err.Number = 0 Then
		Response.Write " <div class='alert alert-success'>Consegna della verifica avvenuta!</div> "
		' CORREZIONE ISTANTANEA

		dim risposta_ideale(100)
		dim risposta(100)
		dim totale 

		querySQL="Select * from Allievi where CodiceAllievo='"&Session("CodiceAllievo")&"';"
		set rsTabella=ConnessioneDB.Execute(QuerySQL) ' mi serve per agganciarmi al file include della correzione

		urlRisRisposte=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"	
		urlRisorsaRisposte=paragrafo&"_"&rsTabella("CodiceAllievo")&".xml"
		urlRisposte=urlRisRisposte&urlRisorsaRisposte
		urlRisposte=Replace(urlRisposte,"\","/")	

		shortUrl="Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
		shortUrl=shortUrl&paragrafo&"_correzione_"&session("CodiceAllievo")&".xml"
		shortUrl=Replace(shortUrl,"\","/")	

		urlRisModello=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
		urlRisorsaModello=paragrafo&".xml"
		urlModel=urlRisModello&urlRisorsaModello
		urlModello=Replace(urlModel,"\","/")
 
		Set objXMLDocM = Server.CreateObject("Microsoft.XMLDOM") ' per il file modello
		objXMLDocM.async = False 
		Set objXMLDocR = Server.CreateObject("Microsoft.XMLDOM") ' per il file delle risposte
    	objXMLDocR.async = False 
		objXMLDocM.load urlModello
	 
			
		
		Set fso = CreateObject("Scripting.FileSystemObject") 
		 %>
		<!-- #include file = "3correggi_verifica_paragrafo_include.asp" -->
		
		<%  

		Set objXMLDocM = Nothing
		Set objXMLDocR = Nothing
		'On Error Resume Next
		If Err.Number = 0 Then%>
		 <br> <div class='alert alert-success'><b>Correzione effettuata  <a href='3visualizza_risultati_verifiche.asp?cartella=<%=cartella%>&cod=<%=Session("CodiceAllievo")%>&paragrafo=<%=paragrafo%>&shorturl=<%=shorturl%>'>Dettagli correzione</a></b> 
		 </div>
		 
       <%  
		Else
			Response.Write Err.Description 
			Err.Number = 0
		End If




		Else
			Response.Write Err.Description 
			Err.Number = 0
		End If
		else
			response.write("<code>Hai già consegnato, guarda il report della correzione nel tuo quaderno</code>")
		end if



%>





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
