<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Inserisci verifica</title>

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

 

<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">

        <%
 

		' connessione al database e inclusione dei menu
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
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
  
 

	%>
	</div>

	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Inserisci verifica </h1>

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
							<a href="#"><%=response.write(Titolo)%></a>
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
							DataOggi=Day(date())&"/"&Month(date())&"/"&Year(date())
  
							QuerySQL="Update Paragrafi Set [Verifica]=1 WHERE ID_Paragrafo='" & CodiceTest & "';"  
						'	response.write(QuerySQL) 
							ConnessioneDB.Execute QuerySQL 
							QuerySQL="Select * from [2ESERCITAZIONI_SINGOLI] where CodiceTest='"&CodiceTest&"';"
							set rsTab=ConnessioneDB.Execute(QuerySQL)
							inserita=false
							if rsTab.EOF then
							inserita=true
							QuerySQL="INSERT INTO [2ESERCITAZIONI_SINGOLI] (Descrizione,Data,Id_Classe,Classifica,TipoVoto,CodiceTest) SELECT '" & Paragrafo & "','" & DataOggi & "','" & Session("Id_Classe") & "'," & 1 & ",'V','"&CodiceTest&"';"					
							'response.write(QuerySQL) 
							ConnessioneDB.Execute QuerySQL 
							Else
							  response.write("<code>Verifica già presente</code>") 
							end if

								urlRis=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Risorse/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
								urlVer=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella &"/Verifiche/Mod_"&right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&"/"
								
								'ulrRisorsa1=right(Id_Mod,len(Id_Mod)-instr(Id_Mod,"_"))&".xml"
								ulrRisorsa1=paragrafo&".xml"
								ulrRisorsa=urlRis&ulrRisorsa1
								ulrRisorsa=Replace(ulrRisorsa,"\","/")	

						

							'	response.write("<br>"&ulrRisorsa&"<br>")
								Set fso = CreateObject("Scripting.FileSystemObject") 
								Set objCreatedFile = fso.CreateTextFile(ulrRisorsa, True)
								objCreatedFile.WriteLine("<Domande>")
							cont=0
							for k=0 to Num-1
								Domanda = LTrim(RTrim(Request.Form("txtFrase"&k)))
								Domanda = ReplaceCar(Domanda) 
								Risposta = LTrim(RTrim(Request.Form("txtModello"&k)))
								Risposta = ReplaceCar(Risposta)  
								IdPrefrase = Request.Form("txtIdFrase"&k)  
								Inverifica=Request.Form("txtVerifica"&k)  
							
								'response.write(k&"---"&Inverifica&"<br>")
								'if strcomp(Trim(Domanda),"")<>0 then ' controllo per le righe vuote 
								if Inverifica then
										cont=cont+1
								'response.write(Domanda&"<br>")
										objCreatedFile.WriteLine("<Domanda>")
										objCreatedFile.WriteLine("	<IdPrefrase>"&IdPrefrase&"</IdPrefrase>")
										objCreatedFile.WriteLine("	<Numero>"&cont&"</Numero>")
										objCreatedFile.WriteLine("	<Testo>"&Domanda&"</Testo>")
										objCreatedFile.WriteLine("		<Risposta>")
										objCreatedFile.WriteLine(Risposta)
										objCreatedFile.WriteLine("		</Risposta>")
										objCreatedFile.WriteLine("</Domanda>")
								end if
								
							next 
								objCreatedFile.WriteLine("</Domande>")
								objCreatedFile.Close

								If  fso.FolderExists(urlVer)= FALSE Then 
								   fso.CreateFolder urlVer
								end if

							Set objCreatedFile=nothing
' 
							
													'On Error Resume Next
											If Err.Number = 0 Then
												if (inserita) then
													Response.Write "<br> <div class='alert alert-success'><b>Inserimento verifica effettuata </b> </div>"
												end if
											Else
												Response.Write Err.Description 
												Err.Number = 0
											End If
								

							   ' creo file dei token da utilizzare in correzione
							   ' in realtà non lo utilizzerò lo lascio ... può tornare utile
								' ulrRisorsaToken=urlRis&paragrafo&"_token.xml"
								' ulrRisorsaToken=Replace(ulrRisorsaToken,"\","/")
								' Set objXMLDocM = Server.CreateObject("Microsoft.XMLDOM") ' per il file modello
								' objXMLDocM.async = False 
								' objXMLDocM.load ulrRisorsa
								' Set RootM = objXMLDocM.documentElement
								' Set NodeListM = RootM.getElementsByTagName("Domanda")
								' Set fso = CreateObject("Scripting.FileSystemObject") 
								' Set objFileToken = fso.CreateTextFile(ulrRisorsaToken, True)
	             				' objFileToken.WriteLine("<Correzioni>")
								' totale=0
								'  For n = 0 to NodeListM.length -1
								' 	Set TestoM = objXMLDocM.getElementsByTagName("Testo")(n)
								' 	Set RispostaM = objXMLDocM.getElementsByTagName("Risposta")(n)
								' 	objFileToken.WriteLine("<Domanda>")
								' 	response.write("<br>"&TestoM.text)
								' 	objFileToken.WriteLine(TestoM.text)
								' 	objFileToken.WriteLine("</Domanda>")
								' 	sreadAll=Replace(RispostaM.text,".","")
								' 	sreadAll=Replace(sreadAll,",","")
								' 	readAll=Replace(sreadAll,chr(13)," ")  ' *****controllare
								' 	risposta_ideale_pre = Split(sreadAll," ")
								
								' 	for each x in risposta_ideale_pre
								' 			if (len(x)>5) then  
								' 				risposta_ideale=Replace(Lcase(Trim(x)),","," ")
								' 				risposta_ideale=Replace(risposta_ideale,chr(13)," ")
								' 				risposta_ideale=Replace(risposta_ideale,vbCr," ")' ***** FORSE RISOLVE IL
								' 				risposta_ideale=Replace(risposta_ideale,vbLf," ")
											 
								' 				risposta_ideale=Replace(risposta_ideale,"."," ")
								' 				risposta_ideale=Replace(risposta_ideale,"-","")
								' 				risposta_ideale=Replace(risposta_ideale,"-"," ")
								' 				risposta_ideale=Replace(risposta_ideale,"("," ")
								' 				risposta_ideale=Replace(risposta_ideale,")"," ")
								' 				risposta_ideale=Replace(risposta_ideale,"perche`","")
								' 				risposta_ideale=Replace(risposta_ideale,"quindi","")
								' 				risposta_ideale=Replace(risposta_ideale,"quando","")
								' 				risposta_ideale=Replace(risposta_ideale,"infatti","")
								' 				risposta_ideale=Replace(risposta_ideale,"dell`","")
								' 				risposta_ideale=Replace(risposta_ideale,"l`","")
								' 				risposta_ideale=Replace(risposta_ideale,"d`","")
								' 				response.write("<br>"&risposta_ideale)
								' 				risposta_ideale_pre2 = Split(risposta_ideale," ")
								' 				if ubound(risposta_ideale_pre2)>0 then
								' 				'response.write("<br>sono dentro")
								' 					for each y in risposta_ideale_pre2
								' 						if y<>"" then
								' 						objFileToken.WriteLine("<Modello>")
								' 						objFileToken.WriteLine(y)
								' 						objFileToken.WriteLine("</Modello>")
								' 						end if
								' 					next 
								' 				else
								' 					objFileToken.WriteLine("<Modello>")
								' 					objFileToken.WriteLine(risposta_ideale)
								' 					objFileToken.WriteLine("</Modello>")
								' 				end if
								' 			end if
								' 	next
								' next 
								' objFileToken.WriteLine("<Correzioni>")
								' objFileToken.close
								' set objFileToken=nothing
							%>

							
							</div>
							<br>
							

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
