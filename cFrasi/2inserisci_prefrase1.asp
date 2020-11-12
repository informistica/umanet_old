<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Inserisci prefrase</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<meta charset="utf-8">

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
	<script src="js/plugins/gritter/jquery.gritter.min.js"></script>
<!-- Theme framework -->
	<script src="js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="js/application.min.js"></script>
	<!-- Just for demonstration -->
	<script src="js/demonstration.min.js"></script>



	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />


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
     <body class='theme-<%=session("stile")%>'>
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
 TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Sottoparagrafo=Request.QueryString("Sottoparagrafo")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
 %>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Inserisci compito</h1>

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
						<li>
							<a href="#">Inserisci compito</a>
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
				       <h3> <i class="icon-comments"></i> <%=Capitolo%>: <%=TitoloParagrafo%>  <% if SottoParagrafo<>"" then response.write(" - "&Sottoparagrafo) end if%></h3>
			          </div>
				      <div class="box-content">

 <%





Scadenza=Request.QueryString("Scadenza")
Num=Request.QueryString("Num")
txtDomande = Request.Form("MyTextArea")

if Scadenza <>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
  Scadenza=cdate(Request.QueryString("Scadenza"))
end if

Capitolo=Request.QueryString("Capitolo")
Paragrafo=Request.QueryString("Paragrafo")
Modulo=Request.QueryString("Modulo")
BoxApro=Request.QueryString("BoxApro")
Segnalibro=Request.QueryString("Segnalibro")


' se inserisco in più sessioni devo accodare le ultime frasi dietro alle prime, quindi mi serve la posizione raggiunta per proseguire da li

function ReplaceCar(sInput)
dim sAns
  sAns = Replace(sInput,chr(224),"a"&Chr(96))
  sAns = Replace(sAns,chr(225),"a"&Chr(96))
  sAns = Replace(sAns,chr(232),"e"&Chr(96))
  sAns = Replace(sAns,chr(233),"e"&Chr(96))
  sAns = Replace(sAns,chr(236),"i"&Chr(96))
  sAns = Replace(sAns,chr(237),"i"&Chr(96))
  sAns = Replace(sAns,chr(242),"o"&Chr(96))
  sAns = Replace(sAns,chr(243),"o"&Chr(96))
  sAns = Replace(sAns,chr(249),"u"&Chr(96))
  sAns = Replace(sAns,chr(250),"u"&Chr(96))
  sAns = Replace(sAns, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
  sAns=  Replace(sAns,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql
  sAns=  Replace(sAns,chr(58),Chr(44)) ' sostituisco : con , per non disturbare la creazione del file
 ' sAns=  Replace(sAns,"&","e")
  sAns=  Replace(sAns,"/","-")
  sAns=  Replace(sAns,"\","-")
 ' sAns=  Replace(sAns,"?",".")
  sAns=  Replace(sAns,"*","x")
  'sAns=  Replace(sAns,"<","_")
  'sAns=  Replace(sAns,">","_")

ReplaceCar = sAns
end function

if CodiceSottopar<>"" then
	QuerySQL="Select count(*) from preFrasi where Id_Paragrafo='"&Paragrafo&"' and Id_Sottoparagrafo='"&CodiceSottopar&"';"
else
    QuerySQL="Select count(*) from preFrasi where Id_Paragrafo='"&Paragrafo&"';"
end if
set rsTabella=ConnessioneDB.Execute (QuerySQL)
if rsTabella(0)>0 then
		if CodiceSottopar<>"" then
		  QuerySQL="Select max(Posizione) from preFrasi where Id_Paragrafo='"&Paragrafo&"'and Id_Sottoparagrafo='"&CodiceSottopar&"';"
		else
			QuerySQL="Select max(Posizione) from preFrasi where Id_Paragrafo='"&Paragrafo&"';"
		end if
	set rsTabella=ConnessioneDB.Execute (QuerySQL)
	contPos=rsTabella(0)
else
	 contPos=0
end if
'response.write(QuerySQL& "<br>" & "conPos="&contPos)
 cont=0


 ' se ho lasciato vuota la text area faccio inserimento una ad una
if  txtDomande="" then
		for k=1 to Num
			   Immagine=Request.Form("chkImmagine"&k)
			   Domanda = Request.Form("txtDomanda"&k)

      if ((instr(Domanda,"tp://")<>0) or (instr(Domanda,"tps://")<>0)) Then
      else
       Domanda =  ReplaceCar(Domanda)
      end if
			if Domanda<>"" then ' controllo per le righe vuote
					if (strcomp(Immagine,"si")=0) then
					' se è prevista l'immagine
						img=1
					else
						img=0
					end if
					'Esecuzione della query per
		'QuerySQL="INSERT INTO preDomanda (Id_Mod, Id_Paragrafo,Quesito,Eseguita) SELECT '" & Modulo & "','" & Paragrafo "','" & Domanda & "'," & 0 & "';"
				if Scadenza <>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then
				   QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & ",'" & Scadenza & "'," & img & "," & cFile & ",'" & CodiceSottopar & "';"
				   else
				'   QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & "," & img & "," & cFile & ",'" & CodiceSottopar & "';"  fine_anno
				  QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & ",'" & fine_anno & "'," & img & "," & cFile & ",'" & CodiceSottopar & "';"
				   end if
				   response.write(QuerySQL& "<br>")
				   'end if
				   '  Set objFSO = CreateObject("Scripting.FileSystemObject")
		'        				url1="C:\Inetpub\umanetroot\anno_2012-2013\logpreFrasi.txt"
		'        				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
		'        				objCreatedFile.WriteLine(querySQL)
		'        				objCreatedFile.Close
				   ConnessioneDB.Execute QuerySQL
			end if
		next
else
     Scadenza=Request.Form("date3")
        'strText = MyTextArea.Value
		strText = txtDomande
        arrLines = Split(strText, vbCrLf)
    k=1
	For Each strLine in arrLines
	      img=0
		  cFile=0
        if instr(strLine,"$")=0 and instr(strLine,"#")=0 then ' senza immagine nè file
		  img=0
		  cFile=0
		  Domanda=strLine
		else ' immagine
		     if instr(strLine,"$")<>0 then ' immagine
		      img=1
		      Domanda=left(strLine,instr(strLine,"$")-1)
		     end if
			 if instr(strLine,"#")<>0 then ' file
		      cFile=1
		       if img=0 then
			      Domanda=left(strLine,instr(strLine,"#")-1)
				end if
		     end if


		end if
  if ((instr(Domanda,"tp://")<>0) or (instr(Domanda,"tps://")<>0)) Then
  else
			Domanda =  ReplaceCar(Domanda)
  end if
				 if Scadenza <>"" and not (strcomp(Scadenza,"gg/mm/aaaa")=0) then

				    QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & ",'" & Scadenza & "'," & img & "," & cFile & ",'" & CodiceSottopar & "';"
				 else
				 ' QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & "," & img & "," & cFile & ",'" & CodiceSottopar & "';"

				  QuerySQL="  INSERT INTO preFrasi (Id_Mod, Id_Paragrafo,Quesito,Eseguita,Posizione,Scadenza,Img,Files,Id_Sottoparagrafo)  SELECT '" & Modulo & "','" & Paragrafo & "', '" & Domanda & "'," & 0 & "," & contPos+k & ",'" & fine_anno & "'," & img & "," & cFile & ",'" & CodiceSottopar & "';"
			     end if
				  '  response.write(QuerySQL)
				   ConnessioneDB.Execute QuerySQL

		response.write server.htmlencode(Domanda) & "<br>"
       k=k+1
	Next


end if%>
<%

   ' notifico l'inserimento agli studenti della classe
		   Segnalibro=Request.QueryString("Segnalibro")
   		   id=Request.QueryString("id")

		   Avviso="Ho inserito nuovi compiti (" & Num & " - f) "
		   Azione=domino&homesito&"/home_app.asp?divid="&Session("divid")&"&id_classe="&Session("id_Classe")&"&classe="&Session("Classe")&"&cartella="&Session("Cartella")
		  Azione=replace(Azione,"\","/")
		  ' Azione= "https://localhost/anno_2012-2013/UECDL/home_app.asp?divid=quarta&id_classe=5COM&classe=3PC&cartella=3PC"
		   Azione = Azione & "&dividApro="& BoxApro
		   Azione = Azione & "#" & Segnalibro
			Id_Classe=Session("Id_Classe")
			'divid=Request.QueryString("divid")
			Avviso = Replace(Avviso, Chr(34), "'")' sostituisco gli apici " con l'apice singolo
			Avviso=  Replace(Avviso,"'",Chr(96)) ' sostituisco l'apice ' con quello storto per non disturbare la sintassi sql


Commentatore= Session("Cognome") & " " & left(Session("Nome"),1) &"."
' disabilito avviso di inserimento compiti
 '  QuerySQL="  INSERT INTO AVVISI2 (Testo,Data,Azione,Commentatore)  SELECT '" & Avviso & "','" & FormatDateTime(now(),2)   & "', '" & Azione & "', '" & Commentatore & "';"
'  ' ConnessioneDB.Execute (QuerySQL)
'   QuerySQL="select max(ID_Avviso) from AVVISI2"
'   rsTabella=ConnessioneDB.Execute (QuerySQL)
'   maxID=rsTabella(0)
'   QuerySQL="  INSERT INTO AVVISI_CLASSE (Id_Avviso,Id_Classe)  SELECT '" & maxID & "','" & Id_Classe  & "';"
' '  ConnessioneDB.Execute (QuerySQL)

	On Error Resume Next
		If Err.Number = 0 Then%>
	<span class="alert-success">
	<%

		Response.Write "Inserimento avvenuto! "
	Else
	%>
	<span class="alert-error">
	<%

		Response.Write Err.Description
		Err.Number = 0
	End If


   %>
   </span>
	</font>

		  <h5><a href="2inserisci_prefrase.asp?Capitolo=<%=Capitolo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&CodiceSottopar=<%=CodiceSottopar%>">Continua ...</a></h5>
		<p>&nbsp;</p>

				<p><p>

				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h5><a href="../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>&dividApro=<%=BoxApro%>#<%=BoxApro%>"> Torna al Libro </a></h5>


        <!--    <a href="#modal-1" role="button" class="btn notify" data-notify-title="Success!" data-notify-message="The user has been successfully edited.">Basic notification</a>  -->












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
