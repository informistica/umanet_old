<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Valutazione frasi</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

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
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>

	<!--[if lte IE 9]>
		<script src="../../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->




   <script language="javascript" type="text/javascript">
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../home.asp"
//location.href=window.history.back();
 }
 </script>
</head>


<% Response.Buffer = true
   On Error Resume Next
    if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
   <body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

  <% end if %>

  <%

     Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")
   Modulo=Request.QueryString("Modulo")
   Cartella=Request.QueryString("Cartella")

    %>
	<div id="navigation">

        <%


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        <!--#include file="../service/gestione_errori.asp" -->



	</div>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Valutazione frasi </h1>

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
							<a href="#">Valutazioni</a>
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
		   			   <div class="box-content">



  <%
    Capitolo=Request.QueryString("Capitolo")
   Paragrafo=Request.QueryString("Paragrafo")
   NumRec=clng(Request.Form("TxtNUMREC"))
   set objFSO=Server.CreateObject("Scripting.FileSystemObject")
  ' response.write(numrec)
  on error resume next
 Notifica=Request.Form("txtNotifica")  ' se è valorizzata ignoro i feedback'
response.write("<br>NumRec="&NumRec)
  for k=1 to NumRec ' per scorrere tutto il form e fare un update ad ogni ciclo
   Domanda = Request.Form("txtDomanda"&k)
   ID=clng(Request.Form("txtCodiceDomanda"&k))

   Spiegazione=Request.Form("txtSpiegazione"&k)
   'TestoDomandaPlus=Request.Form("TestoDomandaPlus")


   VAL=clng(Request.Form("txtVAL"&k))
   INQUIZ=clng(Request.Form("txtINQUIZ"&k))
   Segnalata=Request.Form("txtSegnalata"&k)
   Segno=Request.Form("txtSegno"&k)
   Data=Request.Form("txtDataDomanda"&k)
   CodiceAllievo=Request.Form("txtCodiceAllievo"&k)
   Motivazione=Request.Form("txtSegnalazione"&k)
   daValutare=Request.Form("cbVal"&k)
   response.write("<br>daValutare"&k&"="&daValutare)
   
if daValutare<>"" then
				'se segno=1 vuol dire che è feedback positivo'
				if Segnalata="" then
					Segnalata=0
				elseif Segnalata=1 and Segno=1 Then
						Segnalata=2
						elseif Segnalata=1 and Segno=0 Then
								Segnalata=1
							else
							Segnalata=0
				end if

				' per la spiegazione della domanda
				'  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"& Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&ID&".txt"
					'url=Replace(url,"\","/")

				' response.write("<br>" &Request.Form("url"&k) )
				url=Request.Form("url"&k)
				url_feedback=left(url,instr(url,".")-1)
					url_feedback=url_feedback&"_feedback.txt"


				if session("Admin")=true then
					QuerySQL ="UPDATE Frasi SET Chi = '" & Domanda & "', Voto = '" & VAL & "', In_Quiz = '" & INQUIZ &"', Segnalata='" &Segnalata&"', Data='" &CDate(Data) &"' WHERE CodiceFrase =" &ID&";"
					'response.write("<br>" & QuerySQL)
					ConnessioneDB.Execute(QuerySQL)
					end if
				'response.write(QuerySQL)  <br> <%
				'CREAZIONE FILE DI TESTO PER AGGIORNARE LA SPIEGAZIONE DELLA DOMANDA E , nel caso di domanda plus, il testo della domanda plus


				if clng(Segnalata)=1 or clng(Segnalata)=2  then


					' se è segnalata aggiorno il file della spiegazione
					'objFSO.DeleteFile url
					'  response.Write("<br>Cancello : " &url)
					'Set objCreatedFile = objFSO.CreateTextFile(url, True)
					'' Write a line with a newline character.
					'objCreatedFile.WriteLine(Spiegazione)
					'	  response.Write("<br>Creo : " &Spiegazione)
					''Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
					'objCreatedFile.Close

					' se è segnalata creo file per il feedback

					' se esiste già lo cancello e poi lo ricreo per eventuali aggiornamenti

					if day(date()) < 10 then
					giorno="0" & day(date())
					else
					giorno=day(date())
					end if
					if len(year(date()) ) = 2 then
					anno="20"& year(date())
					elseif len(year(date()) ) =  3 then
					anno="2"& year(date())
					else
					anno=year(date())
					end if
					if month(date()) < 10 then
					mese="0" & month(date())
					else
					mese=month(date())
					end if

					DataAvviso = giorno & "/" & mese& "/" & anno


					if (motivazione<>"") and (Notifica="") then
						dim objFSO
						set objFSO=Server.CreateObject("Scripting.FileSystemObject")
						if objFSO.FileExists(url_feedback) then
							objFSO.DeleteFile url_feedback
						end if

						Set objCreatedFile = objFSO.CreateTextFile(url_feedback, True)
						'response.write(url_feedback)
						'  response.Write("<br><br>Creo file feedback : " &url_feedback)
						if clng(Segnalata)=1  Then
							response.Write("<br><font color='red'>Contenuto feedback :</font> "& server.htmlencode(Motivazione))
						End if
						if clng(Segnalata)=2  Then
							response.Write("<br><font color='green'> Contenuto feedback :</font> "& server.htmlencode(Motivazione))
						End if
						objCreatedFile.WriteLine(Motivazione)
						objCreatedFile.Close
						 
							if objFSO.FileExists(url_feedback) then
									'response.write("<br>Creato file di feedback:"&url_feedback)
									'ed invio notifica sul quaderno dello stud

									Testo=Motivazione
									Azione="<a  target=blank href=2inserisci_valutazione_frase.asp?cla="&Session("Id_Classe")&"&cod="&CodiceAllievo&"&CodiceFrase="&ID&"&cartella="&cartella&"&classe="&Session("Id_Classe")&"&CodiceTest="&CodiceTest&"&Paragrafo="&Paragrafo&"&MO="&Modulo&"&Capitolo="&Capitolo&">Ho segnalato una tua frase !</a>"
									Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."
									QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo, Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Testo& "','" & Azione & "','" & DataAvviso & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
									if (strcomp(CodiceAllievo,Session("CodiceAllievo"))<>0) and (Notitifca="") then ' evito di notificare a me stesso e non invio notitiche se non richiesto
									ConnessioneDB.Execute(QuerySQL)
									end if
							else
									response.write("<br>Impossibile creare file :"&url_feedback)
									response.write ("<br>"&Err.Description)
							end if
						 

						set objFSO = Nothing
						set objCreatedFile = Nothing
					end if

					

				end if

				'response.write("250")

				'Create the FSO.
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
				'CANCELLA LA VECCHIA VERSIONE DEL FILE11
				'response.write(Cartella)
				'response.write(url)

				' url1="C:\Inetpub\umanetroot\Anno_2010-2011_ITC\logFile.txt"
				'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
				'				objCreatedFile.WriteLine(url)
				'				objCreatedFile.Close
				'

				'se è stata segnalata aggiorno il file della spiegazione
				' LA MODIFICA DEL FILE DI TESTO DELLA SPIEGAZIONE E'DISABILITATA PERCHé , per generare URL per cancellare la vecchia versione del file
				' devo fare in modo che il Paragrafo vari a seconda della domanda, perchè nel form chiamante posso avewre domande dello stesso modulo
				' ma di paragrafi diversi, quindi DOVREI TROVARE IL MODO DI RICAVARE PER OGNI DOMANDA IL PARAGRAFO DI APPARTENENZA COSì da poter
				' generare URL, per ora la lascio così

				'objFSO.DeleteFile url
				'Set objCreatedFile = objFSO.CreateTextFile(url, True)
				'' Write a line with a newline character.
				'objCreatedFile.WriteLine(Spiegazione)
				''Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
				'objCreatedFile.Close
				' per aggiornare la domanda plus
				'if Tipodomanda=1 then
				'	objFSO.DeleteFile url4
				'	Set objCreatedFile = objFSO.CreateTextFile(url4, True)
				'	' Write a line with a newline character.
				'	objCreatedFile.WriteLine(TestoDomandaPlus)
					'Use objCreatedFile and objOpenedFile to manipulate the corresponding files.
				'	objCreatedFile.Close
				'end if
	end if ' if daValutare
next
On Error Resume Next
If Err.Number = 0 Then

Response.Write "<span class=alert-success><b><br><br>Modifica avvenuta!</b></span>"
Else
Response.Write Err.Description
Err.Number = 0
End If
%>
<h4><a href='../cClasse/home_app.asp?id_classe=<%=Session("Id_Classe")%>'> Torna al libro </a></h4>



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
