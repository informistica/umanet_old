<%@ Language=VBScript %>
<!doctype html>

<%  if session("admin") = false then
		response.redirect("../../../../index.html")
		end if
		%>

<html>

<head>

	<title>Admin</title>

	<%
   on error resume next
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
     Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection")
	 Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection")
	 classe=Session("Cartella")
	 if classe="" then
	  'classe=request.QueryString("classe")
	  classe=session("Cartella")
	 end if
		%>




	<!-- #include file = "../var_globali.inc" -->
	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

	<!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
	<!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->

	<!-- #include file = "../service/controllo_sessione.asp" -->

	<!-- #include file = "../service/formatta_data_LO.asp" -->
	<!-- #include file = "../service/replacecar.asp" -->

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

	<!-- Datepicker new-->
	<link rel="stylesheet" href="../../css/plugins/datepicker/datepicker.css">




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

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->
	<!-- CKEditor -->
  	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>

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
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
-->

	<!-- Datepicker -->

	<!-- <script src="../js/plugins/datepicker/bootstrap-datepicker.it.js"></script> -->

	<script src="../../js/jquery-ui.js"></script>
	<script src="../../js/datapicker_it.js"></script>



	<script language="javascript" type="text/javascript">

	/*	function modifica(id) {

			document.Aggiorna.action = "modifica_avviso.asp?tipoAvviso=1&i=" + id
			document.Aggiorna.submit();
		}
*/
		function Avviso(id) {
			if (id == 1)
				window.alert("Nessun modulo da visualizzare, inserisci moduli");
			else if (id == 2)
				window.alert("Inserisci i moduli1");

		}

		function PopUpWindow(w, h) {
			var winl = (screen.width - w) / 2;
			var wint = (screen.height - h) / 2;

			window.open('../upload_resize/ex2_imgclasse.asp', '../upload_resize/ex2_imgclasse.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top=' + wint + ',left=' + winl)
		}
		// -->


		function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
			var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
			upload.focus();
		}


		function validate3() {


			if (votazioni.Txt1.value == "") {
				alert("Non hai inserito il codice della votazione ");
				votazioni.txt1.setfocus();
				return 0;
			}
			else {
				document.votazioni.action = "votazioni_impresa_spoglio.asp";
				document.votazioni.submit();


			}

		}
	</script>

	<script language="javascript" type="text/javascript">
		function cancella_avviso() {

			if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {
				document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=1&Id_Classe=<%=Id_Classe%>";
				//document.dati.action = "../home.asp"
				document.Aggiorna.submit();
			}
		}


		function validate2() {
			var espressione = /^[0-9]{2}\/[0-9]{2}\/[0-9]{4}$/;
			var stringa = insavviso.txtAvviso.value;
			var scad = insavviso.txtScadenza.value;

			if (stringa == "") {
				alert("Non hai inserito il testo dell'avviso ");
				insavviso.txtAvviso.setfocus();
				return 0;
			}
			else
				if (CKEDITOR.instances.txtDescrizione.getData()== "") {
					alert("Non hai inserito la spiegazione del compito");
					//insavviso.txtAzione.setfocus();
					return 0;
				}
				else
					if (!espressione.test(scad)) {
						alert("Non hai inserito la data di scadenza");
						return 0;
					} else {
						var capitolo = $("#selcap").val();
						var paragrafo = $("#selpar").val();
						var testo= CKEDITOR.instances.txtDescrizione.getData();
						testo=encodeURIComponent(testo);
						if ($("#selsottopar").val() != "Seleziona un sottoparagrafo" && $("#selsottopar").val() != null && $("#selsottopar").val() != "Nessun sottoparagrafo disponibile") {
							var sottopar = $("#selpar").val() + "." + $("#selsottopar").val();
						}

						document.insavviso.action = "inserisci_avviso.asp?Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&posizione=<%=posizione%>&divid=<%=divid%>&cap=" + capitolo + "&par=" + paragrafo + "&sottopar=" + sottopar;
						document.insavviso.submit();


					}

		}
	</script>


	<%

	Dim objFSO,objCreatedFile
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim sRead, sReadLine, sReadAll, objTextFile
	Set objFSO = CreateObject("Scripting.FileSystemObject")


%>



</head>





<body class='theme-<%=session("stile")%>'>

	<div id="navigation">

		<!-- #include file = "../include/navigation.asp" -->

	</div>




	<div class="container-fluid" id="content">

		<!-- #include file = "../include/menu_left.asp" -->

		<div id="main">
			<div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Pannello di Amministrazione </h1>
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
							<a href="#more-login.html">Home</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-files.html">Admin</a>

						</li>

					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>


				<div class="row-fluid">
					<div class="box">
						<div class="box-title">
							<h3>
								<i class="icon-reorder"></i>
								Impostazioni
							</h3>
							<div class="actions">
								<a href="#" class="btn btn-mini content-refresh"><i class="icon-refresh"></i></a>
								<a href="#" class="btn btn-mini content-remove"><i class="icon-remove"></i></a>
								<a href="#" class="btn btn-mini content-slideUp"><i class="icon-angle-down"></i></a>
							</div>
						</div>
						<div class="box-content">
							<div class="accordion" id="accordion2">
								<div class="accordion-group">
									<div class="accordion-heading">

										<a class="accordion-toggle"
											href="../cClasse/quaderno.asp?DataClaq=<%=inizio_anno%>&DataClaq2=<%=fine_anno%>&id_classe=<%=Id_classe%>&cod=<%=Session("CodiceAllievo")%>">
											<center>Quaderno di Admin</center>
										</a>
									</div>

								</div>
								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse2">
											<center>Accesso utente</center>
										</a>
									</div>

									<% Id_Classe=Request.QueryString("Id_Classe")
 QuerySQL="Select * from Setting where Id_Classe='" & Id_Classe &"'"
 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

  %>

									<div id="collapse2" class="accordion-body collapse">
										<div class="accordion-inner">

											<div class="box box-bordered">
												<div class="box-title">
													<h3><i class="icon-list-ul"></i> Element settings</h3>
												</div>


												<!-- inizio include-->

												<form method="POST" class='form-horizontal form-striped'
													action="config_accesso_utente.asp?Id_Classe=<%=Id_Classe%>&divid=<%=divid%>">

													<!--
 <div class="control-group">
											<label for="textfield" class="control-label">Datepicker</label>
											<div class="controls">
												<input type="text" name="textfield" id="datepicker" class="input-medium datepick">
												<span class="help-block">As dropdown</span>
											</div>
										</div>-->

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Ogni studente visualizza solo il suo quaderno">
															</i>
															<b>&nbsp;Privato</b> </label>
														<div class="controls">
															<% if (rsTabella1.fields("Privato")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckPrivato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckPrivato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckPrivato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckPrivato" checked="true"
																value="0">No
															<% end if%>
														</div>
													</div>



													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Ad ogni compito caricato viene assegnato automaticamente 1 punto">
															</i>
															&nbsp;<b> Valutato</b>

														</label>
														<div class="controls">
															<% if (rsTabella1.fields("Valutato")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckValutato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckValutato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckValutato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckValutato" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita lo svolgimento dei quiz a risposta multipla">
															</i>
															&nbsp; <b>Quiz abilitati </b>
														</label>
														<div class="controls">
															<% 'response.write(rsTabella1.fields("TestAbilitato"))
									   %>
															<% if ( rsTabella1.fields("TestAbilitato")=1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckAbilitato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckAbilitato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckAbilitato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckAbilitato" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita il controllo qualità dei quiz"> </i>
															&nbsp; <b>Verifica quiz</b>
														</label>
														<div class="controls">
															<% 'response.write(rsTabella1.fields("TestAbilitato"))
									   %>
															<% if ( rsTabella1.fields("ValidaQuiz")=1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckValidaQuiz" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckValidaQuiz" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckValidaQuiz" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckValidaQuiz" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di utilizzare la chat">
															</i>
															&nbsp;<b> Chat abilitata </b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("ChatAbilitata")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckChatAbilitata" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckChatAbilitata" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckChatAbilitata" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckChatAbilitata" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di trasferire testi">
															</i>
															&nbsp;<b> Copia/Incolla </b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("CIAbilitato")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckCIAbilitato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckCIAbilitato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckCIAbilitato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckCIAbilitato" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita il controllo per impedire copia ed incolla">
															</i>
															&nbsp;<b> Controllo JS</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("JSAbilitato")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckJSAbilitato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckJSAbilitato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckJSAbilitato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckJSAbilitato" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita possibilit&agrave; di bonus per risposte lunghe">
															</i>
															&nbsp;<b> Bonus</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("DVAbilitato")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckDVAbilitato" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckDVAbilitato" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckDVAbilitato" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckDVAbilitato" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Voto massimo in classifica (8,9,10)"> </i>
															&nbsp; <b>Voto max</b>
														</label>
														<div class="controls">

															<p><input type="text" name="TxtVVMax" size="1"
																	value="<%=rsTabella1.fields("ScalaValutaz")%>"
																	class="input-mini">
																<b>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di votare con le stelline (non attivo)">
															</i>
															&nbsp; <b>Voto attivo </b>
														</label>
														<div class="controls">
															<%
										  'response.write("Voto attivo "& rsTabella1.fields("Id_Classe"))
										  if (rsTabella1.fields("VotoAttivo")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckVotoAttivo" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckVotoAttivo" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckVotoAttivo" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckVotoAttivo" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Mostra i dettagli delle stelline"> </i>
															&nbsp; <b>Voto palese </b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("VotoPalese")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckVotoPalese" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckVotoPalese" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckVotoPalese" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckVotoPalese" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="MAX Mi piace/Non mi piace    (Voto = 5 + Max Voto )">
															</i>
															&nbsp; <b>Max Voto </b>
														</label>
														<div class="controls">
															<p><input type="text" name="TxtMaxStelline" size="1"
																	value="<%=rsTabella1.fields("MaxStelline")%>"
																	class="input-mini">
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Voto in proporzione a parametro (ancora non attiva)">
															</i>
															&nbsp;<b> Runner</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("Runner")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckLepre" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckLepre" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckLepre" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckLepre" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di registrarsi">
															</i>
															&nbsp;<b> Registrazione</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("Registra")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckReg" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckReg" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckReg" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckReg" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di condividere nodi">
															</i>
															&nbsp;<b> Rete di Nodi</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("Nodi")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckNodi" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckNodi" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckNodi" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckNodi" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>

														<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Abilita la possibilit&agrave; di rispondere alle domande scadute">
															</i>
															&nbsp;<b> Recupero attivo</b>
														</label>
														<div class="controls">
															<% if (rsTabella1.fields("RecuperoAttivo")= 1)  then  %>
															<INPUT TYPE="RADIO" NAME="CheckRecupero" checked="true"
																value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckRecupero" value="0">No
															<% else %>
															<INPUT TYPE="RADIO" NAME="CheckRecupero" value="1">Si
															<INPUT TYPE="RADIO" NAME="CheckRecupero" checked="true"
																value="0">No

															<% end if %>
														</div>
													</div>



													<div class="control-group">
														<label class="control-label">
															<i class="glyphicon-circle_info" rel="tooltip"
																data-placement="right"
																title="Classe che si sta configurando"> </i>
															&nbsp;<b> Classe </b>
														</label>
														<div class="controls">
															<select name="txtClasse" disabled="disabled">

																<%'visualizzo la combo per la scelta della classe
                                               querySQL="Select * from Classi"
                                               set rsTabella=ConnessioneDB.execute(QuerySQL)
                                              do while not rsTabella.eof
                                                if rsTabella.fields("Id_Classe")= Id_Classe then %>
																<option selected
																	value="<%=rsTabella.fields("ID_Classe")%>">
																	<%=rsTabella.fields("Classe")%> </option>
																<% else %>
																<option value="<%=rsTabella.fields("ID_Classe")%>">
																	<%=rsTabella.fields("Classe")%> </option>
																<% end if%>
																<% rsTabella.movenext()
                                                   loop %>
															</select>
														</div>
														<hr>
														<label class="control-label">
															<b>Aggiorna</b>&nbsp;
														</label>
														<div class="controls">
															<p><input type="submit" class="btn btn-primary"
																	value="Invia" name="B1">
														</div>
													</div>
													</p>
												</form>
											</div>
											</b>
										</div>
									</div>
								</div>


								<!-- fine include-->



								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse3">
											<center>Classi</center>
										</a>
									</div>
									<div id="collapse3" class="accordion-body collapse">
										<div class="accordion-inner">

											<form method="POST" class='form-horizontal form-striped'
												action="inserisci_classe.asp" target="_blank">


												<h4>Gestione classi</h4>&nbsp;
												</b>


												<% QuerySQL="SELECT * FROM Classi order by ID_AS asc "

	    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	   ' carico il vettore delle date di valutazione
		 %>
												<table class="table table-hover table-nomargin table-striped">
													<thead>
														<tr>
															<th class='hidden-480'><b>Id_Classe</b></th>
															<th><b>Classe</b></th>
															<th><b>Nome</b></th>
															<th class='hidden-480'><b>Posizione</b></th>
															<th><b>Modifica</b></th>
															<th><b>Visibile</b></th>
															<th><b>Cancella</b></th>
														</tr>
														<thead>
															<%
		do while not rsTabella.eof%>
															<tr>
																<td class='hidden-480'>
																	<%=rsTabella.fields("Id_Classe")%></td>
																<td><a
																		href="../cClasse/home_app.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>">
																		<%=rsTabella.fields("Classe")%></td>
																		<td><a
																		href="../cClasse/home_app.asp?id_classe=<%=rsTabella.fields("Id_Classe")%>">
																		<%=rsTabella.fields("Nome")%></td>
																<td class='hidden-480'>
																	<%=rsTabella.fields("Posizione")%></td>

																<td><a href='#modal-1'
																		onClick='modifica("<%=rsTabella.fields("Id_Classe")%>","<%=rsTabella.fields("Classe")%>")'
																		data-toggle='modal'><i
																			style='text-decoration:none'
																			class='icon-pencil'
																			title='Modifica classe'></i></a></td>
																<td><%
																
																if (rsTabella.fields("Visibile")=1) then 
																	tipo="open"
																else 
																	tipo="close"
																end if
																%> 
																<a  onClick="cambia_visibilita('<%=rsTabella.fields("Id_Classe")%>','<%=rsTabella.fields("Visibile")%>');"> <i id="occhio_<%=rsTabella.fields("ID_Classe")%>" class="icon-eye-<%=tipo%>" title="Modifica "></i></a>
																
																</td>			
																<td><a onClick="return window.confirm('Vuoi veramente cancellare la classe ?');"
																		href="cancella_classe.asp?cancella=1&Id_Classe=<%=rsTabella.fields("Id_Classe")%>&Classe=<%=rsTabella.fields("Classe")%>&divid=<%=divid%>"><i
																			class="icon-trash"></i></td>
															</tr>

															<% 'posizione=rsTabella.fields("Posizione")
			      rsTabella.movenext()
		loop %>
												</table>
												<p><input type="submit" value="Inserisci nuova classe" name="B1"
														class="btn"><br>

												</p>
											</form>


										</div>
									</div>
								</div>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse4">
											<center>Periodi valutazione</center>
										</a>
									</div>
									<div id="collapse4" class="accordion-body collapse">
										<div class="accordion-inner">


											<form method="POST" class='form-horizontal form-striped'
												action="config_periodi_valutazione.asp?Id_Classe=<%=Id_Classe%>&divid=<%=divid%>">
												<h4>Configura periodi di valutazione</h4>&nbsp;

												<% QuerySQL="SELECT * FROM [dbo].[3PERIODI] Where ID_Classe='"& ID_Classe &"' order by Data;"

	    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	   ' carico il vettore delle date di valutazione
		 %><br>

												<table class="table table-hover table-nomargin table-striped">
													<tr>
														<thead>
															<th>Data</th>
															<th>Iniziale</th>
															<th>Cancella</th>
													</tr>
													</thead>

													<% if  rsTabella.eof  then %>

													<tr>
														<td><%=inizio_anno%></td>

														<td><INPUT TYPE="RADIO" NAME="CheckPeriodo" checked="true"
																value="0"></td>
														<td>&nbsp;</td>
													</tr>

													<%else%>

													<tr>
														<td><%=inizio_anno%></td>

														<td><INPUT TYPE="RADIO" NAME="CheckPeriodo" checked="false"
																value="0"></td>
														<td>&nbsp;</td>
													</tr>

													<% end if%>
													<%i=1
		do while not rsTabella.eof


		  %>
													<tr>
														<td><%=rsTabella.fields("Data")%></td>
														<% if rsTabella.fields("Iniziale")=1 then %>
														<td><INPUT TYPE="RADIO" NAME="CheckPeriodo" checked="true"
																value="<%=i%>"></td>
														<%else%>
														<td><INPUT TYPE="RADIO" NAME="CheckPeriodo" value="<%=i%>"></td>
														<%end if%>

														<td><a
																href="cancella_data.asp?cancella=1&data=<%=rsTabella.fields("Data")%>&idperiodo=<%=rsTabella.fields("ID_Periodo")%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>"><i
																	class="icon-trash"></i></a></td>
													</tr>

													<% i=i+1
			    DataCla=rsTabella("Data")
			    rsTabella.movenext()
		loop %>

													<tr>
														<td colspan="3"><input type="hidden" name="txtDataCla" size="10"
																value="<%=DataCla%>">


															<p>Data: <input type="text" name="txtDataVal"
																	id="datepicker" class="input-medium datepick" /></p>

															<br><br>

															<p><input type="submit" value="Aggiorna" name="B1"
																	class="btn">
														</td>
													</tr>
												</table>
												<!-- <div class="controls">
		<input type="text" name="textfield" id="datepicker1" class="input-medium datepick">

											</div>-->

											</form>


										</div>
									</div>
								</div>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse5">
											<center>Moduli didattici</center>
										</a>
									</div>
									<div id="collapse5" class="accordion-body collapse">
										<div class="accordion-inner">
											<%
		 querySQL="Select * from Classi where Id_Classe='"&Id_Classe&"';"
	   set rsTabella=ConnessioneDB.execute(QuerySQL)
	    classe=rsTabella.fields("Classe") ' serve per la query di inserimento della data, mi prende la classe selezionata
		 cartella=rsTabella.fields("Cartella")
		   session("classe")=classe
		%>

											<form method="POST" target="_blank" class='form-horizontal form-striped'
												action="inserisci_modulo.asp?Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&cartella=<%=cartella%>&posizione=<%=posizione%>&divid=<%=divid%>">
												<fieldset style="margin: 0 auto 0 auto;">
													<LEGEND><b><%  response.write("Configurazione moduli didattici  ") %>
														</b></LEGEND><br>
													<select name="txtClasse" disabled="disabled">

														<%'visualizzo la combo per la scelta della classe


		 %>
														<option selected value="<%=rsTabella.fields("ID_Classe")%>">
															<%=rsTabella.fields("Classe")%> </option>

													</select>
													<b>Classe</b><br>
													<% QuerySQL="SELECT * FROM MODULI_CLASSE Where ID_Classe='"& ID_Classe &"';"

	    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	   ' carico il vettore delle date di valutazione
		 %><br>
													<table class="table table-hover table-nomargin table-striped">
														<thead>
															<tr>
																<th class='hidden-480'><b>ID_Mod</b></th>
																<th><b>Modulo</b></th>
																<th><b>Modifica</b></th>
																<th class='hidden-480'><b>Cancella</b></th>
															</tr>
														</thead>
														<tbody>
															<%
		do while not rsTabella.eof
		' controllo se il modulo è condiviso con qualche altra classe
		QuerySQL="SELECT * FROM CLASSI_MODULI_CONDIVISI Where ID_Modulo='"& rsTabella.fields("ID_Mod") &"' and Id_Classe<>'"& Id_Classe &"';"
	    Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
		cond=""
		while not rsTabella1.eof
		   cond =cond &rsTabella1("Classe")&"/"
		rsTabella1.movenext
		wend

		%>
															<tr>
																<td class='hidden-480'><%=rsTabella.fields("ID_Mod")%>
																</td>
																<td><%=rsTabella.fields("Titolo")%>
																	<%if cond<>"" then
			   response.write(" (Condiviso con " &cond&")")
			   end if%>
																</td>

																<td><a
																		href="modificamodulo.asp?ID_Mod=<%=rsTabella.fields("ID_Mod")%>&Classe=<%=Classe%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>"><img
																			src="../../img/Next.gif" width="14"
																			height="13"></a></td>
																<td class='hidden-480'><a
																		onClick="return window.confirm('Vuoi veramente cancellare il modulo?');"
																		href="cancella_modulo.asp?cancella=1&Id_Mod=<%=rsTabella.fields("ID_Mod")%>&Id_Classe=<%=Id_Classe%>&Classe=<%=Classe%>"><i
																			class="icon-trash"></i></a></td>
															</tr>

															<% 'posizione=rsTabella.fields("Posizione")
			      rsTabella.movenext()
		loop %>
														</tbody>
													</table>
													<p><input type="submit" value="Inserisci nuovo modulo" name="B1"
															class="btn"><br>
											</form>

											<form method="POST" class='form-horizontal form-striped' target="_blank"
												action="inserisci_modulo.asp?umanet=1&Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&posizione=<%=posizione%>&divid=<%=divid%>">
												<% QuerySQL="SELECT * FROM MODULI_CLASSE_UMANET Where ID_Classe='"& ID_Classe &"';"

	    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
	   ' carico il vettore delle date di valutazione
		 %><br>
												<table class="table table-hover table-nomargin table-striped">
													<thead>
														<tr>
															<th class='hidden-480'><b>ID_Mod</b></th>
															<th><b>Modulo</b></th>
															<th><b>Modifica</b></th>
															<th class='hidden-480'><b>Cancella</b></th>
														</tr>
													</thead>
													<tbody>

														<%
		do while not rsTabella.eof%>
														<tr>
															<td class='hidden-480'><%=rsTabella.fields("ID_Mod")%></td>
															<td><%=rsTabella.fields("Titolo")%></td>
															<td><a
																	href="modificamodulo.asp?ID_Mod=<%=rsTabella.fields("ID_Mod")%>&Classe=<%=Classe%>&Id_Classe=<%=Id_Classe%>&byUmanet=1"><img
																		src="../../img/Next.gif" width="14"
																		height="13"></a></td>
															<td class='hidden-480'><a
																	onClick="return window.confirm('Vuoi veramente cancellare il modulo?');"
																	href="cancella_modulo.asp?cancella=1&Id_Mod=<%=rsTabella.fields("ID_Mod")%>&Id_Classe=<%=Id_Classe%>"><i
																		class="icon-trash"></i></a></td>
														</tr>

														<% 'posizione=rsTabella.fields("Posizione")
			      rsTabella.movenext()
		loop %>

													</tbody>
												</table>
												<br>
												<input type="submit" value="Inserisci nuovo modulo Umanet" name="B2"
													class="btn"><br>
											</form>


										</div>
									</div>
								</div>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse6">
											<center>Diario</center>
										</a>
									</div>
									<div id="collapse6" class="accordion-body collapse">
										<div class="accordion-inner">



											<div class="box box-bordered">
												<div class="box-title">
													<h3><i class="icon-th-list"></i>
														<a
															href="../cSocial/default0.asp?scegli=3&amp;id_classe=<%=Id_Classe%>&amp;divid=<%=divid%>&amp;cartella=<%=cartella%>"><span></span>&nbsp;Vai
															a interrogazioni</a></h3>

												</div>

												<div class="box-title">
												<input type="text" value="" id="txturlFeedback" name="txturlFeedback" class="input-xxlarge" placeholder="Inserisci url della pagina feedback nel diario">
												<input type="button" value="Inserisci" class="btn" onclick="inserisciUrl();">
													<h3><i class="icon-th-list"></i>
												     <span></span>&nbsp;Inserisci url pagina Feedback&nbsp;&nbsp;</h3>

												</div>
												<div class="box-title">
													<h3><i class="icon-th-list"></i> Inserisci nuovo avviso</h3>
												</div>
												<div class="box-content nopadding">



													<form name="insavviso" class='form-horizontal form-striped'
														method="POST"
														action="inserisci_avviso.asp?Id_Classe=<%=Id_Classe%>&classe=<%=classe%>&posizione=<%=posizione%>&divid=<%=divid%>">
														<table class="table table-hover table-nomargin table-striped">
															<tr>
																<td width="10%"><b> Avviso </b></td>
																<td><input type="text" value="" name="txtAvviso"
																		class="input-xlarge">
																</td>
															</tr>
															<tr>
																<td width="10%"><b> Descrizione </b></td>
																<td><textarea class='ckeditor' class="input-block-level" rows="2" cols="40" name="txtDescrizione" id="txtDescrizione"></textarea>
																</td>
															</tr>
															<tr>
																<td>Data</td>
																<td>
																	<!-- <input type="text" name="date3" id="sel3" size="10" value="gg/mm/aaaa">
   			 <input type="reset" value=" ... " onClick="return showCalendar('sel3', '%d/%m/%Y');">
  			  <input type="reset" value=" No " onClick="document.dati.date3.value='gg/mm/aaaa';">-->

																	<input type="text" name="date3" id="datepicker1"
																		class="input-medium datepick" />


																</td>
															</tr>
															<tr>
																<td>Scadenza (come compiti)</td>
																<td>
																	<!-- <input type="text" name="date3" id="sel3" size="10" value="gg/mm/aaaa">
   			 <input type="reset" value=" ... " onClick="return showCalendar('sel3', '%d/%m/%Y');">
  			  <input type="reset" value=" No " onClick="document.dati.date3.value='gg/mm/aaaa';">-->

																	<input type="text" name="txtScadenza"
																		id="datepicker2"
																		class="input-medium datepick" />


																</td>
															</tr>
															<tr>
																<td colspan="2">
																	<div class="control-group">
																		<label for="select"
																			class="control-label">Modulo</label>
																		<div class="controls">
																			<select id="selcap"
																				onchange="caricaparagrafi()">
																				<option>Seleziona un modulo</option>

																				<% if byUmanet="" then
 querySQL="SELECT DISTINCT (Titolo) AS TitoloMod, Id_Classe,ID_Mod,Cartella,Classe,Posizione FROM MODULI_NOT_UMANET WHERE Id_Classe='"& Id_Classe &"' and Visibile=1 " &_
" ORDER BY MODULI_NOT_UMANET.Posizione;"
else
 querySQL="SELECT DISTINCT (Titolo) AS TitoloMod, Id_Classe,ID_Mod,Cartella,Classe,Posizione FROM MODULI_UMANET1 WHERE Id_Classe='"& Id_Classe &"' and Visibile=1" &_
" ORDER BY MODULI_UMANET1.Posizione;"
end if
'response.write(querySQL)
'response.write("byUmanet="&byUmanet)
Set rsTabellaMod = ConnessioneDB.Execute(QuerySQL)
			posmodulo=1
			do while not rsTabellaMod.EOF
			response.write("<option value='"&rsTabellaMod("ID_Mod")&"'>"&posmodulo&"-"&rsTabellaMod("TitoloMod")&"</option>")
			rsTabellaMod.movenext
			posmodulo=posmodulo+1

			loop


%>

																			</select>
																		</div>
																	</div>


																	<div class="control-group">
																		<label for="select"
																			class="control-label">Paragrafo</label>
																		<div class="controls">
																			<select id="selpar" disabled
																				onchange="caricasottoparagrafi()">
																				<option>Seleziona un paragrafo</option>
																			</select>

																		</div>
																	</div>

																	<div class="control-group">
																		<label for="select"
																			class="control-label">Sotto-Paragrafo</label>
																		<div class="controls">
																			<select id="selsottopar" disabled>
																				<option>Seleziona un sottoparagrafo
																				</option>
																			</select>

																		</div>
																	</div>

																	<!--<iframe src="../cMessaggi/compilapreavviso.asp" name="postmessage" id="postmessage" width="100%" height="60%" frameborder="0" SCROLLING="no" border="0" class="iframe">
      ' </iframe>-->

																</td>
															</tr>
															<tr>
																<td colspan="2"><br>&nbsp;<input type="checkbox"
																		name="cbEmail"
																		title="Selezionare per inviare un email alla classe">
																	&nbsp; Invia avviso per email
																	&nbsp;&nbsp;&nbsp;<br><br></td>
															</tr>

															<tr>
																<td align="center" colspan="2">
																	<input type="button" value="Inserisci nuovo avviso"
																		name="B1" class="btn" onClick="validate2();">
																</td>
															</tr>
														</table>
													</form>
												</div>
											</div>

											<div class="box box-bordered">
												<div class="box-title">
													<h3><i class="icon-th-list"></i> Avvisi pubblicati</h3>
												</div>
												<div class="box-content nopadding">







													<!-- #include file = "../cClasse/studente_domande_include/2_lavagna_admin_1.asp" -->
												</div>
											</div>






										</div>
									</div>
								</div>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse7">
											<center>Chat e forum </center>
										</a>
									</div>
									<div id="collapse7" class="accordion-body collapse">
										<div class="accordion-inner">

											<fieldset style="margin: 0 auto 0 auto;">
												<LEGEND><b><%  response.write("Configurazione della Chat  ") %> </b>
												</LEGEND><br>
												<ul>
													<li><a href="../ChatRoom/resetta_chat.asp"> Resetta la Chat </a>
													</li>
													<li> <a href="../ChatRoom/genera_include.asp"> Rigenera file di
															inclusione immagini</a> </li>

												</ul>
												</p>
												<% 'end if
	%>
											</fieldset>

											<br>
											<fieldset style="margin: 0 auto 0 auto;">
												<LEGEND><b><%  response.write("Configurazione del Forum  ") %> </b>
												</LEGEND><br>
												<ul>

													<li> <a href="../cSocial/genera_include.asp"> Rigenera file di
															inclusione immagini </a> </li>

												</ul>
												</p>
												<% 'end if
	%>
											</fieldset>
										</div>
									</div>
								</div>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse8">
											<center>Profili utente </center>
										</a>
									</div>
									<div id="collapse8" class="accordion-body collapse">
										<div class="accordion-inner">
											<fieldset style="margin: 0 auto 0 auto;">
												<LEGEND><b><%  response.write("Gestione profili") %> </b></LEGEND><br>
												<ul>
													<li><a target="_blank"
															href="consulta_profili_quiz.asp?id_classe=<%=id_classe%>&divid=<%=divid%>">
															Consulta profili InQuiz </a></li>
													<li><a target="_blank"
															href="consulta_profili_new.asp?id_classe=<%=id_classe%>&divid=<%=divid%>">
															Consulta profili </a></li>
													<li><a target="_blank"
															href="consulta_email.asp?id_classe=<%=id_classe%>&divid=<%=divid%>">
															Consulta email </a></li>
													<li><a href="#"
															onClick="javascript:PopUpWindow(550,200);return false;">
															+Foto classe</a></li>


												</ul>
												</p>
												<% 'end if
	%>
											</fieldset>
										</div>
									</div>
								</div>





								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse9">
											<center>Google Drive</center>
										</a>
									</div>
									<div id="collapse9" class="accordion-body collapse">
										<div class="accordion-inner">
											<form method="POST" class='form-horizontal form-striped'
												action="aggiorna_url_bacheche.asp?classe=<%=classe%>&id_classe=<%=id_classe%>">
												<FIELDSET style="margin-left:16px;">
													<LEGEND><B> Aggiorna Bacheca Studente</B></LEGEND>
													<div class="control-group">
														<label for="textarea" class="control-label">Nome
															documento</label>
														<div class="controls">
															<input type="text" name="txtPostit" id="textfield"
																placeholder="Nome che comparir&agrave; in bacheca"
																class="input-xlarge">

														</div>
													</div>

													<div class="control-group">
														<label for="textarea" class="control-label">Url </label>
														<div class="controls">
															<textarea name="txtAzione" id="textarea" rows="5"
																class="input-block-level">
                                             Incolla elenco url da incorporare (senza lasciare righe vuote)
                                            </textarea>

														</div>
													</div>



													<div class="form-actions">

														<input type="submit" class="btn">Aggiorna</button>



													</div>
												</FIELDSET>
											</form>
										</div>
									</div>


								</div>




								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse10">
											<center>Google Calendar</center>
										</a>
									</div>
									<div id="collapse10" class="accordion-body collapse">
										<div class="accordion-inner">
											<form method="POST" class='form-horizontal form-striped'
												action="aggiorna_url_calendario.asp?classe=<%=classe%>&id_classe=<%=id_classe%>">
												<FIELDSET style="margin-left:16px;">
													<LEGEND><B> Incorpora Calendario </B></LEGEND>
													<%
									   QuerySQL="SELECT Url_calendar " &_
"FROM Classi WHERE ID_Classe='" & Session("Id_Classe") & "';"
'response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)

									  %>

													<div class="control-group">
														<label for="textarea" class="control-label">Codice di
															incorporamento </label>
														<div class="controls">

															<% if (len(rsTabella(0))<5) then%>
															<textarea name="txtCalendar" id="textarea" rows="5"
																class="input-block-level">Incolla codice di incorporamento fornito da google</textarea>

															<%else%>

															<textarea name="txtCalendar" id="textarea" rows="5"
																class="input-block-level"><% response.write(Trim(rsTabella("Url_calendar")))%></textarea>

															<%end if%>



														</div>
													</div>



													<div class="form-actions">

														<input type="submit" class="btn"></button>



													</div>
												</FIELDSET>
											</form>
										</div>
									</div>


								</div>






								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse11">
											<center>Registra Utenti</center>
										</a>
									</div>
									<div id="collapse11" class="accordion-body collapse">
										<div class="accordion-inner">
											<form method="POST" class='form-horizontal form-striped'
												action="registra_utenti.asp?classe=<%=classe%>&id_classe=<%=id_classe%>">
												<FIELDSET style="margin-left:16px;">
													<LEGEND><B> Inserisci generalit&agrave; </B></LEGEND>


													<div class="control-group">
														<label for="textarea" class="control-label">Cognome - Nome
															(separati dal segno - ) </label>
														<div class="controls">

															<textarea name="MyTextArea" rows="5"
																class="input-block-level">
                                            Elenco degli utenti (uno per ogni riga)
                                            </textarea>

														</div>
													</div>

													<div class="form-actions">
														<input type="submit" class="btn">Registra</button>

													</div>
												</FIELDSET>
											</form>
										</div>
									</div>


								</div>


								 <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion1" href="#collapseExport">
												<center>Esporta classe (solo moduli visibili)</center>
											</a>
										</div>
										<div id="collapseExport" class="accordion-body collapse">
											<div class="accordion-inner">
												
                     		<div class="box box-bordered">
									<div class="box-title">
										<h3><i class="icon-th-list"></i> Parametri per esportazione  </h3>
									</div>
									<div class="box-content">
										
										 <form method="POST" class='form-vertical form-bordered' action="esporta_classe_json.asp">
										 <div class="controls">
										
											<input type="hidden" name="txtIDCLASSE" value="<%=id_classe%>"> 
											<input type="text" name="txtPKcorso" class="input-xxsmall">&nbsp;PK corso / courses <br>
											<input type="text" name="txtPKmodulo" class="input-xxsmall">&nbsp;PK first modulo / module  <br>
											<input type="text" name="txtPKparagrafo" class="input-xxsmall">&nbsp;Start PK per paragrafi / content  <br>  
											<input type="text" name="txtPKsottoparagrafo" class="input-xxsmall">&nbsp;Start PK per sottoparagrafi / subcontent   <br> 
											<input type="text" name="txtPKprefrase" class="input-xxsmall">&nbsp;Start PK per prefrasi / prephrase   <br>  
											<input type="Submit" name="btnUpload" value="Esporta"  class="btn"   title="Esporta il modulo in formato json">
										  </div>
										 </form>
									</div>
                                              
							</div>
						</div>
				  </div>




								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse12">
											<center>Democrazia Interna</center>
										</a>
									</div>
									<div id="collapse12" class="accordion-body collapse">
										<div class="accordion-inner">


											<div class="box box-bordered">
												<div class="box-title">
													<h3><i class="icon-list-ul"></i> Votazione&nbsp;</h3>
													<i class="glyphicon-circle_info" rel="tooltip"
														data-placement="right"
														title="Indicare il codice di sessione della votazione"> </i>
												</div>


												<!-- inizio include-->

												<form method="POST" name="votazioni"
													class='form-horizontal form-striped'>




													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>Inserisci il codice</b>
														</label>
														<div class="controls">
															<input type="text" name="Txt1" size="1" value=""
																class="input-mini">
														</div>
													</div>



													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>Scrutina</b>
														</label>
														<div class="controls">
															<input type="button" value="Esegui" name="B1" class="btn"
																onClick="javascript:validate3();">
														</div>
													</div>

												</form>
											</div>

											<ul>
												<li><a target="_blank" href="#"> Crea Seggio Elettorale </a></li>
												<li><a target="_blank" href="#"> Effettua Votazioni </a></li>


											</ul>

										</div>
									</div>


								</div>









								<% if session("DB") = "1" then %>

								<div class="accordion-group">
									<div class="accordion-heading">
										<a class="accordion-toggle collapsed" data-toggle="collapse"
											data-parent="#accordion2" href="#collapse13">
											<center>Gestione Applicazioni</center>
										</a>
									</div>
									<div id="collapse13" class="accordion-body collapse">
										<div class="accordion-inner">


											<div class="box box-bordered">
												<div class="box-title">
													<h3><i class="icon-list-ul"></i> Elenco Applicazioni&nbsp;</h3>
													<!--<i class="glyphicon-circle_info"  rel="tooltip"   data-placement="right" title="Indicare il codice di sessione della votazione"> </i>-->
												</div>

												<form method="POST" name="votazioni"
													class='form-horizontal form-striped'>


													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>RimorchiApp</b>
														</label>
														<div class="controls">
															<a href="../cApp/gestionerimorchiapp.asp"><input
																	type="button" class="btn" value="Gestione"></a>
															<a href="../cApp/approvarimorchiapp.asp"><input
																	type="button" class="btn" value="Approva Frasi"></a>
														</div>
													</div>


													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>Schermo Nero</b>
														</label>
														<div class="controls">
															<a href="../cApp/gestionesns.asp"><input type="button"
																	class="btn" value="Gestione"></a>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>Quiz Legalità</b>
														</label>
														<div class="controls">
															<a href="../cApp/sessionilegalita.asp"><input type="button"
																	class="btn" value="Gestione Sessioni"></a>
														</div>
													</div>

													<div class="control-group">
														<label class="control-label">

															&nbsp; <b>Quiz CNV</b>
														</label>
														<div class="controls">
															<a href="../cApp/sessioniall2.asp?id_app=2&id_test=0"><input
																	type="button" class="btn"
																	value="Gestione Sessioni"></a>
														</div>
													</div>

												</form>

											</div>


										</div>
									</div>


								</div>

								<% end if %>

							</div>
						</div>

					</div>
				</div>
			</div>
		</div>
	</div>
	</div>

	<!--

          <div class="box-content">
								<p>
									<h4>Modals</h4>
									<a href="#modal-4" role="button" class="btn" data-toggle="modal" id="allerta" >Alert</a>
								</p>
							</div>


            	<div id="modal-4" class="modal hide fade" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
		<div class="modal-header">
			<button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
			<h3 id="myModalLabel">Modal header</h3>
		</div>
		<div class="modal-body">
			<p>Messaggio</p>
		</div>
		<div class="modal-footer">
			<button class="btn btn-primary" data-dismiss="modal">Ok</button>
		</div>
	</div>


     -->





	</div>
	</div>


	</div>





	</div>
	</div>

	<!-- #include file = "../include/colora_pagina.asp" -->



	<form id="mod" action="../cClasse/modifica_classe.asp" method="post">
		<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel"
			aria-hidden="true" style="display: none;">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i
						class="icon-remove"></i></button>
				<h3 id="myModalLabel">Modifica classe</h3>
			</div>
			<div class="modal-body">
			 
			<b>Id Classe</b><br><input name="titolo_idclasse" id="titolo_idclasse"
					type="text" disabled class="input-xxlarge" style="width: 97%"><br>
				<b>Nome</b><br><input placeholder="Inserisci il nuovo nome" name="titolomodifica" id="titolomodifica"
					type="text" class="input-xxlarge" style="width: 97%"><br>

			</div>
			<div class="modal-footer">
				<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
				<button type="button" id="inviamodifica" class="btn btn-primary"
					onClick="controllamodifica()">Invia</button>
			</div>
		</div>
	</form>



	<script>

		function modifica(id, classe) {

			
			document.getElementById("titolomodifica").value = classe;
			document.getElementById("titolo_idclasse").value = id;
			
			document.getElementById("mod").action = "../cClasse/modifica_classe.asp?id=" + id;
		}

		function controllamodifica() {
			var nome = document.getElementById("titolomodifica").value.trim();
			var idclasse = document.getElementById("titolo_idclasse").value.trim();

			if (nome == "") {
				alert("Il nome della classe è obbligatorio");
			} else {
				//document.getElementById("inviamodifica").type = "submit";
				var url="modifica_classe_ajax.asp?idclasse="+idclasse+"&nome="+nome;
				var xhttp = new XMLHttpRequest();
				xhttp.onreadystatechange = function() {
					if (xhttp.readyState == 4 && xhttp.status == 200) {
						var risposta=xhttp.responseText;
						//alert("?"+risposta);
						if (risposta==1)
							document.getElementById("occhio_"+idclasse).ClassName="icon-eye-open";
						else
							document.getElementById("occhio_"+idclasse).ClassName="icon-eye-close";
						}
					};
					xhttp.open("GET", url, true);
					xhttp.send();
				}
		}



	 
	function cambia_visibilita(idclasse,visibile)
	{	//alert(idclasse + " " + visibile);
	 var url="modifica_classe_ajax.asp?idclasse="+idclasse+"&visibile="+visibile;
	 var xhttp = new XMLHttpRequest();
	 xhttp.onreadystatechange = function() {
		if (xhttp.readyState == 4 && xhttp.status == 200) {
			var risposta=xhttp.responseText;
			alert(risposta);
			document.getElementById("occhio_"+idclasse).className=risposta;
						 
			}
		};
		xhttp.open("GET", url, true);
		xhttp.send();
	}

		function caricaparagrafi() {

			var capitolo = $("#selcap").val();
			//alert(capitolo);

			if (capitolo != "Seleziona un modulo" && capitolo != null) {


				$.ajax({
					method: "POST",
					url: "carica_paragrafi.asp",
					dataType: "html",
					data: { byUmanet: "<%=byUmanet%>", modulo: capitolo }
				}) /* .ajax */
					.done(function (ans) {

						//alert(ans);
						$("#selpar").html("<option>Seleziona un paragrafo</option>" + ans);

					}) /* .done */
					.error(function (jqXHR, textStatus, errorThrown) {
						alert(jqXHR + "\n" + textStatus + ": " + errorThrown);
					});

				document.getElementById("selpar").disabled = false;
				document.getElementById("selsottopar").disabled = true;
				$("#selsottopar").html("<option>Seleziona un sottoparagrafo</option>");


			} else {
				document.getElementById("selpar").disabled = true;
				document.getElementById("selsottopar").disabled = true;
				$("#selpar").html("<option>Seleziona un paragrafo</option>");
				$("#selsottopar").html("<option>Seleziona un sottoparagrafo</option>");
			}

		}


		function caricasottoparagrafi() {

			var capitolo = $("#selcap").val();
			var paragrafo = $("#selpar").val();
			//alert(capitolo);

			if (paragrafo != "Seleziona un paragrafo" && paragrafo != null) {


				$.ajax({
					method: "POST",
					url: "carica_sottoparagrafi.asp",
					dataType: "html",
					data: { byUmanet: "<%=byUmanet%>", modulo: capitolo, paragrafo: paragrafo }
				}) /* .ajax */
					.done(function (ans) {

						if (ans != "") {
							$("#selsottopar").html("<option>Seleziona un sottoparagrafo</option>" + ans);
							document.getElementById("selsottopar").disabled = false;

						} else {
							$("#selsottopar").html("<option>Nessun sottoparagrafo disponibile</option>");
							document.getElementById("selsottopar").disabled = true;
						}

					}) /* .done */
					.error(function (jqXHR, textStatus, errorThrown) {
						alert(jqXHR + "\n" + textStatus + ": " + errorThrown);
					});

			} else {
				document.getElementById("selsottopar").disabled = true;
				$("#selsottopar").html("<option>Seleziona un sottoparagrafo</option>");
			}

		}



function inserisciUrl() {

var url = $("#txturlFeedback").val();
//alert(capitolo);

if (url != null) {
	$.ajax({
		method: "POST",
		url: "url_feedback.asp",
		dataType: "html",
		data: { url: url,idclasse: '<%=Id_Classe%>' }
	}) /* .ajax */
		.done(function (ans) {
			alert(ans);
		}) /* .done */
		.error(function (jqXHR, textStatus, errorThrown) {
			alert(jqXHR + "\n" + textStatus + ": " + errorThrown);
		});

	 
	 

} else {
	alert("Non hai inserito l'url");
}

}



	</script>




</body>

</html>