<!doctype html>
<html>
<head>
<link rel="shortcut icon" href="../favicon.ico" />

<script src="../js/google.js"></script><!--<meta charset="utf-8">-->
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <meta charset="UTF-8">
	<title>Quaderno dello Studente</title>

	<link rel="shortcut icon" href="../favicon.ico" />

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- Easy pie  -->
	<link rel="stylesheet" href="../../css/plugins/easy-pie-chart/jquery.easy-pie-chart.css">
	<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">
	<!-- Theme CSS -->
	<link rel="stylesheet" href="../../css/style.css">
	<!-- Color CSS -->
	<link rel="stylesheet" href="../../css/themes.css">

     <!-- Notify -->

	<link rel="stylesheet" href="../../css/plugins/gritter/jquery.gritter.css">

     <link href="../../../guida/css/pageguide.css" rel="stylesheet">
     <!-- Le styles -->
   <!-- <link href="../../../guida/docs/lib/bootstrap/css/bootstrap.css" rel="stylesheet">
    <link href="../../../guida/docs/lib/bootstrap/css/bootstrap-responsive.css" rel="stylesheet">

-->

	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>

	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>
	<!-- jQuery UI -->

     <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
     <!-- Notify -->
	<script src="../../js/plugins/gritter/jquery.gritter.min.js"></script>

	<!-- Validation -->
	<script src="../../js/plugins/validation/jquery.validate.min.js"></script>
	<script src="../../js/plugins/validation/additional-methods.min.js"></script>
	<!-- Sparkline -->
	<script src="../../js/plugins/sparklines/jquery.sparklines.min.js"></script>
	<!-- Easy pie -->
	<script src="../../js/plugins/easy-pie-chart/jquery.easy-pie-chart.min.js"></script>
	<!-- Flot -->
	<script src="../../js/plugins/flot/jquery.flot.min.js"></script>
	<script src="../../js/plugins/flot/jquery.flot.resize.min.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>
    <!--
    <script src="../../js/plugins/validation/jquery.validate.min.js"></script>
	<script src="../../js/plugins/validation/additional-methods.min.js"></script>

    -->
     <script language="javascript" type="text/javascript" >

 function validate2() {

	 //continua a dare errore TypeError: document.frm0.txtNewCodiceAllievo is undefined, il form è definito nel file di inclusione 2_modifica_login_1.asp, rinuncio faccio il controllo lato server prima di inserire in db

 alert(document.frm0.txtNewCodiceAllievo.value);
 if (document.frm0.txtNewCodiceAllievo.value=="")
	{
	   alert("Non hai inserito lo username !");

	}
else
 if (frm0.txtNewPwd.value=="")
	{
	   alert("Non hai inserito la nuova password");
	}
 else
  if (frm0.txtNewPwd1.value=="")
	{
	   alert("Non hai inserito la conferma password");

	}else
	 if (frm0.txtNewPwd1.value != frm0.txtNewPwd.value)
	{
	   alert("Le password non coincidono");

	}
	else

	{
	    document.frm0.action = "modifica_pwd_new.asp?stato=<%=stato%>&cla=<%=cla%>&id_classe=<%=id_classe%>&divid=<%=divid%>" ;
		document.frm0.submit();


    }

}


    </script>
	<script>
	$(window).ready(function () {
	   // utilizza la guida
	  //  $('#msg').click();

	  // event.stopPropagation();

	});

</script>
	<!--[if lte IE 9]>
		<script src="../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

	<!-- Favicon
	<link rel="shortcut icon" href="img/favicon.ico" />-->
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />

</head>

<body class='theme-<%=Session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">
     <%

function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")

  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")

ReplaceCar = sAns

end function

	if Session("CodiceAllievo")="" or Session("Id_Classe")="" then %>
				<script language="javascript" type="text/javascript">
				    window.alert("Sessione  scaduta, effettua nuovamente il Login!");
                    location.href="../../home.asp";
				</script>
				<%
				response.Redirect "../../home.asp"

				 %>

<% end if%>

		<!-- #include file = "studente_domande_include/4_quaderno.asp" -->

        <!-- #include file = "../var_globali.inc" -->

 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

		<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->

                 <!-- #include file = "../cClasse/studente_domande_include/1_periodi_date.asp" -->

		<!-- #include file = "../include/navigation.asp" -->

        <!-- #include file = "../extra/test_server.asp" -->

		<!-- #include file = "../include/formattaDataCla.inc" -->

        <%


		  QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato=rsTabellaCI("CIAbilitato")
	'ScalaValutaz=rsTabellaCI("ScalaValutaz")
	rsTabellaCI.close
    Dim esecuzione
    set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito



	' PRELEVO IN ANTICIPO IL CONGOME NOME NEL CASO LA QUERY 2 NON TROVI NULLA IN QUEL PERIODO E QUINDI RESTITUISCA NULL
		  cod=Request.QueryString("cod")
		QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)
CodiceAllievo=cod
cognome = rsTabella("Cognome")
nome = rsTabella("Nome")
Probabilita=rsTabella("Probabilita")

'prelevo le classi degli anni scorsi Id_Classe che andrà in Or per il caricamento dei compiti svolti

QuerySQL="SELECT * " &_
" FROM stud_as_classe " &_
" WHERE stud_as_classe.Id_Stud='" & cod & "' and Id_As=1" ' and Id_As=1 andrà reso parametrico altrimenti carica solo un anno prima

Set rsTabellaPassato = ConnessioneDB.Execute(QuerySQL)
if not rsTabellaPassato.eof then
 Id_ClassePassato=rsTabellaPassato("Id_Classe")
 end if

set rsTabellaPassato = nothing
		%>


	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->
			<div id="main">
				<div class="container-fluid">
					<div class="page-header">
					<!--<div class="breadcrumbs">
                       <ul>
                            <li>
                                <a href="#">Home</a>
                                <i class="icon-angle-right"></i>
                            </li>
                            <li>
                                <a href="#">Studente</a>
                                <i class="icon-angle-right"></i>
                            </li>

                        </ul>
						<div class="close-bread">
							<a href="#">
								<i class="icon-remove"></i>
							</a>
						</div>
					</div>-->

					<!-- NOVITA' UMANET!!!! -->

					<% datafinenov = "30/04/2017" ' data di fine novità... il seguente if controlla se la data di oggi è minore della data indicata in questa var
					%>

					<% 'Session("DataCla") & " " & Session("DataCla2")
					%>

				<% if DateDiff("d", now(), datafinenov) > 0 then %>

					<div class="alert">
			<button type="button" class="close" data-dismiss="alert"><i class="icon-remove"></i></button>
			<strong>Novità su Umanet!</strong><br>- Aggiunta una sezione <b>Eccezioni</b> sul tuo Quaderno per consultare quali frasi sono state riaperte, oltre la data di scadenza, dal docente.
			<br>- Implementata la possibilità di modificare le frasi e le immagini già inserite direttamente dal tuo Quaderno
		</div>

		<div class="alert alert-success">
			<button type="button" class="close" data-dismiss="alert"><i class="icon-remove"></i></button>
			<strong>Centro Messaggi!</strong> E' entrato in funzione il Centro Messaggi per la gestione dei
			messaggi privati tra due utenti e la lettura e cancellazione delle notifiche.
			<a style="text-decoration:none" href="../cMessaggi/centro_messaggi.asp">Clicca qui per visitare la pagina</a>.
		</div>

		<% end if %>
		<!-- FINE NOVITA' UMANET!!!! -->

		<!-- #include file = "../include/navigation_small.asp" -->
		<!--
				<div class="box">
					<div class="visible-phone">
						<div class="navbar">
							<div class="navbar-inner">
								<ul class="nav">
								<li><a href="#">Home</a></li>
								<li><a href="#">Link</a></li>
								<li><a href="#">Link</a></li>
								</ul>
							</div>
						</div>
					</div>
				</div>-->

                     <div class="box">

                     <div class="box-title">

				      <h3> <a name="#top"><i class="icon-folder-open"></i> </a><b>Quaderno di <%=ReplaceCar(cognome)%>  &nbsp;<%=ReplaceCar(left(nome,1)&".") %></h3>&nbsp;</b>
                     <% if session("admin")=true then%>

						<select id="studente" onchange="aggiorna_studente();">
							 <option>Seleziona studente</option>
					 <%  QuerySQL="Select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='" & Session("Id_Classe")&"' order by Cognome, Nome;"
						 Set rsTabellaStud = ConnessioneDB.Execute(QuerySQL)
						 do while not rsTabellaStud.eof%>
						 <option value="<%=rsTabellaStud("CodiceAllievo")%>"><%=rsTabellaStud("Cognome")%>&nbsp;<%=left(rsTabellaStud("Nome"),1)&"."%></option>
						 <%rsTabellaStud.movenext
						 loop

					 %>

					 </select>&nbsp; Aggiorna
					 <%end if%>
			          </div><br>
                      <!-- #include file = "studente_domande_include/1_periodi.asp" -->

<input id="periodo" type="button" class="btn"  style="width:60px;height:25px;" value="Invia" name="B1" onClick="aggiornaStud()">
 <input type="checkbox"  name="cbPS" value="1" checked="true" title="Deseleziona per escludere i Punti Social dalla classifica">  <b>
	Includi PS
   </b>
</form>

            <% 'response.write Session("DataCla") & "  " & Session("DataCla2") & "<br>" & Session("DataClaq") & "  " & Session("DataClaq2") & "<br>"
			%>

					</div>

					</div>

					<hr>

					<div class="row-fluid">
						<div class="span12">

                            <div class="box-title">
				        <h4><a name="#"> <i class="icon-reorder"></i></a> <b> Attivit&agrave;</b> </h4>
			          </div>

    <% 'carico i dati per il rewind

	QuerySQL="SELECT  dbo.anni_scolastici.Nome, dbo.stud_as_classe.Id_Stud,dbo.stud_as_classe.Id_Classe " &_
	" FROM  dbo.anni_scolastici INNER JOIN dbo.stud_as_classe ON dbo.anni_scolastici.ID_AS = dbo.stud_as_classe.Id_As" &_
	" where dbo.stud_as_classe.Id_Stud='"&CodiceAllievo&"';"
   Set rsTabella = ConnessioneDB.Execute(QuerySQL)

   QuerySQL2="SELECT  *" &_
	" FROM  dbo.Classi" &_
	" where dbo.Classi.ID_Classe<>'"&id_classe&"';"
  Set rsTabellaPromuoviti = ConnessioneDB.Execute(QuerySQL2)


   %>

       <div class="bs-docs-example">
            <ul id="myTab2" class="nav nav-tabs">
                                  <li id="bacheca" class="active"><a href="#profileMsg" data-toggle="tab" title="Messaggi dalla bacheca">Bacheca</a></li>


                                   <li id="adiario">
                                   <a title="La mia bacheca personale"  href="../cSocial/default0.asp?scegli=0&bacheca=<%=cod%>&nome=<%=nome%>&cognome=<%=cognome%>&id_classe=<%=id_classe%>&divid=<%=Session("divid")%>&cartella=<%=cartella%>" >
                                  Diario</a></li>


                                     <li id="apost" class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown" title="I miei commenti nelle discussioni">iPost <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownPostLavagna" data-toggle="tab" title="I miei commenti">Bacheca</a></li>
                                      <li><a href="#dropdownPostForum" data-toggle="tab" title="I miei commenti">Forum</a></li>
                                      <li><a href="#dropdownPostDiario" data-toggle="tab" title="I miei commenti">Diario</a></li>
																			<li><a href="#dropdownPostInterrogazioni" data-toggle="tab" title="I miei commenti">Interrogazioni</a></li>

                                      <li><a href="../cMessaggi/centro_messaggi.asp" title="Chat">Chat</a></li>

                                    </ul>
                                    </li>

                                     <li id="report" class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">Report <b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownQuiz" data-toggle="tab">Quiz</a></li>
									  <li><a href="#dropdownVerifiche" data-toggle="tab">Verifiche</a></li>
                                      <li><a href="#dropdownCrediti" data-toggle="tab">Crediti</a></li>
                                      <li><a href="#dropdownCronologia" data-toggle="tab">Classifiche</a></li>
                                      <li> <a href="../cMessaggi/centro_messaggi.asp" class='more-messages'>Vai al centro messaggi <i class="icon-arrow-right"></i></a>  </li>


                                    </ul>
                                    </li>

									 <% if 1=2 then   ' nascondo tutto perchè il 01/07/2018 ho deciso di tenere una classe unica per tutti i tre anni senza cambiare con promuoviti o time machine, si fa tutto con i periodi di valutazione
									 ' e vengono nascosti i moduli degli anni scorsi da Admin , così nel quaderno si hanno tutti i compiti dei tra anni
									 %>
											<%if session("Admin")=true or 1=1 then%>
											<li id="rewind" class="dropdown">
											<a href="#" class="dropdown-toggle" data-toggle="dropdown">Time Machine <b class="caret"></b></a>
											<ul class="dropdown-menu">

											 <%do while not rsTabella.EOF %>
											<li> <a href="promuoviti.asp?indietro=1&id_classe=<%=rsTabella("Id_Classe")%>&CodiceAllievo=<%=CodiceAllievo%>" class='more-messages'>Vai al <i class="icon-arrow-right"><%=rsTabella("Nome")%></i></a>  </li>
												<%rsTabella.movenext
												loop%>

											</ul>
											</li>

											  <li id="forward" class="dropdown">
											<a href="#" class="dropdown-toggle" data-toggle="dropdown">Forward <b class="caret"></b></a>
											<ul class="dropdown-menu">
											<li> <%'response.write(QuerySQL2)
											%></li>
											<%do while not rsTabellaPromuoviti.EOF %>
											<li> <a href="promuoviti.asp?avanti=1&id_classe=<%=rsTabellaPromuoviti("ID_Classe")%>&CodiceAllievo=<%=CodiceAllievo%>&Classe=<%=rsTabellaPromuoviti("Classe")%>" class='more-messages'>Promuoviti in <i class="icon-arrow-right"><%=rsTabellaPromuoviti("Classe")%></i></a>  </li>
												<%rsTabellaPromuoviti.movenext
												loop%>

											</ul>
											</li>
											 <% end if%>
									<% end if%>
                                  <% if session("admin") = false then 'se la sessione non è amministratore mostro solo il report delle eccezioni
								  %>

										<li id="eccezioni"><a href="#dropdownEccezioni" data-toggle="tab" title="Report Eccezioni">Eccezioni</a></li>


								 <% end if

										if session("admin") = true then
								  %>


                                        <li class="dropdown">
                                    <a href="#" class="dropdown-toggle" data-toggle="dropdown">MYnd<b class="caret"></b></a>
                                    <ul class="dropdown-menu">
                                      <li><a href="#dropdownProfilo" data-toggle="tab">Profilo</a></li>
									  <li><a href="../cUtenti/login256.asp?identita=1&CodiceAllievo=<%=CodiceAllievo%>&Cartella=<%=Cartella%>&id_classe=<%=id_classe%>" >Assumi identità</a></li>
                                      <li><a href="#dropdownLogin" data-toggle="tab">Login</a></li>
                                      <li><a href="#dropdownContatti" data-toggle="tab">Contatti</a></li>
                                      <li><a href="#dropdownEccezioni" data-toggle="tab">Eccezioni</a></li>
                                       <li><a href="../cFrasi/2inserisci_valutazioni_recuperoaccount.asp?CodiceAllievo=<%=CodiceAllievo%>&Cartella=<%=Cartella%>&id_classe=<%=id_classe%>" >Recupero frasi</a></li>




                                       <% if cod <> "informistica" then %>
									   <li> <A onClick="return window.confirm('Vuoi veramente sospendere questo account ?');" HREF="sospendi_studente.asp?CodiceAllievo=<%=cod%>"><i class="icon-eye-close" ></i> Sospendi</a></li>
                                       <li> <A onClick="return window.confirm('Vuoi veramente cancellare questo account ?');" HREF="cancella_studente.asp?CodiceAllievo=<%=cod%>"><i class=" icon-trash" ></i> Rimuovi</a></li>
										<% end if %>
                                                <!--
                                      <li><a href="#dropdownVisualizzazioni" data-toggle="tab">Visualizzzioni</a></li>
                                      -->
                                    </ul>
                                    </li>




                                  <!--
                                       <li ><a href="#profileProfilo" data-toggle="tab">Profilo</a></li>   -->

                                       <%end if%>

                            </ul>



                            <div id="myTabContent2" class="tab-content">

							 <% if session("admin") = false then %>

							 <div class="tab-pane fade" id="dropdownEccezioni">

                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Scadenze
								</h4>
							</div>

								  <!-- #include file = "studente_domande_include/2_modifica_eccezioni_stud.asp" -->

							</div>
                              </div>

							 <%end if %>


                             <% if session("admin") = true then %>


							  <div class="tab-pane fade" id="dropdownLogin">

                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Login
								</h4>
							</div>

                             <%  cod=Request.QueryString("cod")
		QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)

CIAbilitato = rsTabella("CIAbilitato")
cognome = rsTabella("Cognome")
nome = rsTabella("Nome")  %>

								<!-- #include file = "studente_domande_include/2_modifica_login_1.asp" -->





							</div>

                              </div>



                                <div class="tab-pane fade" id="dropdownContatti">

                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Contatti
								</h4>
							</div>

								  <!-- #include file = "studente_domande_include/2_modifica_contatti_1.asp" -->

							</div>
                              </div>


                              <div class="tab-pane fade" id="dropdownProfilo">

                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Profilo
								</h4>
							</div>

								  <!-- #include file = "studente_domande_include/2_modifica_profilo_1.asp" -->

							</div>
                              </div>



                                  <div class="tab-pane fade" id="dropdownEccezioni">

                              <div class="box-content nopadding">
                              <div class="box-title">
								<h4>
									<i class="icon-user"></i>
									Modifica Scadenze
								</h4>
							</div>

								  <!-- #include file = "studente_domande_include/2_modifica_eccezioni.asp" -->

							</div>
                              </div>






							 <% end if%>


                              <div class="tab-pane fade" id="profileProfilo">



     			 <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Modifica Login
								</h3>
							</div>
							<div class="box-content nopadding">
<%QuerySQL="SELECT * " &_
" FROM Allievi " &_
" WHERE Allievi.CodiceAllievo='" & cod & "'"

Set rsTabella = ConnessioneDB.Execute(QuerySQL)
cognome = rsTabella("Cognome")
nome = rsTabella("Nome") %>


                           <!-- #include file = "studente_domande_include/2_modifica_login_1.asp" -->



							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                              </div>



                              <div class="tab-pane fade in active" id="profileMsg">

								<!-- #include file = "studente_domande_include/2_messaggi_1.asp" -->

<%QuerySQL="SELECT * FROM Classi WHERE Id_Classe='"&id_classe&"'"





					Set rsTabella = ConnessioneDB.Execute(QuerySQL)%>



                                <div class="box box-color box-bordered">
								<div class="box-title">
									<h3>
										<i class="icon-reorder"></i>

                                         Messaggi alla Classe<a  href="../cSocial/default0.asp?scegli=2&id_classe=<%=rsTabella("Id_Classe")%>&divid=<%=divid%>&cartella=<%=rsTabella("cartella")%>"> <i style="color:#FFF" title="Vai alla lavagna" class="icon-circle-arrow-right"></i></a>
									</h3>
								</div>
									<ul class="timeline">
              <% If (rsTabellaAvvisi2.BOF=True And rsTabellaAvvisi2.EOF=True) and (rsTabellaDiario.BOF=True And rsTabellaDiario.EOF=True) and (rsTabellaForum.BOF=True And rsTabellaForum.EOF=True) then %>
              <table class="table table-hover table-nomargin">
									<thead>
										<tr>
											<th>Non ci sono messaggi</th>
										</tr>
									</thead>
									</table>

             <%else%>
           <%

		    k=0
		     do while not rsTabellaDiario.EOF and k<3
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon red">
											<i class="glyphicon-log_book" title="Messaggi dal Diario"></i>
								</div>
								<div class="date"><%=left(rsTabellaDiario("DatePosted"),2) & ". " &monthname(mid(rsTabellaDiario("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                        <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=2&ID=" & rsTabellaDiario("ID") & "&RCount=" & rsTabellaDiario("ReplyCount")& "&TParent=" & rsTabellaDiario("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe  & "&categoria="&rtrim(rsTabellaDiario("Descrizione"))&"&id_categoria="&rsTabellaDiario("ID_Categoria")& "&bacheca="&rsTabellaDiario("Bacheca")& "&cognome="&rsTabellaDiario("Cognome")& "&nome="&rsTabellaDiario("Nome")& "'>"  & replaceCar(rsTabellaDiario("Topic")) & "</A>"%>

                                      </div>
								</div>
							</div>
					<div class="line"></div>
				</li>
				<%rsTabellaDiario.movenext
				loop
		   ' bacheca, per la classe docenti doc non la visualizzo
		   if strcomp(classe,"DOC")<>0 then
		     k=0
		     do while not rsTabellaAvvisi2.EOF and k<3
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon blue">
											<i class="icon-desktop" title="Messaggi dalla bacheca <%=classe%>"></i>
								</div>
								<div class="date"><%=left(rsTabellaAvvisi2("DatePosted"),2) & ". " &monthname(mid(rsTabellaAvvisi2("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                         <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=1&ID=" & rsTabellaAvvisi2("ID") & "&RCount=" & rsTabellaAvvisi2("ReplyCount")& "&TParent=" & rsTabellaAvvisi2("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe& "&categoria="&rtrim(rsTabellaAvvisi2("Descrizione"))&"&id_categoria="&rsTabellaAvvisi2("ID_Categoria")&  "&bacheca="&rsTabellaAvvisi2("Bacheca")& "&cognome="&rtrim(rsTabellaAvvisi2("Cognome"))& "&nome="&rtrim(rsTabellaAvvisi2("Nome"))& "&visibile="&rsTabellaAvvisi2("Visibile")& "&privato="&rsTabellaAvvisi2("Privato")& "&zip="&rsTabellaAvvisi2("Zip")&"'>"  & replaceCar(rsTabellaAvvisi2("Topic")) & "</A>"%>

                                      </div>
								</div>
							</div>
					<div class="line"></div>
				</li>
				<%rsTabellaAvvisi2.movenext
				loop
			end if

			 k=0
		     do while not rsTabellaForum.EOF and k<3
               k=k+1%>
                   <li>
						<div class="timeline-content">
							<div class="left">
								<div class="icon <%=Session("stile")%>">
											<i class="icon-comments" title="Messaggi dal forum"></i>
								</div>
								<div class="date"><%=left(rsTabellaForum("DatePosted"),2) & ". " &monthname(mid(rsTabellaForum("DatePosted"),4,2),true)%></div>
								</div>
								<div class="activity">
									<div class="user">
                                        <% response.write "<A HREF='../cSocial/ShowMessage.asp?scegli=0&ID=" & rsTabellaForum("ID") & "&RCount=" & rsTabellaForum("ReplyCount")& "&TParent=" & rsTabellaForum("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe & "&categoria="&rtrim(rsTabellaForum("Descrizione"))&"&id_categoria="&rsTabellaForum("ID_Categoria")& "&bacheca="&rsTabellaForum("Bacheca")& "&cognome="&rsTabellaForum("Cognome")& "&nome="&rsTabellaForum("Nome")& "'>"  & replaceCar(rsTabellaForum("Topic")) & "</A>"%>

                                      </div>
								</div>
							</div>
					<div class="line"></div>
				</li>
				<%rsTabellaForum.movenext
				loop

			 end if
				%>

									</ul>
								</div>










     <%
	 ' se  sono nel mio quaderno non visualizzo la casella per invio messaggio personale
	 if strcomp(cod,Session("CodiceAllievo"))<>0 then %>
								<br>
                               <div class="box box-color box-bordered">
								<div class="box-title">
									<h3>
										<i class="icon-reorder"></i>
										Invia messaggio personale
                                        <a href="../cMessaggi/centro_messaggi.asp" class='more-messages'>  <i class='icon-circle-arrow-right' style="color:#FFF"></i></a>
									</h3>
								</div>


   <div class="accordion" id="accordionMsg2">
									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle"  data-toggle="collapse" data-parent="#accordionMsg2" href="#collapseTwo">
												(+) messaggio
											</a>
										</div>
										<div id="collapseTwo" class="accordion-body collapse">
											<div class="accordion-inner">

                                                  <form  class='form-horizontal' name="frmDocument" action="../cMessaggi/inserisci_messaggio_personale.asp?CodiceAllievo=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq%>&cbEmail=1" METHOD = "POST">

    <br>Messaggio: <br>

    <textarea class="input-block-level" name="txtMessaggio" ></textarea>
    <br> <br>
    <!--<p> <input type="checkbox"  name="cbEmail" title="Selezionare per inviare un email allo studente">   Notifica per email &nbsp;&nbsp;&nbsp;<br>-->
      <p> <input class="btn" type="submit" value="Inserisci"><br>
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>


											</div>
										</div>
									</div>

								</div>

								</div>
  <%else %>
  <br><hr>
   <%end if%>
                              </div>




                             <div class="tab-pane fade" id="dropdownQuiz">

                                 <!-- #include file = "studente_domande_include/2_quiz_1.asp" -->

                            </div>
                          </div>
                      </div>


					   <div class="tab-pane fade" id="dropdownVerifiche">

                                 <!-- #include file = "studente_domande_include/2_verifiche_1.asp" -->

                            </div>
                          </div>
                      </div>

                       <div class="tab-pane fade" id="dropdownPostLavagna">

                                <!-- #include file = "studente_domande_include/2_lavagna_1.asp" -->

                            </div>
                          </div>
                      </div>


                       <div class="tab-pane fade" id="dropdownPostForum">

                                <!-- #include file = "studente_domande_include/2_forum_1.asp" -->

                            </div>
                          </div>
                      </div>


                       <div class="tab-pane fade" id="dropdownPostDiario">

                                <!-- #include file = "studente_domande_include/2_diario_1.asp" -->

                            </div>
                          </div>
                      </div>

											<div class="tab-pane fade" id="dropdownPostInterrogazioni">

															 <!-- #include file = "studente_domande_include/2_interrogazioni_1.asp" -->

													 </div>
												 </div>
										 </div>


                        <div class="tab-pane fade" id="dropdownCrediti">

                                 <!-- #include file = "studente_domande_include/2_crediti_1.asp" -->

                            </div>
                          </div>
                      </div>




                      <%
					  cod=Request.QueryString("cod")

					  %>
                       <div class="tab-pane fade" id="dropdownCronologia">

                                 <!-- '#include file = "studente_domande_include/2_cronologia_1.asp" -->

                            </div>
                          </div>
                      </div>





                   </div>

  <hr>
   <div class="box-title">
				        <h3> <a name="#"><i class="icon-reorder"></i></a>  Compiti<small title="Punti totalizzati"> (Pt.)</small></h3>
			          </div>
 <div class="row-fluid">


</div>

 <div class="bs-docs-example">


	<div id="compitispec">


		<button class="btn btn-primary" style="width:100%; border-radius:5px; line-height:40px" onclick="caricacompiti()"><b><h4>Clicca qui per caricare i tuoi compiti</h4></b></button>

	</div>


        <p> <span class="invisible">
	   <a id="msg" href="#modal-4" role="button" data-notify-time="3000" class="btn notify" data-notify-title="Utilizza la Guida!" data-notify-message="INFORMAZIONI ALLA TUA DESTRA ">

	   </a></span>






		 <!-- #include file = "../include/colora_pagina.asp" -->






	</body>
    <br><br><br><br><hr>
      <!-- #include file = "../include/footer.asp" -->

        <!-- #include file = "../cGuide/g_quaderno.asp" -->



 <script language="javascript" type="text/javascript">
function cancella_avviso() {

	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();
	 }
}


 function aggiornaStud() {
	 // alert (DataClaq);
	 var DataClaq=document.dati.txtData.value;
	 var DataClaq2=document.dati.txtData2.value;
	// alert (DataClaq);
	 // alert (DataClaq2);
		with (document.dati) {

		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=Session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daForm=1";
		 else
		   document.dati.action = "?divid=<%=Session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=" +DataClaq+ "&DataClaq2="+ DataClaq2 +"&daForm=1";

	    }
		document.dati.submit();
}


</script>


<script language="javascript" type="text/javascript">
function stampa() {
    document.dati.action = "../cFrasi/7_stampa_schede_frasi_elenco.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Classe%>&Query=<%=QueryTuttoCap%>";
		//document.dati.action = "../../home.asp"
		document.dati.submit();
}
 </script>

<script type="text/javascript">

function caricacompiti(){

	var query = window.location.search.substring(1);
	//alert(query);
	var url = "https://www.umanetexpo.net/expo2015Server/UECDL/script/cClasse/compiti_old.asp?"+query;
	//alert(url);

	var testo;
	var stato1, stato2;

	//$("#compitispec").html("Attendere prego...");
	$("#compitispec").html("<img src='taoloader.gif'> Caricamento...");

	//eseguo chiamata http
					var xhttp = new XMLHttpRequest();
					xhttp.onreadystatechange = function() {

						stato1=xhttp.readyState;
						stato2=xhttp.status;

						if(stato1==4 && stato2==200){

						testo = xhttp.responseText;

						$("#compitispec").empty().append(testo);

						}


					};

					xhttp.open("GET", url, true);
					xhttp.send();



}


$(window).load(function () {

	   $('#<%=box_apri%>').click();
	   $('#<%=box_apri1%>').click();
	    $('#<%=box_apri2%>').click();
		$('#<%=box_apri3%>').click();
	    $('#<%=box_apri4%>').click();
	   $("body").addClass("theme-"+"<%=stile%>").attr("data-theme","theme-"+"<%=stile%>");



	  // event.stopPropagation();

	});


/*$(".red").click(function(event){

   // alert("Hai cliccato sull'Elemento");
	document.location = "script/aggiorna_stile.asp?stile=red"
});
*/

</script>


  <script>
  function aggiornaimpostazioni(){

	var CI = document.getElementById("CIvero").checked;
	var prob = document.getElementById("probabilita").value;
	if(CI){CI=1;}else{CI=0;}


	window.location.href = "aggiornaimpostazioni.asp?CI="+CI+"&cod=<%=CodiceAllievo%>&probabilita="+prob;

}


 function aggiorna_studente(){

	var stud = document.getElementById("studente").value;

	window.location.href = "quaderno.asp?daStud=1&DataClaq=<%=DataCla%>&DataClaq2=<%=DataCla2%>&id_classe=<%=id_classe%>&classe=<%=classe%>&cod="+stud;


}

</script>

<% if Session("Cambio") <> "" then
cambio = Session("Cambio")
else
cambio = 0
end if
%>

<script>
	$( document ).ready(function() {
		var cambio = <%=cambio%>

	var t = setTimeout(function(){

	if(cambio==1){
		alert("Sei ora connesso con l'utente <%=CodiceAllievo%>");
	}

	clearTimeout(t);

	},200);

	});
</script>

<% Session("cambio") = 0
%>

  <script type="text/javascript" src="../js/refresh_session.js"></script>
	</html>
