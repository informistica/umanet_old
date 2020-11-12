<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Convalida test</title>

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
<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>
<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">

	<div id="navigation">

        <%


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
						<h1> <i class="icon-comments"></i> Convalida test </h1>

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
							<a href="#">Quiz</a>
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
				        <h3> <i class="icon-reorder"></i>  Risultati </h3>
			          </div>
				      <div class="box-content">






				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">



		    <div class="box-content">

                         <%
 ' set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito
  DataTest=formatDateTime(Request.QueryString("DataTest"),2)
  gira_data=Day(DataTest)&"/"&Month(DataTest)&"/"&Year(DataTest)

    if day(DataTest) < 10 then
    giorno="0" & day(DataTest)
	else
	giorno=day(DataTest)
    end if

	if len(year(DataTest) ) = 2 then
	anno="20"& year(DataTest)
	elseif len(year(DataTest) ) =  3 then
	anno="2"& year(DataTest)
	else
	anno=year(DataTest)
	end if
    if month(DataTest) < 10 then
    mese="0" & month(DataTest)
	else
	mese=month(DataTest)
    end if

	 pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO"))
	  if (left(pathEnd1,10)<>"D:\inetpub") then
		 locale=1
	  else
		 locale=0
	  end if
	 ' response.write(left(pathEnd1,10))
	'response.write("locale="&locale)
                if locale=1 then

				%>
					  <%DataAvviso = giorno & "/" & mese& "/" & anno  %>
					 <% else
					' response.write("online")
					 %>
						 <% DataAvviso = mese & "/" & giorno& "/" & anno  %>
					 <% end if %>





 <% CodiceTest=Request.QueryString("CodiceTest")
  TitoloTest=Request.QueryString("TitoloTest")
  classe=Request.QueryString("classe")
  id_classe=Request.QueryString("id_classe")
  SessioneQuiz=Request.QueryString("SessioneQuiz")
  tipoTest=Request.QueryString("tipoTest")

  if tipoTest=0 then
     tipoDesc="(V/F)"
  else if tipoTest=1 then
          tipoDesc="(Singola)"
	    else
		   tipoDesc="(Multipla)"
		end if
  end if



 QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati1.Risultato, Risultati1.Data,Risultati1.Ora, Risultati1.CodiceTest,Risultati1.Risultato*8/100 as [PUNTI],Risultati1.ID_R " &_
" FROM Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo " &_
" WHERE Risultati1.Tipo="&  tipoTest &  " and Risultati1.CodiceTest='"&  CodiceTest & "' AND Risultati1.Sessione="&SessioneQuiz &_
" ORDER BY Allievi.Cognome Asc, Risultati1.Ora; "

 response.Write(QuerySQL)
 ' end if
 ' response.Write(QuerySQL)
%>

  <form method="POST" class="form-horizontal" name="Aggiorna" action="aggiorna_punteggio_pulisci_test2.asp?SessioneQuiz=<%=SessioneQuiz%>&tipoTest=<%=tipoTest%>&classe=<%=classe%>&xQuiz=1&id_classe=<%=id_classe%>&DataTest=<%=DataTest%>&TitoloTest=<%=TitoloTest%>">
<LEGEND class="sottotitoloquaderno2"><B> <%=TitoloTest%> - <%=tipoDesc%> - <%=DataAvviso%> </B></LEGEND>
<%
Set rsTabella = ConnessioneDB.Execute(QuerySQL) %>
<table class="table-condensed table-bordered">
<tr><th>Cognome</th><th>Nome</th><th width="10%">Punteggio</th><th width="10%">P</th><th>Data</th><th>Ora</th><th><a onClick="elimina_test();" href="#">Elimina</a></b>&nbsp;<a onClick="uncheckTutti();" href="#"> (-)</a></th><th><a onClick="aggiorna_test();" href="#">Aggiorna</a></b></th></tr>

<% ' response.write(rsTabella.fields("Data") )
   if rsTabella.eof then
   %>
   <tr><td colspan="7">Vuota</td></tr>
   <%
   end if
   i=1
   ' prelevo i dati da inserire dalla query sui risultati
   do while not rsTabella.eof %>

   <tr><td><input type="hidden" value="<%=rsTabella.fields("CodiceAllievo")%>" name="txtStud<%=i%>" /><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td><td><input type="text" class="input-mini" value="<%=rsTabella.fields("Risultato")%>" name="txtPunti<%=i%>" /></td><td><%=rsTabella.fields("Risultato")%></td><td><%=rsTabella.fields("Data")%></td><td><%=left(rsTabella.fields("Ora"),5)%></td><td> <input type="checkbox"  name="cbDelete<%=i%>" title="<%=i%>" value="<%=i%>" checked="true" ></td><td><input type="text" title="Cambia risultato"  class="input-mini"  size="4" name="txtCambia<%=i%>" value=""/> </td></tr>
   <%  rsTabella.movenext
   i=i+1
   loop
   rsTabella.close
%>
</table>

<br />
 <input type="hidden" value="<%=i-1%>" name="txtnumRec" />
 <input type="checkbox"  name="cbScrutini" title="Selezionare se il voto deve contribuire al calcolo della media per lo scrutinio">   Registra anche per scrutini &nbsp;&nbsp;&nbsp;<br><br/>
 <select name="txtTipoVoto">
			<option selected value="S">Scritto</option>
            <option value="O">Orale</option>
            <option value="P">Pratico</option>

	</select>

  					<p > primo  <input TYPE="RADIO"   name="VF" value=1> <br>
                   <p>  secondo  <input TYPE="RADIO"  name="VF" value=2> <br>
                   <p>  terzo  <input TYPE="RADIO"  name="VF" value=3 checked="checked"> <br>
                   <p>  quarto  <input TYPE="RADIO"  name="VF" value=4> <br>


 <br>
<input class="btn btn-primary" type="submit" value="Convalida Quiz" />
</form>

               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6>

                <h3> <i class="icon-reorder"></i>  Risultati con Sessione 0  </h3>

   <%  ' mopstro i quiz di quelli che hanno inviato con sessione 0 per fare i furbi e vedere la correzione

                QuerySQL="SELECT Allievi.CodiceAllievo, Allievi.Nome, Allievi.Cognome, Risultati1.Risultato, Risultati1.Data,Risultati1.Ora, Risultati1.CodiceTest,Risultati1.Risultato*8/100 as [PUNTI],Risultati1.ID_R " &_
" FROM Allievi INNER JOIN Risultati1 ON Allievi.CodiceAllievo = Risultati1.CodiceAllievo " &_
" WHERE Risultati1.Tipo="&  tipoTest &  " and   Risultati1.Data >='" & DataAvviso & "' AND Risultati1.CodiceTest='"&  CodiceTest & "' AND Risultati1.Sessione=0"&_
" ORDER BY Allievi.Cognome Asc, Risultati1.Ora; "

Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySQL)  %>
<table class="table-condensed table-bordered">
<tr><th>Cognome</th><th>Nome</th><th width="10%">Punteggio</th><th width="10%">P</th><th>Data</th><th>Ora</th><th><a onClick="elimina_test();" href="#">Elimina</a></b>&nbsp;<a onClick="uncheckTutti();" href="#"> (-)</a></th><th><a onClick="aggiorna_test();" href="#">Aggiorna</a></b></th></tr>

<% ' response.write(rsTabella.fields("Data") )
   if rsTabella.eof then
   %>
   <tr><td colspan="7">Vuota</td></tr>
   <%
   end if
   i=1
   ' prelevo i dati da inserire dalla query sui risultati
   do while not rsTabella.eof %>

   <tr><td><input type="hidden" value="<%=rsTabella.fields("CodiceAllievo")%>" name="txtStud<%=i%>" /><%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td><td><input type="text" class="input-mini" value="<%=rsTabella.fields("Risultato")%>" name="txtPunti<%=i%>" /></td><td><%=rsTabella.fields("Risultato")%></td><td><%=rsTabella.fields("Data")%></td><td><%=left(rsTabella.fields("Ora"),5)%></td><td> <input type="checkbox"  name="cbDelete<%=i%>" title="<%=i%>" value="<%=i%>" checked="true" ></td><td><input type="text" title="Cambia risultato"  class="input-mini"  size="4" name="txtCambia<%=i%>" value=""/> </td></tr>
   <%  rsTabella.movenext
   i=i+1
   loop
   rsTabella.close
%>
</table>



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

function elimina_test() {


	  if (confirm("Vuoi eliminare tutti i risultati selezionati ?")) {
    document.Aggiorna.action = "cancella_test.asp?SessioneQuiz=<%=SessioneQuiz%>&tipoTest=1&CodiceTest=<%=CodiceTest%>&id_classe=<%=id_classe%>&DataTest=<%=DataAvviso%>&TitoloTest=<%=TitoloTest%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();
	 }
}
function aggiorna_test() {

	  if (confirm("Vuoi aggiornare tutti i risultati selezionati ?")) {
   // document.Aggiorna.action = "#";
	  document.Aggiorna.action = "cancella_test.asp?aggiorna=1&SessioneQuiz=<%=SessioneQuiz%>&tipoTest=1&CodiceTest=<%=CodiceTest%>&id_classe=<%=id_classe%>&DataTest=<%=DataAvviso%>&TitoloTest=<%=TitoloTest%>";

		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();
	 }
}

 function uncheckTutti() {
	with (document.Aggiorna) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'checkbox')
		elements[i].checked = false;
		}
	}
}


  </script>

 </html>
