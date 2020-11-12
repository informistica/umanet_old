<!doctype html>
<html>
<head>
<script src="../js/google.js"></script><meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 <meta charset="UTF-8">

	<title>Classifica</title>

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
	<!-- jQuery UI -->
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
	<!-- dataTables -->
	<link rel="stylesheet" href="../../css/plugins/datatable/TableTools.css">
	<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">
	<!-- Theme CSS -->

    	 <link rel="stylesheet" href="../../css/style-themes.css">

        <link rel="stylesheet" href="../../css/docs.css">


	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>

	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
<!--	<script src="../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>


    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

    <!-- dataTables -->
	<script src="../../js/plugins/datatable/megaDatatable.min.js"></script>

	<!-- Chosen -->
	<script src="../../js/plugins/chosen/chosen.jquery.min.js"></script>



	<!-- Theme framework -->
	 <script src="../../js/eak_app_dem.min.js"></script>




      <script language="javascript" type="text/javascript">
	  function SessioneScaduta(){
            window.alert("Sessione  scaduta, effettua nuovamente il Login!");
             location.href="../home.asp";
             }
      </script>

	<!--[if lte IE 9]>
		<script src="js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />

</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

	<div id="navigation">


<!--#include file = "../service/controllo_sessione.asp"-->

<%

 		Set ConnessioneDB0 = Server.CreateObject("ADODB.Connection")  ' per il DBClassifica
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection") ' per il forum
		Set ConnessioneDB2 = Server.CreateObject("ADODB.Connection") ' per lavagna
		Set ConnessioneDB3 = Server.CreateObject("ADODB.Connection") ' per diario


		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

		<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_lavagna.inc" -->
        <!-- #include file = "../stringhe_connessione/stringa_connessione_diario.inc" -->

                 <!-- #include file = "../cClasse/studente_domande_include/1_periodi_date.asp" -->

		<!-- #include file = "../include/navigation.asp" -->
        <!-- #include file = "../extra/test_server.asp" -->
		<!-- #include file = "../include/formattaDataCla.inc" -->
        <!-- #include file = "../cUtenti/adovbs.inc" -->





      <%

' non si capisce per quale cazzo di motivo non funziona il controlla sessione come in home_app.asp ??????
' non va il window.alert dentro if , se lo metto fuori funziona

'daStud=Request.QueryString("daStud")
'daMenu=Request.QueryString("daMenu")
'DataCla=request.form("txtData")
'DataCla2=request.form("txtData2")
'DataClaq=request.QueryString("DataClaq")
'DataClaq2=request.QueryString("DataClaq2")
' DataClaq="01/09/2017"
' DataClaq2="01/10/2018"
' if DataClaq="" then
  ' DataClaq=DataClaDefault
' end if
' if DataClaq2="" then
  ' DataClaq2=DataCla2Default
' end if
'if daMenu<>"" then
'    DataCla=request.QueryString("DataClaq")
'    DataCla2=request.QueryString("DataClaq2")
'end if
'if daStud<>"" then
'   DataClaq= DataCla
'   DataClaq2=DataCla2
'end if
'
'
'
'Session("DataClaq")=DataClaq
'Session("DataClaq2")=DataClaq2



'' se è la prima chiamata il valore del form sopra la classifica è nullo
if (DataCla<>"") and (DataCla2<>"") then
	Session("DataCla")=DataCla
	Session("DataCla2")=DataCla2 ' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
else
   Session("DataCla")= Session("DataClaq")
   Session("DataCla2")= Session("DataClaq2")
   DataCla=Session("DataCla")
   DataCla2=Session("DataCla2")
end if

	'Response.AddHeader "Refresh", "600"	 per ridurre il carico sul DBCopiatestonline, implementerò il blocco della pagina Libro che ha il ruolo di mantenere attiva la ssessione
  Cartella=Request.QueryString("Cartella")
  ClasseProfili=Request.QueryString("classe")
  TitoloCapitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")


function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
  sAns=  Replace(sAns,"&","e")
  sAns=  Replace(sAns,"/","-")
  sAns=  Replace(sAns,"\","-")
  sAns=  Replace(sAns,"?",".")
  sAns=  Replace(sAns,"*","x")
  sAns=  Replace(sAns,"<","_")
  sAns=  Replace(sAns,">","_")

ReplaceCar = sAns
end function



	'SetLocale(1040)  ' imposta il formato data corretto
    QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato=rsTabellaCI("CIAbilitato")
	ScalaValutaz=rsTabellaCI("ScalaValutaz")
	rsTabellaCI.close
    Dim esecuzione
    set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito




%>



	</div>
	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->


<%
On Error Resume Next

function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sInput,"è","&egrave;")

  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
'
 ' sAns=  Replace(sInput,"&igrave;","i")
'  sAns=  Replace(sAns,"&egrave;","e'")
'  sAns=  Replace(sAns,"&ugrave;","u'")
'  sAns=  Replace(sAns,"?","&ograve;")
'  sAns=  Replace(sAns,"&agrave;","a'")
'

ReplaceCar = sAns
'ReplaceCar = sInput

end function


 QuerySQL="Select count(*) from Allievi where Id_Classe='" & Session("Id_Classe")&"';"
	Set rsTabellaT = ConnessioneDB.Execute(QuerySQL)
	numstud=rsTabellaT(0)
	 'response.write(numstud)
	rsTabellaT.close


Dim periodi() ' vettore delle date per il calcolo della classifica, più avanti farò il redim
Dim vetstud(400) ' massimo numero di studenti possibile
'Dim vetstud(cint(numstud))
vetstud(0)="?"

xEstrazione=request.querystring("xEstrazione")
id_classe=request.querystring("id_classe")

classe=request.querystring("classe")
divid=request.querystring("divid")
if divid="" then
   divid=Session("divid")
end if
divid2=request.querystring("divid")

PS=request.querystring("PS") ' vale 1 se devo mostrare anche i Punti Social chiamato da javasscript
if PS="" then ' per la prima chiamata mostrio i PS
   PS=1
end if
cod=Request.QueryString("cod")





 %>

			<div id="main">
				<div class="container-fluid">
				<!-- #include file = "../include/navigation_small.asp" -->
					<div class="page-header">
						<div class="box">
                          <div class="box-title">
                             <h2> <i class="icon-user"></i> Classifica </h2>
                             <%'response.write("num="&numstud)
							 %>
                          </div>
                        </div>

					</div>
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
                            <li>
                                <a href="#">Quaderno</a>
                            </li>
                        </ul>
						<div class="close-bread">
							<a href="#">
								<i class="icon-remove"></i>
							</a>
						</div>
					</div>
                    -->














					<div class="row-fluid">
						<div class="span12">





                        <b>
                        <% 'if (request.form("txtData")<>"") then response.write("Caklcola data") end if
                          ' response.Write("Classifica al " &day(date())&"/"&month(date())&"/"&year(date()))
						  %> </b>

                    <!-- #include file = "studente_domande_include/1_periodi.asp" -->
                  <!--  <input type="button"  style="width:60px;height:25px;" value="Invia" name="B1" onClick="aggiorna()"> -->
                    <button class="btn"  onClick="aggiorna()">Invia</button>
                     <input type="checkbox"  name="cbPS" value="1" checked="true" title="Deseleziona per escludere i Punti Social dalla classifica">  <b>
                        Includi PS
                     </p>
                    </form>


                         <!-- #include file = "studente_domande_include/1_classifica_new.asp" -->


							<div class="box box-color box-bordered <%=Session("stile")%>">
								<div class="box-title">
									<h3>
										<i class="icon-table"></i>
										Punteggi attivit&agrave;&nbsp;
										<% if Session("DB") <> 1 then %>
										<a id="nmax0" href="javascript:void(0)" onclick="window.localStorage.setItem('nmax',1); location.reload()" style="text-decoration:none; color:white; white-space:nowrap"><i style="font-size:14px" class="icon-resize-vertical"></i></a>
										<a id="nmax1" href="javascript:void(0)" onclick="window.localStorage.setItem('nmax',0); location.reload()" style="text-decoration:none; color:white;  white-space:nowrap"><i style="font-size:14px" class="icon-reply"></i></a>
										<% end if %>
									</h3>
								</div>

								<div class="box-content nopadding">

								<% if Session("DB") = 1 then %>
									 <table class="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped">
                                   <% else %>
								   <table id="tabcla">
								   <%end if%>

										<thead>
											<tr>
												<th  title="Posizione"><b>N.</b></th>
                                                <th><b>Cognome Nome</b>
                                                <th title="Totale"><center><b>TOT</b></center></th>
                                                <th title="Punti Domande"><center><b>PD</b></center></th>
                                                <th title="Punti Nodi"><center><b>PN</b></center></th>
                                                <th title="Punti Frasi"><center><b>PF</b></center></th>
                                                <th title="Punti Metafore"><center><b>PM</b></center></th>
                                                <th><center><b title="Punti Crediti">PC</b></center></th>
                                               <% if cint(PS)<>0 then%>
                                                <th><center><b title="Punti Social (Risposte nel Forum)">PS</b></center></th>
                                               <%end if%>
                                                 <th><center><b title="Punti Laboratorio (Consegne in Bacheca)">PL</b></center></th>
																								 <th><center><b title="Punti Elexpo (Feedback ricevuti)">PE</b></center></th>
																								 <th><center><b title="Punti Interrogazioni (Risposte orali)">PO</b></center></th>
                                                <th title="Voto Virtuale"><center><b>VV</b></center></th>
                                                <th title="Percentuale rispetto al massimo"><center><b>%</b></center></th>
																								<%if strcomp(Session("DB"),"1")=0 then%>
																								<th title="Raggruppamento"><center><b>Tags</b></center></th>
																								<%end if%>

											</tr>
										</thead>
										<tbody>




                                             <%


			'	dim objFSO,objCreatedFile
'				Const ForReading = 1, ForWriting = 2, ForAppending = 8
'				Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				'url="DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/log_calcola.txt"
'			    url="C:\Inetpub\umanetroot\expo2015Server\logCla.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'
'


	 i=0
	 rsTabella.movefirst
	 max=rsTabella("TOT")


	 do while not rsTabella.eof
	 CodiceAllievo=rsTabella("CodiceAllievo")

'if (rsTabella("TOT")=0) then
'	     rsTabella("TOT")=1
'	end if


if (i=0) then
 classe_riga="success"
else
		if (strcomp(rsTabella("CodiceAllievo"),Session("CodiceAllievo")) = 0)  then
			classe_riga="info"
		else
			 if ((fix((rsTabella("TOT")*ScalaValutaz/max)*10)/10)<6) and (strcomp(ucase(cartella),"EXPO")<>0) then
				  classe_riga="error"
			 else
			  classe_riga=""
			end if
		end if
end if

	Url_imgcla = rsTabella("Url_img")

	if strcomp(Url_imgcla&"","")=0 then
    urlimgcla = "../../img/no-avatar.jpg"
	else

	urlcla= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&cartella&"/Profili/thumb" ' vuole il percorso relativo della cartella

					  urlcla=Replace(urlcla,"\","/")
					  urlimgcla=urlcla&"/"& Url_imgcla
	end if

     if cint(PS)=0 then   %>

     <tr class="<%=classe_riga%>"><td style="vertical-align:middle"><%=i+1%></td><td><img src="<%=urlimgcla%>" style="width:28px; height:28px">&nbsp;&nbsp;<a style="text-decoration:none" href="quaderno.asp?umanet=0&daStud=1&divid=<%=divid%>&DataClaq=<%=DataCla%>&DataClaq2=<%=DataCla2%>&id_classe=<%=id_classe%>&classe=<%=classe%>&cod=<%=rsTabella("CodiceAllievo")%>"><%=replaceCar(rsTabella("Cognome"))&" "&replaceCar(rsTabella("Nome"))%></a></td><td style="vertical-align:middle"><center><%=rsTabella("TOT")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PD")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PN")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PF")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PM")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PC")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PL")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PE")%></center></td> <td style="vertical-align:middle"><center><%=rsTabella("PO")%></center></td> <td style="vertical-align:middle"><center><%=fix((rsTabella("TOT")*ScalaValutaz/max) * 10) / 10 %></center></td><td style="vertical-align:middle"><center><%=round(fix((rsTabella("TOT")*100/max) * 10) / 10 ) %></center></td>	<%if strcomp(Session("DB"),"1")=0 then%> <td style="vertical-align:middle"><center><%=rsTabella("Tags")%></center></td><%end if%></tr>

      <%else%>
				<tr class="<%=classe_riga%>"><td style="vertical-align:middle"><%=i+1%></td><td><img src="<%=urlimgcla%>" style="width:28px; height:28px">&nbsp;&nbsp;<a style="text-decoration:none" href="quaderno.asp?umanet=0&daStud=1&divid=<%=divid%>&DataClaq=<%=DataCla%>&DataClaq2=<%=DataCla2%>&id_classe=<%=id_classe%>&classe=<%=classe%>&cod=<%=rsTabella("CodiceAllievo")%>"> <%=ReplaceCar(rsTabella("Cognome"))%>&nbsp;<%=ReplaceCar(rsTabella("Nome"))%>  </a></td><td style="vertical-align:middle"><center><%=rsTabella("TOT")%>  </center></td><td style="vertical-align:middle"><center><%=rsTabella("PD")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PN")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PF")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PM")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PC")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PS")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PL")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PE")%></center></td><td style="vertical-align:middle"><center><%=rsTabella("PO")%></center></td> <td style="vertical-align:middle"><center><%=fix((rsTabella("TOT")*ScalaValutaz/max) * 10) / 10 %></center></td><td style="vertical-align:middle"><center><%=round(fix((rsTabella("TOT")*100/max) * 10) / 10 ) %></center></td>	<%if strcomp(Session("DB"),"1")=0 then%> <td style="vertical-align:middle"><center><%=rsTabella("Tags")%></center></td><%end if%></tr>
     <%end if%>
      <%
		' aggiungo al vettore che servirà per estrarre a sorte per l'orale


i=i+1
vetstud(i)=rsTabella.fields("Cognome")
strstud = strstud & rsTabella.fields("Cognome") & ","

rsTabella.movenext

'href="quaderno.asp?daStud=1&DataClaq="&DataCla&"&DataClaq2="&DataCla2&"&id_classe="&id_classe&"&classe="&classe&"&cod="&rsTabella("CodiceAllievo")
'objCreatedFile.WriteLine(href)



loop
' numero studenti per quella classe
NumStud=i
rsTabella.close()



				'objCreatedFile.Close


%>











										</tbody>
									</table>
								</div>
							</div>
						</div>
					</div>

                     <div class="row-fluid">

					</div>

                <%
				'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logClassifica_518.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 518"
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				%>



                <div class="row-fluid">

                     <!-- #include file = "studente_domande_include/1_report.asp" -->

					</div>


               <%
		'	   Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logClassifica_536.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 518"
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				%>




                    <%
if (session("Admin")=true) and (id_classe <> "") then %>
                    <div class="row-fluid">

                     <!-- #include file = "studente_domande_include/1_altro.asp" -->

					</div>
			<%end if%>
			</div>
		</div>

        <%'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logClassifica_557.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 518"
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				%>

		<% if Session("DB") <> 1 then %>
		<script>
											var x = window.localStorage.getItem('nmax');
											if(x=="" || x == "undefined" || x == null || x == "null"){
											x=0;
											}

											if(x==0){
											document.getElementById("nmax0").style.display="inline-block";
											document.getElementById("nmax1").style.display="none";

											if( /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ) {
												document.getElementById("tabcla").className="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped";
												document.getElementById("nmax0").style.display="none";
												document.getElementById("nmax1").style.display="none";

											} else {
												document.getElementById("tabcla").className="table table-hover table-nomargin table-bordered dataTable-fixedcolumn dataTable-scroll-x table-striped";
											}


											}else{
											document.getElementById("nmax1").style.display="inline-block";
											document.getElementById("nmax0").style.display="none";
											document.getElementById("tabcla").className="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped";
											}
										</script>

		<% end if %>
		 <!-- #include file = "../include/colora_pagina.asp" -->
	</body>

<%'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logClassifica_567.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				QuerySQL="riga 518"
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
				%>



<script type="text/javascript">

function aggiorna() {

		with (document.dati) {

		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=divid%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&daForm=1";
		 else
		   document.dati.action = "?divid=<%=divid%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&daForm=1";

	    }
		document.dati.submit();
}

 function aggiornaStud() {

		with (document.dati) {

		if (elements["cbPS"].checked == true)
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=1&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		 else
		   document.dati.action = "?divid=<%=session("divid")%>&classe=<%=classe%>&id_classe=<%=id_classe%>&PS=0&cod=<%=cod%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";

	    }
		document.dati.submit();
}

 </script>

 <script language="javascript" type="text/javascript">
function cancella_avviso() {

	  if (confirm("Vuoi cancellare tutti gli avvisi selezionati ?")) {
    document.Aggiorna.action = "cancella_avviso.asp?tipoAvviso=0&CodiceAllievo=<%=CodiceAllievo%>&Id_Classe=<%=Id_Classe%>&DataClaq=<%=DataClaq%>&DataClaq2=<%=DataClaq2%>";
		//document.dati.action = "../home.asp"
		document.Aggiorna.submit();
	 }
}

  </script>


	<script type="text/javascript">

	function caricaeccezioni (id_classe){
//alert(id_classe);

var url = "studente_domande_include/2_modifica_eccezioni_classe.asp?id_classe="+id_classe;

var testo;
var stato1, stato2;

$("#compitispec").html("Attendere prego...");
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
	</script>

  <script type="text/javascript" src="../js/refresh_session.js"></script>

	</html>
