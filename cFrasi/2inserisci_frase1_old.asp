<%@ Language=VBScript %>

<!doctype html>
  <meta https-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />
<html>
<head>
<script src="js/google.js"></script>
<script language="javascript" type="text/javascript">
    function showText2() {
        window.alert("La sessione � scaduta, effettua nuovamente il Login!")
        location.href = "../../home.asp"
        //location.href=window.history.back();
    }
 </script>



     <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
  <!--  <meta charset="utf-8">-->
  <meta https-equiv="Content-Type" content="text/html; charset=ISO-8859-1" />


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
	<!-- Theme framework -->
	<script src="../../js/eak_app_dem.min.js"></script>

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />

 <title>Inserimento</title>

   <!-- #include file = "../include/header.asp" -->
</head>

<body class='theme-<%=session("stile")%>'>
	<div id="navigation">

        <%

 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>     <!-- #include file = "../var_globali.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		  <!-- #include file = "../include/navigation.asp" -->
          <!--#include file = "../service/controllo_sessione.asp"-->

  <%

ID_Prefrase=Request.QueryString("ID_Prefrase")
Quesito=Request.QueryString("Quesito")
Capitolo=Request.QueryString("Capitolo")
Paragrafo=Request.QueryString("Paragrafo")
Img=Request.QueryString("Img")
CodiceSottopar = Request.QueryString("CodiceSottopar")
Sottoparagrafo=Request.QueryString("Sottoparagrafo")

  'url1= Request.QueryString("url1")
	'   url2= Request.QueryString("url2")
	 '  url3= Request.QueryString("url3")

	 'prendo url via Form come il testo (Sintesi) -> ho sistemato l'invio POST

 url1=Request.Form("txtImg1")
 url2=Request.Form("txtImg2")
 url3=Request.Form("txtImg3")

%>

	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">

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
							<a href="#more-files.html">Libro</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html"> Inserisci frase</a>
                            <i class="icon-angle-right"></i>
						</li>
                        <li>
							<a href="#more-blank.html"> <%Response.write (Capitolo)  %></a>
						</li>
					</ul>
					<div class="close-bread">
						<a href="#"><i class="icon-remove"></i></a>
					</div>
				</div>

				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-comments"></i><%Response.write (Paragrafo)  %>
                         <%if CodiceSottopar<>"" then%>
                              /<%=response.write(Sottoparagrafo)%>
                            <%end if%>
                         </h3>
			          </div>
				      <div class="box-content">

					 <% %>


                       <!--#include file = "controlla_inserimento.asp"-->
                       <%If True Then  ' non � stata trovata quindi la inserisco' ho annullato il controllo con l'inserimento del salvataggio bozza 4/4/2019 %>

    <div id="container">
	<div class="risultati_test" >
	<font color=#FF0000 size="4">

   <% Response.Buffer=True


   Nome=Request.QueryString("Nome")
   Cognome=Request.QueryString("Cognome")
   CodiceTest = Request.QueryString("CodiceTest")

   if (CodiceTest="") then
        CodiceTest=Request.Cookies("Dati")("CodiceTest")
   end if
   DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
	'DataTest = Now()
   Cartella=Request.QueryString("Cartella")
   Tipo=Request.QueryString("Tipo") ' tipo di domanda 0 normale 1 estesa
   StringaConnessione= Request.Cookies("Dati")("StrConn")
   prefrase=Request.QueryString("prefrase") ' serve per capire il chiamante e quindi sapere se alla fine devo redirectare ad home_ver o home_app
    ID_Prefrase=Request.QueryString("ID_Prefrase") ' serve per controllare se � gi� stata inserita
   by_UECDL=Request.QueryString("by_UECDL")
   'Apertura della connessione al database

                            'Lettura dei dati memorizzati nei cookie.
  ' CodiceTest = Request.Cookies("Dati")("CodiceTest")


   CodiceCorso = Request.Cookies("Dati")("CodiceCorso")
  ' CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
    CodiceAllievo = Session("CodiceAllievo")
   CodiceCap=Request.Cookies("Dati")("CodiceCap")
   Num=Request.QueryString("Num")
   Capitolo=Request.QueryString("Capitolo")

		Paragrafo=Request.QueryString("Paragrafo")
		Modulo=Request.QueryString("Modulo")
		DataTest=Day(date())&"/"&Month(date())&"/"&Year(date())
	'DataTest = Now()

	   Chi = Request.Form("txtChi")
	   Chi = Replace(Chi, Chr(34), "'")
	   Chi=  Replace(Chi,"'","''")
  	   Sintesi= Request.Form("S1")

	   'response.write("<script>alert('"&Chi&"')</script>")




	   if Chi="" then
	   Chi=Request.QueryString("Quesito")
	   end if

		'Elimino questa verifica perch� ho sistemato l'invio da POST quindi i dati arrivano

		' if Sintesi="" then
	    ' Sintesi=Request.QueryString("Sintesi")
	    ' end if


	     'QuerySQL="INSERT INTO Frasi_Img (Id_Frase,Url,Nome) SELECT " & ID & ",'" & urldb & "','" & imgname & "';"

	 '  response.write(Sintesi) &"<br>????<br>" &url1 &"Chi?"&Chi
   	  ' Sintesi= Replace(Sintesi, Chr(34), "'")
 '  Sintesi=  Replace(Sintesi,"'","''")




   if ( (len(Chi)=0)  ) then

   errore=2

   end if

 if (errore=0) then


 %>

 <!--#include file="2inserisci_frase1_include.asp"-->


 <%
			 if (prefrase<>"") then 'se sono stato chiamato da compilaprefrase devo ritornare ad home_app
			    %>

					<%' se sono stato chiamato da compilaprefrase di home_app_uecdl devo tornare li
					   if by_UECDL<>"" then %>
                   <!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
    <h5><a href="../U-ECDL/home_uecdl_app.asp?id_classe=<%=Session("Id_Classe")%>&divid=<%=Session("divid")%>"> Torna al Libro U-ECDL ... </a></h5>
                <%else%>
                    <!-- REDIRECT INTELLIGENTE  -->


                <%end if%>

   <h5><a href="2compilaprefrase.asp?Cartella=<%=Cartella%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&CodiceTest=<%=CodiceTest%>&Modulo=<%=Modulo%>&prefrase=<%=prefrase%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>"> Torna alle domande... </a></h5>


				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
 				<%
			 else
 				%>
<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h5><a href="../../home_ver.asp?id_classe=<%=Session("Id_Classe")%>"> Torna alla pagina Verifica... </a></h5>
				<%

			end if
else
	  if (errore=1) then
	'	 response.write("Controlla che il numero della risposta esatta sia compreso tra 1 e 4")
	  end if
	  if (errore=2) then%>
		<button class="btn btn-danger">Controlla che non ci siano campi lasciati vuoti</button>

	  <%end if %>

		<% if not Session("containsert") > 1 and Abs(DateDiff("s",Session("firstinsert"),Time())) < 10 then  %>
		<a href="#" onClick="history.go(-1);return false;">Indietro</a>
		<% end if %>
	  <%
end if

'end if ' di If not(rsTabella.BOF=True And rsTabella.EOF=True) Then non � stata trovata la frase la inserisco
%>



                        </div>
			        </div>
			      </div>
			    </div>
			</div>


		</div>
          <!-- #include file = "../include/colora_pagina.asp" -->
	  <!--fine main-->

	</body>

	 </html>
