<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Valutazioni frasi</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	<meta charset="utf-8">

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


    <script src="https://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>
 <script src="../../js/datapicker_it.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />


       <script src="../js/privacy.js"></script>
	
<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>

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

  <script type="text/javascript">
function checkTutti() {
	numcb=0;
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		if (elements[i].type == 'checkbox')
		    {
		     elements[i].checked = true;
			 numcb=numcb+1;
			}
		}
	}
	document.dati.txtNUMREC.value=numcb;
}
function uncheckTutti() {
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'checkbox')
		elements[i].checked = false;
		}

	}
	document.dati.txtNUMREC.value=0;

}
function aggiorna(nome) {

		with (document.dati) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMREC.value=parseInt(txtNUMREC.value)+1;
		 else
		    txtNUMREC.value=parseInt(txtNUMREC.value)-1;
	    }
}
function aggiorna2(nome) {

		with (document.dati) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)+1;

		 else
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)-1;
	    }
}

 </script>

 <script type="text/javascript">
//assegna la valutazione solo se il record è selezionato per la valutazione
function valutaTutti(voto) {
	var stringa,stringa2;
	var voto=document.dati.txtVoto.value;

	numcb=1;

		for (var i=0; i < document.dati.elements.length; i++) {
			stringa=document.dati.elements[i].name;
			stringa2='txtVAl'+numcb;

		if (stringa.search(stringa2) == 0)
		     {
			if (document.dati.elements["cbVal"+numcb].checked == true) document.dati.elements[i].value = voto;
			numcb=numcb+1;

		 	}
		}
}

function segnalaTutti() {
	var stringa,stringa2;

	numcb=1;

		for (var i=0; i < document.dati.elements.length; i++) {
			stringa=document.dati.elements[i].name;
			stringa2='txtSegnalata'+numcb;

		if (stringa.search(stringa2) == 0)
		     {

           document.getElementById("txtSegnalata"+numcb).checked = true;
           document.getElementById("txtSegnalazione"+numcb).value = document.getElementById("txtFeedback").value;

			numcb=numcb+1;

		 	}
		}
}

function desegnalaTutti() {
	var stringa,stringa2;

	numcb=1;

		for (var i=0; i < document.dati.elements.length; i++) {
			stringa=document.dati.elements[i].name;
			stringa2='txtSegnalata'+numcb;

		if (stringa.search(stringa2) == 0)
		     {

           document.getElementById("txtSegnalata"+numcb).checked = false;
           document.getElementById("txtSegnalazione"+numcb).value = "";

			numcb=numcb+1;

		 	}
		}
}


  function selezionatutti(id) {
	//per modificare tutte le date di un form impostandole uguale al valore della textbox passata per parametro
    //document.dati.date3.value="11/11/1111";
	// document.dati.txtScadenza1.value="19/11/2010";

    var el = document.getElementById(id);
    var idtext=1;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=0; i < elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'txtDataDomanda'+idtext)
		    {
		    elements[i].value = el.value;
			idtext=idtext+1;
			}
	 }
	 return true;
    }
 }

 </script>


<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>

<%
  Response.Buffer = true
  On Error Resume Next









  ' dichiarazione delle variabili per contenere i parametri (codice del corso, codice del test, titolo del test) passatti dalla pagina menu

  Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
 'StringaConnessione= Response.Cookies("Dati")("StrConn")

   Set ConnessioneDB = Server.CreateObject("ADODB.Connection") %>

    <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
	<%
  ' definizione dei valori delle variabili leggendoli dall'oggetto Request utilizzando il metodo QueryString("Nome parametro")
  CodiceSottopar=Request.QueryString("CodiceSottopar")
  CodiceAllievo=Request.QueryString("cod")
  'cla=Request.QueryString("cla")
  Codice_Test=Request.QueryString("CodiceTest")
  CodiceFrase=Request.QueryString("CodiceFrase")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("CodiceTest")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Modulo=Request.QueryString("Modulo")
  Cartella=Request.QueryString("Cartella")
  NumRec=Request.QueryString("NumRec") ' è la variabile i contatore per scorrere il form e fare update
  idprefrase=Request.QueryString("idprefrase")
  '-----
  Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("ID_MOD")
  ID_PAR=Request.QueryString("ID_PAR")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  TutteCap=Request.QueryString("TutteCap") ' vale 1 se devo visualizzare tutte le domande  dello studente CAPITOLO
  TuttePar=Request.QueryString("TuttePar") ' vale 1 se devo visualizzare tutte le domande  dello studente PARAGRAFO
  if left(Cartella,1)<>"" then ' DA SISTEMARE NELLE QUERY PER I GRUPPI !!!!!!!!!!!!!
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if

  BoxApro=Request.QueryString("BoxApro")

  			'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'                url1="C:\inetpub\umanetroot\expo2015Server\logFilTutte.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'

  'per selezionare il periodo della
DataClaq=Request.QueryString("DataClaq")
DataClaq2=Request.QueryString("DataClaq2")
if DataClaq="" then
DataClaq=Session("DataClaq")
end if
if DataClaq2="" then
DataClaq2 =Session("DataClaq2")' per rendere visibile la data alle pagine che devono fare il redirect a studente.asp
 end if

  '----
 FraseScelta=Request.QueryString("FraseScelta")


 function ReplaceCar(sInput)
dim sAns

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


'if MO<>"" then
' Modulo=MO
'end if
'
if (FraseScelta<>"") then ' se sono chiamata da 2scegli_valutazioni_frasi visualizzo la stessa frase per tutti gli studenti
   'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where  Chi='"&FraseScelta&"' and ID_Mod<>'6C' AND Cartella = '" & Cartella &"' order by Data asc, Ora asc;"
   QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where  Id_Prefrase='"&idprefrase&"' order by Data asc, Ora asc;"

else



	if (CodiceAllievo<>"") then  ' se sono stata chiamata dalla pagina studente_domande, valuterò solo le domande di quello studente
		 if (Nulle<>"") then ' se devo mostrare sollo quelle con voto=0
				  if (Tutte<>"") then ' visualizzo tutte le frasi
					  QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where Voto=0 and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' AND  Cartella = '" & Cartella &"';"
					else ' visualizzo solo quelle non valutate voto=0
						 QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_MOD='"& ID_MOD &"' and Voto=0 and CodiceAllievo='"&CodiceAllievo&"' AND  Cartella = '" & Cartella &"';"
				   end if
		else
				if (Data<>"") then ' se devo mostrare sollo quelle dopo una certa data, sono chiamata da me stessa, non è il caso in cui sono chiamata da studente_domande
					if (TutteTutte<>"") then ' tutte quelle dopo una certa data
						 QuerySQL="SELECT MODULO_PARAGRAFO_FRASI1.*, MODULO_PARAGRAFO_FRASI1.Data FROM MODULO_PARAGRAFO_FRASI1 WHERE MODULO_PARAGRAFO_FRASI1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"  and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' AND  Cartella = '" & Cartella &"';"
				     else
					        if (TuttePar<>"") then ' tutte quelle dopo una certa data e di un parasgrafo
						 QuerySQL="SELECT MODULO_PARAGRAFO_FRASI1.*, MODULO_PARAGRAFO_FRASI1.Data FROM MODULO_PARAGRAFO_FRASI1 WHERE  ID_Paragrafo='"& ID_PAR &"' and MODULO_PARAGRAFO_FRASI1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"  and CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' AND  Cartella = '" & Cartella &"';"

					        else
					    ' solo quelle del modulo dopo una certa data
						        if (TutteCap<>"") then
						 QuerySQL="SELECT MODULO_PARAGRAFO_FRASI1.*, MODULO_PARAGRAFO_FRASI1.Data FROM MODULO_PARAGRAFO_FRASI1 WHERE MODULO_PARAGRAFO_FRASI1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &"  and ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"'AND Cartella = '" & Cartella &"';"
				                end if
						  end if

					 end if
				else
				   if (TuttePar<>"") then
						QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where  CodiceAllievo='"&CodiceAllievo&"' and ID_Mod<>'6C' AND  Cartella = '" & Cartella &"' and ID_Paragrafo='"&ID_PAR &"'" &_
						  " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_

	 						" ;"
					else
						 QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_MOD='"& ID_MOD  &"' and CodiceAllievo='"&CodiceAllievo&"' AND  Cartella = '" & Cartella &"'" &_
						  " and (Data>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (Data<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"&_

						 " ;"
					end if
				end if
			  end if
	else
		if (Gruppi<>"") and (Nulle<>"") then ' mostro le domande per gruppo solo quelle con voto =0 NB SISTEMARE Classe !!!!!!!!!
	  'response.write("QUI")
		QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &" and Voto=0;"
		else
		   if (Gruppi<>"") then
			   QuerySQL="SELECT * FROM 1_GRUPPI_DOMANDE1 Where Gruppi1.Classe="& Classe &";"
		   else
			  if (Nulle<>"") then
					QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_Paragrafo='"& Paragrafo & "' and SotPar='"&CodiceSottopar&"' and Voto=0  order by Moduli.In_Umanet,Moduli.Posizione,Paragrafi.Posizione,Id_Prefrase asc;"
			 else
				 if (Data<>"") then

				   QuerySQL="SELECT MODULO_PARAGRAFO_FRASI1.*, MODULO_PARAGRAFO_FRASI1.Data FROM MODULO_PARAGRAFO_FRASI1 WHERE MODULO_PARAGRAFO_FRASI1.Data>=" & mid(Data,4,2)&"/" &left(Data,2)&"/"& right(Data,4)  &" AND ID_Paragrafo='"& Paragrafo & "' and SotPar='"&CodiceSottopar&"'  order by Moduli.In_Umanet,Moduli.Posizione,Paragrafi.Posizione,Id_Prefrase asc;"
				else
				  QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_Paragrafo='"& Paragrafo & "' and SotPar='"&CodiceSottopar&"' order by Moduli.In_Umanet,Moduli.Posizione,Paragrafi.Posizione,Id_Prefrase asc;"
				end if
			  end if
		  end if

		end if

	end if
end if

'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url1="C:\Inetpub\umanetroot\Anno_2012-2013_2\logFile1.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close

'QuerySQL="SELECT * FROM MODULO_PARAGRAFO_FRASI1 Where ID_Paragrafo='"& Paragrafo &"'"
'response.write("<br><br><br><br>    					" & QuerySQL)
Set rsTabellaNew = ConnessioneDB.Execute(QuerySQL)
'QueryPrima=	QuerySQL
QueryPrima=	QuerySQL


	'objCreatedFile.WriteLine("358"&QueryPrima)


'per il copia incolla ed il privato
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"

 'objCreatedFile.WriteLine("364"&QuerySQL)

	Set rsTabellaCI = ConnessioneDB.Execute(QuerySQL)
	CIAbilitato=rsTabellaCI("CIAbilitato")
	Privato=rsTabellaCI.fields("Privato")
	rsTabellaCI.close

if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=True) or (Privato=0) then  ' else è alla fine



Set objFSO = CreateObject("Scripting.FileSystemObject")
%>


<%  if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
  <% else %>
		   <% if (CIAbilitato=0) then ' disabilito copia incolla oncontextmenu="return false" ondragstart="return false onselectstart="return false""
		   %>
        <body class="theme-<%=session("stile")%>"  data-layout-sidebar="fixed" data-layout-topbar="fixed">
        <%else%>
         <body class="theme-<%=session("stile")%>" data-layout-sidebar="fixed" data-layout-topbar="fixed">

        <%end if%>
  <% end if %>





	<div id="navigation">




        <!-- #include file = "../var_globali.inc" -->

  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->



	</div>

 <%
 Capitolo=Request.QueryString("Capitolo")
 Paragrafo=Request.QueryString("Paragrafo")
 %>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>Valutazioni frasi</h1>
                    <%'response.write("jj"&QueryPrima)
					%>
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
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  <%=Capitolo%>: <%=TitoloParagrafo%></h3>
			          </div>

					  <center><h5 style="color:red">Da questa pagina non puoi eseguire modifiche sulle frasi: per effettuare le modifiche devi aprire le frasi singolarmente.</h5></center>
				      <div class="box-content">

 <%
' response.write(QueryPrima)
 %>

	 <%
   if (session("Admin")=true) then %>
	<form  method="POST" class="form-vertical"  action="2inserisci_valutazioni_frasi.asp?Nulle=1&Tutte=<%=Tutte%>&Gruppi=<%=Gruppi%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>"><b>Seleziona domande</b><br><br>
	<b>Da valutare </b>
	 <input type="submit" value="Voto=0" name="B1" class="btn"> </p>
	</form>

<form method="POST" class="form-vertical" action="2inserisci_valutazioni_frasi.asp?Gruppi=<%=Gruppi%>&Tutte=<%=Tutte%>&Cartella=<%=Cartella%>&CodiceTest=<%=Codice_Test%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&TitoloParagrafo=<%=TitoloParagrafo%>&Modulo=<%=Modulo%>&TitoloParagrafo=<%=TitoloParagrafo%>&CodiceAllievo=<%=CodiceAllievo%>&ID_MOD=<%=ID_MOD%>">
<i class="icon-calendar"></i>  <b>Data da:</b>
<%if data<>"" then%>
<input type="text" name="txtDATA" value="<%=Data%>" class="input-small datepick">
<% else%>
<input type="text" name="txtDATA" value="gg/mm/aaaa" class="input-small datepick">
<% end if%>

 <input type="submit" value="Invia" name="B1" class="btn"> </p>

</form>

</div>
<div id="segnalazioni"></div>
<%end if %>





 <div class="immagini" style="height:auto; width:auto; border:none;" >
  <form name="dati" class="form-vertical" method="POST"  action="2inserisci_valutazioni_frasi1.asp?NumRec=<%=i%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=TitoloParagrafo%>&Cartella=<%=Cartella%>&Modulo=<%=Modulo%>" >

<!----><p align="center">



	<%


	TitoloParagrafo1=TitoloParagrafo
	i=1
	consegnato=""
  segnalati=""
  segnalazioni=""
  segnalati_pos=""  ' segnalazioni positive = 2'
  segnalazioni_pos=""
	'response.write(QuerySql)
	'objCreatedFile.WriteLine("506")
	' apro il file di testo che conterrà gli url delle domande da modificare
    do while not rsTabellaNew.eof 'and i<10
%>  <!-- <p><hr style="width:80%" align="center"><br>-->
<!--<div class="hr"><hr /></div><br>-->

			<% if StrComp(TitoloParagrafo1, rsTabellaNew("TitPar")) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")

                   else
                    'i=0
                       TitoloParagrafo1= rsTabellaNew("TitPar")
					    %>
                        <font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (TitoloParagrafo1) %></b></font> <!-- stampa il titolo-->


                <%end if %>



<fieldset><legend><h4>
			 <b><%=UCASE(rsTabellaNew(2))%>&nbsp;<%=left(UCASE(rsTabellaNew("Nome")),1)&"."%></b> &nbsp; &nbsp; &nbsp;</h4></legend>


             <div class="control-group">

				  <div class="controls">





              <b>Codice Frase </b> <input class="input-mini"  type="text" name="txtCodiceDomanda<%=i%>"  tabindex="<%=(2*i)%>" value="<%=rsTabellaNew.Fields("CodiceFrase")%>" > &nbsp; &nbsp;
             <b>Data </b> <input type="text" name="txtDataDomanda<%=i%>" class="input-small"   value="<%=rsTabellaNew.Fields("Data")%>" size="8" maxlength="250"> &nbsp; &nbsp;
             <b>Ora </b>
            <b> <input type="text" class="input-mini"  name="txtOraDomanda<%=i%>"  value="<%=left(rsTabellaNew.Fields("Ora"), 5)%>" size="5" maxlength="250"> <br><br>
             <b>Chi </b>   <input class="input-xxlarge" type="text" name="txtDomanda<%=i%>"  value="<%=rsTabellaNew.Fields("Chi")%>" tabindex="<%=(7*i)+1%>"   maxlength="250">
          <INPUT TYPE="HIDDEN" NAME="txtCodiceAllievo<%=i%>" VALUE="<%=rsTabellaNew("CodiceAllievo")%>">

	</b></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->

	 <%
	    Paragrafo=rsTabellaNew(0)

		Modulo=rsTabellaNew.fields("ID_Mod")
url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceFrase")&".txt"
 url_feedback=left(url,instr(url,".")-1)
	url_feedback=url_feedback&"_feedback.txt"

	'url_feedback=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceFrase")&"_feedback.txt"
    url=Replace(url,"\","/")
	url_feedback=Replace(url_feedback,"\","/")

'response.write(url_feedback)





'response.write(url)

    urliFrame="https://www.umanetexpo.net"&homesito&"/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Cartella &"/" &Modulo&"_Frasi/"&Modulo&"_"&Paragrafo&"_"&rsTabellaNew.Fields("CodiceFrase")&".txt"
	' leggo spiegazione



	if objFSO.FileExists(url) then

		Set objTextFile = objFSO.OpenTextFile(url, ForReading)
		sReadAll="" 'pulisco sReadAll -> altrimenti rimane la vecchia spiegazione
		sReadAll = objTextFile.ReadAll
		'sReadAll=url
		objTextFile.Close
	else
	    response.write("Il file non esiste:"&url)
	  sReadAll=""
	end if

'response.write(rsTabellaNew.Fields("Segnalata")&"<br>")

	if (rsTabellaNew.Fields("Segnalata")=1) or (rsTabellaNew.Fields("Segnalata")=2)   then
    if (rsTabellaNew.Fields("Segnalata")=1)  then
    segnalati=segnalati&"'"&rsTabellaNew("CodiceAllievo")&"'"&","
    segnalazioni=segnalazioni&"<br>"&rsTabellaNew("Voto")&") " &rsTabellaNew("Cognome")&" "&left(rsTabellaNew("Nome"),1)&". : "
	   f="<span style='color:red'>Segnalata</span>"
     Else
     segnalati_pos=segnalati_pos&"'"&rsTabellaNew("CodiceAllievo")&"'"&","
     segnalazioni_pos=segnalazioni_pos&"<br><b>pt."&rsTabellaNew("Voto")&") " & rsTabellaNew("Cognome")&" "&left(rsTabellaNew("Nome"),1)&". :</b> "
      f="<span style='color:green'>Segnalata</span>"
     end if
	   'response.write("<br>Segnalata:"&rsTabellaNew.Fields("CodiceFrase"))
		' leggo feedback
		'Response.write("<br>"&url_feedback)
		if objFSO.FileExists(url_feedback) then
		
			Set objTextFile = objFSO.OpenTextFile(url_feedback, ForReading)
			feedback="" 'pulisco feedback -> altrimenti rimane la vecchia feedback
			feedback = objTextFile.ReadAll
        if (rsTabellaNew.Fields("Segnalata")=1) then
          segnalazioni=segnalazioni&feedback
        end if
        if (rsTabellaNew.Fields("Segnalata")=2) then
         segnalazioni_pos=segnalazioni_pos&feedback
        end if
		''	feedback=url_feedback
			objTextFile.Close
    else
       feedback="Nessun feedback"
    end if
	else
	 f="<span>Segnalata</span>"
	 url_feedback=""
	 feedback=""

	end if



	%><br>
	<b>Frase</b>
	<div class="container" style="word-wrap: break-word;">
	<input  type="text" name="url<%=i%>" value="<%=url%>" size="0" class="hidden">

	<p>

	<% if sReadAll = "" then
				sReadAll1 = "File spiegazione mancante. Elimina e reinserisci la frase nel tuo quaderno."
				dis = true
			 else
				sReadAll1=sReadAll
				dis = true 'lo imposto true perché tanto in questa pagina non si possono comunque fare modifiche
			end if

	%>

   <% 
   if instr(sReadAll1,"<script>")<>0 then
	   sReadAll1=Replace(sReadAll1,"<script>","")
	   sReadAll1=Replace(sReadAll1,"</script>","")
	end if
   sReadAll1 = ltrim(sReadAll1)
   sReadAll1 = rtrim(sReadAll1)
sReadAll1=sReadAll1
   lunghezza=1+round((len(sReadAll))/110)%>
    <%' if CIAbilitato=0 then   ' se lo impedisco metto la textarea altrimenti iframe
	%>
<%Response.write(ReplaceCar(sReadAll1))%>
	<!--<textarea class="input-block-level" rows="<%=lunghezza%>"   tabindex="<%=(2*i)+1%>" <% if dis = true then response.write("disabled") end if %> name="txtSpiegazione<%=i%>"><%Response.write(ReplaceCar(sReadAll1))%> </textarea></p>
		-->
		<%	 ' if clng(rsTabellaNew(6))=1 then  ' se la domanda è plus prima della spiegazione metto anche il testo prelvato dal file
			'		response.write(sReadAll1)
			 'end if
			' finito=false
			 'sReadAll1=sReadAll
'			 inizio=1
'			' do
'
'			  posApice=instr(inizio,sReadAll1,"'")
'			  if posApice=0 then
'			      finito=true
'			   else
'			     if  (strcomp (Mid(sReadAll1,posApice-1,1),"e")<>0) then ' se il carattere prcedente  ' non è una e allora sostituisco con "
'			         sReadAll1=replace(sReadAll1,"'",chr(34)&";",posApice)
'				 end if
'				 inizio=posApice+1
'			   end if
'			 ' response.write("Apice "& posApice)
'
'
'			' loop until finito=true

		 	' sReadAll1=replace(sReadAll1,"';",chr(34)&";")
			' sReadAll1=replace(sReadAll1,"<'","<"&chr(34))
			 ' sReadAll1= replace(sReadAll,"e"&chr(34),"e'")
'			 sReadAll1= replace(sReadAll,"o"&chr(34),"o'")
'			 sReadAll1= replace(sReadAll,"a"&chr(34),"a'")
'			 sReadAll1= replace(sReadAll,"u"&chr(34),"u'")
			' sReadAll1= replace(sReadAll,"'",chr(34))

			' ho spostato la chiusura della textarea sulla stessa riga dell'apertura perché in questo modo non compaiono gli spazi bianchi

			%>

         <%'else
		 %>

            <center>
             <%'response.write(lunghezza)

			 %>
           <!--   <iframe style="" src="<%=urliFrame%>"  width= 800px  height=<%=lunghezza*20%>px ></iframe>
              -->
          <%'end if
		  %>
 </center>
 <%'inserisco le eventuali immagini
if rsTabellaNew("Img")=1 then%>

 <%     QuerySQL1="Select * from Frasi_Img where Id_Frase="& rsTabellaNew("CodiceFrase")&";"
	   url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&Cartella&"/"&Modulo&"_Frasi/Img" ' vuole il percorso relativo della cartella
       url=Replace(url,"\","/")

	   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1)
   	   do while not rsTabella1.eof
	   'response.write(url&"/"& rsTabellaNew("Url")&"<br>")

	   urlimg=url&"/"& rsTabella1("Url") ' aggiungo al percorso il nome del file

  	 urldelete=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Cartella&"/"&Modulo&"_Frasi/Img/"&rsTabella1("Url")  ' per cancellare l'immagine.jpg
	   urldelete=Replace(urldelete,"\","/")

	   'response.write("urlimg="&urlimg)
	   %>

      <% nome=right(rsTabella1("Nome"),len(rsTabella1("Nome"))- instr(ucase(rsTabella1("Nome")),"C:\FAKEPATH\"))
		 nome=left(nome,len(nome)-4)
		 flag=0
		 %>

       <p align="center">
	   <%  gdoc="false"
	     if ((instr(rsTabella1("Url"),"docs.google.com")<>0) or (instr(rsTabella1("Url"),"drive.google.com")<>0) or (instr(rsTabella1("Url"),"colab.research.google.com")<>0) )  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
         gdoc="true"
		 response.write("<a href='"& rsTabella1("Url") &"' target='_blank'>apri url google drive</a>")
		  %>
	 <% else%>
		   <% if ((instr(rsTabella1("Url"),"tp://")<>0) or (instr(rsTabella1("Url"),"tps://")<>0)) and ((instr(rsTabella1("Url"),".jpg")<>0) or (instr(rsTabella1("Url"),".jpeg")<>0) or (instr(rsTabella1("Url"),".png")<>0) or (instr(rsTabella1("Url"),".gif")<>0))  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
			 'response.write(rsTabella1("Url"))
			 flag=1
			 %>
			 <a href="<%=rsTabella1("Url")%>" target="_blank"><img src="<%=rsTabella1("Url")%>" border="1"></a> <br>
		   <%else
				if (instr(rsTabella1("Url"),"tp://")<>0) or (instr(rsTabella1("Url"),"tps://")<>0) and ((instr(rsTabella1("Url"),".htm")<>0) or (instr(rsTabella1("Url"),".html")<>0) or (instr(rsTabella1("Url"),".php")<>0) )  then ' nb è voluto il tp:// invece di https:// perchp altrimenti essendo all'inizio restituisce 0 che è come se non fosse presente
					pagina=1
					flag=2
					%>  <a href="<%=rsTabella1("Url")%>" target="_blank"><%=rsTabella1("Url")%></a> <br><%
				else
		   flag=3
		  ' response.write("urlimg1="& urlimg)
		  %>

		   <img src="<%=urlimg%>" border="1"> <br>
		        <%end if%>
		   <%end if%>
		<%end if%>

      <%' response.write("flag="&flag)
	  %><br>
      <!--<a href="cancella_immagine.asp?urldb=<%=rsTabella1("Url")%>&urlimg=<%=urldelete%>&CodiceAllievo=<%=Session("CodiceAllievo")%>"><img src="../../img/elimina_small.jpg" width="10" height="10" title="Elimina" onClick="return window.confirm('Vuoi veramente cancellare questa immagine?');"></a></p>-->
    <%

	'pulisco sReadAll e sReadAll1


	rsTabella1.movenext
	   loop

%> <%
end if
%>

</div>

<%if (session("Admin")=true) then


%>
 <p><input class="input-mini" type="text" name="txtVAl<%=i%>" id="txtVAl<%=i%>" value="<%=rsTabellaNew.Fields("Voto")%>" size="1"  ><b>
	Valutazione </b> &nbsp;<i onclick="azzera('txtVAl<%=i%>',<%=i%>)" class="icon-remove"></i>

   <br>


       <span title="Feedback all'autore"><b><%=f%></b></span>

                                <% if (rsTabellaNew.Fields("Segnalata")<>0)  then%>
										        	 <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"  id="txtSegnalata<%=i%>" checked="true" value="1" onclick="segno_segnalazione(0,<%=i%>);">Si
                                <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"  value="0"  onclick="segno_segnalazione(1,<%=i%>);">No
                                <% else %>
                                <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"  id="txtSegnalata<%=i%>" value="1" onclick="segno_segnalazione(0,<%=i%>);">Si
                                <INPUT TYPE="RADIO" name="txtSegnalata<%=i%>"   checked="true" value="0" onclick="segno_segnalazione(1,<%=i%>);">No
									        	<% end if %><br>
                            <div id="divSegno<%=i%>"  style="display:none"><b>Segno</b>
                            <% if (rsTabellaNew.Fields("Segnalata")=2)  then%>
                           <INPUT TYPE="RADIO" name="txtSegno<%=i%>" checked="true" value="1">(+)
                            <INPUT TYPE="RADIO" name="txtSegno<%=i%>"  value="0">(-)
                            <% else %>
                            <INPUT TYPE="RADIO" name="txtSegno<%=i%>" value="1">(+)
                            <INPUT TYPE="RADIO" name="txtSegno<%=i%>"   checked="true" value="0">(-)
                         <% end if %>
                         </div><br>


										<textarea class="input-block-level" rows="2" cols="40" placeholder="Feedback allo studente" id="txtSegnalazione<%=i%>" name="txtSegnalazione<%=i%>"><%=feedback%></textarea>
      <p><input type="checkbox"  name="cb<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b>
	Seleziona per la stampa </b><br>
      <p><input type="checkbox"  name="cbVal<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna2('cbVal<%=i%>');">  <b>
	Seleziona per la valutazione </b><br>
	<!-- <input type="text" name="txtINQUIZ<%=i%>" value="<%=rsTabellaNew.Fields("In_Quiz")%>" size="1" ><b> In Quiz </b></p>
  <!--Definisce i due bottoni del form -->


  <% else
   if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) then %>

   <p><input type="checkbox"  name="cb<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cb<%=i%>');">  <b>
	Seleziona per la stampa </b><br>
	

	<!-- <input type="text" disabled="disabled" name="txtINQUIZ<%=i%>" value="<%=rsTabellaNew.Fields("In_Quiz")%>" size="1"><b> In Quiz </b> -->
	</p>
    <!--<p><input type="submit" value="Invia" name="B1"> </p>  Definisce i due bottoni del form -->

<% end if
end if


    i=i+1
	'response.write(i)
	%><br><%

	  consegnato=consegnato&"'"&rsTabellaNew("CodiceAllievo")&"'"&","
    rsTabellaNew.movenext
loop

  consegnato=left(consegnato,len(consegnato)-1)
'response.write("819")
if len(segnalati)>=1 then
  segnalati=left(segnalati,len(segnalati)-1)
end if
'response.write("821")

%>
	</div>
				</div>

</legend>


 


<%

if consegnato="" then
	response.write("<br>Nessuna consegna")
else%>
	
<%	q="select count(*) as NC from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") ;"
 ' response.write(q)
  set rsNC= ConnessioneDB.execute(q)
  noconsegne=rsNC("NC")
 
 %>
<hr> <b>Mancata consegna: (<%=noconsegne%>)</b>
 <% q="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and Attivo=1 and CodiceAllievo not in ("&consegnato&") order by Cognome;"
  'response.write(q)
  response.write("<br>")
  set rsTabellaNC= ConnessioneDB.execute(q)
  do while not rsTabellaNC.eof
  response.write(rsTabellaNC("Cognome")&" "&left(rsTabellaNC("Nome"),1)&".<br>")
  rsTabellaNC.movenext
  loop
end if

 if segnalati="" then
 	response.write("<br><span style='color:red'>Nessuna segnalazione negativa</span>")
 else
'   q="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and CodiceAllievo  in ("&segnalati&") order by Cognome;"
'   'response.write(q)
   response.write("<br><span style='color:red'><b>Segnalazioni(-):</span></b>")
'   set rsTabellaNC= ConnessioneDB.execute(q)
'   do while not rsTabellaNC.eof
'   response.write(rsTabellaNC("Cognome")&" "&left(rsTabellaNC("Nome"),1)&".<br>")
'   rsTabellaNC.movenext
'   loop
    response.write(segnalazioni)
 end if

 if segnalati_pos="" then
 	response.write("<br><span style='color:green'>Nessuna segnalazione positiva</span>")
 else
'   q="select Cognome,Nome,CodiceAllievo from Allievi where Id_Classe='"&id_classe&"' and CodiceAllievo  in ("&segnalati&") order by Cognome;"
'   'response.write(q)
   response.write("<br><span style='color:green'><b>Segnalazioni(+):</span></b>")

    response.write(segnalazioni_pos)
 end if


'response.write(consegnato)
%>



<hr>

<img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa(0);">
<%
if Session("Admin")=true then %>
  <a onClick="stampa(1);" title='Stampa consegna verifica'> 1</a>
<% end if

%>
&nbsp;
<b>Stampa <input class="input-mini" type="text" name="txtNUMREC" value="<%=i-1%>">Frasi</b></p>
<input class="btn" type="button" value="Seleziona tutti" onClick="checkTutti()">
<input  class="btn" type="button" value="Deseleziona tutti" onClick="uncheckTutti()"><br><hr>
 <% if Session("Admin")=true then%>
<b>Voto</b><input class="input-mini" type="text"   name="txtVoto">
<input type="button" class="btn" value="Valuta tutti" onClick="valutaTutti()">
<input type="text" name="txtNUMVAL" value="<%=i-1%>" size="1" class="input-mini"><hr>
<input type="button" class="btn" value="Segnala tutti" onClick="segnalaTutti()">
<input type="text" name="txtFeedback" id="txtFeedback" size="1" class="input-xlarge">
<input type="button" class="btn" value="Desegnala tutti" onClick="desegnalaTutti()"><br>
<input type="checkbox"  name="txtNotifica" id="txtNotifica" >  
Non inviare notifica dei feedback <hr>


 <br>
  <i class="icon-calendar"></i> <b>Consegnati</b>
    <input type="text"  name="date3" id="datepicker" class="input-small datepick" /></p>
  <input class="btn" type="button" value="Tutti" title="Attribuisci a tutti la stessa data di consegna" onClick="selezionatutti('datepicker')">

  <hr>
<input type="submit" value="Invia" name="B1" class="btn-primary"> </p>
<HR>

<%end if%>

</form>

















                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->


		</div> <!--fine main-->
        </div>




	</body>
    <% else%>
<BODY onLoad="showText();"> </BODY>
  <% ' torna all'homepage
  ' Response.Redirect "studente_domande.asp?cla="&cla
   end if


 '  objCreatedFile.Close


'response.write("<script language='javascript' type='text/javascript'> $(document).ready(function() { document.getElementById('segnalazioni').innerHTML='"&Server.HTMLEncode(segnalazioni)&"'; });</script>")


    %>
 <script language="javascript" type="text/javascript">

function stampa(tipo) {
	//tipo 1 per stampare la consegna di una verifica (Sintesi e verifica -> = -> Consegna verifica)
	if (tipo == 0)
	document.dati.action = "7_stampa_schede_frasi_elenco_sint.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&CodiceTest=<%=Request.QueryString("ID_PAR")%>&tutto=0&FraseScelta=<%=idprefrase%>";
	else
	  document.dati.action = "7_stampa_schede_frasi_elenco_una.asp?CodiceAllievo=<%=CodiceAllievo%>&Modulo=<%=Modulo%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Cartella=<%=Cartella%>&CodiceTest=<%=Request.QueryString("ID_PAR")%>&tutto=0&FraseScelta=<%=idprefrase%>";
	
	//&QuerySQL=<%=QueryPrima%>

		//document.dati.action = "../../home.asp"
		document.dati.submit();
}

function segno_segnalazione(s,id)
{
if (s==0)
 document.getElementById("divSegno"+id).style.display='block';
 else
   document.getElementById("divSegno"+id).style.display='none';
}


function azzera(ids,idx) {

   //var i = document.getElementById(ids).value;
   //document.getElementById(ids).value = parseInt(i) + 1;

	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
			if (elements[i].id == 'txtVAl'+idx)
		    {
		      
			 elements[i].value=0;
			 
			 
			}
			if (elements[i].id == 'txtSegnalata'+idx)
		    {
		     
			 elements[i].checked=true;
			 segno_segnalazione(0,idx);
			 
			 
			}	
			if (elements[i].id == 'txtSegnalazione'+idx)
		    {
		  
			 elements[i].value="Penalizzazione per ritardata consegna";
		 
			idx=idx+1;
			}	
		}
		
	}
}


 </script>

 </html>
