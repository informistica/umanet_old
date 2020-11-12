<%@ Language=VBScript %>
<%

Response.AddHeader "Refresh", "500"



'Response.Buffer = true
 daAvviso=request.QueryString("daAvviso") ' mi serve per sapere se devo aprire il div per indicare i compiti assegnati
   'dividApro=cint(request.QueryString("dividApro"))
   dividA=request.QueryString("dividApro")
   capApro=request.QueryString("capApro") 'identifica la section del capitolo da aprire
   sottocapApro=request.QueryString("sottoCapApro") 'identifica la section del sottoparagrafo da aprire
   dividApro= right(dividA,len(dividA)- instr(dividA,"#"))
 sottocapApro = Replace(sottocapApro,".","")

   id_classe=request.QueryString("id_classe")
   cartella=request.QueryString("cartella")
    id_materia=request.QueryString("id_materia")
	visibili=request.querystring("visibili")
   Session("Id_Classe")=id_classe
   Session("Cartella")=cartella
    Session("cartella")=cartella


%>
<!doctype html>
<html>
<head>
 <meta charset="UTF-8">
<script src="../js/google.js"></script>

 <script language="javascript" type="text/javascript">

function closeAction() {

	  if (confirm("Vuoi uscire ?")) {
        window.close();
	 }
	 else
	 return 0;
}
</script>
<!--	<meta charset="utf-8">-->

    <title>Libro</title>
    <%
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	%>
 <!-- #include file = "../include/header.asp" -->
 <!-- #include file = "../var_globali.inc" -->
 <!-- #include file = "../extra/test_server.asp" -->
 <!-- #include file = "../include/cambia_sessione.asp" -->
 <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->


<style>

a:hover {
	text-decoration:none;
}

</style>
<link rel="shortcut icon" href="../../favicon.ico" />
</head>

<body onUnLoad="alert 'ciao';" class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

<% Dim box_apri
  ' box_apri="toggleBersaglio"&dividApro
  box_apri=dividApro

  if box_apri="" then
    box_apri="0"
  end if

 ' box_apri=1
 ' box_apri2="collapseTrenew10"
  ' box_apri1="toggleBersaglionew10" ' apre il titolo del sottoparagrafo
  ' box_apri2="Naviga10"	' apre il tab di navigazione

%>
<%
function ReplaceCar(sInput)
dim sAns

  sAns=  Replace(sInput,"è","&egrave;")
  sAns=  Replace(sAns,"ì","&igrave;")
  sAns=  Replace(sAns,"ù","&ugrave;")
  sAns=  Replace(sAns,"ò","&ograve;")
  sAns=  Replace(sAns,"à","&agrave;")
  sAns=  Replace(sAns,"'",Chr(96))

ReplaceCar = sAns
end function

%>
	<div id="navigation">
     <!-- #include file = "../service/controllo_sessione.asp" -->
		  <!-- #include file = "../include/navigation.asp" -->

	</div>




	<div class="container-fluid" id="content">
      <!-- #include file = "../include/menu_left.asp" -->

       <%

	    QuerySQL="Select * from Setting where Id_Classe='" & id_classe &"'"
 Set rsTabellaSetting = ConnessioneDB.Execute(QuerySQL)
 ValidaQuiz=rsTabellaSetting("ValidaQuiz")

	   if session("CodiceAllievo")<>"" then
	  ' se la classe ? appena inserita non ha moduli quindi rimando ad admin per aggiungerli
					  QuerySQL="SELECT Classe from Classi where Id_Classe='"&id_classe&"';"
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					Classe=rsTabella("Classe")

					 QuerySQL="SELECT Stile from Allievi where CodiceAllievo='"&Session("CodiceAllievo")&"';"
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					stile=rsTabella("Stile")
					Session("stile")=stile
					'stile="pink"

					 QuerySQL="SELECT count(*) from MODULI_NOT_UMANET where Id_Classe='"&id_classe&"';"
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					numMod=rsTabella(0)





					if numMod=0 then
					     if   session("Admin")=true then
					         response.Redirect "../cAdmin/admin.asp?dividApro=config_moduli&Id_Classe="&id_classe&"&divid="&divid&"&idmsg=1"

						else
						      response.Redirect request.serverVariables("HTTP_REFERER")
						end if


					else

					    if visibili="" then
						QuerySQL="SELECT Id_Classe, Titolo, TitPar, ID_Mod, ID_Paragrafo,Cartella,URL,URL_OL,Classe,URL_L,URL_O,Posizione from MODULI_NOT_UMANET  where Id_Classe='"&id_classe&"' and Visibile=1 order by PosMod, PosPar ;"

						else
						  QuerySQL="SELECT Id_Classe, Titolo, TitPar, ID_Mod, ID_Paragrafo,Cartella,URL,URL_OL,Classe,URL_L,URL_O,Posizione from MODULI_NOT_UMANET  where Id_Classe='"&id_classe&"' order by PosMod, PosPar ;"

						end if
                    Set rsTabella = ConnessioneDB.Execute(QuerySQL)
					end if



					 dim objFSO,objCreatedFile
					Const ForReading = 1, ForWriting = 2, ForAppending = 8
					Dim sRead, sReadLine, sReadAll, objTextFile
				    Set objFSO = CreateObject("Scripting.FileSystemObject")

				  url="C:\inetpub\umanetroot\expo2015Server\UECDL\piani_di_lavoro\piano_" & left(Classe,1+len(Classe)-instr(Classe,"$"))& ".txt"
				  Set objCreatedFile = objFSO.CreateTextFile(url, True)



				'objCreatedFile.WriteLine(QuerySQL)


	  %>

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
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#"><b>Libro</b></a>

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
								<h3>
									<i class="icon-reorder"></i>
									Libro classe <%=left(Classe,1+len(Classe)-instr(Classe,"$"))%>
                  <%response.write("<br>"&url)%>
									<% if visibili="" then %>
									<a href="home_app.asp?id_classe=<%=id_classe%>&cartella=<%=cartella%>&id_materia=<%=id_materia%>&visibili=1"><font size="-1"><i class="icon-eye-close" title="Mostra i moduli degli anni scorsi" ></i></a></font>
									<%else%>
									<a href="home_app.asp?id_classe=<%=id_classe%>&cartella=<%=cartella%>&id_materia=<%=id_materia%>&visibili="><font size="-1"><i class="icon-eye-open" title="Nascondi i moduli degli anni scorsi" ></i></a></font>
									<%end if%>
								</h3>
							</div>
							<div class="box-content">

							  <div class="accordion accordion-widget" id="accordionContenitore">


<%
i=0
k=1 ' conta i moduli inseriti mi serve come indice per le ancore al modulo dal quaderno
capitolo=rsTabella(1)
iddiv=1
pospar=1 ' posizione del paragrafo all'interno del modulo serve per individuare il box da aprire
rsTabella.movefirst

do while not rsTabella.eof


if (i=0) then ' Titolo del Modulo
%>

 <div class="accordion-group">
  <div class="accordion-heading">
    <a id="cap<%=k%>" class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordionContenitore" href="#capitolo<%=k%>" <% if session("Admin")=true then %> title="<%=k%>" <%end if%>>
	  <%=k &") " %>	<%=ReplaceCar(rsTabella(1))%>
	   <%objCreatedFile.WriteLine()%>
	  <%objCreatedFile.WriteLine("Capitolo:"&ReplaceCar(rsTabella(1)))%>
	</a>
	</div>
	 <div id="capitolo<%=k%>" class="accordion-body collapse">
	   <div class="accordion-inner">

   <section id="collapse">


         <%if (instr(rsTabella("URL_OL"),"https")<>0) or (instr(rsTabella("URL_OL"),"http")<>0) then %>
            <a  title="Risorse introduttive" href="<%=rsTabella("URL_OL")%>" target="_blank"><i class="icon-cloud"></i>&nbsp;&nbsp;-&nbsp; </a></a>
		<%else%>
		    <%if rsTabella("URL")<>"" then %>
            <% riferimento=homesito& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" &rsTabella("Cartella")%>/Risorse/Mod_<%=right(rsTabella(3),len(rsTabella(3))-instr(rsTabella(3),"_"))%>/<%=rsTabella("URL_O")%>
            <%else
			 riferimento="#"
			 end if %>
             <a  title="Risorse introduttive" href="<%=riferimento%>" target="_blank"> <i class="icon-cloud"></i>&nbsp;&nbsp;-&nbsp;</a>
         <%end if %>
         <% if ValidaQuiz=1 then%>
          <a  title="Apprendimento del Capitolo" target="_blank" href="scegli_azione_app.asp?Cartella=<%=rsTabella.fields("Cartella")%>&Stato=1&Stato0=1&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Modulo=<%=rsTabella(3)%>" ><i class="icon-book"></i>  </a>
          &nbsp;&nbsp;- <%end if%>
            <a title="Mettiti alla prova su tutto il capitolo" target="_blank" href="scegli_azione_test.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella("Cartella")%>&Stato=1&Stato0=1&Tutti=1&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;&nbsp;<i class="icon-edit"></i>

			<% if Session("Admin") = true then %>
			&nbsp;&nbsp;&nbsp;-
			<a data-placement="right"   rel="tooltip" title="Valuta Frasi" href="../cFrasi/2scegli_valutazioni_frasi.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&tutto=1"><%=response.write(" (=")%></a>
				&nbsp; <i class="icon-reply"></i>
            <a data-placement="right"   rel="tooltip" title="Modifica Frasi" href="../cFrasi/2modificaprefrase.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&Capitolo=<%=rsTabella(1)%>&Modulo=<%=rsTabella(3)%>&tutto=1"><%=response.write(" m) ")%></a>&nbsp;

			&nbsp;-

			<a data-placement="right"   rel="tooltip" title="Valuta Domande" href="../cDomande/inserisci_valutazioni.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&iddiv=<%=iddiv%>&tutto=1">&nbsp;<%=response.write("( = ")%>&nbsp;</a>

             <i class="icon-question-sign"></i>
             <a data-placement="right"   rel="tooltip" title="Modifica Domande" href="../cDomande/modificapredomande.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&iddiv=<%=iddiv%>&tutto=1"><%=response.write(" m")%>)&nbsp;</a>

			 <% end if %>




     <!-- Se ho appena scritto il titolo inizio una nuova bs-docs-example per contenere tutti i id="accordion<=iddiv> dei paragrafi -->
      <div class="bs-docs-example"> <!-- Cornice esterna che contiene Argomenti-->
              <div class="accordion-widget" id="accordion<%=iddiv%>">

 <%end if %>

  <%' adesso vedo se il paragrafo ha sottoparagrafi, se non ne ha metto il solito R F D N
		  ' altrimenti metto il titolo del sottoparagrafo dentro il quale mettero R F D N
 QuerySQL="SELECT  * from ParagrafiSottoparagrafi2  where Id_Paragrafo='"&rsTabella("ID_Paragrafo") &"';"
   Set rsTabellaSottopar = ConnessioneDB.Execute(QuerySQL)
	 if not rsTabellaSottopar.eof then %>




                  <div class="accordion-group">
                  <div class="accordion-heading">
                    <a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordionnew<%=k%>" href="#collapsenew<%=iddiv%>" id="par<%=iddiv%>" <% if session("Admin")=true then %> title="<%=iddiv%>" <%end if%>>
                         <%=i+1%>.&nbsp;<span title="Paragrafo"><%=ReplaceCar(rsTabella(2))%>  </span>
						  <%objCreatedFile.WriteLine("Capitolo:"&ReplaceCar(rsTabella(2)))%>
                    </a>

                  </div>
                  <div id="collapsenew<%=iddiv%>" class="accordion-body collapse">

                    <div class="accordion-inner">

                  <a   rel="popover" data-trigger="hover" data-content="Apri pagina del libro"   title="Leggi Risorsa paragrafo(R)" href="<%=rsTabella("URL_O")%>" target="_blank">&nbsp;<i class="icon-cloud"></i></a>&nbsp;&nbsp;&nbsp;


                <% if ValidaQuiz=1 then%>
                    <a  rel="popover" data-trigger="hover" data-content="Leggi e vota Frasi,Domande,Nodi" title="Apprendimento del Paragrafo" target="_blank"    name="<%=iddiv%>" href="scegli_azione_app.asp?Cartella=<%=rsTabella.fields("Cartella")%>&Stato=0&Stato0=0&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;&nbsp;<i class="icon-book"></i> &nbsp;
					</a> <%end if%>
                     &nbsp;

                      <a rel="popover" data-trigger="hover" data-content="Crea o svolgi Quiz" title="Mettiti alla prova  sul Paragrafo" target="_blank" href="scegli_azione_test.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella("Cartella")%>&Stato=0&Stato0=0&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;&nbsp;<i class="icon-edit"></i>


				  <!--
				  <br><br>

                  </b>&nbsp;  <a rel="popover" data-trigger="hover" data-content="Crea frase utilizzando parole chiave" title="Rispondi con una frase (F)" target=blank href="../cFrasi/2compilaprefrase.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prefrase=1&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-reply"></i></span></a>&nbsp;&nbsp; &nbsp;
				 &nbsp;&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;<a rel="popover" data-trigger="hover" data-content="Crea nodo della rete concettuale"  Title="Compila Nodo (N)" target=blank href="../cNodi/1compilaprenodo.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prenodo=1&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>" ><span style="text-transform: uppercase;"> <i class="glyphicon-snowflake"></i></span></a>&nbsp;&nbsp;&nbsp;&nbsp;
                    <a rel="popover" data-trigger="hover" data-content="Crea quiz"   title="Svolgi Domanda (D)" target=blank href="../cDomande/compilapredomanda.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&predomanda=1"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-question-sign"></i></span></a>

				  -->



          <% p=0
            do while not rsTabellaSottopar.eof
                        %>
                  <div class="accordion-group">
                     <div class="accordion-heading">
                            <a class="accordion-toggle collapsed " data-toggle="collapse" id="Sottoparagrafo<%=i+1%><%=p+1%>" data-parent="#accordionnew<%=k%><%=p%>" href="#collapseTrenew<%=iddiv%><%=k%><%=p%>" >
                            <%=i+1%>.<%=p+1%>&nbsp;<span title="Sottoparagrafo<%=i+1%><%=p+1%>"><%=ReplaceCar(rsTabellaSottopar("Titolo"))%></span>
                            </a>
							 <%objCreatedFile.WriteLine(ReplaceCar(rsTabellaSottopar("Titolo")))%>

                          </div><!-- fine <div class="accordion-heading"> -->

                          <div id="collapseTrenew<%=iddiv%><%=k%><%=p%>" class="accordion-body collapse">

						<ul id="myTab22" class="nav nav-tabs">
                              <li  class="active" ><a href="#home2<%=iddiv%><%=p%>" data-toggle="tab">Compiti</a></li>
                              <li><a href="#profile2<%=iddiv%><%=k%><%=p%>" data-toggle="tab" id="Naviga<%=k%><%=p%>">Naviga</a></li>
                          </ul>
                          <div id="myTabContent22" class="tab-content">
                              <div class="tab-pane fade in active" id="home2<%=iddiv%><%=p%>">
                                <p><b>
								<a href="#" data-rel="tooltip" data-placement="bottom" title="Tooltip on bottom">
                                 </b>
                                 <% if rsTabellaSottopar("URL")<>"" then %>
                                     <%if (instr(rsTabellaSottopar("URL"),"https")<>0) or (instr(rsTabellaSottopar("URL"),"http")<>0) then %>
                                         <a rel="popover" data-trigger="hover" data-content="Apri pagine del libro"  title="Leggi Risorsa Sottoparagrafo(R)" href="<%=rsTabellaSottopar.fields("URL")%>" target="_blank">&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; <i class="icon-cloud"></i></a>&nbsp; &nbsp;
                                      <% else%>
            							<a rel="popover" data-trigger="hover" data-content="Apri pagine del libro"  title="Leggi Risorsa Sottoparagrafo (R)" href="<%=homesito& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/Risorse/Mod_" &  right(rsTabella(3),len(rsTabella(3))-instr(rsTabella(3),"_")) &"/"& rsTabellaSottopar.fields("URL")%>" target="_blank">&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; <i class="icon-cloud"></i> </a>&nbsp;
        						      <%end if%>
								<%end if%>




                  </b>&nbsp;  <a rel="popover" data-trigger="hover" data-content="Crea frase utilizzando parole chiave" title="Rispondi con una frase (F)" target=blank href="../cFrasi/2compilaprefrase.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prefrase=1&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-reply"></i></span></a>&nbsp;&nbsp; &nbsp;



				 &nbsp;&nbsp; &nbsp; &nbsp;&nbsp; &nbsp; &nbsp;&nbsp;<a rel="popover" data-trigger="hover" data-content="Crea nodo della rete concettuale"  Title="Compila Nodo (N)" target=blank href="../cNodi/1compilaprenodo.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prenodo=1&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>" ><span style="text-transform: uppercase;"> <i class="glyphicon-snowflake"></i></span></a>&nbsp;&nbsp;&nbsp;&nbsp;




                    <a rel="popover" data-trigger="hover" data-content="Crea quiz"   title="Svolgi Domanda (D)" target=blank href="../cDomande/compilapredomanda.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&predomanda=1"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-question-sign"></i></span></a>



						<%if (session("Admin")=true) then %>
                    <p></p> <p></p>

                     <% if rsTabellaSottopar("URL")<>"" then %>
                     &nbsp; &nbsp; &nbsp;&nbsp;
                   <a data-placement="right"   rel="tooltip" title="Modifica Risorse"  href="../cAdmin/modificamodulo.asp?ID_Mod=<%=rsTabella.fields("ID_Mod")%>&Classe=<%=Classe%>&Id_Classe=<%=Id_Classe%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>">(m)</a>
                     <% end if%>
                      &nbsp;&nbsp;&nbsp;
						<a data-placement="right"   rel="tooltip"  title="Aggiungi preFrasi" href="../cFrasi/2inserisci_prefrase.asp?Segnalibro=<%=k%>&BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>">(<%=response.write(" + ")%>&nbsp;</a>

						<a data-placement="right"   rel="tooltip" title="Valuta Frasi" href="../cFrasi/2scegli_valutazioni_frasi.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>"><%=response.write(" = ")%>&nbsp;</a>

                    <a data-placement="right"   rel="tooltip" title="Modifica Frasi" href="../cFrasi/2modificaprefrase.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>"><%=response.write(" m ")%>)</a>&nbsp;



                            &nbsp;
						<a data-placement="right"   rel="tooltip"  Title="Aggiungi preNodi" href="../cNodi/1inserisci_prenodo.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>">(<%=response.write("+")%>&nbsp;</a>

							<a data-placement="right"   rel="tooltip"  Title="Valuta Nodi" href="../cNodi/2scegli_valutazioni_nodi.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>"><%=response.write("=")%>&nbsp;</a>
					 	 <a data-placement="right"   rel="tooltip" title="Modifica Nodi" href="../cNodi/2modificaprenodo.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>"><%=response.write(" m ")%>)&nbsp;</a>

                     <a data-placement="right"   rel="tooltip" title="Aggiungi preDomande" href="../cDomande/inserisci_predomande.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>">(&nbsp;<%=response.write("+")%>&nbsp;</a>

							<a data-placement="right"   rel="tooltip" title="Valuta Domande" href="../cDomande/inserisci_valutazioni.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>">&nbsp;<%=response.write("=")%>&nbsp;</a>

                              <a data-placement="right"   rel="tooltip" title="Modifica Domande" href="../cDomande/modificapredomande.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottopar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>&iddiv=<%=iddiv%>"><%=response.write(" m ")%>)&nbsp;</a>


					<% end if %>



                                </b></p>
                                <br>
                              </div>
                              <div class="tab-pane fade  " id="profile2<%=iddiv%><%=k%><%=p%>">
                                <p><% if ValidaQuiz=1 then%>
                                 <a rel="popover" data-trigger="hover" data-content="Leggi e vota Frasi,Domande,Nodi" title="Apprendimento del Sottoparagrafo"  name="<%=iddiv%>" target="_blank" href="scegli_azione_app.asp?Cartella=<%=rsTabella.fields("Cartella")%>&Stato=2&Stato0=2&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottoPar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>">&nbsp;&nbsp;<i class="icon-book"></i>Apprendimento &nbsp;
					</a>
                    - <%end if%>
                    <a rel="popover" data-trigger="hover" title="Mettiti alla prova sul Sottoparagrafo"    data-content="Crea o svolgi Quiz" target="_blank" href="scegli_azione_test.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella("Cartella")%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&Sottoparagrafo=<%=rsTabellaSottopar("Titolo")%>&CodiceSottoPar=<%=rsTabellaSottopar("ID_Sottoparagrafo")%>">&nbsp;&nbsp;<i class="icon-edit"></i>Verifica
							</a>
                                </p>
                                <p>
                                </p><br>
                              </div>
                         </div>




                          </div><!-- fine collapse(treuno)-->
                        </div> <!-- fine accordino group-- da Descrizione capitolo in giù >-->

                         <% p=p+1
						   rsTabellaSottopar.movenext()
						   Loop
						%>




                    </div><!-- fine accordion inner-->
                  </div>
                </div>




	 <% else ' non ha sottoparagrafi
	 %>

                    <!--Inizia un nuovo paragrafo -->
                  <div class="accordion-group"> <!--Contenitore per il titolo del paragrafo-->
                  <div class="accordion-heading"> <!-- titolo del paragrafo-->
                    <a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion<%=iddiv%>" href="#collapse<%=iddiv%>" name="<%=iddiv%>" id="par<%=iddiv%>" <% if session("Admin")=true then %> title="<%=iddiv%>"<%end if%> >
                       <%=i+1%>. <%=ReplaceCar(rsTabella(2))%>
					    <%objCreatedFile.WriteLine(ReplaceCar(rsTabella(2)))%>

                    </a>
                  </div>



                  <div id="collapse<%=iddiv%>" class="accordion-body collapse">
                           <ul id="myTab2" class="nav nav-tabs">
                              <li class="active" ><a href="#home<%=iddiv%>" data-toggle="tab">Compiti</a></li>
                              <li ><a href="#profile<%=iddiv%>" id="Naviga2<%=iddiv%>" data-toggle="tab">Naviga</a></li>
                          </ul>
                          <div id="myTabContent2" class="tab-content">
                              <div class="tab-pane fade in active " id="home<%=iddiv%>">
                                <p><b>
<a href="#" data-rel="tooltip" data-placement="bottom" title="Tooltip on bottom">
                                 </b>
                                      <%if (instr(rsTabella("URL_O"),"https")<>0) or (instr(rsTabella("URL_O"),"http")<>0) then %>
                                         <a rel="popover" data-trigger="hover" data-content="Apri pagine del libro"  title="Leggi Risorsa (R)" href="<%=rsTabella.fields("URL_O")%>" target="_blank">&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; <i class="icon-cloud"></i></a>&nbsp; &nbsp;
                                      <% else%>
            							<a rel="popover" data-trigger="hover" data-content="Apri pagine del libro"  title="Leggi Risorsa (R)" href="<%=homesito& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&cartella&"/Risorse/Mod_" &  right(rsTabella(3),len(rsTabella(3))-instr(rsTabella(3),"_")) &"/"& rsTabella.fields("URL_O")%>" target="_blank">&nbsp;&nbsp;&nbsp; &nbsp; &nbsp; <i class="icon-cloud"></i> </a>&nbsp;
        						      <%end if%>





                  </b>&nbsp;  <a rel="popover" data-trigger="hover" data-content="Crea frase utilizzando parole chiave" title="Rispondi con una frase (F)" target=blank href="../cFrasi/2compilaprefrase.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prefrase=1"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-reply"></i></span></a>&nbsp;&nbsp; &nbsp; &nbsp;&nbsp; &nbsp;&nbsp;&nbsp; &nbsp;
                    <a rel="popover" data-trigger="hover" data-content="Crea nodo della rete concettuale"  Title="Compila Nodo (N)" target=blank href="../cNodi/1compilaprenodo.asp?Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&prenodo=1" ><span style="text-transform: uppercase;"> <i class="glyphicon-snowflake"></i></span></a>
                     &nbsp;&nbsp; &nbsp;
                   <a rel="popover" data-trigger="hover" data-content="Crea quiz"   title="Svolgi Domanda (D)" target=blank href="../cDomande/compilapredomanda.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella(5)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>&CodiceTest=<%=rsTabella(4)%>&predomanda=1"><span style="text-transform: uppercase;">&nbsp;&nbsp; &nbsp; &nbsp; &nbsp;  <i class="icon-question-sign"></i></span></a>



						<%if (session("Admin")=true) then %>
                    <p></p> <p></p>
					&nbsp; &nbsp; &nbsp;&nbsp;
                   <a data-placement="right"   rel="tooltip" title="Modifica Risorse"  href="../cAdmin/modificamodulo.asp?ID_Mod=<%=rsTabella.fields("ID_Mod")%>&Classe=<%=Classe%>&Id_Classe=<%=Id_Classe%>&divid=<%=divid%>">(m)</a>

                      &nbsp;&nbsp;&nbsp;
						<a data-placement="right"   rel="tooltip"  title="Aggiungi preFrasi" href="../cFrasi/2inserisci_prefrase.asp?Segnalibro=<%=k%>&BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">(<%=response.write(" + ")%>&nbsp;</a>

						<a data-placement="right"   rel="tooltip" title="Valuta Frasi" href="../cFrasi/2scegli_valutazioni_frasi.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>"><%=response.write(" = ")%>&nbsp;</a>

                    <a data-placement="right"   rel="tooltip" title="Modifica Frasi" href="../cFrasi/2modificaprefrase.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>"><%=response.write(" m ")%>)</a>&nbsp;



						<a data-placement="right"   rel="tooltip"  Title="Aggiungi preNodi" href="../cNodi/1inserisci_prenodo.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">(<%=response.write("+")%>&nbsp;</a>

							<a data-placement="right"   rel="tooltip"  Title="Valuta Nodi" href="../cNodi/2scegli_valutazioni_nodi.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>"><%=response.write("=")%>&nbsp;</a>
					 	 <a data-placement="right"   rel="tooltip" title="Modifica Nodi" href="../cNodi/2modificaprenodo.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>"><%=response.write(" m ")%>)&nbsp;</a>

                         <a data-placement="right"   rel="tooltip" title="Aggiungi preDomande" href="../cDomande/inserisci_predomande.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">(&nbsp;<%=response.write("+")%>&nbsp;</a>

							<a data-placement="right"   rel="tooltip" title="Valuta Domande" href="../cDomande/inserisci_valutazioni.asp?BoxApro=<%=iddiv%>&Cartella=<%=rsTabella(5)%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(4)%>&TitoloParagrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;<%=response.write("=")%>&nbsp;</a>

                              <a data-placement="right"   rel="tooltip" title="Modifica Domande" href="../cDomande/modificapredomande.asp?BoxApro=<%=iddiv%>&Id_Stud=<%=Id_Stud%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>"><%=response.write(" m ")%>)&nbsp;</a>

                            &nbsp;



					<% end if %>



                                </b></p>
                                <br>
                              </div>
                              <div class="tab-pane fade " id="profile<%=iddiv%>">
                                <p><% if ValidaQuiz=1 then%>
                                 <a rel="popover" data-trigger="hover" data-content="Leggi e vota Frasi,Domande,Nodi" title="Apprendimento del paragrafo"  name="<%=iddiv%>" href="scegli_azione_app.asp?Cartella=<%=rsTabella.fields("Cartella")%>&Stato=0&Stato0=0&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;&nbsp;<i class="icon-book"></i>Apprendimento &nbsp;
					</a>
                    -<%end if%>
                    <a rel="popover" data-trigger="hover" title="Crea e svolgi Quiz"   data-content="Verifica apprendimento del paragrafo" href="scegli_azione_test.asp?id_classe=<%=id_classe%>&Cartella=<%=rsTabella("Cartella")%>&CodiceTest=<%=rsTabella(4)%>&Capitolo=<%=rsTabella(1)%>&Paragrafo=<%=rsTabella(2)%>&Modulo=<%=rsTabella(3)%>">&nbsp;&nbsp;<i class="icon-edit"></i>Verifica
							</a>
                                </p>
                                <p>
                                </p><br>
                              </div>
                         </div>

                  </div>




                </div>

                 <% end if ' if not rsTabellaSottopar.eof then
				 %>

<%


	i=i+1
	iddiv=iddiv+1
	capitolo=rsTabella(1)
	rsTabella.movenext




	if not rsTabella.eof then
		c=rsTabella(1)
	  '  response.write(capitolo & " " & c)
			    if StrComp(capitolo, c) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")
                   else
                    i=0
					k=k+1 ' conta i moduli inseriti mi serve come indice per le ancore al modulo dal quaderno
                   ' Response.Write("Le due stringhe sono diverse")
			       %>
                   </section>
				   </div>
				</div>
				</div>
				  <%
                end if
         end if
		loop
		set esecuzione=nothing ' libero l'oggetto
		%>


          <%else%>


          <%end if%>

<%
objCreatedFile.Close
response.write("<script>alert('Creato il file:"&url&"');</script>")
'response.redirect url
%>







							</div>
						</div>
					</div>
				</div>





			</div>


		</div> <!--fine main-->
        </div>
        <script type="text/javascript" src="../../js/personalizza.js"></script>
		<script type="text/javascript">



$(window).load(function () {
/*alert("#Sottoparagrafo<%=sottocapApro%>");*/


	   $('#cap<%=capApro%>').click();
	   $('#par<%=box_apri%>').click();
	   $('#Sottoparagrafo<%=sottocapApro%>').click();



	 /* $('#Sottoparagrafo4.3').click();
	  $('#<%=box_apri%>').click();
	   $('#<%=box_apri1%>').click();
	   $('#<%=box_apri2%>').click();
	    */
	   $("body").addClass("theme-"+"<%=stile%>").attr("data-theme","theme-"+"<%=stile%>");



	  // event.stopPropagation();

	});


/*$(".red").click(function(event){

   // alert("Hai cliccato sull'Elemento");
	document.location = "script/aggiorna_stile.asp?stile=red"
});
*/

</script>
	</body>

	</html>
