<%@ Language=VBScript %>
<!doctype html>
<html>
<head>

   <title>Valuta metafore</title>

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

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../img/apple-touch-icon-precomposed.png" />


       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
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

  Cartella=Request.QueryString("Cartella")
  Capitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  TitoloParagrafo=Request.QueryString("TitoloParagrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
  Cognome=Session("Cognome")
  Nome=Session("Nome")
  CodiceSottopar = Request.QueryString("CodiceSottopar")
   BoxApro=Request.QueryString("BoxApro")
  soloimg=Request.QueryString("soloimg")
'--------------
  'Data=Request.Form("txtDATA")
  Nulle=Request.QueryString("Nulle") ' per selezionare solo le domande ancora da valutare con valutazione=0
  CodiceAllievo=Request.QueryString("CodiceAllievo")
  ID_MOD=Request.QueryString("Modulo")
  Tutte=Request.QueryString("Tutte") ' vale 1 se devo visualizzare tutte le domande  dello studente
  if left(Cartella,1)<>"" then
     Classe=clng(left(Request.QueryString("Cartella"),1))
  end if
'---------

 %>


	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-reply"></i> Valuta metafore</h1>

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
							<a href="#">Valuta metafore</a>

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
				        <h3> <i class="icon-reorder"></i> <%=Capitolo%>: <% if soloimg="" then response.write(TitoloParagrafo) end if%>
                  &nbsp;<a href="#modal-1" title="Interroga"  data-toggle="modal" onclick="uncheckTutti();"><i class="icon-question-sign"></i></a>
                </h3>
			          </div>
				      <div class="box-content">




	<%


Select Case CodiceTest%>
   <% Case Cartella&"_U_2_3" 'Topolino%>
    <% Case Cartella&"_U_2_5" 'Navigazione%>
<%
	QuerySQL="SELECT * " &_
"FROM preNavigazione WHERE  Id_Paragrafo='" & CodiceTest & "' and Id_Mod='"&ID_MOD & "' "

%>

	 <% Case Cartella&"_U_2_8" 'Navigazione 
	' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
  end Select


	 
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
'response.write(QuerySql & " " &Paragrafo )

if rsTabella.eof then%>
<span class="alert-error">
Non ci sono compiti da valutare<br>

</span>
<a href="javascript:history.back()">	Indietro </a>
<%else%>

<ol>
<%
i=0
'paragrafo=rsTabella(2)

 
'QuerySQL="SELECT Id_Classe, Titolo, TitPar, ID_Mod, ID_Paragrafo,Cartella,URL,URL_OL,Classe,URL_L,URL_O,Posizione from MODULI_UMANET1  where Id_Classe='"&id_classe&"' order by PosMod, PosPar ;"
					'response.write(QuerySQL)
 '                   Set rsTabella = ConnessioneDB.Execute(QuerySQL)

do while not rsTabella.eof
	'if (i=0) or (StrComp(capitolo, rsTabella(0)) <> 0) then'

	 %>
			  <% if rsTabella("img")=1  then
   				      image="  <i class='icon-picture' title='richiede immagine'></i>"
					  else
					  image=""
					 end if %>

					<li>
 
					<a href="sintesi_metafore_classe.asp?Id_Premetafora=<%=rsTabella("ID_Premetafora")%>&id_classe=<%=id_classe%>&Cartella=<%=Cartella%>&classe=<%=classe%>&CodiceTest=<%=CodiceTest%>&Modulo=<%=Modulo%>&Paragrafo=<%=Paragrafo%>"><%=image%>&nbsp;<%=server.htmlencode(rsTabella.fields("Quesito"))%>
					</a> </li>

				<%


	i=i+1

	cap=rsTabella(1)
	'response.write(capitolo)
	rsTabella.movenext
	if not rsTabella.eof then
		c=rsTabella(1)
	  '  response.write(capitolo & " " & c)
			    if StrComp(cap, c) = 0 then
                  ' Response.Write("Le due stringhe sono uguali")

                   else
                    i=0
                   ' Response.Write("Le due stringhe sono diverse")
			       %>
			       </ol>
  </div>
				  <%
                end if
         end if
		loop%>

<% end if%>


 <br>
















                      </div>
			        </div>
			      </div>
			    </div>
			</div>
             <!-- #include file = "../include/colora_pagina.asp" -->


		</div> <!--fine main-->
        </div>



<%  QuerySQL="SELECT count(*)" &_
" FROM Allievi  " &_
" WHERE Id_Classe ='" & Session("Id_Classe") & "'" &_
" ; "
Set rsTabellaCount = ConnessioneDB.Execute(QuerySQL)
RCount=rsTabellaCount(0)
 %>

        <form id="mod" name="studenti"  method="post"   onSubmit = 'return validate_feedback()'>
         <div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none; ">
           <div class="modal-header">
             <button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
             <h3 id="myModalLabel">Interroga</h3><button type="button" id="inviamodifica" class="btn btn-primary" onClick="estrai()" style="vertical-align:top">Sorteggia</button>
             <input type="text" name="txtSTUD" id="txtSTUD" value="" size="50"> &nbsp;&nbsp;  <!--Definisce i due bottoni del form -->


           </div>
           <div class="modal-body">


             <div class="control-group">
                <label for="textfield" class="control-label"><B>Domanda</B></label>
               <div class="controls">
                 <select id="Domanda" NAME = "Domanda" onchange="carica_domanda();">
                <option>Seleziona domanda</option>
                <%  rsTabella.movefirst
                  do while not rsTabella.EOF %>
                  <option value="<%=rsTabella("Quesito")%>"><%=rsTabella("Quesito")%></option>
                 <%rsTabella.movenext
                  loop
                %>
              </select>
              <br>
               </div>




               <div class="controls">
                 <table class="table table-hover table-nomargin table-condensed">
                 <%  QuerySQL="SELECT Cognome,Nome,CodiceAllievo,Interrogazioni" &_
                 " FROM Allievi  " &_
                 " WHERE Id_Classe ='" & Session("Id_Classe") & "' and Attivo=1" &_
                 " ORDER BY Allievi.Cognome Asc; "
                 Set rsTabella = ConnessioneDB.Execute(QuerySQL) %>
                 <thead>
                 <tr><th><b><span id="seleziona" onclick="uncheckTutti();">Seleziona</span></b></th><th><b>Studente</b></th><th><b>Valutazione</b></th><th><b>Risposta</b></th><th><b>Risposte</b></th><tr>
                 </thead>
                 <tbody>
                 <%
                    i=1
                    strstud=""
                    do while not rsTabella.eof %>
                        <tr>
                              <td style="width:10%"><input name="stud_<%=i%>" id="stud_<%=i%>" value='<%=rsTabella.fields("CodiceAllievo")%>' type='checkbox' checked='false'></td>
                            <td><%=rsTabella.fields("Cognome") & " " & left(rsTabella.fields("Nome"),1) &"." %></td>
                            <td>
                              <select id="Punteggio_<%=i%>" NAME = "Punteggio_<%=i%>">
                             <option value=-3>Seleziona punteggio</option>
                             <%  p=-3
                               do while p<=3%>
                               <option value="<%=p%>"><%=p%></option>
                              <%p=p+1
                               loop
                             %>
                             </select>
                            </td>
                            <td> <a id="spiegazioni_<%=i%>" data-original-title="Spiegazione" href="#" class="btn" rel="popover" data-trigger="hover" title="" data-placement="left" data-content="">
                          <center>  <i class="icon-question-sign"></i></center></a></td>
                            <td><span id="interrogazioni_<%=i%>"><%=rsTabella.fields("Interrogazioni")%></span></td>
                        </tr>
                    <%  'rsTabella.movenext
                  ''  vetstud(i)=rsTabella.fields("Cognome")
                    strstud = strstud & rsTabella.fields("Cognome") & " "& left(rsTabella("Nome"),1)&"." & ","
                    rsTabella.movenext
                    i=i+1

                    loop
                    RCount=i
                    NumStud=RCount
                    rsTabella.close
                 %>
               </tbody>
             </table><hr>
                 <INPUT TYPE = "button" NAME = "SubmitReply" VALUE = "Assegna punteggio" onClick="validate_feedback(<%=i%>)" class="btn">

               </div>

             </div>




           </div>
           <div class="modal-footer">
             <button id ="chiudi" class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>

           </div>
         </div>
       </form>

	</body>
<script>

function getRandomArbitrary(min, max) {
  return Math.random() * (max - min) + min;
}

function estrai(){

	var vettstud = new Array();
	var strstud = "null,<%=Left(strstud,Len(strstud)-1)%>";

	vettstud = strstud.split(",");

	var num = Math.round(getRandomArbitrary(1,<%=NumStud%>));
  if (typeof num == 'undefined') num=1;
	$("#txtSTUD").fadeOut(1000);
	var t = setTimeout(function(){ document.getElementById("txtSTUD").value = vettstud[num]; $('#stud_'+(num)).click(); clearTimeout(t); }, 1250);
	$("#txtSTUD").fadeIn(2000);

}


function uncheckTutti() {
  document.getElementById("txtSTUD").value="";
  //document.getElementById("Domanda").selectedIndex = "0";
 with (document.studenti) {
   for (var i=0; i < elements.length; i++) {
   if (elements[i].type == 'checkbox')
      elements[i].checked = false;
   if ((elements[i].type == 'select-one') && (elements[i].id != 'Domanda'))
      elements[i].selectedIndex="0";
   }
 }
}

function carica_domanda() {
   var maxlen=0;
   var id_maxlen=0;
      if (document.getElementById('Domanda').value!='Seleziona domanda')   {
        var url="../cDomande/9_carica_spiegazioni_ajax.asp?quesito="+document.getElementById('Domanda').value+"&id_classe=<%=id_classe%>";

         var xhttp = new XMLHttpRequest();
         xhttp.onreadystatechange = function() {
         if (xhttp.readyState == 4 && xhttp.status == 200) {
          var risposta=xhttp.responseText;
          console.log(risposta);
          var spiegazione;
              if (risposta!=""){
                testoJSON=JSON.parse(risposta);
                num=testoJSON["num"];
               for (var i=1;i<=num;i++){
                   spiegazione=testoJSON[i];
                  // alert(spiegazione.length);
                   if (spiegazione.length>maxlen) {
                     maxlen=spiegazione.length;
                     id_maxlen=i;
                  }
                   document.getElementById("spiegazioni_"+i).setAttribute("data-content",spiegazione);
                }
              } else {
                  alert("Errore nel caricamento risposte");
              }
             //alert(maxlen+'---'+id_maxlen);
              $('#stud_'+(id_maxlen)).click();
          }
         };
         xhttp.open("GET", url, true);
         xhttp.send();

      }

}


function validate_feedback(n) {
    var selezionato=false;
    var domanda=false;
    var pronto=true;
    var url="../cSocial/invia_feedback_frasi.asp?scegli=3&RCount=<%=RCount%>&paragrafo=<%=TitoloParagrafo%>&id_categoria=<%=id_categoria%>&id_classe=<%=id_classe%>&ThreadId=<%=threadparent%>&ParentId=<%=parentmessage%>";
    var i;

    if (document.getElementById('Domanda').value=='Seleziona domanda')   {
      alert('Non hai selezionato la domanda');
      pronto=false;
    }
    else {
        url=url+"&domanda="+document.getElementById('Domanda').value;
    }
    for (i=1;i<n;i++){
     if  (document.getElementById('stud_'+i).checked==true)
       if (document.getElementById('Punteggio_'+i).selectedIndex!='0'){
           selezionato=true;
           url=url+'&'+document.getElementById('stud_'+i).value +'='+document.getElementById('Punteggio_'+i).value;
           }
        else
          selezionato=false;
    }
   if  (selezionato==false)
      {
      alert("Non hai selezionato lo studente o il punteggio  ");
      pronto=false;
   }
 else if (pronto==true)
 {
/*
   document.studenti.action = url;
  document.studenti.submit();*/
//alert(url);
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  if (xhttp.readyState == 4 && xhttp.status == 200) {
   var risposta=xhttp.responseText;
   if (risposta=="Modifica avvenuta"){
     alert('Punteggio assegnato');
    // $('#seleziona').click();
   var k=0;
    with (document.studenti) {
      for (var i=0; i < elements.length; i++) {
           if (elements[i].type == 'checkbox'){
              k+=1;
              if (elements[i].checked){
              //  alert(k+'&'+document.getElementById('interrogazioni_'+k).innerHTML);
                document.getElementById('interrogazioni_'+k).innerHTML=parseInt(document.getElementById('interrogazioni_'+k).innerHTML)+1;
                }
          }
      }
    }
    uncheckTutti();
     }
   else {
     alert(risposta);
     }
   }
  };
  xhttp.open("GET", url, true);
  xhttp.send();
 }
}

</script>

 </html>
