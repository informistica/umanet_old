<%@ Language=VBScript %>

<%
 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 divid=Session("divid")
 id_classe= Session("Id_Classe")
 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
 session("scegli")=scegli

 categoria=request.QueryString("categoria")
 id_categoria=request.QueryString("id_categoria")
 if categoria="" then
    categoria=session("categoria")
 end if
 if id_categoria="" then
    id_categoria=session("id_categoria")
 end if

 if strcomp(scegli,"0")=0 then
 bacheca=request.QueryString("bacheca")
 Session("bacheca")=bacheca
 end if
 'if session("Admin")=true then
' cognome=Session("Cognomebacheca")
' nome=Session("Nomebacheca")
' else
' if cognome="" or nome="" then
' cognome=request("cognome")
' nome=request("nome")
' end if

 cognome=request("cognome")
 nome=request("nome")
 if (session("CodiceAllievo")="") or (session("Id_Classe")="") then response.Redirect("../../home.asp")

%>
	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->

<!-- #include file = "../service/controllo_sessione.asp" -->
   <!-- #include file = "../var_globali.inc" -->


<%
Function prepStringForSQL(sValue)

Dim sAns
sAns = Replace(sValue, Chr(39), "''")

sAns = "'" & sAns & "'"
prepStringForSQL = sAns

End Function

sName = Session("Cognome") & " " & left(Session("Nome"),1)&"."





Function isBlank(Value)

if isNull(Value) then
	bAns = true
else
	bAns = trim(Value) = ""
end if
isBlank = bAns

end function

Function FixNull(Value)
if isNull(Value) then
	sAns = ""
else
	sAns = trim(Value)
end if

FixNull = sAns
end function



ID=request.QueryString("ID")

%>
<!doctype html>
<html>
<head>


   <title>Nuova attivit&agrave;</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	    <meta charset="UTF-8">
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

<!-- Bootstrap -->

<link rel="stylesheet" href="../../css/bootstrap2.min.css">
<!--	<link rel="stylesheet" href="../../css/bootstrap.min.css">
	<!-- Bootstrap responsive -->
<!--<link rel="stylesheet" href="../../css/bootstrap-responsive.min.css">-->



    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
<!--	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui.css">
	<link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery.ui.theme.css">-->


    <!-- Theme CSS -->
   <!-- <link rel="stylesheet" href="../../css/style.css">-->
	<!-- Color CSS -->
	<!--<link rel="stylesheet" href="../../css/themes.css">-->

     <link rel="stylesheet" href="../../css/style-themes.css">

	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>

    <!-- jQuery UI -->
    <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>
<!--	<script src="../../js/plugins/jquery-ui/jquery.ui.core.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.widget.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.mouse.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.resizable.min.js"></script>
	<script src="../../js/plugins/jquery-ui/jquery.ui.sortable.min.js"></script>-->



	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>

	<!-- Theme framework -->

    <script src="../../js/eak_app_dem.min.js"></script>
<!--	<script src="../../js/eakroko.min.js"></script> -->
	<!-- Theme scripts -->
	<!--<script src="../../js/application.min.js"></script>-->
	<!-- Just for demonstration -->
	<!--<script src="../../js/demonstration.min.js"></script> -->

	<!--[if lte IE 9]>
		<script src="../../js/plugins/placeholder/jquery.placeholder.min.js"></script>
		<script>
			$(document).ready(function() {
				$('input, textarea').placeholder();
			});
		</script>
	<![endif]-->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../social/img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../social/img/apple-touch-icon-precomposed.png" />

    <!-- CKEditor -->
	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>




<script src="https://ajax.microsoft.com/ajax/jQuery/jquery-1.4.4.min.js" type="text/javascript"></script>
<script src="_assets/js/jquery.zclip.js"></script>
<script src="include/copiaincolla.js"></script>
 <script type="text/javascript">
function scambia(state) {
	if (state==1)
	{
		document["bottiglia"].src = "smilies/on_1.gif";
	}
	else
	{
		document["bottiglia"].src = "smilies/on_2.gif";
	}
}
</script>





   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />-->

 <!--#include file = "include/Validate.inc"-->

<%
 select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="bacheca"
  case "2"
    session("social")="diario"
  case "3"
    session("social")="interrogazioni"

 end select %>

	<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->
</head>


<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
	<div id="navigation">


		<!-- #include file = "../include/navigation.asp" -->


	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Nuovo argomento </h1>

					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->
                    </div>
				</div>
                <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
				<div class="breadcrumbs">
					<ul>



                         <%
						     cartella=rsTabella.fields("cartella") ' per passarlo ex2_imgAll.asp
						     select case scegli
							 case "0"
								 session("social")="forum"
							 %>
                             <li>
							<a href="#">Umanet</a>
							<i class="icon-angle-right"></i>
						</li>
							<li>
							<a   href="../cSocial/default0.asp?scegli=0&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Forum</a>
						    </li>



							 <%
							 case "1"
							 %>
                             <li>
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>
							 <li>
							<a  href="../cSocial/default0.asp?scegli=1&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>">&nbsp;Bacheca</a>
						    </li>
							 <%
							  case "2"
							 %>
                             <li>
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>
							 <li>
							<a   href="../cSocial/default0.asp?scegli=2&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Diario</a>
						    </li>

							 <%

               case "3"
              %>
                            <li>
             <a href="#">Classe</a>
             <i class="icon-angle-right"></i>
           </li>
              <li>
             <a   href="../cSocial/default0.asp?scegli=3&amp;id_classe=<%=rsTabella.fields("Id_Classe")%>&amp;divid=<%=divid%>&amp;cartella=<%=rsTabella.fields("cartella")%>"><span></span>&nbsp;Interrogazioni</a>
               </li>

              <%
							 end select %>


                        <li> <i class="icon-angle-right"></i>


                       <%select case scegli
							 case "0"
								 session("social")="forum"
							 %>



							<a title="Torna alle Discussioni" href="default0.asp?scegli=0&id_classe=<%=id_classe%>&cartella=<%=cartella%>&bacheca=<%=bacheca%>&nome=<%=request("nome")%>&cognome=<%=request("nome")%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>
							 <%
							 case "1"
							 %>


							 <li>
							<a title="Torna alle Discussioni" href="default0.asp?scegli=1&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>
							 <%
							  case "2"
							 %>
							 <li>
							<a title="Torna alle Discussioni" href="default0.asp?scegli=2&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                            <i class="icon-angle-right"></i>
						    </li>

							 <%

               case "3"
              %>
              <li>
             <a title="Torna alle Discussioni" href="default0.asp?scegli=3&id_classe=<%=id_classe%>&cartella=<%=cartella%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"><%=categoria%></a>
                           <i class="icon-angle-right"></i>
               </li>

              <%
							 end select %>


						 <li>
							<a title="Torna alla Discussione" href="#">Nuova</a>
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
				   <!--   <div class="box-title">
				        <h3> <i class="icon-reorder"></i> ... </h3>
			          </div>-->
				     <!-- <div class="box-content">-->



				<div class="row-fluid">
				  <div class="span12">
				 <!--   <div class="box">-->






		  <!--  <div class="box-content"> -->

              	<div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-th-list"></i> Nuovo argomento</h3>
							</div>
							<div class="box-content nopadding">



              <FORM Name = "InputForm"   METHOD = "POST" class='form-horizontal form-bordered'>

<INPUT TYPE = "Hidden" NAME = "MessageType" VALUE = "NEW">
<INPUT TYPE = "Hidden" NAME = "CodBacheca" VALUE = "<%=bacheca%>">


<div class="control-group">
<label for="textfield" class="control-label"><b>Nome</b></label>
  <div class="controls">

  <INPUT   TYPE = "Hidden"  NAME="Name"  value='<%=Session("Cognome") & " " & left(Session("Nome"),1)& "."%>' >
	  <INPUT TYPE = "TEXT" disabled="true" NAME="Name1"  value='<%=Session("Cognome") & " " & left(Session("Nome"),1)& "."%>' class="input-xlarge">
  </div>
</div>

<div class="control-group">
<label for="textfield" class="control-label"><B>Argomento:</B></label>
  <div class="controls">
	  <INPUT TYPE = "TEXT"  NAME = "Topic" class="input-xlarge" placeholder="Titolo del post">
  </div>
</div>

<div class="control-group">
<label for="textfield" class="control-label"><B>Abstract:</B></label>
  <div class="controls">
	  <INPUT TYPE = "TEXT"  NAME = "Breve" class="input-xlarge" placeholder="Descrizione breve del post" value="...">
  </div>
</div>

 <div class="control-group">
<label for="textfield" class="control-label"><B>Messaggio:</B></label>
  <div class="controls">
    <% if sMsg="" then %>
	  <textarea class='ckeditor span12' rows="5" NAME = "MESSAGE" cols="40" placeholder="Text input" >
      <% if (request("MESSAGE")<>"") then
	   response.write(request("MESSAGE"))
	   end if
	  %>

      </textarea>
      <%else%>
        <textarea class='ckeditor span12' rows="5" NAME = "MESSAGE" cols="40" placeholder="Text input"><%=sMsg%></textarea>
      <%end if%>
  </div>
</div>

<div class="control-group">
<label for="textfield" class="control-label"><B>Azione:</B></label>
  <div class="controls">

	   <INPUT TYPE = "TEXT"  NAME = "AZIONE1" class="input-xxlarge" placeholder="Incolla URL da collegare al Compito">

  </div>
</div>


 <div class="control-group"><center>
<span class="sottotitolo"><a title="Carica foto" href="#" onClick="javascript:PopUpWindow(600,300,<%=scegli%>);return false;">  <img src="img/caricaimg.png" width="39" height="39"></a>
</span>  &nbsp;&nbsp;&nbsp;&nbsp;
 <span class="sottotitolo"><a title="Carica file"  href="#" onClick="javascript:PopUpWindow2(600,300,<%=scegli%>);return false;"> <img src="img/caricafile.jpg" width="35" height="33"></a>
</span><br>
 	<% if session("CaricatoFile")=true then %>
	 Risorse:

			   <B><%=session("NomeFileForum2")%></B>
    <%end if%>
  </center>



 <div class="accordion" id="accordion3">
									<div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseSmile"><center>
												 <img title='Inserisci emoticons' src="smilies/icon-smilie.gif" align="absmiddle" class="image"></center>
											</a>
										</div>
										<div id="collapseSmile" class="accordion-body collapse">
											<div class="accordion-inner">








                  <ul id="myTab2" class="nav nav-tabs">

                                    <li class="active"><a href="#profileEm" data-toggle="tab">Emoticons</a></li>
                                    <li class="active"><a href="#profileCo" data-toggle="tab">Connessioni</a></li>
                                    <li class="active" ><a href="#profileNa" data-toggle="tab">Navigazione</a></li>
                                    <li class="active"><a href="#profileUe" data-toggle="tab">Umanet Explorer</a></li>
                                    <li class="active"><a href="#profileIn" data-toggle="tab">Interfacce</a></li>

                            </ul>
                            <div id="myTabContent2" class="tab-content">

                              <div class="tab-pane fade in active" id="profileEm">


                     <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Emoticons
								</h3>
							</div>
							<div class="box-content nopadding">
							     <!--#include file = "include/smilies.inc"-->
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                           </div>






                               <div class="tab-pane in active " id="profileCo">



     			 <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Connessioni
								</h3>
							</div>
							<div class="box-content nopadding">
								 <!--#include file = "include/connessioni_percezioni.inc"-->
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                              </div>



                               <div class="tab-pane fade in active" id="profileNa">

                                <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Navigazione
								</h3>
							</div>
							<div class="box-content nopadding">
								 <!--#include file = "include/navigazione_browser.inc"-->
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                              </div>



                               <div class="tab-pane fade in active" id="profileIn">



     			 <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Interfacce
								</h3>
							</div>
							<div class="box-content nopadding">
								 <!--#include file = "include/interfacce.inc"-->
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                              </div>

                                 <div class="tab-pane fade in active" id="profileUe">



     			 <!----Inizio -->
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Umanet Explorer
								</h3>
							</div>
							<div class="box-content nopadding">
								 <!--#include file = "include/umanet_explorer.inc"-->
							</div>
						</div>
					</div>
				</div>
                 <!-- >fine form -->

                              </div>



                   </div>

                      </div>

											</div>
										</div>


									</div>




 <!-- Incorporamento url obsoleto
  <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseUrl"><center>

                                                <img src="../../img/link.jpg" width="35" height="25" title="Incorpora utilizzando i link">
                                                </center>
											</a>
										</div>
										<div id="collapseUrl" class="accordion-body collapse">
											<div class="accordion-inner">

											<TEXTAREA class="input-block-level" NAME = "INCORPORA" placeholder="Incolla URL della pagina da incorporare"></TEXTAREA>

                     						 </div>

										</div>

                                     </div>

  -->

     <!--   </div>   -->

                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseMail"><center>

                                                <img id="notify-email" class="imground" title='Notifica per email'  src="../../img/icon_mail.jpg" width="50px" height="40px" align="absmiddle" style="border-color:none;">
                                                </center>
											</a>
										</div>

										<div id="collapseMail" class="accordion-body collapse">
											<div class="accordion-inner">

											  <input type="checkbox"  name="cbEmail2" id="cbEmail1" title="Selezionare per inviare un email alla classe">   Notifica per email a tutta la classe &nbsp;&nbsp;&nbsp;<br><br/><br>
                                              <input type="checkbox"   id="cbEmailProf1" name="cbEmailProf" title="Selezionare per inviare un email al prof.">   Notifica per email al prof. &nbsp;&nbsp;&nbsp;
                                               <input type="checkbox"   id="cbAnonimo" name="cbAnonimo" title="Seleziona per creare una discussione anonima">   Discussione anonima &nbsp;&nbsp;&nbsp;
											  <% if session("Admin")=true then%>
                                               <input type="checkbox"   id="cbZip" name="cbZip" title="Selezionare per prevedere la consegna di siti in archivi .zip">   Consegna html in .zip &nbsp;&nbsp;&nbsp;
      										   <input type="checkbox"   id="cbNascosto" name="cbNascosto" title="Seleziona per nascondere il post">   Nascondi il post &nbsp;&nbsp;&nbsp;
                                              
												<% end if%>
                     						 </div>
										</div>

                                       <%
									 '  response.write("bacheca="&bacheca)
									   if bacheca<>"" then %>


                                       <%  QuerySQL="SELECT count(*)" &_
" FROM Allievi  " &_
" WHERE Id_Classe ='" & Session("Id_Classe") & "';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
numStud=rsTabella(0) %>


                                          <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseCondividi">
                                            	<center>
                                                <img id="notify-condividi" class="imground" title='Condividi con '  src="img/condividere_small.jpg"width="50px" height="40px" align="absmiddle" style="border-color:none;">
                                                </center>
											</a>
										</div>

										<div id="collapseCondividi" class="accordion-body collapse">
											<div class="accordion-inner">


                                            <LEGEND><B>Condividi con </B></LEGEND>
       <b>Gruppo da </b> <input type="text" name="txtNUMREC" class="input-mini" value="<%=numStud%>" size="1"> <br>
      <table class="table table-hover table-nomargin table-condensed">
		<tr align="center"><th  colspan="2">

       </th></tr>






<%  QuerySQL="SELECT Cognome,Nome,CodiceAllievo" &_
" FROM Allievi  " &_
" WHERE Id_Classe ='" & Session("Id_Classe") & "'" &_
" ORDER BY Allievi.Cognome Asc; "
Set rsTabella = ConnessioneDB.Execute(QuerySQL) %>

 <tr><td>&nbsp;</td><td>&nbsp;</td><td>
<a title="Rendi pubblico" onClick="checkTutti(<%=numStud%>);" >
        Tutti&nbsp;&nbsp;&nbsp;</a>
<a title="Rendi privato" onClick="uncheckTutti();" >&nbsp;&nbsp;&nbsp;Nessuno</a>

</td></tr>
<tr><td><b>Cognome</b></td><td><b>Nome</b></td><td><b>Condividi</b></td><tr>

<!--<tr><td colspan="2">&nbsp;</td></tr>-->

<% ' response.write(rsTabella.fields("Data") )
   if rsTabella.eof then
   %>
   <tr><td colspan="7">Vuota<%=QuerySQL%></td></tr>
   <%
   end if
   i=1
   ' prelevo i dati da inserire dalla query sui risultati
   do while not rsTabella.eof %>
   <input type="hidden" value="<%=rsTabella.fields("CodiceAllievo")%>" name="txtStud<%=i%>" />
   <input type="hidden" value="<%=rsTabella.fields("Cognome") & " " & left(rsTabella.fields("Nome"),1) &"." %>" name="txtStud2<%=i%>" />
   <tr><td>
   <%=rsTabella.fields("Cognome")%></td><td><%=rsTabella.fields("Nome")%></td><td align="center"> <input type="checkbox"  name="cbCondividi<%=i%>" title="<%=i%>" value="<%=i%>" checked="true"  onChange="aggiorna('cbCondividi<%=i%>');" ></td> </tr>
   <%  rsTabella.movenext
   i=i+1
   loop
   rsTabella.close
%>
</table>









                     						 </div>
										</div>

							<%end if%>



										</div>
                                    </div>


















<P>
<CENTER>
<!-- <INPUT TYPE = "button"  VALUE = "Sostituisci" onClick="sostituisci();">-->
<br>
<INPUT TYPE = "button" NAME = "SubmitMessage" VALUE = "Pubblica" onClick="newpost();" class="btn">
</FORM>

</div>
<!--</div> -->









                   <!--   </div> -->
			        </div>
			      </div>
			  <!--  </div>-->











                      </div>
			        </div>
			      </div>
			    </div>
			</div>


		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->



	</body>

     <script type="text/javascript">


$(window).load(function () {

	   $('#notify-email').click();
	   $('#cbEmail1').click();
	   $('#cbEmailProf1').click();


	    //event.stopPropagation();

	});

</script>

<script language="javascript" type="text/javascript">
function newpost() {
	    sostituisci();
	    var autore=InputForm.Name.value;
		var commento=InputForm.Topic.value;
		//var abstract=InputForm.Breve.value;
		var messaggio=InputForm.MESSAGE.value;
		if (autore=="")
	     {
		   alert("Non hai scritto l'autore ");
		   return 0;
		}
		 else
		 if (commento=="")
		{
		   alert("Non hai scritto l'argomento! ");

		   return 0;
		}
		// else
		// if (abstract=="")
		//{
		 //  alert("Non hai scritto la descrizione breve ");

	//	   return 0;
	//	}
	// else
		// if (messaggio=="")
		//{
		  // alert("Non hai scritto il messaggio.");

		  // return 0;
		//}
	else
	{
        document.InputForm.action = "PreviewMessage.asp?scegli=<%=scegli%>&SubmitMessage=1&codBacheca=<%=bacheca%>&numRec=<%=i-1%>&numStud=<%=numStud%>&cognome=<%=request.QueryString("cognome")%>&nome=<%=request.QueryString("nome")%>&byChiamante=1";
		document.InputForm.submit();
	}

}
 </script>


  <script type="text/javascript">


 function aggiorna(nome) {
	// window.alert ("ciao");
		with (document.InputForm) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMREC.value=parseInt(txtNUMREC.value)+1;
		 else
		    txtNUMREC.value=parseInt(txtNUMREC.value)-1;
	    }
}

 function uncheckTutti() {
	with (document.InputForm) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'checkbox')
			  elements[i].checked = false;

		}
	 txtNUMREC.value=parseInt(0);
	}
}

function checkTutti(numStud) {
	with (document.InputForm) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		if (elements[i].type == 'checkbox')
		    elements[i].checked = true;


		}
		 txtNUMREC.value=parseInt(numStud);
	}
}


function addsmile(codice) {

		with (document.InputForm) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina

		   // Name.value= Name.value + codice;

		  MESSAGE.value= MESSAGE.value + codice;

	    }
}

 //fa una sostituzione alla volta di " quindi la richiamo 8 volte
 function sostituisci2()
 {
	var msg=InputForm.MESSAGE.value;
	 with (document.InputForm) {
		  MESSAGE.value= msg.replace(String.fromCharCode(34),"");
	   }
 }
 function sostituisci() {
		for (var i=1;i<=8;i++)
		{
		 sostituisci2();
		 }
}



 //si può togliere probabilmente

 function validate2() {
	var stringa=frmDocument.flname.value;
	if (stringa.search(".jpg") == -1)
	{
	   alert("L'immagine deve essere in formato .jpg");
	   frmDocument.imgname.setfocus();
	   return 0;
	}
 else

 if (frmDocument.imgname.value=="")
	{
	   alert("Non hai inserito il nome dell'immagine.");
	   frmDocument.imgname.setfocus();
	   return 0;
	}
	else
	{
	    document.frmDocument.action = "admin/Upload/confirm_update.asp?scegli=<%=scegli%>&Quesito=<%=Quesito%>&ID_Prefrase=<%=ID_Prefrase%>&by_UECDL=<%=by_UECDL%>&prefrase=<%=prefrase%>&Tipo=<%=Tipo%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&Num=<%=Num%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&AggRisFrase=1&Img=1&by_UPLOAD=<%=by_UPLOAD%>&ID=<%=Id_Frase%>&contDomande=<%=contDomande%>";
		document.frmDocument.submit();


    }

}

function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

/*
    switch(s)
	{
	case 0:

window.open('../upload_resize/ex2_imgforum.asp','../upload_resize/ex2_imgforum.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl);
 break;

	case 1:

window.open('../upload_resize/ex2_imglavagna.asp','../upload_resize/ex2_imglavagna.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl);
   break;

	case 2:

window.open('../upload_resize/ex2_imgdiario.asp','../upload_resize/ex2_imgdiario.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl);
 break;
}*/

window.open('../upload_resize/ex2_imgAll.asp?cartella=<%=cartella%>','../upload_resize/ex2_imgAll.asp?cartella=<%=Session("Cartella")%>', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl);

}
// -->
function PopUpWindow2(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;
	//alert(s);
	 switch(s)
	{
	case 0 :
   window.open('../upload_file/db-file-to-disk.asp?daForum=1','../db-file-to-disk.asp?daForum=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=930,height=340,top='+wint+',left='+winl);break;
	case 1 :
   window.open('../upload_file/db-file-to-disk.asp?daLavagna=1','../db-file-to-disk.asp?daLavagna=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=930,height=340,top='+wint+',left='+winl);break;
	case 2 :
   window.open('../upload_file/db-file-to-disk.asp?daDiario=1','../db-file-to-disk.asp?daDiario=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=930,height=340,top='+wint+',left='+winl);break;
 }

}

function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>

   <script type="text/javascript" src="../js/refresh_session.js"></script>

 </html>
