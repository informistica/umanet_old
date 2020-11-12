<%@ Language=VBScript %>
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


 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 divid=Session("divid")
 id_classe= Session("Id_Classe")
 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario

 categoria=request.QueryString("categoria")
 id_categoria=request.QueryString("id_categoria")
 if strcomp(scegli,"0")=0 then
 bacheca=request.QueryString("bacheca")
 Session("bacheca")=bacheca
 end if


 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
 session("scegli")=scegli
select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="lavagna"
  case "2"
    session("social")="diario"
    case "3"
      session("social")="interrogazioni"

 end select



 if (session("CodiceAllievo")="") or (session("Id_Classe")="") then response.Redirect("../../home.asp")

%>
   <!-- #include file = "../var_globali.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->
   <!--#include file = "include/format_message.asp"-->


<%

'Response.AddHeader "Refresh", "900" ' dopo un ora di inattività
Function prepStringForSQL(sValue)

Dim sAns
sAns = Replace(sValue, Chr(39), "''")

sAns = "'" & sAns & "'"
prepStringForSQL = sAns

End Function

sName = Session("Cognome") & " " & left(Session("Nome"),1)&"."





MessageID = Request("MessageID")
if MessageID = "" then MessageID = request.QueryString("MessageID")
ThreadID = request("ThreadID")
if ThreadID = "" then ThreadID = request.QueryString("ThreadID")
sOrigAuthor = request("OrigAuthor")
if sOrigAuthor = "" then sOrigAuthor = request.QueryString("OrigAuthor")

' lo faccio da default con session("Zip")
'QuerySQL="Select Zip from FORUM_MESSAGES where ID="&ThreadID&";"
'set rs = conn.execute (QuerySQL)
'Session("zipFile")=rs(0)
Session("IDTHREAD")=ThreadID

'sOrigMessageFormat=request.QueryString("sOrigMessageFormat")

 ' divid=request("divid")
 ' cartella=request("cartella")
 ' id_classe=request("id_classe")
  divid=Session("divid")

  id_classe= Session("Id_Classe")
  bacheca=request.querystring("bacheca")
  RCount=request("RCount")

bValid = MessageID <> ""
if bValid then



sSQL = "SELECT TOPIC,COMMENTS,CodiceAllievo,Urlimg FROM FORUM_MESSAGES WHERE ID = " & Request("MessageID")
cmd.CommandText = sSQL
set rs = cmd.Execute

    CodiceAllievoOrig = rs("CodiceAllievo")
	sOrigMessage = replace(rs("comments"), vbcrlf, "<BR>")
	sOrigMessageFormat=sOrigMessage
	'sOrigMessage = replace(sOrigMessage, "  ", "&nbsp; ")
	sOrigTopic = rs("Topic")
	sUrlimg=rs("Urlimg")
	conn.execute sSQL

	 iCategory = request("Category")
%>
<!--#include file = "database_cleanup.inc"-->
<%




iTopic = Request("Category")

end if 'bValid
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
<!-- inclusione dei fogli di stile e javascript per il layout grafico-->
<script src="../../js/google.js"></script><title>Feedback</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />
	 <meta charset="UTF-8">

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

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
  <!-- CKEditor -->
  	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!--Chiamata periodica a pagina di refresh-->
  <script type="text/javascript" src="../js/refresh_session.js"></script>

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />




<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

<script>
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

function checkTutti() {
 numcb=0;
 with (document.InputForm) {
   for (var i=0; i < elements.length; i++) {
   if (elements[i].type == 'checkbox')
       {
        elements[i].checked = false;
      numcb=numcb+1;
     }
   }
 }
 //document.dati.txtNUMREC.value=numcb;

 /*

 document.getElementById('cbEmail0').checked=false
 document.getElementById('cbEmail1').checked=false
 document.getElementById('cbEmail2').checked=false
 document.getElementById('cbEmailProf').checked=false

 */

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


window.open('../upload_resize/ex2_imgAll.asp','../upload_resize/ex2_imgAll.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl);

}
// -->
function PopUpWindow2(w,h,s) {
 var winl = (screen.width - w) / 2;
 var wint = (screen.height - h) / 2;
 //alert(s);
  switch(s)
 {
 case 0 :
  window.open('../upload_file/db-file-to-disk.asp?daForum=1','../db-file-to-disk.asp?daForum=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=400,height=250,top='+wint+',left='+winl);break;
 case 1 :
  window.open('../upload_file/db-file-to-disk.asp?daLavagna=1','../db-file-to-disk.asp?daLavagna=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=400,height=250,top='+wint+',left='+winl);break;
 case 2 :
  window.open('../upload_file/db-file-to-disk.asp?daDiario=1','../db-file-to-disk.asp?daDiario=1', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=400,height=250,top='+wint+',left='+winl);break;
}

}

function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
   var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
   upload.focus();
 }
</script>


   <!-- Copia incolla-->
<script src="_assets/js/jquery.zclip.js"></script>
<script src="include/copiaincolla.js"></script>
  <script src="_assets/js/jquery-ui.js"></script>
<script src="../../js/datapicker_it.js"></script>

</head>

<%
 select case scegli
 case "0"
     session("social")="forum"

 case "1"

    session("social")="lavagna"
  case "2"
    session("social")="diario"
    case "3"
      session("social")="interrogazioni"

 end select %>

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
 						<h1> <i class="icon-comments"></i> Risposta </h1>

 					</div>
 					<div class="pull-right">
                      <!-- se mi interessa devo includere
                          include pull_right.asp-->
                     </div>
 				</div>
                 <!--Barra per sapere la pagina in cui sono eventualmente fa anche da menu-->
 				<div class="breadcrumbs">
 					<ul>
                          <%select case scegli
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
 							<a title="Torna alla Discussione" href="#">Rispondi</a>
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
 				    <!--  <div class="box-content">-->
 				<div class="row-fluid">
 				  <div class="span12">
 				   <!-- <div class="box">-->
 		  <!--  <div class="box-content"> -->
  <div class="box box-bordered">
 <div class="box-title">
    <h3><i class="icon-th-list"></i> In Risposta a :</h3>
 </div>
   <% 'url= "../../Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/img_lavagna/thumb" ' vuole il percorso relativo della cartella
   url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("Cartella")&"/img_"&Session("Social")&"/img" ' vuole il percorso relativo della cartella

 		   url=Replace(url,"\","/")
 		   urlimg=url&"/"& sUrlimg
 		   'response.write(urlimg)
 		   'Session("Caricata")=false
 		   if strcomp(sUrlimg&"","")<>0 then
 		   %>
           <br> <center><img class="imground" src="<%=urlimg%>" border="1"> </center> <br>
    		  <%end if %>


  <div class="control-group"><br>
 <label for="textfield" class="control-label"><b>Post:</b></label>
   <div class="controls">

  <!-- <p><textarea  rows="6" name="S1"  id="S1"  class="input-block-level">-->
   <%=ReplaceCar(FormatMessage(sOrigMessageFormat)) %>
  <!-- </textarea></p>-->


   </div>
 </div>


               	<div class="box box-bordered">
 							<div class="box-title">

                 <%if (strcomp(categoria,"Feedback")=0) and (session("admin")=true) then%>
                 	<h3><i class="icon-th-list"></i> Nuovo feedback</h3>
                 <%else%>
 								<h3><i class="icon-th-list"></i> Nuovo commento</h3>
                 <%end if%>

 							</div>
 							<div class="box-content nopadding">

 <%if (strcomp(categoria,"Feedback")=0) and (session("admin")=true) then%>
 <!--<FORM NAME = "InputForm" ACTION = "Invia_feedback.asp?Reply=1&scegli=<%=scegli%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>" onSubmit = 'return validate_feedback()' METHOD = "POST" class='form-horizontal form-bordered'>
-->
<form name="InputForm" METHOD = "POST" class='form-horizontal form-bordered'>
 <%else%>
 <FORM NAME = "InputForm" ACTION = "PreviewMessage.asp?Reply=1&scegli=<%=scegli%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>" onSubmit = 'return validate()' METHOD = "POST" class='form-horizontal form-bordered'>
 <%end if%>
 <INPUT TYPE = "HIDDEN" NAME = "ThreadId" VALUE="<%= ThreadID %>"></INPUT>
 <INPUT TYPE = "HIDDEN" NAME = "ParentId" VALUE="<%= MessageID %>"></INPUT>
 <INPUT TYPE="HIDDEN" NAME="OrigAuthor" VALUE="<%= sOrigAuthor %>">
 <INPUT  TYPE="HIDDEN" NAME="CodiceAllievo" VALUE="<%= CodiceAllievoOrig %>">
 <INPUT TYPE = "Hidden" NAME = "CodBacheca" VALUE = "<%=bacheca%>">




 <div class="control-group">
 <label for="textfield" class="control-label"><b>Nome</b></label>
   <div class="controls">

   <INPUT   TYPE = "Hidden"  NAME="Name"  value='<%=Session("Cognome") & " " & left(Session("Nome"),1)& "."%>' >
 	  <INPUT TYPE = "TEXT" disabled="true" NAME="Name1"  value='<%=Session("Cognome") & " " & left(Session("Nome"),1)& "."%>' class="input-xlarge">
   </div>
 </div>
 <%if (strcomp(categoria,"Feedback")=0) and (session("admin")=true) then%>
 <div class="control-group">
 <label for="textfield" class="control-label"><B>Feedback</B></label>
   <div class="controls">
		<select id="Polarita" NAME = "Polarita" onchange="carica_feedback();">
				<option value='Seleziona polarità'>Seleziona polarità</option>
			<%  'QuerySQL="Select * from Feedback_Polarita order by id"
			'QuerySQL="Select * from Feedback_Polarita where id=4" ' se voglio avere solo una categoria generale
			QuerySQL="Select * from Feedback_Polarita"
			  Set rsTabellaFeed = ConnessioneDB.Execute(QuerySQL)
			  do while not rsTabellaFeed.eof
			  %>
			  <option  value=<%=rsTabellaFeed("id")%>> <%=rsTabellaFeed("Nome")%></option>
			 <%rsTabellaFeed.movenext
			  loop
			%>
			</select>

     <select id="Topic" NAME = "Topic" disabled	>
        <option selected>Seleziona feedback</option>
    <%  QuerySQL="Select Segno,Descrizione from Feedback order by Segno, Posizione"
      Set rsTabellaFeed = ConnessioneDB.Execute(QuerySQL)
      do while not rsTabellaFeed.eof
      if rsTabellaFeed("Segno")="-" Then
        colore="red"
      else
        colore="green"
      end if
      %>
      <option style="color:<%=colore%>;" value="(<%=rsTabellaFeed("Segno")%>)<%=rsTabellaFeed("Descrizione")%>">(<%=rsTabellaFeed("Segno")%>)&nbsp;<%=rsTabellaFeed("Descrizione")%></option>
     <%rsTabellaFeed.movenext
      loop
    %>
    </select>
    <select id="Punteggio" NAME = "Punteggio" disabled>
   <option>Seleziona punteggio</option>
   <%  p=-5
     do while p<6%>
     <option value="<%=p%>"><%=p%></option>
    <%p=p+1
     loop
   %>
 </select>
 <br>
 <fieldset>
   <legend></legend>
   <INPUT TYPE = "TEXT"  NAME="Topic1" id="Topic1"  value='' placeholder='Inserisci un nuovo tipo di feedback'  class="input-xlarge">
 <font color=red>(KO)</font>
 <input name="newfeedneg" id="newfeedneg" value='1' type='checkbox'   title="Aggiungi al database come feedback negativo">
 <font color=green>(OK)</font>
   <input name="newfeedpos" id="newfeedpos" value='1' type='checkbox'   title="Aggiungi al database come feedback positivo">
  </fieldset>


   </div>
 </div>
 <div class="control-group">
 <label for="textfield" class="control-label"><B>Invia a </B></label>
   <div class="controls">
     <table class="table table-hover table-nomargin table-condensed">
     <%  QuerySQL="SELECT Cognome,Nome,CodiceAllievo" &_
     " FROM Allievi  " &_
     " WHERE Id_Classe ='" & Session("Id_Classe") & "' and Attivo=1" &_
     " ORDER BY Allievi.Cognome Asc; "
     Set rsTabella = ConnessioneDB.Execute(QuerySQL) %>
     <thead>
     <tr><th><b>Seleziona</b></th><th><b>Studente</b></th><tr>
     </thead>
     <tbody>
     <%
        i=1
        do while not rsTabella.eof %>
            <tr>
                	<td style="width:10%"><input name="stud_<%=i%>" id="stud_<%=i%>" value='<%=rsTabella.fields("CodiceAllievo")%>' type='checkbox'></td>
                <td><%=rsTabella.fields("Cognome") & " " &  rsTabella.fields("Nome")  %></td>
                
            </tr>
        <%  rsTabella.movenext
        i=i+1
        loop
        RCount=i
        rsTabella.close
     %>
   </tbody>
 </table>
   </div>

 </div>


 <%else%>
 <div class="control-group">
 <label for="textfield" class="control-label"><B>Argomento:</B></label>
   <div class="controls">
 	  <INPUT TYPE = "TEXT"  NAME = "Topic" class="input-xlarge"  VALUE="Re: <%=sOrigTopic %>">
   </div>
 </div>
  <% end if%>
  <div class="control-group">
 <label for="textfield" class="control-label"><B>Messaggio:</B></label>
   <div class="controls">
     <%if (strcomp(categoria,"Feedback")=0) and (session("admin")=true) then
         stile="input-block-level"
         else
         stile="ckeditor"
       end if%>
     <% if sMsg="" then %>
 	  <textarea class='<%=stile%>' rows="5" NAME = "MESSAGE" cols="40" placeholder="Inserisci commento" ></textarea>
       <%else%>
         <textarea class='<%=stile%>' rows="5" NAME = "MESSAGE" cols="40" placeholder="Inserisci commento"><%=sMsg%></textarea>
       <%end if%><hr>

			  <%if (strcomp(categoria,"Feedback")=0) and (session("admin")=true) then%>
           <INPUT TYPE = "button" class="btn-primary" NAME = "SubmitReply" VALUE = "Invia feedback" onClick="validate_feedback(<%=i%>)" >
          	<hr>
           <INPUT TYPE = "button" class="btn-error"  VALUE = "Consulta feedback" onClick='location.href="<%=Request.ServerVariables("HTTP_REFERER")%>"' >
				<%else%>

				<%end if%>


   </div>

 </div>

 <%if not ((strcomp(categoria,"Feedback")=0) and (session("admin")=true)) then%>

   <% if (strcomp(ucase(session("CodiceAllievo")),"OSPITE")<>0) then %>
  <div class="control-group"><center>
  <% if Session("Zip")<>1 then %>
 <span class="sottotitolo"><a title="Carica foto" href="#" onClick="javascript:PopUpWindow(600,300,<%=scegli%>);return false;">  <img src="img/caricaimg.png" width="39" height="39"></a>
 </span>  &nbsp;&nbsp;&nbsp;&nbsp;
 <%end if%>

  <span class="sottotitolo"><a title="Carica file"  href="#" onClick="javascript:PopUpWindow2(650,300,<%=scegli%>);return false;"> <img src="img/caricafile.jpg" width="35" height="33"></a>
  <!--&nbsp;&nbsp;&nbsp;&nbsp;
  <span class="sottotitolo"><a title="Condividi da Drive"  onclick="apridrive()"> <img src="https://www.shadowsplace.net/wp-content/uploads/2015/01/Google-Drive-icon.png" width="35" height="33"></a>-->
 </span><br>
  	<% if session("CaricatoFile")=true then %>
 	 Risorse:

 			   <B><%=session("NomeFileForum2")%></B>
     <%end if%>

     <% if Session("Zip")=1 then %>

   <code>
       Questa discussione prevede la consegna di un sito .zip   <br> il nome della cartella compressa <b>non deve avere spazie bianchi</b><br> e la cartella deve contenere il file <b>index.html</b>
        </code>

  <% end if%>


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
                                     <li class="active" ><a href="#profileUe" data-toggle="tab">Umanet Explorer</a></li>
                                     <li class="active"><a href="#profileIn" data-toggle="tab">Interfacce</a></li>

                             </ul>
                             <div id="myTabContent2" class="tab-content">

                               <div class="tab-pane fade  in active" id="profileEm">


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

                             <img src="smilies/on_20.gif" align="absmiddle"  title=':zz' onclick='javascript:addsmile(":zz")'>

 							  <!--#include file = "include/smilies.inc"-->
 							</div>
 						</div>
 					</div>
 				</div>
                  <!-- >fine form -->

                            </div>






                                <div class="tab-pane fade  in active" id="profileCo">



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



                                <div class="tab-pane fade  in active" id="profileNa">

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



                                <div class="tab-pane fade  in active" id="profileIn">



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

                                  <div class="tab-pane fade  in active" id="profileUe">



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
      <!--   </div>   -->

                                         <div class="accordion-heading">
 											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseMail"><center>

                                                 <img id="notify-email" class="imground" title='Notifica per email'  src="../../img/icon_mail.jpg" width="50px" height="40px" align="absmiddle" style="border-color:none;">
                                                 </center>
 											</a>
 										</div>

 										<div id="collapseMail" class="accordion-body collapse">
 											<div class="accordion-inner">
                                             <label style="color:#000"><b>Notifica per email</b> </label>

                                             <input type="checkbox" onClick="document.getElementById('cbEmail00').checked=true; checkTutti();" value="yes" name="cbEmail00" id="cbEmail00" title="Selezionare per non inviare notifiche ">  Nessuno &nbsp;&nbsp;&nbsp;  <br>
                                             <input type="checkbox" onclick="document.getElementById('cbEmail00').checked=false" value="no" name="cbEmail0" id="cbEmail0" title="Selezionare per inviare un email solo al mittente">   solo al mittente &nbsp;&nbsp;&nbsp;  <br>
                                                 <input type="checkbox" onclick="document.getElementById('cbEmail00').checked=false" value="no" name="cbEmail1" id="cbEmail1" title="Selezionare per inviare un email a chi partecia alla discussione">   Notifica per email a chi ha commentato &nbsp;&nbsp;&nbsp;
 											<br>  <input type="checkbox" onclick="document.getElementById('cbEmail00').checked=false" value="no"  name="cbEmail2" id="cbEmail2" title="Selezionare per inviare un email alla classe">    a tutta la classe &nbsp;&nbsp;&nbsp;<br>
                                                <input type="checkbox" onclick="document.getElementById('cbEmail00').checked=false" value="no"   name="cbEmailProf" id="cbEmailProf" title="Selezionare per inviare un email al prof.">   al prof. &nbsp;&nbsp;&nbsp;


                      						 </div>

 										</div>

                                         <% if session("Admin")=true then %>
                                          <div class="accordion-heading">
 											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion3" href="#collapseCompito"><center>

                                                  <i class="icon-pencil"></i>
                                                 </center>
 											</a>
 										</div>
                                         <div id="collapseCompito" class="accordion-body collapse">
 											<div class="accordion-inner">
                                             <label style="color:#000"><b>Inserisci come compito in </b> </label>
                                               <iframe src="../cMessaggi/compilapreavviso.asp" name="postmessage" id="postmessage" width="100%" height="40%" frameborder="0" SCROLLING="no" border="0" class="iframe">
       </iframe>


                                                <input type="checkbox"  checked="false"    name="cbCompito" id="cbCompito" title="Selezionare per inserire il compito nel paragrafo">   Inserisci compito &nbsp;&nbsp;&nbsp;
                                                <input type="checkbox"  checked="false"    name="cbImg" id="cbImg" title="Selezionare per richiedere upload immagine">   Richiede immagine &nbsp;&nbsp;&nbsp;
                                                <input type="checkbox"  checked="false"    name="cbFile" id="cbFile" title="Selezionare per richiedere upload file">   Richiede file &nbsp;&nbsp;&nbsp;
                                                 <input type="text" name="date3" id="datepicker1" class="input-medium datepick" value"gg/mm/aaaa"/>




                      						 </div>

 										</div>


                                         <% end if%>

 										</div>

    <% end if '  if (strcomp(ucase(session("CodiceAllievo")),"OSPITE")<>0)
%>
<P>
<CENTER>
<!-- <INPUT TYPE = "button"  VALUE = "Sostituisci" onClick="sostituisci();">-->
<br>

<INPUT TYPE = "button" NAME = "SubmitReply" VALUE = "Invia" onClick="newpreview()" class="btn">
<br>
<%
    end if  'if not ((strcomp(categoria,"Feedback")=0) and (session("admin")=true)) then
    %>
 									</div>

 <script language="javascript" type="text/javascript">
 function newpreview() {
 	    sostituisci();
         document.InputForm.action = "PreviewMessage.asp?scegli=<%=scegli%>&Reply=1&RCount=<%=RCount%>&byChiamante=1&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>";
 		document.InputForm.submit();
 }
  </script>

   <script type="text/javascript">


 $(window).load(function () {

 	   $('#notify-email').click();
 	    $('#cbEmail00').click();
 	//   $('#cbEmail1').click();
 	//   $('#cbEmailProf').click();


 	   // event.stopPropagation();

 	});

 </script>


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

   <script type="text/javascript" src="../js/refresh_session.js"></script>
 	<script language="javascript" type="text/javascript">

 function newpost() {
 	    sostituisci();
 	    var autore=InputForm.Name.value;
 		var commento=InputForm.Topic.value;
 		var messaggio=InputForm.MESSAGE.value;
 		if (autore=="")
 	     {
 		   alert("Non hai scritto l'autore ");
 		   return 0;
 		}
 		 else
 		 if (commento=="")
 		{
 		   alert("Non hai scritto l'argomento ");

 		   return 0;
 		}
 	// else
 		// if (messaggio=="")
 		//{
 		  // alert("Non hai scritto il messaggio.");

 		  // return 0;
 		//}
 	else
 	{

     document.InputForm.action = "PreviewMessage.asp?Reply=1&scegli=<%=scegli%>&SubmitMessage=1&codBacheca=<%=bacheca%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>";
 		document.InputForm.submit();
 	}

 }




function carica_feedback() {
var idpoli = $("#Polarita").val();
//alert(capitolo);

if (idpoli != "Seleziona polarità" && idpoli != null) {


	$.ajax({
		method: "POST",
		url: "carica_feedback.asp",
		dataType: "html",
		data: { id: idpoli }
	}) /* .ajax */
		.done(function (ans) {

			//alert(ans);
			$("#Topic").html("<option>Seleziona un feedback</option>" + ans);

		}) /* .done */
		.error(function (jqXHR, textStatus, errorThrown) {
			alert(jqXHR + "\n" + textStatus + ": " + errorThrown);
		});

	document.getElementById("Topic").disabled = false;
	document.getElementById("Punteggio").disabled = false;
	 

} else {
	document.getElementById("Topic").disabled = true;
	document.getElementById("Punteggio").disabled = true;
	$("#selpar").html("<option>Seleziona un paragrafo</option>");
	$("#selsottopar").html("<option>Seleziona un sottoparagrafo</option>");
}

}





 function validate_feedback(n) {
   //alert("n="+n);
     var selezionato=false;
     var pronto=true;
     var i;
	 if ((document.getElementById('Polarita').value=='Seleziona polarità')  && (document.getElementById('Polarita').value=='')) {
       alert('Non hai selezionato la polarità');
       pronto=false;
     }
     else if ((document.getElementById('Topic').value=='Seleziona feedback')  && (document.getElementById('Topic1').value=='')) {
       alert('Non hai selezionato il feedback');
       pronto=false;
     }
     else if ((document.getElementById('Topic1').value!='')  && ((document.getElementById("newfeedneg").checked==false) && (document.getElementById("newfeedpos").checked==false))) {
     alert('Non hai selezionato il segno del nuovo feedback');
     pronto=false;
   } else {
   }

     if (document.getElementById('Punteggio').value=='Seleziona punteggio'){
       alert('Non hai selezionato il punteggio');
       pronto=false;
     }

     for (i=1;i<n;i++){
      if  (document.getElementById('stud_'+i).checked==true)
         selezionato=true;
     }
 		if  (selezionato==false)
 	     {
 		   alert("Non hai selezionato nessun studente  ");
 		   pronto=false;
 		}
 	else if (pronto==true)
 	{

     document.InputForm.action = "invia_feedback.asp?scegli=<%=scegli%>&RCount=<%=RCount%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>";
 		document.InputForm.submit();
 	}

 }

  </script>
   <script type="text/javascript" src="../js/refresh_session.js"></script>

 	</body>

 </html>
