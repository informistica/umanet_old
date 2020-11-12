<%@ Language=VBScript %>
<%   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
on error resume next
' DA SISTEMARE
' questa versione permette la modifica della categoria e di nascondere le categorie agli studenti,
' ad esempio quelle degli anni passati.
' ma ha il problema che in alcuni casi allo studente compare la bacheca o il diario vuoti, senza nessuna categoria

Response.AddHeader "Refresh", "600"
 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
 id_social=cint(scegli)
 select case scegli
 case "0"
     session("social")="forum"
	 icon="icon-group"
 case "1"
    session("social")="lavagna"
	 icon="icon-bullhorn"
 case "2"
    session("social")="diario"
	 icon="icon-book"
   case "3"
      session("social")="interrogazioni"
     icon="icon-question-sign"
 end select %>



<%
  divid=Session("divid")
  cartella=request.querystring("cartella")
  id_classe=Session("id_classe")
  bacheca=request.querystring("bacheca")
  id_categoria=request.querystring("id_categoria")
  session("id_categoria")=id_categoria
  categoria=request.querystring("categoria")
  session("categoria")=categoria
  cognome=request.querystring("cognome")
  nome=request.querystring("nome")
  if bacheca<>"" then
     Session("Bacheca")=bacheca
	 if Session("CognomeBacheca")="" then ' per evitare circolo vizioso
	 	'cognome=request.querystring("cognome")
	 	'nome=request.querystring("nome")
		 Session("CognomeBacheca")=cognome
	    Session("NomeBacheca")=nome
	 else

	     ' cognome=Session("CognomeBacheca")
	 	 ' nome =Session("NomeBacheca")
	 end if

  else
     Session("Bacheca")=""
	 Session("CognomeBacheca")=""
	 Session("NomeBacheca")=""
  end if
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


%>
<!doctype html>
<html>
<head>

  <title> Home <%
  if strcomp(ucase(session("social")),"LAVAGNA")=0 then
    response.write("BACHECA")
  else
  response.write(ucase(session("social")))
  end if %>  </title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	   <meta charset="utf-8"/>
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />


	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <!-- dataTables -->
	<link rel="stylesheet" href="../../css/plugins/datatable/TableTools.css">
<!-- chosen -->
	<link rel="stylesheet" href="../../css/plugins/chosen/chosen.css">

     <link rel="stylesheet" href="../../css/style-themes.css">
        <link rel="stylesheet" href="../../css/docs.css">

	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
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


	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />






   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />-->

  <script language="javascript" type="text/javascript">
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!!")
location.href="../../home.asp"
//location.href=window.history.back();
 }
 </script>

<%
' Session("DB")=1
'
'		  Session("Cognome") = "Ospite"
'		  Session("Nome") = "Ospite"
'		  Session("CodiceAllievo")="ospite"
'		  Session("Username")= "ospite" ' per la chat dopo disastro
'		 ' Session("DataTest") = DataTest
'		  Session("stile")="blue"
'
'		  Session("cartella")="Expo"
'		  session("Id_Classe")="6COM"
'		  Session("Admin")=False
'		  session("ID_Materia")="materia_1"
'
'		    app=1
'  			materia="Umanet 1"
' 			' cartella="Expo"
' 			 id_materia=1
' 			 Session("idxMat") =id_materia
'			 Session("Materia")="Umanet 1"
'			 session("DBCopiatestonline")="ok"
	Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
%>

</head>
 		<!-- #include file = "../var_globali.inc" -->
     	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    	<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->

<%if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY <!--onLoad="showText2();"-->> </BODY>
  <% else %>
<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">

<%end if%>

	<div id="navigation">

        <%

  Cartella=Request.QueryString("Cartella")
  TitoloCapitolo=Request.QueryString("Capitolo")
  Paragrafo=Request.QueryString("Paragrafo")
  Modulo=Request.QueryString("Modulo")
  CodiceTest = Request.QueryString("CodiceTest")
  'CodiceAllievo = Request.Cookies("Dati")("CodiceAllievo")
 ' Cognome=Session("Cognome")
  'Nome=Session("Nome")
   ' CognomeBacheca=request.QueryString("cognome")
 ' NomeBacheca=request.QueryString("nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")

		%>

  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->



	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="<%=icon%>"></i>
							<%
                          if strcomp(ucase(session("social")),"LAVAGNA")=0 then
                            response.write("BACHECA")
                          else
                          response.write(ucase(session("social")))
                          end if %>
                      </h1>
                      <% if session("DB")=1 and strcomp(session("Username"),"ospite")<>0 then%>
                        <a title="Condividi link alla pagina" href="#" onClick="javascript:PopUpWindow(600,400,<%=scegli%>);return false;"><i class="glyphicon-share_alt"> </i> <small>Condividi</small> </a>
                      <% end if%>

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

                         <%
							if id_categoria<>"" then%>
							<li><i class="icon-angle-right"></i>
							<a   href="#"><span></span><%=categoria%></a>
						    </li>
							<%end if
							%>

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
				        <h3> <i class="icon-reorder"></i>
                        <%if Bacheca<>"" and (strcomp(Bacheca,Session("CodAdmin"))<>0) then %>
                        BACHECA DI.   <%=request("cognome") & " " & request("nome")%>
                        <%else%>
                         <% if id_categoria<>"" then%>
                         ATTIVITA' IN <b>"<%=ucase(categoria)%>" </b>
                         <%else%>
                          AREE DI PROGETTO
                         <%end if%>
						  <%=left(classe,1+len(classe)-instr(classe,"$")) %>

                        <%end if%>

                        </h3>
			          </div>
				      <div class="box-content">


				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">

          	<div id="modal-3" class="modal hide fade" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true">
			<div class="modal-header">
				<button type="button" class="close" data-dismiss="modal" aria-hidden="true">×</button>
				<h3 id="myModalLabel">Modal header</h3>
			</div>
			<div class="modal-body">
				<p>One fine body…</p>
			</div>
			<div class="modal-footer">
				<button class="btn" data-dismiss="modal" aria-hidden="true">No</button>
				<button class="btn btn-primary" data-dismiss="modal">Yes</button>
			</div>
		</div>


                 <center>


<FORM class='form-horizontal form-striped' ACTION = 'forum_search.asp?scegli=<%=scegli%>&divid=<%=divid%>&id_classe=<%=id_classe%>&cartella=<%=cartella%>&bacheca=<%=bacheca%>'><b> Cerca nel Forum: <img src="img/icon_aim.gif"></b>
    <input type="text" name="search" size="25">
 <input type="submit" value="Cerca" name="searchbutton" class="btn" >
</Form>




 <%



iPageSize = 20
iPage = cint(Request.QueryString("Page"))
if iPage = 0 then iPage = 1


'sSQL = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"';"
'cmd.CommandText = sSQL
 'cmd.CommandText = "MESSAGETHREADS"
'cmd.CommandType = 4
'rs.open cmd, , 1, 3


if (id_categoria="") then
' faccio la query per selezionare le catgorie
 	  sSQL = "select count(*) from CAT_THREADS where Id_Classe='"&Id_Classe&"'  and Id_Social="&id_social
	  sSQL1 = "select * from CAT_THREADS where Id_Classe='"&Id_Classe&"'and Id_Social="&id_social &" order by ID_Categoria desc"

else

'if Session("Bacheca")= true  then
 if Bacheca<>""  then
 ' devo visualizzare la bacheca di uno studente
     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Bacheca='"&Session("Bacheca") &"' or CodiceAllievo='"&Session("Bacheca") &"' and Id_Social=0 ;"

	 sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Bacheca='"&Session("Bacheca") &"' and Bacheca<>'"&Session("CodAdmin")&"' or CodiceAllievo='"&Session("Bacheca") &"' and Topic<>'' or ID in (Select Id_Post from Condividi WHERE CodiceAllievo='"&Session("Bacheca")&"') order by LASTTHREADPOST desc;"

  else
     if strcomp(scegli,"0")=0 then
  ' se va in forum devo distinguere per le bacheche
     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and Bacheca='"&Session("CodAdmin") &"' and comments<>'InizializzaDB' and Id_Social=0 and Id_Categoria="& cint(id_categoria)&";"
	  sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and Bacheca='"&Session("CodAdmin") &"' and comments<>'InizializzaDB' and Topic<>'' and Id_Social=0 and Id_Categoria="& cint(id_categoria)&" order by LASTTHREADPOST desc;"
	else
     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"'  and comments<>'InizializzaDB' and Id_Social="& id_social &" and Id_Categoria="& cint(id_categoria)&";"
	sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Topic<>'' and Id_Social="&id_social&" and Id_Categoria="& cint(id_categoria)&" order by LASTTHREADPOST desc;"
 	end if
  end if
end if

'response.write(sSQL&"<br>"&scegli)
'response.write(sSQL1&"<br>"&scegli)

cmd.CommandText = sSQL
set rs = cmd.Execute
conn.execute sSQL
lTotalRecords=rs(0)
rs.close
' response.write("<br>lTotalRecords="&lTotalRecords)
rs.Open sSQL1, conn, 1,3


 
if not rs.Eof and not rs.bof then
'response.write("pippo:"&iPageSize)
' calcola il numero di pagine necessarie in base al numero di post da mostrare
iTotalPages = int(lTotalRecords / iPageSize)
 
'response.write("ciao"&iTotalPages)
	if lTotalRecords MOD iPageSize <> 0 then iTotalPages = iTotalPages + 1
	' se basta una pagina
		if lTotalRecords <=  iPageSize then
			rs.MoveFirst
			bOnePage = true
			lPageEnd = lTotalRecords
			lPageStart = 1
			iTotalPages = 1
		else
			lPageStart = ((iPage - 1) * iPageSize) + 1
			lPageEnd = lPageStart + (iPageSize - 1)


			if lPageEnd >= lTotalRecords Then
				lPageEnd = lTotalRecords
				bLastPage = true
			end if
			' posiziona il recordset in base alla pagina da visualizzare
			if iPage > 1 then
				rs.AbsolutePosition = ((iPage - 1) * iPageSize) + 1
			else
			' se ce una sola pagina va all'inizio
				rs.MoveFirst
			end if
		end if

	else
		bNoRecords = true

	end if

%>

</SELECT> </FORM>




<% if id_categoria<>"" then%>

<FORM ACTION = "new_post.asp?scegli=<%=scegli%>&bacheca=<%=bacheca%>&cognome=<%=request.QueryString("cognome")%>&nome=<%=request.QueryString("nome")%>" METHOD = "POST">
<INPUT TYPE = "SUBMIT" VALUE = "Nuova attivit&agrave;" class="btn-primary"></TD></FORM>
<%else%>
 <%If session("Admin")=true then%>
<FORM ACTION = "new_categorie.asp?scegli=<%=scegli%>" METHOD = "POST">
<INPUT TYPE = "SUBMIT" VALUE = "Nuova categoria" class="btn-primary"></TD></FORM>
  <%end if%>
<%end if%>





<P>
<%
if not bNoRecords then
    response.write "<P><B>Pagina " & iPage & " di " & iTotalPages & "</B><P>"
end if
%>

<!---->

<% if bNoRecords then%>
<table class="table table-hover table-nomargin table-bordered table-striped">
<%else%>
  <!-- parte che da errore per i menu  dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped">    -->
  <table class="table table-hover table-nomargin table-bordered dataTable dataTable-fixedcolumn dataTable-scroll-x table-striped">
<%end if%>
<thead>

<TR>
<% if id_categoria<>"" then%>
<Th><B> N. </B></Th>
<Th><B> Info </B></Th>
<Th><B> Argomento </B></Th>
<Th><B> Autore </B></Th>

<Th  class='hidden-480' ALIGN = CENTER><center><B>Risposte</center></FONT></B></Th>
<Th class='hidden-480' ><B>Ultimo post </B></Th>
<Th class='hidden-480' ><B>Visualizzazioni</B></Th>
<%else%>
<Th><B> N. </B></Th>
<Th><B> Descrizione </B></Th>
 <Th  class='hidden-480' ALIGN = CENTER><center><B>Progetti</center></FONT></B></Th>
 <Th ><B>Ultimo post </B></Th>
<%end if%>





<%
if (Session("Admin")=true) or (strcomp(bacheca,Session("CodiceAllievo"))=0) then%>
<Th  class='hidden-480'><B><center>Modifica - Elimina</B></center></Th></TR></thead>
<% else%>
</TR></thead>
<%end if %>
<%if bNoRecords then
 if id_categoria<>"" then
	  if Session("Admin")=true then
	 response.write "<tr><TD COLSPAN = 6><B>Non ci sono attivit&agrave;</B></TD></tr>"
	   else
		response.write "<tr><TD COLSPAN = 5><B>Non ci sono attivit&agrave;</B></TD><tr>"
	   end if
  else
     if Session("Admin")=true then
	 response.write "<tr><TD COLSPAN = 6><B>Non ci sono categorie di attivit&agrave;</B></TD><tr>"
	   else
		response.write "<tr><TD COLSPAN = 5><B>Non ci sono  categorie di attivit&agrave;</B></TD><tr>"
	   end if
  end if
else
numrow=1
for lCtr = lPageStart to lPageEnd

 if (not rs.Eof) and ( (rs("Attiva")=1) or (Session("Admin")=true))  then

		' conto il numero di discussioni per ogni categoria
 sSQLC="select count (*) from forum_messages where Id_Categoria="&rs(0) &" and ParentMessage=0 and Id_Classe='"&Id_Classe&"'"
' response.write(sSQLC)
 cmd.CommandText = sSQLC
 set rs1 = cmd.Execute
' rs1.Open sSQLC, conn, 1,3
 numDis=rs1(0)

 ' conto il numero di visualizzazioni totali dei post della categoria
 sSQLC="select sum (Visualizzazioni) from forum_messages where Id_Categoria="&rs(0) &" and Id_Classe='"&Id_Classe&"'"
 'response.write(sSQLC)
 cmd.CommandText = sSQLC
 set rs1 = cmd.Execute
' rs1.Open sSQLC, conn, 1,3
 numVisTot=rs1(0)



'leggo data ultimo ultimo post
 		' conto il numero di discussioni per ogni categoria
 'sSQLC="select max (DatePosted) from forum_messages where Id_Categoria="&rs(0) &" and ParentMessage=0 and Id_Classe='"&Id_Classe&"'"
  sSQLC="select max (DatePosted) from forum_messages where Id_Categoria="&rs(0) &" and Id_Classe='"&Id_Classe&"'"
 'response.write(sSQLC)
 cmd.CommandText = sSQLC
 set rs1 = cmd.Execute
' rs1.Open sSQLC, conn, 1,3
 lastPost=rs1(0)
 rs1.close



	 if id_categoria<>"" then
	   if (Session("Admin")=true) or (rs("Visibile")<>0) or (rs("Attiva")=0) then
		 response.write "<tr class="&classe_riga&"><td>"&numrow&"</td>"
		 response.write("<td>  <a data-original-title='"&rs("Topic")&"' href='#' class='btn' rel='popover' data-trigger='hover' title='' data-placement='right' data-content='"&rs("Abstract")&"'><center>  <i class='icon-question-sign'></i></center></a></td>")

		 response.write" <TD><A HREF='ShowMessage.asp?categoria=" & categoria &"&id_categoria=" & id_categoria &"&nome=" & request.QueryString("nome") &"&cognome="&request.QueryString("cognome")&"&scegli="&scegli&"&bacheca="&bacheca&"&ID=" & rs("ID") & "&Zip=" & rs("Zip")&"&RCount=" & rs("ReplyCount")& "&TParent=" & rs("ID")& "&divid=" & divid & "&id_classe=" & id_classe & "&visibile=" & rs("Visibile") & "&privato=" & rs("Privato") & "'>"  & rs("Topic") & "</A></FONT></TD>"
		 response.write " <TD>"
		if session("Admin")=true then ' se sono admin visualizzo il codice autore post
	   response.write "<A title='" & rs("CodiceAllievo") &"' HREF = '#'>" & rs("AuthorName") & "</A>"
	 else
	 response.write "<A HREF = '#'>" & rs("AuthorName") & "</A>"
	 end if

		response.write "</FONT></TD>"


	response.write "</TD><TD  class='hidden-480' ALIGN = CENTER><center>" & rs("ReplyCount") & "</FONT></center></TD>"
	response.write "</TD><TD>" & rs("LastThreadPost") & "</FONT></TD>"
	response.write "</TD><TD>" & rs("Visualizzazioni") & "</FONT></TD>"

	if (Session("Admin")=true) or (strcomp(ucase(bacheca),ucase(Session("CodiceAllievo")))=0) then
	ID=rs("ID")
	%>
	<TD  class='hidden-480' align=center><center><A HREF="DeleteMessage.asp?discussione=1&scegli=<%=scegli%>&bacheca=<%=bacheca%>&ID=<%=ID%>&cognome=<%=Request("cognome")%>&nome=<%=Request("nome")%>&categoria=<%=categoria%>&id_categoria=<%=id_categoria%>"> <i class=" icon-trash" onClick="return window.confirm('Vuoi veramente cancellare questa discussione ?');"  ></i></a></center></TD></TR>
	<%
	else%>
	</TR>
	<%end if
  end if ' if (Session("Admin")=true) or (rs("Visibile")<>0) then

else
 response.write "<tr class="&classe_riga&"> <TD><A HREF='default0.asp?nome=" & request.QueryString("nome") &"&cognome="&request.QueryString("cognome")&"&categoria="&rtrim(rs(2))&"&id_categoria="&rs(0)&"&id_classe=" & id_classe & "&cartella="&cartella&"&scegli="&scegli&"&bacheca="&bacheca&"'>"&lCtr&"</A></TD>"
	 response.write " <TD><A title='" & rs(0) &"' HREF = 'default0.asp?nome=" & request.QueryString("nome") &"&cognome="&request.QueryString("cognome")&"&categoria="&rtrim(rs(2))&"&id_categoria="&rs(0)&"&id_classe=" & id_classe & "&cartella="&cartella&"&scegli="&scegli&"&bacheca="&bacheca&"'>" & rs(2) & "</A></TD>"
	  response.write " <TD style='text-align:center'>"& numDis&" </TD>"
	  response.write " <TD style='text-align:center'>"&lastPost&"</TD>"


		if (Session("Admin")=true) or (strcomp(bacheca,Session("CodiceAllievo"))=0) then
		'response.write("id_cat="&id_categoria)
		response.write "<TD  class='hidden-480' ALIGN = CENTER><center>"&_
    "<a href='#modal-1' onClick='modifica("&rs(4)&","&rs(0)&","""&rs(2)&""")' data-toggle='modal'><i style='text-decoration:none' class='icon-pencil' title='Modifica categoria'></i></a>"&_
    "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&_
    "<a href='#modal-2' onClick='elimina("&rs(0)&","""&rs(2)&""")' data-toggle='modal'><i style='text-decoration:none' class='icon-trash' title='Elimina categoria'></i></a>"
	 if rs("Attiva")=0 then
	    response.write "&nbsp;&nbsp;<i class='icon-eye-close' title='Non visibile agli studenti'></i>"
	 else
	    response.write()
	 end if
	 response.write("</center></TD></TR>")
		else
		response.write "</TR>"
		end if

end if ' if id_categoria



numrow=numrow+1
rs.movenext
end if ' eof
 Next
end if
response.write "</TABLE>"

if bOnePage = false and bNoRecords = false then

response.write "<br><br><TABLE id=zebra_forum1 WIDTH = WIDTH = '35%'  ><TR><TD WIDTH = '50%'>"

if iPage > 1 then
sPrevQuery = "Page=" & iPage - 1
response.write "<A HREF = 'default.asp?scegli="&scegli&"&divid="&divid &"&cartella="&cartella&"&id_classe="&id_classe&"&"& sPrevQuery & "'><B><< Pagina precedente</B></A></p>"
  else
response.write "&nbsp;"
end if

response.write "</TD><TD VALIGN = TOP NOWRAP>"

'iTotalPages
if (bLastPage = false) then

sNextQuery = "Page=" & iPage + 1
response.write "<A HREF = 'default.asp?scegli="&scegli&"&divid="&divid &"&cartella="&cartella&"&id_classe="&id_classe&"&"& sNextQuery & "'><B>Pagina successiva >></B></A></p>"
else
response.write "&nbsp;"
end if

response.write "</TD></TR></TABLE>"
response.write "<P><CENTER><FONT SIZE =-1>"

for iCtr = 1 to iTotalPages
sPageQuery = "Page=" & iCtr & sQuery
if iCtr <> iPage then
 response.write "<A HREF = 'default.asp?scegli="&scegli&"&divid="&divid &"&cartella="&cartella&"&id_classe="&id_classe&"&"& sPageQuery & "'>"
else

 response.write "<B>"
end if
response.write iCtr

if iCtr <> iPage then
response.write "</A>"
else
response.write "</B>"
end if
if iCtr < iTotalPages then response.write "&nbsp;&nbsp;|&nbsp;&nbsp;"


Next
response.write "</FONT></CENTER>"
end if
%>











               <h6 align="center"><a href="#" onClick="javascript:window.close();"> Chiudi </a></h6>
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



	<form id="mod" action="modifica_categoria.asp" method="post">
			<div id="modal-1" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
				<div class="modal-header">
					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
					<h3 id="myModalLabel">Modifica categoria</h3>
				</div>
				<div class="modal-body">
				     <b>Id</b><br><input placeholder="ID" name="idcat" id="idcat" type="text" disabled class="input-xxsmall" style="width: 25%"><br>
					<b>Descrizione</b><br><input placeholder="Inserisci il nuovo nome" name="titolomodifica" id="titolomodifica" type="text" class="input-xxlarge" style="width: 97%"><br>
					<b>Attiva</b><br><input placeholder="1 Attiva 0 Disattiva" name="attiva" id="attiva" type="text" class="input-xxsmall" style="width: 25%"><br>

				</div>
				<div class="modal-footer">
					<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
					<button type="button" id="inviamodifica" class="btn btn-primary" onClick="controllamodifica()">Invia</button>
				</div>
			</div>
		</form>

    <form name="eli" id="eli" action="elimina_categoria.asp" method="post">
  			<div id="modal-2" class="modal hide" tabindex="-1" role="dialog" aria-labelledby="myModalLabel" aria-hidden="true" style="display: none;">
  				<div class="modal-header">
  					<button type="button" class="close" data-dismiss="modal" aria-hidden="true"><i class="icon-remove"></i></button>
  					<h3 id="myModalLabel">Elimina categoria</h3>
  				</div>
  				<div class="modal-body">
  					<b>Descrizione</b><br><input placeholder="Inserisci il nuovo nome" name="titoloelimina" id="titoloelimina" type="text" disabled class="input-xxlarge" style="width: 97%"><br>

  				</div>
  				<div class="modal-footer">
  					<button class="btn" data-dismiss="modal" aria-hidden="true">Chiudi</button>
  					<button type="button" id="inviamodifica" class="btn btn-danger" onClick="conferma_elimina()">Elimina</button>
  				</div>
  			</div>
  		</form>



		<script>

		function modifica(attiva,id,categoria){
		 //alert(categoria);
		    document.getElementById("idcat").value=id
			document.getElementById("titolomodifica").value=categoria
			document.getElementById("attiva").value=attiva
			document.getElementById("mod").action="modifica_categoria.asp?id="+id;
		}
    function elimina(id,categoria){
     //alert(categoria);
      document.getElementById("titoloelimina").value=categoria
      document.getElementById("eli").action="elimina_categoria.asp?id="+id;
    }


		function controllamodifica(){
			var titolo = document.getElementById("titolomodifica").value.trim();
			var attiva = document.getElementById("attiva").value.trim();
			 


			if ((titolo == "") || (attiva=="")) {
				alert("Il nome della categoria e lo stato sono obbligatorio");
			}else
			{
				document.getElementById("inviamodifica").type="submit";
			}

		}
  function conferma_elimina(){
  var domanda = confirm("Sei sicuro di voler cancellare la categoria?");
  if (domanda === true) {
    document.eli.submit();
  }else{
    alert('Operazione annullata');
  }

    }



		</script>

		<script>
function PopUpWindow(w,h,s) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

window.open('share.asp?scegli='+s,'share.asp?scegli='+s, 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=460,top='+wint+',left='+winl);

}
</script>

	</body>


 </html>
