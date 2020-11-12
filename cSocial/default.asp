<%@ Language=VBScript %>
<%   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
'on error resume next
Response.AddHeader "Refresh", "600"
 scegli=request.QueryString("scegli") ' 0 = forum 1=lavagna 2=diario
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


<!--#include file = "../service/controllo_sessione.asp"-->



<%
  divid=Session("divid")
  cartella=request.querystring("cartella")
  id_classe=Session("id_classe")
  bacheca=request.querystring("bacheca")
  if bacheca<>"" then
     Session("Bacheca")=bacheca
	 if Session("CognomeBacheca")="" then ' per evitare circolo vizioso
	 	cognome=request.querystring("cognome")
	 	nome=request.querystring("nome")
		 Session("CognomeBacheca")=cognome
	 Session("NomeBacheca")=nome
	 else
	      cognome=Session("CognomeBacheca")
	 	  nome =Session("NomeBacheca")
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

  <title> Default <%=ucase(session("social"))%>  </title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />

	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">

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
	<!-- Form -->
	<script src="../../js/plugins/form/jquery.form.min.js"></script>
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
function showText2() {window.alert("La sessione è scaduta, effettua nuovamente il Login!")
location.href="../../home.asp"
//location.href=window.history.back();
 }
 </script>



</head>

<%if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
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
  Cognome=Session("Cognome")
  Nome=Session("Nome")
   ' CognomeBacheca=request.QueryString("cognome")
 ' NomeBacheca=request.QueryString("nome")
  by_UECDL=Request.QueryString("by_UECDL")
  dividA=request.QueryString("dividApro")
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
     	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
    	<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->



	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "../include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="<%=icon%>"></i>  <%=ucase(session("social"))%>  </h1>

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
							<a href="#">Classe</a>
							<i class="icon-angle-right"></i>
						</li>




                        <%select case scegli
							 case "0"
								 session("social")="forum"
							 %>
							<li>
							<a href="#">Forum</a>
						    </li>
							 <%
							 case "1"
							 %>
							 <li>
							<a href="#">Lavagna</a>
						    </li>
							 <%

							  case "2"

							 %>
							 <li>
							<a href="#">Diario</a>
						    </li>
							 <%

               case "3"
              %>
              <li>
             <a href="#">Interrogazioni</a>
               </li>
              <%
							 end select %>


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
                        BACHECA DI   <%=request("cognome") & " " & request("nome")%>
                        <%else%>
                         ATTIVITA'
                           CLASSE <%=cartella%>
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

<%'if session("Admin")=true then%>
<FORM class='form-horizontal form-striped' ACTION = 'forum_search.asp?scegli=<%=scegli%>&divid=<%=divid%>&id_classe=<%=id_classe%>&cartella=<%=cartella%>&bacheca=<%=bacheca%>'><b> Cerca nel Forum: <img src="img/icon_aim.gif"></b>
    <input type="text" name="search" size="25">
 <input type="submit" value="Cerca" name="searchbutton" class="btn">
</Form>
<%'end if%>



 <%



iPageSize = 20
iPage = cint(Request.QueryString("Page"))
if iPage = 0 then iPage = 1


'sSQL = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"';"
'cmd.CommandText = sSQL
 'cmd.CommandText = "MESSAGETHREADS"
'cmd.CommandType = 4
'rs.open cmd, , 1, 3




'if Session("Bacheca")= true  then
 if Bacheca<>""  then
 ' devo visualizzare la bacheca di uno studente
     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Bacheca='"&Session("Bacheca") &"' or CodiceAllievo='"&Session("Bacheca") &"' and Id_Social=0;"

	 sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Bacheca='"&Session("Bacheca") &"' and Bacheca<>'"&Session("CodAdmin")&"' or CodiceAllievo='"&Session("Bacheca") &"' and Topic<>'' or ID in (Select Id_Post from Condividi WHERE CodiceAllievo='"&Session("Bacheca")&"');"

  else
     if strcomp(scegli,"0")=0 then
  ' se va in forum devo distinguere per le bacheche
     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and Bacheca='"&Session("CodAdmin") &"' and comments<>'InizializzaDB' and Id_Social=0;"

	  sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and Bacheca='"&Session("CodAdmin") &"' and comments<>'InizializzaDB' and Topic<>'' and Id_Social=0;"
	else

     sSQL = "select count(*) from MESSAGETHREADS where Id_Classe='"&Id_Classe&"'  and comments<>'InizializzaDB' and Id_Social="&cint(scegli)&";"

	sSQL1 = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"' and comments<>'InizializzaDB' and Topic<>'' and Id_Social="&cint(scegli)&";"
 	end if


  end if

''response.write(sSQL&"<br>"&scegli)
'response.write(sSQL1&"<br>"&scegli)

cmd.CommandText = sSQL
set rs = cmd.Execute
conn.execute sSQL
lTotalRecords=rs(0)
rs.close
' response.write("<br>lTotalRecords="&lTotalRecords)
rs.Open sSQL1, conn, 1,3



if not rs.Eof and not rs.bof then

' calcola il numero di pagine necessarie in base al numero di post da mostrare
iTotalPages = int(lTotalRecords / iPageSize)
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





                  <FORM ACTION = "new_post.asp?scegli=<%=scegli%>&bacheca=<%=bacheca%>" METHOD = "POST">

<INPUT TYPE = "SUBMIT" VALUE = "Nuova attivit&agrave;" class="btn-primary"></TD></FORM>






<P>
<%
if not bNoRecords then
    response.write "<P><B>Pagina " & iPage & " di " & iTotalPages & "</B><P>"
end if
%>

<table class='table table-hover table-nomargin table-bordered table-striped'>


<thead>
<TR>
<Th><B> Argomento </B></Th>
<Th><B> Autore </B></Th>
<Th  class='hidden-480' ALIGN = CENTER><center><B>Risposte</center></FONT></B></Th>
<Th ><B>Ultimo post </B></Th>

<%
if (Session("Admin")=true) or (strcomp(bacheca,Session("CodiceAllievo"))=0) then%>
<Th  class='hidden-480'><B><center>Elimina</B></center></Th></TR></thead>
<% else%>
</TR></thead>
<%end if
if bNoRecords then
  if Session("Admin")=true then
 response.write "<TD COLSPAN = 5><B>Non ci sono attivit&agrave;</B></TD>"
   else
    response.write "<TD COLSPAN = 4><B>Non ci sono attivit&agrave;</B></TD>"
   end if
else
 for lCtr = lPageStart to lPageEnd
' if (lCtr mod 2) = 0  then
'	    classe_riga="zebra-dispari"
'	else
'	    classe_riga=""
'end if

if not rs.Eof  then
 if id_categoria<>"" then
   if (Session("Admin")=true) or (rs("Visibile")<>0) then
	 response.write "<tr class="&classe_riga&"> <TD><A HREF='ShowMessage.asp?nome=" & request.QueryString("nome") &"&cognome="&request.QueryString("cognome")&"&scegli="&scegli&"&bacheca="&bacheca&"&ID=" & rs("ID") & "&Zip=" & rs("Zip")&"&RCount=" & rs("ReplyCount")& "&TParent=" & rs("ID")& "&divid=" & divid & "&id_classe=" & id_classe & "&visibile=" & rs("Visibile") & "&privato=" & rs("Privato") & "'>"  & rs("Topic") & "</A></FONT></TD>"
	 response.write " <TD>"
 	if session("Admin")=true then ' se sono admin visualizzo il codice autore post
   response.write "<A title='" & rs("CodiceAllievo") &"' HREF = '#'>" & rs("AuthorName") & "</A>"
 else
 response.write "<A HREF = '#'>" & rs("AuthorName") & "</A>"
 end if

	response.write "</FONT></TD>"

	response.write "</TD><TD  class='hidden-480' ALIGN = CENTER><center>" & rs("ReplyCount") & "</FONT></center></TD>"
	response.write "</TD><TD>" & rs("LastThreadPost") & "</FONT></TD>"
	if (Session("Admin")=true) or (strcomp(ucase(bacheca),ucase(Session("CodiceAllievo")))=0) then
	ID=rs("ID")
	%>
	<TD  class='hidden-480' align=center><center><A HREF="DeleteMessage.asp?discussione=1&scegli=<%=scegli%>&bacheca=<%=bacheca%>&ID=<%=ID%>&cognome=<%=Request("cognome")%>&nome=<%=Request("nome")%>"> <i class=" icon-trash" onClick="return window.confirm('Vuoi veramente cancellare questa discussione ?');"  ></i></a></center></TD></TR>
	<%
	else%>
	</TR>
	<%end if
  end if ' if (Session("Admin")=true) or (rs("Visibile")<>0) then

else

end if ' if id_categoria


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



	</body>

 </html>
