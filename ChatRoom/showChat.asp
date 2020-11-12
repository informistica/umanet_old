<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Chat</title>   
   
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
       
       
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  


   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" onLoad="cambiaSessione();">
 

	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
       
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
<%

  set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn

  divid=request.querystring("divid")
  cartella=request.querystring("cartella")
  id_classe=request.querystring("id_classe")
 
 
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
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Chatroom </h1> 
                    
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
							<a href="#">Umanet</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Chat</a>
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
				        <center>
                          <FORM ACTION = 'forum_search.asp'><b> Cerca nelle Chat : <img src="../cSocial/img/icon_aim.gif"></b>
    <input type="text" name="search" size="25">
 <input type="submit" value="Cerca" name="searchbutton" disabled="true">
</Form>

			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			  <div class="box-content"> 
                     
                  
                  
                  
                  
                 
<P>


<%

 
iPageSize = 20
iPage = cint(Request.QueryString("Page"))
if iPage = 0 then iPage = 1


'sSQL = "select * from MESSAGETHREADS where Id_Classe='"&Id_Classe&"';"
'cmd.CommandText = sSQL
 'cmd.CommandText = "MESSAGETHREADS"
'cmd.CommandType = 4
'rs.open cmd, , 1, 3





sSQL = "select count(*) from CHAT_SESSION where Id_Classe='"&Id_Classe&"' ;"
cmd.CommandText = sSQL
set rs = cmd.Execute	
conn.execute sSQL
lTotalRecords=rs(0)

sSQL = "select * from CHAT_SESSION where Id_Classe='"&Id_Classe&"' order by Inizio desc  ;"
cmd.CommandText = sSQL
set rs = cmd.Execute	
conn.execute sSQL

'set rs = cmd.Execute


if not rs.Eof and not rs.bof then
'rs.MoveLast non supportto per le mie query
'lTotalRecords = rs.RecordCount
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

</SELECT></TD></FORM>
</center>
<center>
<FORM class="form-horizontal" onClick="PopUpWindow(409,481)" ACTION = "chatroom.asp?id_classe=<%=id_classe%>&divid=<%=divid%>&cartella=<%=cartella%>" target="ChatWindow2" METHOD = "POST" >

 
<TD align="center">
<%
 QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe") &"'"
 Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
 'response.Write(QuerySQL)
 ' response.write("chat="&rsTabella1("ChatAbilitata"))
 if rsTabella1("ChatAbilitata")=0 then %>
<INPUT disabled="disabled" TYPE = "SUBMIT" class="btn-primary" VALUE = "Inizia nuova Chat"></TD></FORM>
<%else%>
<INPUT  TYPE = "SUBMIT" class="btn-primary" VALUE = "Inizia nuova Chat."></TD></FORM>
</center>
<%end if
rsTabella1.close
set rsTabella1=nothing
connessioneDB.close
set connessioneDB=nothing
%>

     </tr>



</TABLE><P>
<%
if not bNoRecords then
    response.write "<P><B>Pagina " & iPage & " di " & iTotalPages & "</B><P>"
end if
%>
<TABLE class="table table-hover table-nomargin">
<thead>
<TR>
<Th><B>Titolo</B></Th>
<Th><B>Inizio</FONT> </Th>
<Th ALIGN = CENTER><B>Fine</B></Th>
 
<%
if Session("Admin")=true then%>
<Th><B>Elimina </B></Th></TR></thead>
<%
else%>
</TR></thead>
<%end if 
if bNoRecords then
 response.write "<TD COLSPAN = 4><B>Non ci sono chat nello storico</B></TD>"

else
 for lCtr = lPageStart to lPageEnd
 if (lCtr mod 2) = 0  then 
	    classe_riga="zebra-dispari"
	else
	    classe_riga=""
end if	
 response.write "<tr class="&classe_riga&"> <TD><A HREF='ShowChat2.asp?ID_Chat=" & rs("ID_Chat") &"&cartella="&cartella&"'>"  & rs("Titolo") & "</A></FONT></TD>"
 response.write " <TD>" 
 if session("Admin")=true then ' se sono admin visualizzo il codice autore post
   response.write "<A title='" & rs("ID_Chat") &"' HREF = '#'>" & rs("Inizio") & "</A>" 
 else
 response.write "<A HREF = '#'>" & rs("Inizio") & "</A>" 
 end if

response.write "</FONT></TD>"

response.write "</TD><TD ALIGN = CENTER>" & rs("Fine") & "</FONT></TD>"
'response.write "</TD><TD>" & rs("Fine") & "</FONT></TD>"
if Session("Admin")=true then
ID_Chat=rs("ID_Chat")
%>

 
<TD align=center><A onClick="return window.confirm('Vuoi veramente cancellare la Chat?');" HREF="cancella_chat.asp?ID_Chat=<%=ID_Chat%>&nome=<%=rs("Nome")%>"> X</a></TD></TR> 
<%
else%>
</TR> 
<%end if 

 rs.movenext
 Next
end if
response.write "</TABLE>"

if bOnePage = false and bNoRecords = false then

response.write "<TABLE WIDTH = '100%'><TR><TD>&nbsp;</TD></TR><TR><TD WIDTH = '10%'>&nbsp;</TD><TD WIDTH = '60%'>"
 
if iPage > 1 then
sPrevQuery = "Page=" & iPage - 1
response.write "<A HREF = 'default.asp?" & sPrevQuery & "'><B><< Previous Page</B></A>"
  else
response.write "&nbsp;"
end if
		
response.write "</TD><TD VALIGN = TOP NOWRAP>"

if bLastPage = false then
		
sNextQuery = "Page=" & iPage + 1 
response.write "<A HREF = 'default.asp?" & sNextQuery & "'><B>Next Page >></B></A>"
else
response.write "&nbsp;"
end if
response.write "<TD WIDTH = '10%'>&nbsp;</TD>"
response.write "</TD></TR></TABLE>"
response.write "<P><CENTER><FONT SIZE =-1>"

for iCtr = 1 to iTotalPages
sPageQuery = "Page=" & iCtr & sQuery
if iCtr <> iPage then
 response.write "<A HREF = 'ShowChat.asp?" & sPageQuery & "'>"
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
        
        <!-- #include file = "include/colora_pagina.asp" -->
         

			 
	</body>

 </html>

