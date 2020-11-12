<%@ Language=VBScript %>
<%
 Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 divid=Session("divid")
 id_classe= Session("Id_Classe")
 cartella=request.querystring("cartella")
 bacheca=request.querystring("bacheca")
 on error resume next
 %>
 <!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->
 
<!--#include file = "../service/controllo_sessione.asp"-->
 <%
 set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
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

Function prepStringForSQL(sValue)

Dim sAns
sAns = Replace(sValue, Chr(39), "''")


prepStringForSQL = sAns

End Function

if request("searchbutton") <> "Search" then bInValid = false

if not bInvalid then

dim sQueryStrings() 


sOrigQuery = Request("search") 
sSearchString = sOrigQuery & " "

iPos = 1
iCnt = 1
do While len(sSearchString)
  redim preserve sQueryStrings(iCnt)

	'find position of " "
	iPos = instr(iPos, sSearchString, " ") + 1

	'Add individual word to array
	sWord = Mid(sSearchString,1,iPos - 2)
	
	'Handle case where user enters more than 
        'one space between words

	if trim(sWord) <> "" Then
	  sQueryStrings(iCnt) = Mid(sSearchString,1,iPos - 2)
	
          'truncate search string to  eliminate newly added word
	  iCnt = iCnt + 1
	end if
	
	sSearchString = Mid(sSearchString, iPos)

	'reset
	iPos = 1
Loop
iCnt = uBound(sQueryStrings)


sChoice = Request("SearchChoice") '1 specific category from listcode.asp


sSQLString = "SELECT * FROM FORUM_MESSAGES_CLASSI WHERE "




	for iCtr = 1 to iCnt
	
		sSQLString = sSQLString & "(COMMENTS LIKE '%" & sQueryStrings(ictr) & "%' OR TOPIC LIKE '%" & sQueryStrings(ictr) & "%' OR AUTHORNAME LIKE '%" & sQueryStrings(ictr) & "%'  )"
	if iCtr < iCnt then sSQLString = sSQLString & " AND "
	next

sSQLString = sSQLString & " ORDER BY DATEPOSTED, TOPIC, AUTHORNAME"

cmd.CommandType = 1
cmd.CommandText = sSQLString
rs.Open cmd,, 1, 1
iRecCnt = rs.recordcount
if iRecCnt > 0 then
rs.MoveLast
rs.MoveFirst
iRecCnt = rs.recordcount
end if

if iRecCnt= 1 then
	
 sNewURL = "ShowMessage.asp?by_search=1&scegli="& rs("Id_Social")&"&cartella="&rs("Classe")&"&id_classe="&rs("Id_Classe")&"&ID=" & rs("ID") & "&categoria="&rs("Descrizione")&"&id_categoria="&rs("ID_Categoria")& "&Caption=One%20Match%20Found"
 rs.Close
 set rs = nothing
 set cmd = nothing
 conn.close
 Response.Redirect sNewURL
end if

end if 'bInvalid
%>
<!doctype html>
<html>
<head>
   
   <title>Ricerca</title>   
   
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
<body class='theme-<%=session("stile")%>'  data-layout-topbar="fixed">  

	<div id="navigation">
     
        <% 
		
 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        	  
          
         
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i>&nbsp;Cerca nei Social  </h1> 
                    
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
							<a href="#">Social</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#">Cerca</a>
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
				        <h3> <i class="icon-reorder"></i> Risultati della ricerca: <%=iRecCnt%> corrispondenze trovate   </h3>
			          </div>
				      <div class="box-content">
                      
 
 									 
	 <%
if iRecCnt = 0 then
Response.Write "<B>Nessun risultato trovato per " & sOrigQuery & ".<B><P>"
else

Response.Write "<table class='table table-striped table-invoice'>"
Response.Write "<thead>"
Response.Write "<tr>"
Response.Write "<th>Messaggio</th>"
Response.Write "<th>Data</th>"
Response.Write "<th>Autore</th>"	
Response.Write "<th>Classe</th>"										 
Response.Write "</tr>"
Response.Write "</thead>"
Response.Write "<tbody>"


i=0
do while not rs.EOF
 	
response.write "<TR><TD>"
  
Response.Write "<A HREF = 'ShowMessage.Asp?by_search=1&scegli="& rs("Id_Social")&"&cartella="&rs("Classe")&"&id_classe="&rs("Id_Classe")&"&ID=" & rs("ID") & "&categoria="&rs("Descrizione")&"&id_categoria="&rs("ID_Categoria")& "'>" & rs("Topic") & "</A>"
response.write "</TD><TD>"
response.write rs("DatePosted")
response.write "</TD><TD>"
response.write "<A HREF = '#'>" & rs("AuthorName") & "</A>"
response.write "</TD><td>" & rs("Classe") & "</td></TR>"
'Response.write "<BR>"
i=i+1
rs.MoveNext
loop
	
response.write "</TABLE>"	
end if

%>
<!--#include file = "database_cleanup.inc"-->
</td></tr></table><center>
 
	 
		 
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
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

