<%@ Language=VBScript %>
<!doctype html>
<html>
<head>
   
   <title>Chatroom</title>   
   
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

  
<script>
<!-- Hide from older browsers...

var blnSoundOff = false;

//toggle the smilies box
function toggleEmo(objDiv)
{
	var objBox = document.getElementById(objDiv);

	if (objBox.style.display == "none")
		objBox.style.display = "";
	else
		objBox.style.display = "none";

}

function logOut() {
	// window.location="showChat.asp?id_classe=" + <%=Session("Id_Classe")%> + "&divid=" + <%=Session("divid")%> + "&cartella="+<%=Session("Cartella")%>
	 window.location="../../home.asp" ;  
}
// -->
</script>

<script src="../SpryAssets/SpryTabbedPanels.js" type="text/javascript"></script>
<link href="../SpryAssets/SpryTabbedPanels.css" rel="stylesheet" type="text/css" />
 
   
</head>

 
<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
 
	<div id="navigation">
     <% 
'on error resume next

Response.Buffer = True %>
<% Response.Expires = -1 %>
 
<% Response.CacheControl = "Public" %>


        <% 
		
 ' AGGIUNTO DA DEFAUL
dim nomeChat, url1,registra,origine,fso
dim idx,numSmile,numImg,numTot
'Dim esecuzione
'set esecuzione = New TestServer ' oggetto di classe per testare dove gira il sito
   classe=request.QueryString("cartella")
   session("Classe")=classe
   

' necessitano di visibilità globale
nomeChat=year(date()) &"_"& month(date()) &"_" & day(date())&"_"& left(FormatDateTime(now(),4),2)&"_"&right(FormatDateTime(now(),4),2) &".txt" 
nomeChat=Replace(nomeChat,":","_")
homesito="/expo2015/UECDL"  
url1=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Session("cartella") & "/Chatlog/" & nomeChat  
url1=Replace(url1,"\","/")
                
				' dim objFSO,objCreatedFile,url
'				 ' per registrare la chta
'				'Dim sRead, sReadLine, sReadAll, objTextFile
'				Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logChat.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(url1)
'				objCreatedFile.Close	
				


' 
 
 ' carico i codici per le immagini in array per evitare di accedere al db per ogni messaggio perchè non fnziona utenti on line

		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")    
		%> 
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
      
  		<!-- #include file = "../service/controllo_sessione.asp" -->
		<!-- #include file = "../include/navigation.asp" -->
        
        <!--#include file="functions/functions_users.asp"-->
		<!--#include file="functions/functions_chat.asp"-->
        	  
          
         
	</div>
    
  <%
  strUsername=Session("CodiceAllievo")
Session("Username") = strUsername
Session("lastMessage") = lastMessageID()

Session("registra") = False

Session("FormatText") = True


If CheckUsername(strUsername) Then 
'Response.Redirect("../../home.asp")
	%><script> window.close();</script>
<%End If

%>

<%' AGGIUNTO
'Get the array
'E? QUI CHE NN ENTRA E QUINDI DICE NESSUN UTEN ON LINE PERCHE DIMENSIONA A 0 IL VETTORE
If IsArray(Application(ApplicationUsers)) Then
	saryActiveUsers = Application(ApplicationUsers)
Else
	ReDim saryActiveUsers(6, 0)
End If

Call RemoveUnActive()
'INGHIPPO QUA PORCO
'If UBound(saryActiveUsers, 2) = 0 Then
'	'Call Reset()
'
'	Response.Write(vbCrLf & "Nessun utente online porco")
'Else
'	Dim intArrayPass
'
'	For intArrayPass = 1 To UBound(saryActiveUsers, 2)
'		Response.Write(vbCrLf & saryActiveUsers(1, intArrayPass) & "<br>")
'	Next
'End If
%>  
    
    
	<div class="container-fluid" id="content">
    
      <!-- #include file = "../include/menu_left.asp" -->
         
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Chatroom 
 </h1> 
                    
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
							<a href="#more-files.html">Umanet</a>
							<i class="icon-angle-right"></i>
						</li>
						<li>
							<a href="#more-blank.html">Chat</a>
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
				        <h3> <i class="icon-reorder"></i>  Connesso come  <% = strUsername %>  </h3>
			          </div>
				      <div class="box-content">
                      
 
 
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
                   
                   
 
		    <div class="box-content"> 
                     
                      
                      
                      <table class="table ">
 <tr>
  <td bgcolor="#FFFFFF" rowspan="4" bordercolor="#333399">&nbsp;</td>
  <td bgcolor="#FFFFFF" width="70%" class="chatBorderText"> ChatRoom (digita <i>/commands</i> nel message box per vedere i comandi)</td>
  <td bgcolor="#FFFFFF" rowspan="3" width="1"></td>
  <td bgcolor="#FFFFFF" width="150" class='hidden-480'>Utenti Online</td>
  <td bgcolor="#FFFFFF" rowspan="4">&nbsp;</td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" height="450" style="padding: 0px; border: 1px solid #1B467D">
  <iframe src="chat.asp" width="100%" height="140%" id="chatframe" frameborder="0" SCROLLING="no" border="5" class="iframe"></iframe></td>
  <td bgcolor="#FFFFFF" height="530" style="padding: 0px; border: 1px solid #1B467D">
  <iframe src="users.asp" width="100%" height="525" id="users" frameborder="0" SCROLLING="no" border="0" class="iframe"></iframe></td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" height="9"></td>
  <td bgcolor="#FFFFFF" height="9"></td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" colspan="3" height="45" style="padding: 0px; border: 1px solid #1B467D">
  <iframe src="message.asp" name="postmessage" id="postmessage" width="100%" height="100%" frameborder="0" SCROLLING="no" border="0" class="iframe"></iframe></td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" colspan="5" align="right">
  <table cellpadding="0" cellspacing="0" width="100%">
   <tr>
    <td align="left" class="copyright"> <%=Session("CodiceAllievo")%></td>
    <td><input type='checkbox' name='turnoffsound' value='True' checked onClick='chatframe.SoundOption()'> Attiva Suono Notifiche</td>
    <td align="right" class="LogOut"><a href="javascript:logOut()" title="Logout">Logout</a>
     <% if session("Admin")=true then%>
          -<a href="resetta_chat.asp">Reset Chat </a>
     <%end if%>
    </td>
   </tr>
  </table>
  </td>
 </tr>
</Table>

<div id="emoticonsNew" style="display:none; width:75%;">
 <!--#include file = "include/Tabbed_Panels.inc"-->  
 </div>
 
  


<div id="emoticonsNewPersonal" style="display:none; width:75%;">
 <% 'path = "../../3PC/img_social/Baldi/include/Tabbed_Panels.inc"
 'Server.MapPath("../../database/" & Session("DBCopiatestonline"))
 
 'ASSOLUTO
 'path=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("cartella")&_ 
 '"/img_social/include/Tabbed_Panels.inc" 
 
' path=replace(path,"\","/")
 'response.write("1="&path)
 'Lo server.exec lo vuole RELATIVO
path=Server.MapPath("../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"&Session("cartella")&_ 
 "/img_social/include/Tabbed_Panels.inc")
 path=replace(path,"\","/")
' response.write(path)
 response.write("Personal emoticons ... in costruzione")
'path = Server.MapPath("../../materia_1/" & Session("cartella") & "/img_social/include/Tabbed_Panels.inc")
'path = "../../materia_1/" & Session("cartella") & "/img_social/include/Tabbed_Panels.inc"
 
 

 
'Server.Execute(path)
' qua ci va include .... %>


</div>
                      
                      
                      
                      
               
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

