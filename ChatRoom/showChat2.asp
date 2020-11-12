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
	<!-- CKEditor -->
	<script src="../../js/plugins/ckeditor/ckeditor.js"></script>
    
	
	
	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />

  
<script src="../SpryAssets/SpryTabbedPanels.js" type="text/javascript"></script>
<link href="../SpryAssets/SpryTabbedPanels.css" rel="stylesheet" type="text/css" />
 <script type="text/javascript">
 
function addsmile(codice) {
	 
		with (document.frmMessage) { 
		 
		 
		  messaggio.value= messaggio.value + codice;
		 
	    }	
}
 

 </script>
 

 <script src="_assets/js/jquery-1.4.4.min.js" type="text/javascript"></script>
<script src="_assets/js/jquery.zclip.js"></script>
<script src="include/copiaincolla.js"></script>
   <script src="_assets/js/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 
    
   
</head>

<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed" onLoad="cambiaSessione();">
 

	<div id="navigation">
     
        <% 
		
 
	 
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")     
		%> 
        <!-- #include file = "../service/controllo_sessione.asp" -->
       
        <!-- #include file = "../var_globali.inc" --> 
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
           
		<!-- #include file = "../include/navigation.asp" -->
        <!--#include file="functions/functions_chat.asp"-->
        	  
          
         
	</div>
  
  <% 
  cartella=request.QueryString("cartella")
set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn

function inHTML(sReadAll)
   sReadAll=replace(sReadAll,"[color","<font color")
   sReadAll=replace(sReadAll,"[/color]","</font>")  
   sReadAll=replace(sReadAll,"[i]","<i>")
   sReadAll=replace(sReadAll,"[/i]","</i>")
   sReadAll=replace(sReadAll,"[b]","<b>")
   sReadAll=replace(sReadAll,"[/b]","</b>")
   sReadAll=replace(sReadAll,"]",">")
   inHTML=FormatMessage(sReadAll) 
end function



function durata(h,m,s)
 if h>0 then
    durata= h&"h"
 end if
  if m>0 then
    if h>0 then
    	durata=durata&":" & m &"min"
	else
	   durata= m &"min"
	end if
 end if 
 if s>0 then
    if m>0 then
    	durata=m&"min:" & round(m/s) &"sec"
	else
	   durata= s &"sec"
	end if
 end if 
 

end function

   ID_Chat=Request.QueryString("ID_Chat")
   Dim objFSO, objTextFile
  Dim sRead, sReadLine, sReadAll
  Const ForReading = 1, ForWriting = 2, ForAppending = 8
  
   QuerySQL="Select * from CHAT_SESSION where ID_Chat=" & ID_Chat &""
   Set rsTabella0 = ConnessioneDB.Execute(QuerySQL)   
   ore=DateDiff("h",rsTabella0("Inizio"),rsTabella0("Fine")) 
   minuti=DateDiff("n",rsTabella0("Inizio"),rsTabella0("Fine")) 
   secondi=DateDiff("s",rsTabella0("Inizio"),rsTabella0("Fine")) 
 
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & rsTabella0("cartella") & "/Chatlog/" & rsTabella0("Nome")  
   url=Replace(url,"\","/")
   Set objTextFile = objFSO.OpenTextFile(url, ForReading)
   sReadAll = objTextFile.ReadAll
   sReadAll1=sReadAll
   'sReadAll=url
   'response.write(url)
	'response.write(inHTML(sReadAll))
	objTextFile.Close
	'registrataresponse.write(url)
	daShowChat2=1 'serve per include che deve selezionare in base all'inclusione da qui o da nuovo messaggio
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
				        
                          
			          </div>
				      <div class="box-content">
                      
 
 									 
	 
	 
				 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
		   			  <div class="box-content"> 
                     
                  
                  
                  
                  
                 
<P>


 <TABLE class="table">
    <thead>
    <TR >
    <Th><B><FONT COLOR = "RED"><%=rsTabella0("Titolo")%></FONT></B></Th>
   
    <Th ALIGN = CENTER><B><FONT COLOR = "RED"><%=durata(ore,minuti,secondi)%></FONT></B></Th></TR>
    
    <tr style="border-bottom:inset;"><td colspan="2">
	<p><%=Response.write(FormatMessage(sReadAll))%> </p>
    </td>
    </tr></table>
    
   <br><br><center>
  <%if session("Admin")=true then%>
  
  
<div class="accordion" id="accordion5">
<div class="accordion-group">      
                                        <div class="accordion-heading">
											<a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion5" href="#collapseMail5"><center>
												
                                                <i class="icon-edit" title="Modifica"></i>
                                                </center>
											</a>
										</div>
										<div id="collapseMail5" class="accordion-body collapse">
											<div class="accordion-inner">
 

 <a title="Modifica testo" href="#">Modifica</a> 
   
    <form class="form-horizontal" name="frmMessage" action="aggiorna_chat.asp?ID_Chat=<%=ID_Chat%>&nome=<%=rsTabella0("Nome")%>&cartella=<%=cartella%>" METHOD = "POST">
     
    <br><b>Titolo : <br></b><br>
    <input type="text" name="txtTitolo" class="input-xxlarge" value="<%=rsTabella0("Titolo")%>"><br>
    <br><b>Messaggio :</b> <br></div>
    <textarea class='ckeditor span12' name="messaggio" cols="60" rows="40"><%=sReadAll1%></textarea>
    <br> 
    <p>
 
<center>    
<a href="#" onClick="Effect.toggle('dEmo','BLIND'); return false;">
<img title='Inserisci emoticons' src="smilies/icon-smilie.gif" align="absmiddle" style="border-color:blue">
</a> 
</center>
<center>
 
<!--#include file = "include/Tabbed_Panels.inc"-->

 
</center>
<br> <center>
       <input type="submit" value="Aggiorna"><br><br><hr style="width:35%"> </center>
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>


									    
                     						 </div>                       
										</div>
                                     </div>

</div>



   <%end if%>               
                  
                  
                  
                      
                      
               
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
         
<script>
 $(window).load(function () {	   
	  $('#FissaTopBar').click();
	 // $('#FissaSideBar').click();
	  
	  // $('#FissaSideBar').click();
		//alert('Finestra caricata completamente, compresa la grafica');   
	
	});
</script>
			 
	</body>
    

 </html>

