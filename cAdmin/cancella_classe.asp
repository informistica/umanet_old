<%@ Language=VBScript %>
<html>
<head>	
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
    
    <!-- Datepicker new-->
	<link rel="stylesheet" href="../../css/plugins/datepicker/datepicker.css">
    
    


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
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />
       
       
       <!-- PLUpload -->
	<script src="../../js/plugins/plupload/plupload.full.js"></script>
	<script src="../../js/plugins/plupload/jquery.plupload.queue.js"></script>
	<!-- Custom file upload -->
	<script src="../../js/plugins/fileupload/bootstrap-fileupload.min.js"></script>
	<script src="../../js/plugins/mockjax/jquery.mockjax.js"></script>
    
    
    
   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />
-->
  
<!-- Datepicker --> 

<!-- <script src="../js/plugins/datepicker/bootstrap-datepicker.it.js"></script> -->
  
  <script src="../../js/jquery-ui.js"></script>   
 <script src="../../js/datapicker_it.js"></script> 

	<style>
	<!--
	 li.MsoNormal
		{mso-style-parent:"";
		margin-bottom:.0001pt;
		font-size:12.0pt;
		font-family:"Times New Roman";
		margin-left:0cm; margin-right:0cm; margin-top:0cm}
	-->
	</style>
	
	<style>
.loader {
display: block;
position: fixed;
left: 0px;
top: 0px;
width: 100%;
height: 100%;
z-index: 9999;
background: #fafafa url(../image/page-loader.gif) no-repeat center center;
text-align: center;
color: #999;
}
</style>
 
</head>


   <% Response.Buffer=True 
   Dim ConnessioneDB, ConnessioneDB1, rsTabella,rsTabella1, QuerySQL,StringaConnessione,URL,RecSet
   on error resume next
   'Id_Mod=Request.QueryString("Id_Mod")
   Id_Classe=request.querystring("Id_Classe")
   Classe=Request.QueryString("Classe")
    divid=request.QueryString("divid")
   
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
	
       %>   
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
   
<%  
 ' mi servirà per cancellare la cartella risorse                             

'url=Server.MapPath(homesito)&"/"&Classe 
url = Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe
url=Replace(url,"\","/")
'response.write(url)

QuerySQL = "SELECT Url_calendar FROM Classi WHERE ID_Classe = '"&Id_Classe&"';"
response.write "<br>"&QuerySQL
set rsCal = ConnessioneDB.Execute(QuerySQL)

idcal = rsCal("Url_calendar")
response.write "<br>Calendario: "&idcal
 
     QuerySQL ="DELETE   FROM Classi WHERE Id_Classe ='" &Id_Classe&"';"
'response.write(QuerySQL)
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE   FROM Classi_Moduli_Paragrafi WHERE Id_Classe  ='" & Id_Classe & "';"
	  ConnessioneDB.Execute(QuerySQL)
	  

	   QuerySQL ="DELETE  FROM preFrasi WHERE ID_Paragrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM preDomande WHERE ID_Paragrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM preNodi WHERE ID_Paragrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)
	 
	  QuerySQL ="DELETE  FROM Moduli WHERE ID_Mod  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)
      
	    QuerySQL ="DELETE  FROM Sottoparagrafi WHERE ID_Sottoparagrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)
	 
	      QuerySQL ="DELETE  FROM ParagrafiSottoparagrafi WHERE ID_Paragrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)

     QuerySQL ="DELETE  FROM Paragrafi WHERE ID_Paragrafo  LIKE '" & Classe & "%';"
	 ConnessioneDB.Execute(QuerySQL)


Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FolderExists (url) then
    fso.DeleteFolder(url)
	 ' response.write("La cartella è stata cancellata : <br> "& url )
'se esiste la cartella risorse la cancello
else
   response.write("La cartella non esiste : <br> "& url )
   end if

'response.write(url)

' per ogni studente della classe devo cancellare tutto 
QuerySQL =" Select CodiceAllievo FROM Allievi WHERE Classe ='" &Classe&"';"
'response.write(QuerySQL)
set rsTabella = ConnessioneDB.Execute(QuerySQL)
do while not rsTabella.eof
    
	
	
	
	 QuerySQL ="DELETE  FROM Domande WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Nodi WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Frasi WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM M_Desideri WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM M_Navigazione WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM M_Topolino WHERE Id_Stud ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	 QuerySQL ="DELETE  FROM Risultati WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM Risultati1 WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	  QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
 QuerySQL ="DELETE  FROM FILE_FORUM WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
     QuerySQL ="DELETE  FROM ALLIEVI WHERE CodiceAllievo ='" &CodiceAllievo&"';"
	 ConnessioneDB.Execute(QuerySQL)
	
	
	
	
	
	
	rsTabella.movenext
loop
rsTabella.close : set rsTabella=nothing

' quan ci starebbe bene una compattata al db


'On Error Resume Next
If Err.Number = 0 Then

     if Session("DB")=1 then
%>
<div class="loader"></div>

			   <script language="javascript">
				   // window.alert("Classe cancellata!");
					//location.href="../../home.asp"
					 				</script>

      <%else%>
<div class="loader"></div>
 <script language="javascript">
				    //window.alert("Classe cancellata!");
					//location.href="../../home.asp"
					 				</script>
<% end if
' lo tolgo perchè se cancello me stessa come classe non sa come dove ritornare
	'if Request.ServerVariables("HTTP_REFERER") <>"" then 
'			response.Redirect request.serverVariables("HTTP_REFERER") 
'	end if

	   QuerySQL ="DELETE FROM CAT_CAT WHERE Id_Classe ='" & Id_Classe & "';"
	  response.write QuerySQL
	  ConnessioneDB.Execute(QuerySQL)
	  



ConnessioneDB.Close : Set ConnessioneDB = Nothing 
ConnessioneDB1.Close : Set ConnessioneDB1 = Nothing 

%>


<script>
	 $.ajax({
						method: "POST",
						url: "../../../../googleapi/delcalendario.php?calendario=<%=idcal%>",
						dataType: "html",
						data: {  }
					}) /* .ajax */
					.done(function( ans ) {
								
						//alert(ans);
						window.location.href = "../../home.asp";
						
					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});
	 
	 </script>


<%

Else
Response.Write Err.Description 
Err.Number = 0
End If





   %>
	   
<div id=piede_pagina>
				<p><p>
				
				<!-- REDIRECT INTELLIGENTE al posto del Select case Session("Stato") -->
<h3><a href="../../home.asp"> Torna all'Home Page </a></h3> 
	
  
			</div>	 

 <!-- se il login è corretto richima la pagina per inserire le domande del test -->
 
	</body>
	</html>
	