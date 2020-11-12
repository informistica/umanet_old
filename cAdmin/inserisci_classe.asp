<!-- richiama_test.asp -->
<%@ Language=VBScript %>
<%

  Dim Codice_Corso,Codice_Test, Capitolo, Paragrafo,Num,Nome,Cognome,Parag
  Id_Classe=Request.QueryString("Id_Classe")
  divid=request.QueryString("divid")
  Classe=Request.QueryString("Classe")
 ' posizione= Request.QueryString("posizione")
 ' response.Write("jhjhk="&posizione)
  Titolo = Request.Form("TxtTitolo")
  Num = Request.Form("TxtNum") ' numero di paragrafi che si vogliono inserire
  ID_Mod=Request.Form("txtID_Mod")
%>
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
/*
 li.MsoNormal
	{mso-style-parent:"";
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman";
	margin-left:0cm; margin-right:0cm; margin-top:0cm}
*/
</style>
<meta https-equiv="Content-Language" content="it">
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta https-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Inserisci Classe </title>

</head>
  <script language="javascript" type="text/javascript" >
 function validate2() {


 if (frmDocument.txtCla.value=="")
	{
	   alert("Non hai inserito il nome della nuova classe.");
	   frmDocument.txtCla.setfocus();
	   return 0;
	}
 else
 if (frmDocument.txtId_Cla.value=="")
	{
	   alert("Non hai inserito il codice della classe.");
	   frmDocument.imgname.setfocus();
	   return 0;
	}
	else
    if (frmDocument.txtPos.value=="")
	{
	   alert("Non hai inserito la posizione della classe.");
	   frmDocument.txtPos.setfocus();
	   return 0;
	}else
	{
	    document.frmDocument.action = "inserisci_classe.asp?num=1";
		document.frmDocument.submit();


    }

}
 </script>
<body bgcolor="#FFFFFF">
<div id="container">

<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
       %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
     <%
Classe= Request.Form("txtCla")
Id_Classe=Request.Form("txtId_Cla")
Posizione=Request.Form("txtPos")
num=Request.querystring("num")

%>
<% if num<>"" then

 Dim cartelle(13)

   cartelle(0)="Risorse"
   cartelle(1)="Verifiche"
   cartelle(2)="Profili"
   cartelle(3)="Lavagna" ' questo serviva per la lavagna statica si pu� anche rtogliere
   cartelle(4)="img_social"
   cartelle(5)="img_lavagna"
   cartelle(6)="img_forum"
   cartelle(7)="file_lavagna"
   cartelle(8)="file_forum"
   cartelle(9)="Chatlog"
   cartelle(10)="img_diario"
   cartelle(11)="file_diario"
  cartelle(12)="file_interrogazioni"


 QuerySQL="  INSERT INTO Classi (Id_Classe,Classe,Cartella,Posizione,Visibile)  SELECT '" & ID_Classe & "','" & Classe & "', '" & Classe & "', " & Posizione &",1;"
  ConnessioneDB.Execute QuerySQL
   QuerySQL="  INSERT INTO Setting (Privato,Valutato,TestAbilitato,In_Quiz,Max_In_Quiz,Id_Classe,CIAbilitato,DVAbilitato,JSAbilitato)  SELECT " & 1 & "," & 1 & ", " & 1 & "," & 1 & "," & 4 &  ",'" & Id_Classe &  "'," & 0 &  "," & 1 &  "," & 1  &";"
  ConnessioneDB.Execute QuerySQL

   QuerySQL="  INSERT INTO [dbo].[2ESERCITAZIONI_SINGOLI] (Descrizione,Data,Id_Classe)  SELECT 'Iscrizione','12/12/2112','"&Id_Classe &"';"
  ConnessioneDB.Execute QuerySQL

    QuerySQL="  INSERT INTO CAT_CAT (Id_Classe,Descrizione,Id_Social)  SELECT '" & ID_Classe & "','Generale',0;"
  ConnessioneDB.Execute QuerySQL
    QuerySQL="  INSERT INTO CAT_CAT (Id_Classe,Descrizione,Id_Social)  SELECT '" & ID_Classe & "','Programma',1;"
  ConnessioneDB.Execute QuerySQL
    QuerySQL="  INSERT INTO CAT_CAT (Id_Classe,Descrizione,Id_Social)  SELECT '" & ID_Classe & "','Compiti',2;"
  ConnessioneDB.Execute QuerySQL
  QuerySQL="  INSERT INTO CAT_CAT (Id_Classe,Descrizione,Id_Social)  SELECT '" & ID_Classe & "','Interrogazioni',3;"
ConnessioneDB.Execute QuerySQL

    Set fso = CreateObject("Scripting.FileSystemObject")
    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe
	url=Replace(url,"\","/")
	 ' creo la cartella della classe
	if fso.FolderExists (url) then
			 response.Write( "<br>La cartella " & url & " esiste gi�.<br>")
			fso.DeleteFolder (url)
			fso.CreateFolder (url)
		else
			fso.CreateFolder (url)
			response.Write( "<br>La cartella " & url & " � stata creata.<br>")
		end if


  for i=0 to 12 ' creo le cartelle dentro la cartella della classe
		url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/"&cartelle(i)
		url=Replace(url,"\","/")
		if fso.FolderExists (url) then
			' response.Write( "La cartella " & url & " esiste gi�.<br>")

		else
			fso.CreateFolder (url)
			if (i=2)or (i=4) or (i=5) or (i=6) or (i=10) then ' creo le sottocartelle
			    url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/"&cartelle(i)&"/img"
				url=Replace(url,"\","/")
				fso.CreateFolder (url)
				url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/"&cartelle(i)&"/thumb"
				url=Replace(url,"\","/")
				fso.CreateFolder (url)
				if (i=4) then
				  url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/"&cartelle(i)&"/include"
				url=Replace(url,"\","/")
				fso.CreateFolder (url)
				end if
			end if
			'response.Write( "La cartella " & url & " � stata creata.<br>")
		end if
    next

	' copio l'immagine del profilo vuoto
	' urlOrig=Server.MapPath(homesito)&"/img/profilo_vuoto.png"
'	 urlOrig=Replace(urlOrig,"\","/")
'	 urlDest=Server.MapPath(homesito)&"/"&Classe&"/Profili/img"
'	 urlDest=Replace(urlDest,"\","/")
' 	 fso.CopyFile urlOrig, urlDest
'	 urlOrig=Server.MapPath(homesito)&"/img/profilo_vuoto_thumb.png"
'	 urlOrig=Replace(urlOrig,"\","/")
'	 urlDest=Server.MapPath(homesito)&"/"&Classe&"/Profili/thumb/"
'	 urlDest=Replace(urlDest,"\","/")

	 folderorigine=Server.MapPath(homesito)&"/Profili" ' questa non va aggiornata perch� � cartella su root da cui importano gli altri
	 folderorigine=Replace(folderorigine,"\","/")
	 folderdestinazione=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Classe&"/Profili"
 	 folderdestinazione=Replace(folderdestinazione,"\","/")

	  set folder = fso.GetFolder (folderorigine)
	  folder.Copy folderdestinazione,true


    On Error Resume Next
	If Err.Number = 0 Then
		Response.Write "Inserimento della classe avvenuto! "
		if session("DB")=1 then
		'response.Redirect "../../home.asp"
        else
		'response.Redirect "../../home.asp"
		end if
	Else
		Response.Write Err.Description
		Err.Number = 0
	End If%>

     <div class="loader"></div>

	 <script>
	 $.ajax({
						method: "POST",
						url: "../../../../googleapi/addcalendario.php?classe=<%=Classe%>&idclasse=<%=Id_Classe%>",
						dataType: "html",
						data: {  }
					}) /* .ajax */
					.done(function( ans ) {

						window.location.href = "<%=Request.ServerVariables("HTTP_REFERER")%>";

					}) /* .done */
					.error(function( jqXHR, textStatus, errorThrown ){
					alert(jqXHR+"\n"+textStatus+": "+errorThrown);
					});

	 </script>


  <%else
 QuerySQL="SELECT * FROM Classi order by Id_Classe"
 Set rsTabella = ConnessioneDB.Execute(QuerySQL)

  %>

  <FIELDSET><LEGEND><b><%  response.write("Inserisci classe  ") %> </b></LEGEND><br>
<p class="titolo">

   <form method="POST"   name="frmDocument" >
    <div style="border:#FC0 solid 1px; padding:20px; width:30%;">
    <b>Classi gi&agrave; inserite</b><br><br>
     <table id="zebra_stud">
		 <%
		if  rsTabella.eof then
		id_classe_new="1COM"
		pos_classe_new=1
		%>
            <tr><td><b>Classi</b></td> </tr>
		    <tr><td>Nessuna </td>
		<%else
		    id_classe_new="" %>
			<tr><td><b>Id_Classe</b></td><td><b>Classe</b></td><td><b>Posizione</b></td> </tr>
			<%do while not rsTabella.eof%>
				   <tr><td><%=rsTabella.fields("Id_Classe")%></td>
				   <td><%=rsTabella.fields("Classe")%></td>
					<td><%=rsTabella.fields("Posizione")%></td>
			   </tr>

				   <% 'posizione=rsTabella.fields("Posizione")
					  rsTabella.movenext()
			loop
		end if %>
		</table>
    </div><br><br>
      <div style="border:#FC0 solid 1px; padding:20px; width:450px;">
   <br>  <b>Inserisci Nuova Classe</b>
    <br><br>
	<%
     ' determino l'ID della classe in base alle classi gi� inserite
	 if id_classe_new="" then
		  ' incremento di 1 rispetto alla chiave dell'ultima classe inserita
		  QuerySQL="SELECT * FROM Classi where Contatore=(Select max(Contatore) from Classi) ;"
		 ' response.write(QuerySQL)
		  Set rsTabella = ConnessioneDB.Execute(QuerySQL)
		  id_classe_new= 1+cint(left(rsTabella("ID_Classe"),len(rsTabella("ID_Classe"))-instr(rsTabella("ID_Classe"),"COM")))&"COM"

		'classe=rsTabella("Classe")
		 classe="5C$5"
		  'response.write("classe="&classe)
		 ' response.write("<br>instr="&instr(classe,"$"))
		  nome_classe_new="..$"&right(classe,len(classe)-instr(classe,"$"))

		  QuerySQL="SELECT * FROM Classi where Posizione=(Select max(Posizione) from Classi) ;"
		 ' response.write(QuerySQL)
		  Set rsTabella = ConnessioneDB.Execute(QuerySQL)

		  pos_classe_new=1 + rsTabella("Posizione")
	 end if
	 response.write("Nome della classe")%>
    <p><input type="text" name="txtCla" size="10" value="<%=nome_classe_new%>"></p>
    <div style="border:#FC0 solid 1px; padding:20px; width:50%;">
	<% response.write("ID della Classe ")%>
    <p><input type="text" name="txtId_Cla" size="7"  value="<%=id_classe_new%>"></p>
    <%response.write("Posizione della classe")%>
    <p><input type="text" name="txtPos" size="5"  value="<%=pos_classe_new%>"></p>
	 </div>
	  <p><input type="button"  value="Inserisci" onClick="return validate2();"></p>

	</p>

    </form>
   </div>
    <br><br>
       <div style="border:#FC0 solid 1px; padding:20px; width:450px;">
   <b>  Trasferisci Classe </b>
    <!-- Per utilizzare un modulo gi� esistente in altra classe -->
    <div style="overflow:scroll; height:900px;">
   <!--
    <iframe src="seleziona_origine_classe.asp" name="postmessage" id="postmessage" width="100%" height="100%" frameborder="0" SCROLLING="no" border="0" class="iframe"></iframe>
    --->
    </div>
  </div>

<% end if%>
</div>
</body>
</html>
