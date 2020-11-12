<%
 Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
  'Apertura della connessione al database
  
 ' db = Request.QueryString("db")
 ' response.write("DB="& db)
%>
<!-- #include File="resizecheck.asp" -->
 
<!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
<!-- #include file = "../var_globali.inc" -->

<!-- #include file = "../service/controllo_sessione.asp" -->
  

  
<%  
 pippo=request.QueryString("pippo")
'imgPath = Server.MapPath(".") & "\img"	'I suppose your images will be saved in an "img" folder which is child of the current folder
'imgPath = Server.MapPath(homesito) & "\img"

'studPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo")
'Set fso = CreateObject("Scripting.FileSystemObject")
'if not(fso.FolderExists (studPath)) then
'	fso.CreateFolder (studPath) 
'	response.write("<br>Creata dir : " &studPath)
'end if

'imgPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/img"  
 	 imgPath=Server.MapPath(homesito)& "/DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/img"  
	 
     imgPath=Replace(imgPath,"\","/")

	 thumbPath=Server.MapPath(homesito)& "/DB"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/Profili/thumb"
	 thumbPath=Replace(thumbPath,"\","/")
	 session("cartella")=request.QueryString("cartella")
	 Set fso = CreateObject("Scripting.FileSystemObject")
	 if  fso.FileExists (thumbPath&"/"& Session("CodiceAllievo")&".jpg") then
			
			
 
			set OggFile = fso.GetFile (thumbPath&"/"&Session("CodiceAllievo")&".jpg")
			'OggFile.Copy destinazione,true
		    OggFile.Delete
			'response.write("<br>Cancellato dir : " &imgPath)
	 end if

%>
<%
Dim pageUpload
pageUpload = "upload.asp"
If Not CheckResizeLib Then
	pageUpload = "upload.aspx"
End if
%>
<html>
<head>
<title>Aggiorna immagine profilo</title>
<script>
	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>


 
  


<link rel="stylesheet" type="text/css" href="../../stile.css">

</head>
<body> 
<div class="contenuti_forum" style="width:75%">
  
 <% ' per passare Session("DBCopiatestonline") non so perche ma querystring non va

 
'Dim objFSO,objCreatedFile
'Const ForReading = 1, ForWriting = 2, ForAppending = 8
'Dim sRead, sReadLine, sReadAll, objTextFile
'Set objFSO = CreateObject("Scripting.FileSystemObject")
' 
''Create the FSO.
'url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/"&Session("Cartella")&"/db.txt" 
'url=Replace(url,"\","/")
''response.write("ulr spiegazio="&url)
'Set objCreatedFile = objFSO.CreateTextFile(url, True)
'response.write("url="&url)
'line=Session("DBCopiatestonline")
'objCreatedFile.WriteLine(line)
'objCreatedFile.Close
 %>
 
 
<form name="theform" action="save_img_profilo.asp?">
	<!-- I suppose you have a save.asp script that you will use to save all data in this form -->
	 <table id="zebra_forum1" >
	<tr><td>Immagine:&nbsp;</td><td> <input readonly="readonly" name="img" / size="30"><input type="hidden" name="thumb" />
	<input type="button" value="Carica foto" onClick="uploadImgWindow(this.form.name, 'img', 'thumb', '<%= Server.URLEncode(imgPath) %>', '<%= Server.URLEncode(thumbPath) %>', this.form.img.value, 200, 200, 120, 92);" /></td></tr>
    <tr class="invisibile"><td>Nome:&nbsp;&nbsp;&nbsp;</td><td><input type="hidden" name="nomeimg" size="40" value="<%=Session("CodiceAllievo")%>.jpg"></td></tr> 
    <br/>
	<!-- I resize the image with the max with of 191 and the max height of 144.
	I create a thumbnail with the max with of 120 and the max height of 92.-->
	 
     
    </table>
 <%
 uploaded=request.QueryString("uploaded") 
 if uploaded=1 then %>   
    <font color="green"></font> Modifiche salvate, adesso aggiorna la pagina<br><br>
 <%else%>
    <font color="#FF0000">N.B.</font> Dopo il caricamento salva le modifiche e ricarica la pagina<br><br>
 <%end if%>
    <input type="submit"  value="Salva"  />
  <!--   <input type="button" value="Chiudi" onClick="window.close();" />-->
</form>
<%

if uploaded=1 then%>
<font color=green> Caricamento  del file <%=Session("FileName")%> completato</font>
<%
'   url=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/img/" & urlimg
'   url=Replace(url,"\","/")
'
'   response.Write("<font color=green> Immagine caricata! Salva i dati per completare</font>")
'   %>
 <!--  <img src="<%=Session("urluploaded")%>">-->
   <%
'   session("uploaded")=false
'   response.write(session("FileName"))
uploaded=0
end if

%>
 
  
<br>
 
</div>

 
</body>