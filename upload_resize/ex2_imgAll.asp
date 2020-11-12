<%
 Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
  
%>
<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script> 
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
    
<!-- #include File="resizecheck.asp" -->
 
 
  <!-- #include file = "../stringhe_connessione/stringa_connessione_forum.inc" -->
<!-- #include file = "../var_globali.inc" -->
<!-- #include file = "../service/controllo_sessione.asp" -->
  
<%
'scegli=request("scegli")
'session("scegli")=scegli
 select case session("scegli")
 case "0" 
     session("social")="forum"
	 icon="icon-group"
  
 case "1" 
 
    session("social")="lavagna"
	 icon="icon-bullhorn"
  case "2" 
    session("social")="diario"
	 icon="icon-book"
  
 end select 
 
 
'imgPath = Server.MapPath(".") & "\img"	'I suppose your images will be saved in an "img" folder which is child of the current folder
'imgPath = Server.MapPath(homesito) & "\img"

'studPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo")
'Set fso = CreateObject("Scripting.FileSystemObject")
'if not(fso.FolderExists (studPath)) then
'	fso.CreateFolder (studPath) 
'	response.write("<br>Creata dir : " &studPath)
'end if

'imgPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/"  & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/img"  

'response.write(session("Cartella") & "---" &session("cartella") & "-classe=" &session("Classe")&"<br>")
'response.write(request.Cookies("Cartella") & "---" &request.Cookies("cartella") & "-classe=" &request.Cookies("Classe"))

if session("Cartella")="" then
  session("Cartella")=request.QueryString("Cartella")
end if

imgPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("Cartella") &"/img_"&session("social")&"/img"  
imgPath=Replace(imgPath,"\","/")
'thumbPath = Server.MapPath(".") & "\thumb"	'I suppose your images will be saved in a "thumb" folder which is child of the current folder
'thumbPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/thumb"
thumbPath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("Cartella") &"/img_"&session("social")&"/thumb"
thumbPath=Replace(thumbPath,"\","/")
'includePath=Server.MapPath(homesito)& "/Db"&Session("DB")&"/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_forum/include"
'includePath=Replace(includePath,"\","/")
'response.write(imgPath)
' controllo se esistono le directori dello stud, altrimenti le creo
Set fso = CreateObject("Scripting.FileSystemObject")
if not(fso.FolderExists (imgPath)) then
	fso.CreateFolder (imgPath) 
	response.write("<br>Creata dir : " &imgPath)
end if
Set fso = CreateObject("Scripting.FileSystemObject")
if not(fso.FolderExists (thumbPath)) then
	fso.CreateFolder (thumbPath) 
	response.write("<br>Creata dir : " &thumbPath)
end if

'Set fso = CreateObject("Scripting.FileSystemObject")
'if not(fso.FolderExists (includePath)) then
'	fso.CreateFolder (includePath) 
'	response.write("<br>Creata dir : " &includePath)
'end if



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
<title>Carica immagine in <%=session("social")%> </title>
<script>
	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>


 
  


<link rel="stylesheet" type="text/css" href="../../stile.css">

</head>
<body> 
 
 
 

<FORM  name="theform" onSubmit = 'return Validate()' METHOD = "POST" class='form-horizontal form-bordered'> 
<div class="control-group">
<label for="textfield" class="control-label"><B>Immagine:</B></label>
  <div class="controls">
	  
     <input readonly="readonly" name="img"  size="30"  placeholder="Scegli l'immagine da caricare" />
    <input type="hidden" name="thumb" />
    <input type="hidden" name="scegli" value="<%=session("scegli")%>" />
	<input id="caricaimg" type="button" class="btn"  value="Upload" onClick="uploadImgWindow(this.form.name, 'img', 'thumb', '<%= Server.URLEncode(imgPath) %>', '<%= Server.URLEncode(thumbPath) %>', this.form.img.value, 640, 480, 120, 92);" />
    	<!-- I resize the image with the max with of 191 and the max height of 144.
	I create a thumbnail with the max with of 120 and the max height of 92.-->
  </div>
</div>

<div class="control-group">
<label for="textfield" class="control-label"><B>Commento:</B></label>
  <div class="controls">
	  
      <input type="text" name="nomeimg" value="..." class="input-xxlarge" placeholder="">
  </div>
</div>
 
 
 <div class="control-group">
<label for="textfield" class="control-label"><B>Link to:</B></label>
  <div class="controls">
      <input  class="input-xxlarge" placeholder="Inserisci url a cui vuoi collegare questa immagine" type="text" name="linkto" size="40">	  
 </div>
</div>
 
 
    
<%
 uploaded=request.QueryString("uploaded") 
 if session("Caricata")=true then
   uploaded=1
 end if
 
 if uploaded=1 then %>   
    <br><span class="alert-success">Modifiche salvate<span><br><br>
 <%else%>
    <br><span class="alert-danger">N.B.</font> Dopo l'Upload salva le modifiche</span><br>
 <%end if%>   
 
     <div class="control-group">
      <div class="controls">
      <input type="submit" value="Salva e Termina" name="B1" class="btn" onClick="return validate2();"> 
      </div>
    </div>   
   

    
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
uploaded=0%>
<script>window.close();</script>
<%
end if

%>
 
<script language="javascript" type="text/javascript" >

 function validate2() {
	var stringa=theform.nomeimg.value;
	 
 if (stringa=="")
	{
	   alert("Non hai inserito il commento all'immagine");
	   theform.nomeimg.setfocus();
	   return 0;
	}
 else
  
	{
	    document.theform.action = "save_img_All.asp";
		document.theform.submit();
		
		
	 
    }
	
}

 </script>
 
</body>