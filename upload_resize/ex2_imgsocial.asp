<%
 Set ConnessioneDB1 = Server.CreateObject("ADODB.Connection")
%>
<!-- #include File="resizecheck.asp" -->
<!-- #include file = "stringa_connessione_forum.inc" -->
<!-- #include file = "../var_globali.inc" -->
<!-- #include file = "controllo_sessione.asp" -->
  
<%
 
'imgPath = Server.MapPath(".") & "\img"	'I suppose your images will be saved in an "img" folder which is child of the current folder
'imgPath = Server.MapPath(homesito) & "\img"

'studPath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo")
'Set fso = CreateObject("Scripting.FileSystemObject")
'if not(fso.FolderExists (studPath)) then
'	fso.CreateFolder (studPath) 
'	response.write("<br>Creata dir : " &studPath)
'end if

'imgPath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/img"  
imgPath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/img"  
imgPath=Replace(imgPath,"\","/")
'thumbPath = Server.MapPath(".") & "\thumb"	'I suppose your images will be saved in a "thumb" folder which is child of the current folder
'thumbPath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/"&Session("CodiceAllievo") &"/thumb"
thumbPath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/thumb"
thumbPath=Replace(thumbPath,"\","/")
includePath=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")&"/" & Session("cartella") &"/img_social/include"
includePath=Replace(includePath,"\","/")

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

Set fso = CreateObject("Scripting.FileSystemObject")
if not(fso.FolderExists (includePath)) then
	fso.CreateFolder (includePath) 
	response.write("<br>Creata dir : " &includePath)
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
<title>Aggiorna immaginario</title>
<script>
	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>


 
  


<link rel="stylesheet" type="text/css" href="../../stile.css">

</head>
<body><br><br><br>
<div class="contenuti_forum" style="width:60%">
<fieldset><legend>Aggiungi immagine</legend>
<form name="theform">
	<!-- I suppose you have a save.asp script that you will use to save all data in this form -->
	 <table id="zebra_forum1" >
	<tr><td>Immagine:&nbsp;</td><td> <input readonly="readonly" name="img" / size="30"><input type="hidden" name="thumb" />
	<input type="button" value="Upload" onClick="uploadImgWindow(this.form.name, 'img', 'thumb', '<%= Server.URLEncode(imgPath) %>', '<%= Server.URLEncode(thumbPath) %>', this.form.img.value, 200, 200, 120, 92);" /></td></tr><br/>
	<!-- I resize the image with the max with of 191 and the max height of 144.
	I create a thumbnail with the max with of 120 and the max height of 92.-->
	<tr><td>Nome:&nbsp;&nbsp;&nbsp;</td><td><input type="text" name="nomeimg" size="40"></td></tr> 
     <tr><td title="Testo da visualizzare posizionandosi sopra l'immagine">Titolo: </td> <td><input type="text" name="txtTitle" / size="40" > </td></tr>
    <tr><td title="Spiega cosa significa per te questa immagine">Descrizione : </td>
    <td>
    <textarea name="txtDescrizione" size="6" cols="33"></textarea>
    </td><tr>
    <tr><td title="Inserisci url a cui vuoi collegare questa immagine">Link to:&nbsp;</td><td><input type="text" name="linkto" size="40"></td></tr>
    <tr><td>Nella categoria: </td><td> 
	
    <select name="txtCategoria" style="width:auto">

	 <%'visualizzo la combo per la scelta della classe
	   querySQL="Select max(ID_Categoria) from CAT_SOCIAL; "
	  set rsTabella=ConnessioneDB1.execute(QuerySQL)	
	   
	  if rsTabella(0)&"" =""  then
	   numcat=1
	  else
	   numcat=rsTabella(0)+1 
	  end if
	   querySQL="Select * from CAT_SOCIAL where CodiceAllievo='"&Session("CodiceAllievo")&"'"
	  set rsTabella=ConnessioneDB1.execute(QuerySQL)	
	  ' response.write(querySQL)
	 
	  do while not rsTabella.eof %>
	    
		 <option selected  value="<%=rsTabella.fields("ID_Categoria")%>"><%=rsTabella.fields("Testo")%> </option>	
		 
		<% rsTabella.movenext()
		 
		   loop %>
		</select> 
    </td></tr>   
    </table>
 <%
 uploaded=request.QueryString("uploaded") 
 if uploaded=1 then %>   
    <br><font color="green"></font> Modifiche salvate<br><br>
 <%else%>
    <br><font color="#FF0000">N.B.</font> Dopo l'Upload salva le modifiche<br><br>
 <%end if%>
       <input type="button"  value="Salva" onClick="return validate2();" />
     <input type="button" value="Chiudi" onClick="window.close();" />
</form>
<%

if uploaded=1 then%>
<font color=green> Caricamento  del file <%=Session("FileName")%> completato</font>
<%
'   url=Server.MapPath(homesito)& "/Materie/"&Session("ID_Materia")& "/" & Session("cartella") &"/img_social/img/" & urlimg
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
 
</fieldset><br><br>
<fieldset><legend>Nuova categoria</legend>
<form name="nuovaCat" method="post" >
Nome&nbsp;&nbsp;: <input type="text" name="nomeCat" size="20"><br>
<input type="button" value="Aggiungi" onClick="return validate();">
</form>
<a href="../ChatRoom/genera_include_personale.asp"> Genera include...</a>

</fieldset>
<br>
<input type="button" value="Termina" onClick="javascript:window.close();">
</div>

<script language="javascript" type="text/javascript" >

 function validate() {
	var stringa=nuovaCat.nomeCat.value;
	 
 if (stringa=="")
	{
	   alert("Non hai inserito il nome della nuova categoria di immagini");
	   nuovaCat.nomeCat.setfocus();
	   return 0;
	}
 else
 
	{
	    document.nuovaCat.action = "../inserisci_categoria_immaginario.asp?numcat=<%=numcat%>";
		document.nuovaCat.submit();
		
	 
    }
	
}

 function validate2() {
	var stringa=theform.nomeimg.value;
	 
 if (stringa=="")
	{
	   alert("Non hai inserito il nome dell'immagine");
	   theform.nomeimg.setfocus();
	   return 0;
	}
 else
 if (theform.txtTitle.value=="")
	{
	   alert("Non hai inserito il titolo dell'immagine");
	   theform.txtTitle.setfocus();
	   return 0;
	}
 
  else
  if (theform.txtDescrizione.value=="")
	{
	   alert("Non hai inserito la descrizione dell'immagine");
	   theform.nomeimg.setfocus();
	   return 0;
	}
  
  else
	{
	    document.theform.action = "save.asp";
		document.theform.submit();
		
	 
    }
	
}


 </script>
</body>