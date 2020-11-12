<%@ Language=VBScript %>
<!doctype html>
<meta charset="UTF-8">
<html>
<head>   
   <title>Condividi pagina</title>   
</head>
<%
		Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
%>
		<!-- #include file = "../var_globali.inc" --> 
     	<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
        <!-- #include file = "../cAdmin/include_mail.asp" -->
   
     
     
	<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
     <link rel="stylesheet" href="../../css/style-themes.css">
	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script> 
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	
      <!-- Per copiare nella clipboard -->
	 <script src="_assets/js/jquery-1.4.4.min.js" type="text/javascript"></script>
     <script src="_assets/js/jquery.zclip.js"></script>
     <script src="_assets/js/jquery-ui.js"></script>  
 <script>
 
 function validate() {
 if (frmMail.Destinatario.value=="")
	{
	   alert("Non hai inserito l'indirizzo email ");
	   return 0;
	}
	else
	{
	    document.frmMail.action = "share_mail.asp";
		document.frmMail.submit();	 
    }
}
 </script>
<body>
	<%

	scegli=request.QueryString("scegli")
    Url=Request.ServerVariables("HTTP_REFERER")&"&by_ospite=1"
	 
 QuerySQL="INSERT INTO Shared (Url,CodiceAllievo,Data) SELECT '" & Url & "','" & session("CodiceAllievo") & "','" &  now & "';"
	' response.write(QuerySQL)
   ConnessioneDB.Execute (QuerySQL) 
   QuerySQL="select max(ID) from Shared"
   set rsTabella=ConnessioneDB.Execute(QuerySQL) 
   id = rsTabella(0)
   Url=dominio&"/page.asp?id="&id
    	%>
       

    
<FORM Name = "frmMail"  onSubmit = 'return Validate()' METHOD = "POST" class='form-horizontal form-bordered'> 
<INPUT TYPE = "Hidden" NAME = "MessageType" VALUE = "NEW">
<INPUT TYPE = "Hidden" NAME = "CodBacheca" VALUE = "<%=bacheca%>">


 <%session("url")=Url%>
 <div class="control-group">
  <div class="controls">
	 <a href="#" id="copy-description">Ottieni URL di condivisione</a>  
         <p id="description" ><%=Url%> </p> <hr><b>Oppure invialo per email</b>
  </div>
</div>


<div class="control-group">
<label for="textfield" class="control-label"><B>Destinatario:</B></label>
  <div class="controls">
	  <INPUT TYPE = "TEXT"  NAME = "Destinatario" class="input-xlarge" placeholder="Indirizzo email"> 
  </div>
</div>

<div class="control-group">
<label for="textfield" class="control-label"><B>Oggetto:</B></label>
  <div class="controls">
	  <INPUT TYPE = "TEXT"  NAME = "Topic" class="input-xlarge" placeholder=""> 
  </div>
</div>
 
 
 <div class="control-group">
<label for="textfield" class="control-label"><B>Messaggio:</B></label>
  <div class="controls">
   
	  <textarea class="input-block-level" rows="5" NAME = "MESSAGE" cols="40" placeholder="Text input" >
    
	   <%=Url%>  
	  </textarea>
 </div>
</div>
 
 <div class="control-group">
  <div class="controls">
  <input type="submit" value="Invia" name="B1" class="btn" onClick="validate();"> 
  </div>
</div>   
    
</form>
			 
	</body>

<script language="javascript">
$(document).ready(function(){
$('a#copy-description').zclip({
path:'_assets/js/ZeroClipboard.swf',
copy:$('p#description').text()
});
// The link with ID "copy-description" will copy
// the text of the paragraph with ID "description"
$('a#copy-dynamic').zclip({
path:'_assets/js/ZeroClipboard.swf',
copy:function(){return $('input#dynamic').val();}
});
// The link with ID "copy-dynamic" will copy the current value
// of a dynamically changing input with the ID "dynamic"
});
 

</script>
 </html>

