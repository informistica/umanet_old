<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Stampa</title>
<link rel="stylesheet" type="text/css" href="stile_stampa.css">
</head>

<body>
<% Dim objFSO, objTextFile
   Dim sRead, sReadLine, sReadAll
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   tipo=request.QueryString("tipo")
   Modulo=request.QueryString("Modulo")
   Paragrafo=request.QueryString("Paragrafo")
   url=request.QueryString("url")
   n1=Request.Form("txtNUMREC")
    n=request.querystring("N")
   response.Write("N="&n)
    response.Write("<br>N1="&n1)
  select case tipo
    case "Frasi"%>
	<div align="center">
		<font size="4" color="#FF0000"><b>Frasi</b></font>
		<br>
		<p></p><font color=#00E800 ="Verdana" size="4"><b>Modulo : <%Response.write (Modulo) %></b></font>   
		<p></p><font color=#0066FF face ="Verdana" size="3"><b>Paragrafo : <%Response.write (Paragrafo) %></b></font>
        <form>
		 <% for i=1 to n %>
		       <input type="text" value="<%=Request.Form("txtCodiceDomanda"&i)%>" size="6">
              <b>Codice Frase </b> 
			  <input type="text" value="<%=Request.Form("txtData"&i)%>" size="10">
			  <b>Data </b> 
			  <input type="text" value="<%=Request.Form("txtOra"&i)%>" size="5" />
			  <b>Ora </b> 
              <p><input type="text" value="<%=Request.Form("txtChi"&i)%>" size="100">
              <b>Chi</b> <br></p> <!-- crea la variabile di tipo inputbox avente un certo nome -->
  
	           <% 
	             url=Request.Form("url"&i)
  


	          ' url1="C:\Inetpub\wwwroot\Anno_2010-2011_ITC\logFile.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'				objCreatedFile.WriteLine(url)
'				objCreatedFile.WriteLine(Modulo)
'				objCreatedFile.WriteLine(Paragrafo)
'			 
'				objCreatedFile.Close
'


    Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Set objTextFile = objFSO.OpenTextFile(url, ForReading)
	
	' Use different methods to read contents of file.
	'sReadAll = objTextFile.ReadAll
	sReadAll=url
	'response.write(sReadAll)
	'objTextFile.Close	%>
	<b>Frase</b>
	 <p>
	
	<textarea rows="<%=1+round((len(sReadAll))/50)%>"  cols="100">
	         <%Response.write(sReadAll)%> </textarea></p>  
			
			<%next%>
		</form>	
	
	</div>
	<%
	   
	case else
	end select
%>
</body>
</html>
