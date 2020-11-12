 <%@ Language=VBScript %>

	 <span>
<%
response.write(Request.ServerVariables("PATH_INFO")&"<br>")
response.write(Request.ServerVariables("QUERY_STRING")&"<br>")
response.write(Request.ServerVariables("SERVER_NAME")&"<br>")

  
 
%>

 
	 </span>

	 