<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "https://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="https://www.w3.org/1999/xhtml">
<head>
<meta https-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Resetta</title>
</head>

<body>
<!--#include file="functions/functions_chat.asp"-->
<!--#include file="functions/functions_users.asp"-->
<%
Call Reset()

response.write("Chat resettata")
if Request.ServerVariables("HTTP_REFERER") <>"" then 
							response.Redirect request.serverVariables("HTTP_REFERER") 
		 end if 

%>
</body>
</html>
