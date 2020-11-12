<%@ Language=VBScript %>



		<!-- #include file = "../stringhe_connessione/stringa_connessione_social.inc" -->


<% 
on error resume next
ID=Request.QueryString("ID")
TParent=cint(request.QueryString("TParent")) ' IDdel post per aggiornare ReplyCount
QuerySQL ="DELETE  FROM FORUM_MESSAGES WHERE ID =" &ID&";"
conn.Execute(QuerySQL)
'response.write(QuerySQL &"<br>")
QuerySQL="UPDATE FORUM_MESSAGES SET ReplyCount = ReplyCount-1  " &_
	" WHERE ID="&TParent&";"
'	response.write(QuerySQL &"<br>")
conn.Execute(QuerySQL)
response.write("ok")
       %>


