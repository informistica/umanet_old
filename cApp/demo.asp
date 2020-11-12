<!-- richiama_test.asp -->
<%@ Language=VBScript %>

<%
response.addHeader "Access-Control-Allow-Origin", "*"
response.addHeader "Access-Control-Allow-Credentials", "true"

paragrafo = Request.QueryString("paragrafo")

%>

 <p>
<b> AJAX</b>, acronimo di Asynchronous Javascript And XML, è una tecnica di programmazione che vede coinvolti Javascript, l'oggetto XMLHTTP ed un linguaggio di scripting lato server (come, ad esempio, ASP o PHP).
Il suo scopo è quello di effettuare chiamate ad uno script lato server via XMLHTTP sfruttando la velocità lato client di Javascript.
Grazie ad Ajax è possibile inserire il risultato delle elaborazioni lato server all'interno di comuni pagine statiche senza bisogno di alcun refresh di pagina: in sostanza le operazioni vengono eseguite lato-server e poi richiamate lato-client!
 </p>