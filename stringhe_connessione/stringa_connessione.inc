<%
 	Session.Timeout=60
	'on error resume next
	 by_ospite=request.querystring("by_ospite")
	 
	if by_ospite <> "" and strcomp(Session("DB"),"1")=0  then%>
	  <!-- #include file = "login_ospite_expo.asp" -->  
	   
	 <%end if
	 
	 if by_ospite <> "" and strcomp(Session("DB"),"2")=0  then%>
	  <!-- #include file = "login_ospite_doc.asp" -->  
	 end if
	 
	if (session("DBCopiatestonline")="") and by_email="" then%>
	<!-- #include file = "stringa_connessione_refresh.asp" -->  
	
	      <% 
		  if (session("CodiceAllievo")="")   then
		     response.redirect "../service/redirect.asp"
	      end if
	 
	 end if

'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 
				  
'				  	  ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
' " Initial Catalog=Copiaditestonline; User Id=sa; Password=;"

if session("DB")=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
else
ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;" 
end if 
 
  homesito="/expo2015Server/UECDL"   
%>

