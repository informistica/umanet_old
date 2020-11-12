<%
 	Session.Timeout=60
	if (session("DBCopiatestonline")="")  then%>
	<!-- #include file = "stringa_connessione_refresh.asp" -->  
	
	      <% 
		  if (session("DBCopiatestonline")="")   then
		     response.redirect "../service/redirect.asp"
	      end if
	 
	 end if

'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 
				  
if session("DB")=1 then
 ConnessioneDBStory.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
  
else
ConnessioneDBStory.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
  
end if
 
  'homesito="/expo2015Server/UECDL"   
%>
