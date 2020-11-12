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
 ConnessioneDB2.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
else
ConnessioneDB2.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
	" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
end if


 'ConnessioneDB2.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'" Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"
  
 
%>
