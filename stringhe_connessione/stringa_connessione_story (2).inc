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
 
				  
'				  	  ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
' " Initial Catalog=Copiaditestonline; User Id=sa; Password=;"
if session("DB")=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
" Initial Catalog=Copiaditestonline; User Id=informistica; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
	" Initial Catalog=Copiaditestonline; User Id=informistica; Password=123Maurosho;"
else
ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
" Initial Catalog=Copiaditestonline2; User Id=informistica; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
	" Initial Catalog=Copiaditestonline2; User Id=informistica; Password=123Maurosho;"
end if
 'ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'" Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"

 
 ' serve per 1_classifica_new
'sConnString = "Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'	 " Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"
	
 
  homesito="/expo2015Server/UECDL"   
%>
