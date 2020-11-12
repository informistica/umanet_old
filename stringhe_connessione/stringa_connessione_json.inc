 


<%
  
				  
'				  	  ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
' " Initial Catalog=Copiaditestonline; User Id=sa; Password=;"
if request.querystring("DB")=1 then
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
 'ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'" Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"

 
 ' serve per 1_classifica_new
'sConnString = "Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'	 " Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"
	
 
  homesito="/expo2015Server/UECDL"   
%>

