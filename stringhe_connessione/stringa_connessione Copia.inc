 


<%
 	Session.Timeout=60
	'on error resume next
	 
	'Response.AddHeader "Refresh", "3600" ' dopo un ora di inattività
	if (session("DBCopiatestonline")="") and by_email="" then%>
	<!-- #include file = "stringa_connessione_refresh.asp" -->  
	
	      <% 
		  if (session("CodiceAllievo")="")   then
		     response.redirect "../service/redirect.asp"
	      end if
	 
	 end if

'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 
				  
'	ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
' " Initial Catalog=Copiaditestonline; User Id=sa; Password=;"
if session("DB")=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS;"&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
	" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
else
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS;"&_
" Initial Catalog=Copiaditestonline2; User Id=informistica; Password=123Maurosho;"
 sConnString = "Provider=sqloledb; Data Source=SERVERWIN\SQLINFO; "&_
	" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
end if
 'ConnessioneDB.Open	"Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'" Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"

 
 ' serve per 1_classifica_new
'sConnString = "Provider=sqloledb; Data Source=MAUROSHODE6E; "&_
'	 " Initial Catalog=Copiaditestonline; User Id=maurosho; Password=Didamatica2001;"
	
 
  homesito="/expo2015Server/UECDL"   
%>

