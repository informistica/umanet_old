<%
 	Session.Timeout=60
	 if (Session("DBCopiatestonline") ="")   then%>
	<!-- #include file = "stringa_connessione_refresh.asp" -->  
	
	      <% 
		'  response.write(session("DBCopiatestonline") & "<br>" & session("Id_Classe"))
		  if ((session("DBCopiatestonline")=""))  then
		     response.redirect "redirect.asp"
	      end if
	 
	 end if

 

'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
' ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
 '             "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBCopiatestonline")
    
	
	 
	 
    
	'ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
    '          "DBQ=" & Server.MapPath("../database/" & Session("DBCopiatestonline"))
			  
			  
				  
				  
		if session("DB")=1 then
 ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"
  
else
ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
" Initial Catalog=Copiaditestonline2; User Id=utente; Password=123Maurosho;"
  
end if

				  
				  
				  
'ConnessioneDB.Open	"Provider=sqloledb; Data Source=SERVERWIN\SQLEXPRESS; "&_
'" Initial Catalog=Copiaditestonline; User Id=utente; Password=123Maurosho;"

 homesito="/expo2015/UECDL"   
 
%>
