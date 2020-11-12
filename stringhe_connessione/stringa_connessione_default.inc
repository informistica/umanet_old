<%'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 'ConnessioneDB0.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
  '            "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/DBMatProf.mdb"    
 
	Session.Timeout=60		  
			  ConnessioneDB0.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_               
    		   "DBQ=" & Server.MapPath("database/DBMatProf.mdb")

			 
  
 
%>
