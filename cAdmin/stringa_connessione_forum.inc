<%

Session.Timeout=60
 if (Session("DBForum")="") or (session("Id_Classe")="") or (session("CodiceAllievo")="")  then%>
	<!-- #include file = "../stringhe_connessione/stringa_connessione_refresh.asp" -->  
	
	      <% 
		  if (session("DBCopiatestonline")="") then
		     response.redirect "../redirect.asp"
	      end if
	 
	 end if
'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
 ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/expo2015/UECDL/database/" & Session("DBForum")    


'	 ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
 '              "DBQ=" & Server.MapPath("../../database/" & Session("DBForum"))

			  

%>
