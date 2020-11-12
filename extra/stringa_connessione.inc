<%'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
' ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
'              "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2013-2014/ECDL/database/" & Session("DBCopiatestonline")    
    ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
              "DBQ=" & Server.MapPath("../../database/" & Session("DBCopiatestonline"))

 
%>
