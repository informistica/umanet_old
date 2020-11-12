<%'Apre la connessione utilizzando il metodo Open (Tipo di database, percorso)
' ConnessioneDB.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
'              "DBQ=D:/inetpub/vhosts/umanet.net/httpdocs/anno_2010-2011_ITC/ECDL/database/Copiaditestonline.mdb"    
    ConnessioneDB1.Open "DRIVER={Microsoft Access Driver (*.mdb)}; " &_ 
             "DBQ=" & Server.MapPath("../../database/" & Session("DBForum"))

 homesito="/anno_2010-2011_ITC/"   
 
%>
