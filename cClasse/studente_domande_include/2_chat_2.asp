<%

QuerySQL0 = "SELECT count(*) from AVVISI where CodiceAllievo = '"&cod&"' and Visto = 0 and CAST(Testo as ntext) like Azione;"
Set rsTabellaAvvisiChat = ConnessioneDB.Execute(QuerySQL0)
 numMessaggiChat=rsTabellaAvvisiChat(0)
 
  QuerySQL = "SELECT count(DISTINCT CodiceAllievo2) from AVVISI where (CodiceAllievo = '"&cod&"' or CodiceAllievo2= '"&cod&"') and CAST(Testo as ntext) like Azione;"
 Set rsTabellaContattiChat = ConnessioneDB.Execute(QuerySQL)
 x = rsTabellaContattiChat(0)
 'numMessaggiChat=x
 'response.write(x)
 
 stringaelenco = ""
 
 QuerySQL = "SELECT * from AVVISI where (CodiceAllievo = '"&cod&"' or CodiceAllievo2 = '"&cod&"') and CAST(Testo AS ntext) like Azione order by Data desc;"
 Set rsTabellaContattiChat = ConnessioneDB.Execute(QuerySQL)
 
'response.write(QuerySQL)

%>