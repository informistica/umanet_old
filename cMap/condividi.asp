<%@ Language=VBScript %>
<%
  Response.Buffer = true
  'On Error Resume Next  
     
    id = Request.QueryString("id")
	url = Request("url") 
	
	if Session("DB") = "" then
	Session("DB") = Request.QueryString("DB")
	end if
	
	'response.write Session("DB")
	
	%>
		<%Set ConnessioneDB = Server.CreateObject("ADODB.Connection")  %> 
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
		<%
						   QuerySQL="SELECT count(*) FROM Shared WHERE Url = '"&url&"'"						   
						   'response.write(QuerySql&"<br>")
						   set rsTabella = ConnessioneDB.Execute(QuerySQL)
							
							if id = "" and url <> "" then
								if rsTabella(0) = 0 then
									if Session("CodiceAllievo") = "" then
										Session("CodiceAllievo") = "ospite"
									end if
									QuerySQL = "SELECT MAX(ID) as Massimo FROM Shared"
									set rsMax = ConnessioneDB.Execute(QuerySQL)
									
									QuerySQL = "INSERT INTO Shared (ID, Url, CodiceAllievo, Data) VALUES ("&(cInt(rsMax("Massimo"))+1)&", '"&url&"', '"&Session("CodiceAllievo")&"', '"&now()&"')"
									'Response.Write(QuerySQL)
									ConnessioneDB.Execute(QuerySQL)
									Session.Contents.Remove(Session("CodiceAllievo"))
									
									Response.write "https://www.umanetexpo.net/expo2015Server/UECDL/script/cMap/condividi.asp?id="&(cInt(rsMax("Massimo"))+1)&"&DB="&Session("DB")
								else
								
									QuerySQL="SELECT ID FROM Shared WHERE Url = '"&url&"'"	
									set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
									
									Response.write "https://www.umanetexpo.net/expo2015Server/UECDL/script/cMap/condividi.asp?id="&rsTabella1("ID")&"&DB="&Session("DB")
									
								end if
							else
									
									QuerySQL="SELECT Url FROM Shared WHERE ID = '"&id&"'"	
									set rsTabella1 = ConnessioneDB.Execute(QuerySQL)
									Session("UrlCondivisione") = "https://www.umanetexpo.net/expo2015Server/UECDL/script/cMap/condividi.asp?id="&id
									Response.Redirect rsTabella1("Url")
							end if
							
				'Response.AddHeader "REFRESH","2;URL=inserisci_collegamento.asp?Tipo=0&Stato="&Stato&"&Cartella="&Cartella&"&CodiceTest="&CodiceTest&"&Capitolo="&Capitolo&"&Paragrafo="&Paragrafo&"&Modulo="&Modulo
				'Response.Redirect Session("urlmappa")
				
				%>
							