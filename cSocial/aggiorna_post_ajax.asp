<%@ Language=VBScript %>


        <%


  


		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
        <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  		<!-- #include file = "../service/controllo_sessione.asp" -->

    <%
	  id = Request("id")
	  titolo=Request("titolo")
      titolo=Replace(titolo,"'",chr(96))
	  'testo=Request.QueryString("testo")
	  testo=Request("testo")
	  'testo=Request.form("editor1")
	 
      'testo=Replace(testo,"'",chr(96))
	  testo=Replace(testo,"editor1=","")
	  
	 	on error resume next
			 QuerySQL="UPDATE FORUM_MESSAGES SET topic = '" & titolo &"', comments = '" & testo &"' WHERE ID="&id&";"
		  ConnessioneDB.Execute(QuerySQL)
		'  RESPONSE.WRITE(querysql)
		  If Err.Number = 0 Then
	       
			'Response.Write "Modifica avvenuta! "
				stato=1
				messaggio="Modifica avvenuta"
			Else
				stato=0
				messaggio=Err.Description
			Err.Number = 0
			End If

		''response.write(QuerySQL)
'

		response.write(" { ")
		 response.write("""stato"": """&stato&""","  &_
		 """messaggio"": """&messaggio&"""")
		 response.write("}")
'


%>
