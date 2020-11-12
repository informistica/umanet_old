<%@ Language=VBScript %>


        <%

		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
    <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  	<!-- #include file = "../service/controllo_sessione.asp" -->

    <%

    RCount=Request.QueryString("RCount")
    'categoria=Request.QueryString("categoria")
    id_categoria=Request.QueryString("id_categoria")
    id_classe=session("Id_Classe")
    topic=Request.QueryString("Paragrafo")
    bacheca="informistica"
    authorname="Admin A."
    abstract=".."
    parentmessage=Request.QueryString("ParentId")
    threadparent=Request.QueryString("ThreadId")
    comments=Request.QueryString("Domanda")
    privato=0
    visibile=1
    id_social=3 ' metto feedback nel interrogazioni'





    if day(date()) < 10 then
      giorno="0" & day(date())
    else
    giorno=day(date())
      end if
    if len(year(date()) ) = 2 then
    anno="20"& year(date())
    elseif len(year(date()) ) =  3 then
    anno="2"& year(date())
    else
    anno=year(date())
    end if
       if month(date()) < 10 then
      mese="0" & month(date())
    else
    mese=month(date())
    end if

    DataAvviso = giorno & "/" & mese& "/" & anno
    Testo=topic
    Azione="<a  target=blank href=ShowMessage.asp?scegli="&id_social&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&ID="&parentmessage&"&RCount="&RCount&"&categoria=Feedback&id_categoria="&id_categoria&">Hai ricevuto punti per interrogazione !</a>"
    Commentatore=authorname

on error resume next

strText=Request.ServerVariables("QUERY_STRING")
s=split(strText,"&")
  For i=8 to ubound(s)
	 s1=split(s(i),"=")
	 codiceallievo=s1(0)
    

	 punti=s1(1)
      sSQL = "INSERT INTO FORUM_MESSAGES (PARENTMESSAGE,THREADPARENT,AUTHORNAME,CODICEALLIEVO,ID_CLASSE,TOPIC,BACHECA,COMMENTS,Id_Social,DatePosted,Id_Categoria,Abstract,Punti) VALUES (" & parentmessage & "," &  threadparent & ",'" & authorname &"','"& codiceallievo&"','"& id_classe&"','" &topic&"','" & bacheca &"','"& comments & "',"&id_social&",'"&now&"',"&id_categoria&",'"& abstract&"',"&punti&");"
     'response.write(sSQL&"<br>")
   ConnessioneDB.Execute sSQL
     QuerySQL="UPDATE FORUM_MESSAGES set ReplyCount = ReplyCount + 1 where ID="& parentmessage & ";"
     ConnessioneDB.Execute(QuerySQL)

      'ed invio notifica sul quaderno dello stud


       QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo, Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Testo& "','" & Azione & "','" & DataAvviso & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
    ''  response.write(QuerySQL&"<br>")
    ConnessioneDB.Execute(QuerySQL)
    QuerySQL="UPDATE Allievi set Interrogazioni = Interrogazioni + 1 where CodiceAllievo='"& CodiceAllievo & "';"
  ''  response.write(QuerySQL&"<br>")
  ConnessioneDB.Execute(QuerySQL)

   	next

      If Err.Number = 0 Then
				stato=1
				messaggio="Modifica avvenuta"
			Else
				stato=0
				messaggio=Err.Description&"<br>"&Err.Source&"<br>"&Err.Number
			Err.Number = 0
			End If
      response.write(messaggio)

    '' response.redirect Request.ServerVariables("HTTP_REFERER")
 		''response.write(QuerySQL)
'


'


%>
