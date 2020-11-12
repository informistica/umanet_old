<%@ Language=VBScript %>


        <%

		Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
		%>
    <!-- #include file = "../var_globali.inc" -->
 		<!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
  	<!-- #include file = "../service/controllo_sessione.asp" -->

    <%

    RCount=Request.QueryString("RCount")
    categoria=Request.QueryString("categoria")
    id_categoria=Request.QueryString("id_categoria")
    id_classe=session("Id_Classe")

    topic=Request.Form("Topic")
    if instr(topic,"Seleziona un feedback")>0 Then
      topic=Request.Form("Topic1")
    end if
    response.write("instr:"&instr(topic,"Seleziona feedback"))
    response.write(topic)

    querysql="select segno from Feedback where descrizione like '%"&topic&"%'"
    response.write(querysql)
    set rsTab=ConnessioneDB.execute (querysql)
    segno=rsTab("Segno")
    'response.write("<br>Segno="&segno)

    querysql="select nome,id from Feedback_polarita where id="&Request.Form("Polarita")
    set rsTab=ConnessioneDB.execute (querysql)
    nome_polarita=rsTab("nome")
    id_polarita=rsTab("id")
 
  


    nome=Request.Form("Name")
    punti=Request.Form("Punteggio")
    bacheca="informistica"
    authorname=Request.Form("Name")
    parentmessage=Request.Form("ParentId")
    threadparent=Request.Form("ThreadId")

    abstract="feedback"
    privato=0
    visibile=1
    id_social=2 ' metto feedback nel diario'
    comments=cstr(Request("Message"))
    comments=Replace(comments, Chr(39), Chr(96))

    if (Request.Form("newfeedpos")) then
      segno="+"
    end if

    if (Request.Form("newfeedneg")) then
    'response.write("ciao1")
      segno="-"
    end if


    if ((Request.Form("newfeedneg")) or (Request.Form("newfeedpos"))) then
    ' per avere Altro sempre come ultima voce
     'ssql="select id from Feedback where Descrizione='Altro' and Segno='"&segno&"' and id_poli="&Request.Form("Polarita")
     'rsUp=ConnessioneDB.Execute (ssql)
     'response.write(sSQL&"<br>"&rsUp("id"))
    ' ssql="update Feedback set Descrizione='"& topic&"' where ID="&rsUp("id")
     ' response.write(sSQL&"<br>")
     'ConnessioneDB.Execute ssql

     ssql="select max(Posizione) from Feedback where id_poli="&Request.Form("Polarita")
     rsUp=ConnessioneDB.Execute (ssql)
     maxPos=rsUp(0)
     sSQL = "INSERT INTO Feedback (Segno,Descrizione,id_poli,posizione) VALUES ('" & segno & "','"&topic&"',"&id_polarita&","&maxPos+1&");"
     'response.write(sSQL)
     ConnessioneDB.Execute sSQL


    end if

    topic=nome_polarita & " : ("&segno&")"&topic
    
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
    Azione="<a  target=blank href=ShowMessage.asp?scegli="&id_social&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&ID="&parentmessage&"&RCount="&RCount&"&categoria=Feedback&id_categoria="&id_categoria&">Ho creato un feedback per te !</a>"
    Commentatore=authorname
  response.write("<br>RCount"&RCount)

on error resume next
response.write("RCount="&RCount)
   for i=1 to RCount

		''	 QuerySQL="UPDATE FORUM_MESSAGES SET topic = '" & titolo &"', comments = '" & testo &"' WHERE ID="&id&";"
		 '' ConnessioneDB.Execute(QuerySQL)
    '' response.write("<br>"&i&"="&Request.Form("stud_"&i))

  if (Request.Form("stud_"&i)<>"") then
   'response.write("<br>"&i&"="&Request.Form("stud_"&i))
       codiceallievo=Request.Form("stud_"&i)
    ' response.write("<br>ciao"&i)
    ' ''  response.write("<br>"&Request.Form("stud_"&i))
       sSQL = "INSERT INTO FORUM_MESSAGES (PARENTMESSAGE,THREADPARENT,AUTHORNAME,CODICEALLIEVO,ID_CLASSE,TOPIC,BACHECA,COMMENTS,Id_Social,DatePosted,Id_Categoria,Abstract,Punti) VALUES (" & parentmessage & "," &  threadparent & ",'" & authorname &"','"& codiceallievo&"','"& id_classe&"','" &topic&"','" & bacheca &"','"& comments & "',"&id_social&",'"&now&"',"&id_categoria&",'"& abstract&"',"&punti&");"
      response.write(sSQL&"<br>")
      ConnessioneDB.Execute sSQL
      QuerySQL="UPDATE FORUM_MESSAGES SET ReplyCount = ReplyCount+1  " &_
	" WHERE ID="&threadparent&";"
'	response.write(QuerySQL &"<br>")
ConnessioneDB.Execute(QuerySQL)
    '
    '   'ed invio notifica sul quaderno dello stud
    '
    '
        QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo, Azione,Data,CodiceAllievo2,Commentatore) SELECT '" & CodiceAllievo & "','" & Testo& "','" & Azione & "','" & DataAvviso & "','" & Session("CodiceAllievo") & "','" & Commentatore & "';"
        response.write(QuerySQL&"<br>")
        ConnessioneDB.Execute(QuerySQL)


      end if
   	next
'response.write("fine ciclo")
'response.write("<br>Err.Number="&Err.Number)
   messaggio="risposta"
      If Err.Number = 0 Then
	       'RESPONSE.WRITE(querysql)
			'Response.Write "Modifica avvenuta! "
				stato=1
				messaggio="Modifica avvenuta"
			Else
				stato=0
				messaggio=Err.Description&"<br>"&Err.Source&"<br>"&Err.Number
			Err.Number = 0
			End If
      response.write(messaggio)
   response.redirect "ShowMessage.asp?scegli="&id_social&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&Session("bacheca")&"&ID="&parentmessage&"&RCount="&RCount&"&categoria=Feedback&id_categoria="&id_categoria
'response.write("ShowMessage.asp?scegli="&id_social&"&id_classe="&Session("Id_Classe")&"&cartella="&Session("cartella")&"&bacheca="&bacheca&"&ID="&parentmessage&"&RCount="&RCount&"&categoria=Feedback&id_categoria="&id_categoria)
		'response.write(QuerySQL)
'


'


%>
