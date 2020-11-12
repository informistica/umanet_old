<%@ Language=VBScript %>

  <% 
   Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome


   'Apertura della connessione al database
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	ID=request.querystring("ID")
    scegli=request.querystring("scegli")
	 
	CodiceAllievo=session("CodiceAllievo")
	 
	iThreadParent=request.querystring("TParent")
	'MaxStelline=3
	'response.write("aa="& request.querystring("MaxStelline"))
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
%>
   	<!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->
	<!--#include file = "../service/controllo_sessione.asp"-->

<%

QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
MaxStelline=rsTabella("MaxStelline")
	 

segno=request.QueryString("segno") ' 1 mi piace, 0 non mi piace

set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
 
 QuerySQL="Select CodiceAllievo from FORUM_MESSAGES where ID='" &ID &"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
CodiceAllievoPost=rsTabella("CodiceAllievo") 

if (strcomp(ucase(CodiceAllievo),ucase(CodiceAllievoPost))<>0)  then 
'non puoi votare per te stesso

 

            QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
            Set rsTabella = ConnessioneDB.Execute(QuerySQL)
            VotoPalese=rsTabella("VotoPalese")
            if (strcomp(segno,"1"))=0  then
                    ' mi piace quindi cancello eventuali voti negativi
                    QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=-1;"
                   
                    set rs=conn.Execute(QuerySQL)
                '  response.write("<br>"&QuerySQL)
                    'QuerySQL="select count(*) from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&";"
                    'modifico perchè controllo nella discussione il numero di stelline e non nel singolo post
                    QuerySQL="select count(*) from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadParent="&iThreadParent&";"
                    '
                        
                    set rs=conn.Execute(QuerySQL)
                    numVoti=rs(0)


                    if numVoti<MaxStelline then
                                voto=1
                                QuerySQL="INSERT INTO Voti (CodiceAllievo,ThreadParent,ThreadQuote,Data,Voto,Cognome,Nome) SELECT '" & CodiceAllievo & "','" & iThreadParent & "','" & ID & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "';"

                            '  response.write(QuerySQL)
                                conn.Execute(QuerySQL)
                                response.write("Feedback positivo inviato!")
                        
                            if VotoPalese=1 then
                                Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

                                QuerySQL="select * from AVVISI where IdPost=" & ID & " and CodiceAllievo='"&CodiceAllievoPost &"' and CodiceAllievo2='"& Session("CodiceAllievo")&"' and Social="&scegli&";"
                                'response.write(QuerySQL)
                                set rs=ConnessioneDB.Execute(QuerySQL)
                                ' metto solo una notifica anche se ci sono pi� voti, evito di intasare le notifiche quando la classe vota il singolo
                                'response.write("<br>"&Topic)
                                Topic="Ho quotato un tuo post"
                                if rs.eof and rs.bof then
                                QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore,Social,IdPost) SELECT '" & CodiceAllievoPost & "','" &  Topic &"','"& Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "','" & scegli & "'," & ID &";"
                                end if
                                ' ConnessioneDB.Execute(QuerySQL)
                                '  response.write(QuerySQL)
                                    sSQL="select max(ID_Avviso) from AVVISI;"
                                    set rs=ConnessioneDB.Execute(sSQL)
                                    maxIDAvviso=rs(0)
                                    rs.close
                                    Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&maxIDAvviso="&maxIDAvviso&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                    sSQL="Update AVVISI set Azione ='"& Azione&"' where ID_Avviso="&maxIDAvviso&";"
                                    ' ConnessioneDB.Execute(sSQL)
                            end if ' votopalese
                                            
                        else
                    
                            response.write("Hai utilizzato tutti i " & MaxStelline & " feedback disponibili")

                        end if

            else  ' revoco eventuale mi piace
                QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=1;"
               
               
                conn.Execute(QuerySQL)
                QuerySQL="select count(*) from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadParent="&iThreadParent&";"

                    set rs=conn.Execute(QuerySQL)
                    numVoti=rs(0)

                    'if numVoti<3 then

               ' QuerySQL="Delete  from Voti WHERE CodiceAllievo='"& CodiceAllievo &"' and ThreadQuote="&ID&" and Voto=1;"
               ' conn.Execute(QuerySQL)


                ' se non ho utilizzato i MAx voti disponibili aggingo voto=-1
                if numVoti<MaxStelline then
                    voto=-1

                        QuerySQL="INSERT INTO Voti (CodiceAllievo,ThreadParent,ThreadQuote,Data,Voto,Cognome,Nome) SELECT '" & CodiceAllievo & "','" & iThreadParent & "','" & ID & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "';"
                            conn.Execute(QuerySQL)

                        response.write("Feedback negativo inviato!")
                        else
                    response.write("Hai utilizzato tutti i " & MaxStelline & " feedback disponibili")
                        end if


            end if
else
  response.write("Non puoi votare per te stesso!")
end if



 

'response.write(url)

		On Error Resume Next
		If Err.Number = 0 Then
				'Response.Write "Voto avvenuto! "
				'Response.Redirect "ShowMessage.asp?ID="&ID
		Else
				Response.Write Err.Description
				Err.Number = 0
		End If
 
 
 %>