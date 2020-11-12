<%@ Language=VBScript %>

  <% 
   Response.Buffer=True
   Dim ConnessioneDB, ConnessioneDB1, rsTabella, QuerySQL,StringaConnessione,URL,RecSet
   Dim CodiceTest, CodiceAllievo, CodiceCorso,DataTest ,Capitolo,Paragrafo,Nome,Cognome


   'Apertura della connessione al database
   ' Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
	codicemetafora=request("codicemetafora")
    segno=request("segno") ' 1 mi piace, 0 non mi piace
    codicetest=request("CodiceTest")
    ID_Premetafora=request("ID_Premetafora")
    
    commento=request("commento")
    'Cartella=session("Cartella")
    Cartella=request("Cartella")
    CodiceAllievo=session("CodiceAllievo")
    Set ConnessioneDB = Server.CreateObject("ADODB.Connection") 
%>
   	<!--#include file = "../stringhe_connessione/stringa_connessione.inc"-->
	<!--#include file = "../service/controllo_sessione.asp"-->

<%

QuerySQL="Select * from Setting where Id_Classe='" & Session("Id_Classe")&"';"
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
MaxStelline=rsTabella("MaxStelline")
 VotoPalese=rsTabella("VotoPalese")
	


set conn = Server.CreateObject("ADODB.Connection")
set cmd = Server.CreateObject("ADODB.Command")
set rs = Server.CreateObject("ADODB.Recordset")
'conn.mode = 3
conn.open sConnString
set cmd.activeconnection = conn
 'response.write(CodiceTest &"="&Cartella)
    Select Case CodiceTest%>
   <% Case Cartella&"_U_2_3" 'Topolino%>
    <% Case Cartella&"_U_2_5" 'Navigazione%>
<%  ' per non votare per se stessi, prelevo il COdiceAllievo della metafora per cui si vuole votare
	 QuerySQL="Select CodiceAllievo from Elenco_Metafore_Navigazione where CodiceMetafora='" &codicemetafora &"' and Id_Premetafora="&ID_Premetafora&";"
%>

	 <% Case Cartella&"_U_2_8" 'Navigazione 
	' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
  end Select



'response.write(QuerySQL)
Set rsTabella = ConnessioneDB.Execute(QuerySQL)
CodiceAllievoMetafora=rsTabella("CodiceAllievo") 


if (strcomp(ucase(CodiceAllievo),ucase(CodiceAllievoMetafora))<>0)  then 
'non puoi votare per te stesso

 

           

            if (strcomp(segno,"1"))=0  then
                    voto=1


                     Select Case CodiceTest%>
                    <% Case Cartella&"_U_2_3" 'Topolino%>
                        <% Case Cartella&"_U_2_5" 'Navigazione%>
                    <%  
                        ' mi piace quindi cancello eventuali voti negativi
                         QuerySQL1="Delete  from [VotiMetaforaNavigazione] WHERE CodiceAllievo='"& CodiceAllievo &"' and CodiceMetafora="&CodiceMetafora&" and Voto=-1;"
                         QuerySQL2="select count(*) from [VotiMetaforaNavigazione] WHERE CodiceAllievo='"& CodiceAllievo &"' and Id_Premetafora="&ID_Premetafora&";"
                  
                    %>

                        <% Case Cartella&"_U_2_8" 'Navigazione 
                        ' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
                    end Select
                   
                    set rs=conn.Execute(QuerySQL1)  
                    set rs2=conn.Execute(QuerySQL2)
                    numVoti=rs2(0)

                    if numVoti<MaxStelline then
                                voto=1

                                 Select Case CodiceTest%>
                                  <% Case Cartella&"_U_2_3" 'Topolino%>
                                  <% Case Cartella&"_U_2_5" 'Navigazione%>
                                   <%  
                                   QuerySQL="INSERT INTO VotiMetaforaNavigazione (CodiceAllievo,CodiceMetafora,Data,Voto,Cognome,Nome,Id_Premetafora,Commento) SELECT '" & CodiceAllievo & "','" & codicemetafora & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "',"&ID_Premetafora&",'"&commento&"';"
                                   QuerySQL2="select count(*) from [VotiMetaforaNavigazione] WHERE   CodiceMetafora="&CodiceMetafora&"  and Id_Premetafora="&ID_Premetafora&";"
                                    
                                    QuerySQL3="select count(*) from [VotiMetaforaNavigazione] WHERE   CodiceAllievo='"&CodiceAllievo&"'  and Id_Premetafora="&ID_Premetafora&";"
                  
                                    %>

                                  <% Case Cartella&"_U_2_8" 'Navigazione 
                        ' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
                                 end Select
                               


                            '  response.write(QuerySQL3)
                                conn.Execute(QuerySQL)
                                set rsNum=conn.Execute(QuerySQL2)
                                set rsRim=conn.Execute(QuerySQL3)
                                numrim=MaxStelline-rsRim(0)
                                 
                                stato=1
                                msg=rsNum(0)&"-"&numrim
                                response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & msg & """}")
                                'response.write("ok")
                        
                            'if VotoPalese=1 then
                                'Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                'Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

                                'QuerySQL="select * from AVVISI where IdPost=" & ID & " and CodiceAllievo='"&CodiceAllievoPost &"' and CodiceAllievo2='"& Session("CodiceAllievo")&"' and Social="&scegli&";"
                                'response.write(QuerySQL)
                                'set rs=ConnessioneDB.Execute(QuerySQL)
                                ' metto solo una notifica anche se ci sono pi� voti, evito di intasare le notifiche quando la classe vota il singolo
                                'response.write("<br>"&Topic)
                             '   Topic="Ho quotato una tua metafora"
                              '  if rs.eof and rs.bof then
                               ' QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore,Social,IdPost) SELECT '" & CodiceAllievoPost & "','" &  Topic &"','"& Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "','" & scegli & "'," & ID &";"
                               ' end if
                                ' ConnessioneDB.Execute(QuerySQL)
                                '  response.write(QuerySQL)
                                '    sSQL="select max(ID_Avviso) from AVVISI;"
                                 '   set rs=ConnessioneDB.Execute(sSQL)
                                 '   maxIDAvviso=rs(0)
                                 '   rs.close
                                 '   Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&maxIDAvviso="&maxIDAvviso&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                 '   sSQL="Update AVVISI set Azione ='"& Azione&"' where ID_Avviso="&maxIDAvviso&";"
                                    ' ConnessioneDB.Execute(sSQL)
                            ' end if ' votopalese
                                            
                        else
                    
                             stato=0
                                msg="Hai utilizzato tutte le " & MaxStelline & " stelline disponibili"
                                response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & msg & """}")

                        end if

            else  '   if (strcomp(segno,"1"))=0  then revoco eventuale mi piace
                voto=-1
                  Select Case CodiceTest%>
                    <% Case Cartella&"_U_2_3" 'Topolino%>
                        <% Case Cartella&"_U_2_5" 'Navigazione%>
                    <%  
                        ' mi piace quindi cancello eventuali voti negativi
                         QuerySQL1="Delete  from [VotiMetaforaNavigazione] WHERE CodiceAllievo='"& CodiceAllievo &"' and CodiceMetafora="&CodiceMetafora&" and Voto=1;"
                         QuerySQL2="select count(*) from [VotiMetaforaNavigazione] WHERE CodiceMetafora="&CodiceMetafora&" and Id_Premetafora="&ID_Premetafora&";"
                         QuerySQL3="select count(*) from [VotiMetaforaNavigazione] WHERE   CodiceAllievo="&CodiceAllievo&"  and Id_Premetafora="&ID_Premetafora&";"
                  
                    %>

                        <% Case Cartella&"_U_2_8" 'Navigazione 
                        ' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
                    end Select
                   
                    set rs=conn.Execute(QuerySQL1)  
                  
                   ' set rsNum=conn.Execute(QuerySQL2)
                   
                    numVoti=rs(0)
               
               if numVoti<MaxStelline then
                                voto=1

                                 Select Case CodiceTest%>
                                  <% Case Cartella&"_U_2_3" 'Topolino%>
                                  <% Case Cartella&"_U_2_5" 'Navigazione%>
                                   <%  
                                   QuerySQL="INSERT INTO VotiMetaforaNavigazione (CodiceAllievo,CodiceMetafora,Data,Voto,Cognome,Nome,Id_Premetafora,Commento) SELECT '" & CodiceAllievo & "','" & codicemetafora & "','" & now() & "','" & voto & "','" & Session("Cognome") & "','" &  Session("Nome") & "',"&ID_Premetafora&",'"&commento&"';"
                                  QuerySQL2="select count(*) from VotiMetaforaNavigazione where CodiceMetafora="&codicemetafora&""
                                    %>

                                  <% Case Cartella&"_U_2_8" 'Navigazione 
                        ' seleziono solo i compiti relativi al paragrafo che non sono stati ancora svolti
                                 end Select
                               


                            '  response.write(QuerySQL)
                                conn.Execute(QuerySQL)
                                set rsNum=conn.Execute(QuerySQL2)
                                 set rsRim=conn.Execute(QuerySQL3)
                   
                                'response.write("Non mi piace applicato!")
                                numrim=MaxStelline-rsRim(0)
                                 stato=1
                                  msg=rsNum(0)&"-"&numrim
                                response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & msg & """}")
                               ' response.write("ok")
                        
                            'if VotoPalese=1 then
                                'Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                'Commentatore=Session("Cognome") & " " & left(Session("Nome"),1) & "."

                                'QuerySQL="select * from AVVISI where IdPost=" & ID & " and CodiceAllievo='"&CodiceAllievoPost &"' and CodiceAllievo2='"& Session("CodiceAllievo")&"' and Social="&scegli&";"
                                'response.write(QuerySQL)
                                'set rs=ConnessioneDB.Execute(QuerySQL)
                                ' metto solo una notifica anche se ci sono pi� voti, evito di intasare le notifiche quando la classe vota il singolo
                                'response.write("<br>"&Topic)
                             '   Topic="Ho quotato una tua metafora"
                              '  if rs.eof and rs.bof then
                               ' QuerySQL="INSERT INTO Avvisi (CodiceAllievo,Testo,Azione,Data,CodiceAllievo2,Commentatore,Social,IdPost) SELECT '" & CodiceAllievoPost & "','" &  Topic &"','"& Azione & "','" & now() & "','" & Session("CodiceAllievo") & "','" & Commentatore & "','" & scegli & "'," & ID &";"
                               ' end if
                                ' ConnessioneDB.Execute(QuerySQL)
                                '  response.write(QuerySQL)
                                '    sSQL="select max(ID_Avviso) from AVVISI;"
                                 '   set rs=ConnessioneDB.Execute(sSQL)
                                 '   maxIDAvviso=rs(0)
                                 '   rs.close
                                 '   Azione="<a  target=blank href=ShowMessage.asp?byNotifiche=1&scegli="&scegli&"&ID="&ID&"&maxIDAvviso="&maxIDAvviso&"&Zip="&Zip&">Ho quotato un tuo post !</a>"
                                 '   sSQL="Update AVVISI set Azione ='"& Azione&"' where ID_Avviso="&maxIDAvviso&";"
                                    ' ConnessioneDB.Execute(sSQL)
                            ' end if ' votopalese
                                            
                        else
                    
                     '  response.write(QuerySQL)
                               
                                stato=0
                                msg="Hai utilizzato tutte le " & MaxStelline & " stelline disponibili"
                                response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & msg & """}")
                               ' response.write("ok")
                           

                        end if

            end if
else
   stato=0
    msg="Non puoi votare per te stesso!"
     response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & msg & """}")
  ' response.write("Non puoi votare per te stesso!")
end if



 

'response.write(url)

		On Error Resume Next
		If Err.Number = 0 Then
				'Response.Write "Voto avvenuto! "
				'Response.Redirect "ShowMessage.asp?ID="&ID
		Else
				'Response.Write Err.Description
				stato=0
                 response.write(" { "  &_
                                """stato"": """ & stato& """," &_
                                 """msg"": """ & Err.Description & """}")
                Err.Number = 0
		End If
 
 
 %>