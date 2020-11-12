<%@ Language=VBScript %>
<% Response.Buffer=True
Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
 %>
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
<%
idclasse=Request.QueryString("idclasse")
visibile=Request.QueryString("visibile")
nome=Request.QueryString("nome")
 
                on error resume next
                if nome<>"" Then
                         QuerySQL ="update  Classi set Classe='"&nome&"' where ID_Classe='"&idclasse&"';"
                         ConnessioneDB.Execute(QuerySQL) 
                        If Err.Number = 0 Then
                            Response.Write "Nome della classe aggiornato"
                        Else
                            Response.Write Err.Description
                            Err.Number = 0
                        End If
                else
                         QuerySQL ="select Visibile from Classi where ID_Classe='"&idclasse&"';"
                         set rsVisibile=ConnessioneDB.Execute(QuerySQL) 
                        if (strcomp(visibile,"1")=0) and (rsVisibile(0)=1) Then
                            visibilita=0
                            risposta="icon-eye-close"
                        Else
                            visibilita=1
                            risposta="icon-eye-open"
                        end if
                        QuerySQL ="update  Classi set Visibile="&visibilita&" where ID_Classe='"&idclasse&"';"
                        ConnessioneDB.Execute(QuerySQL) 
                        If Err.Number = 0 Then
                            Response.Write risposta
                        Else
                            Response.Write Err.Description
                            Err.Number = 0
                        End If
                       
         
                end if
                        
%>