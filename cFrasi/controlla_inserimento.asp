<%if (session("CodiceAllievo")="") or (session("Id_Classe")="") then %>
	 <BODY onLoad="showText2();"> </BODY>
 <% else %>



			<% ' controllo se è già stata inserita
            response.write(ID_Prefrase)
              ' querySQL="Select * from Frasi where Id_Stud='" & Session("CodiceAllievo") & "' and (Id_Prefrase="&clng(ID_Prefrase)&" or Chi='" &Quesito& "');"
			    querySQL="Select * from Frasi where Id_Stud='" & Session("CodiceAllievo") & "' and (Id_Prefrase="&clng(ID_Prefrase)&");"
             '  response.write(querySQL)
              ' Set objFSO = CreateObject("Scripting.FileSystemObject")
'        				url1="C:\Inetpub\umanetroot\anno_2012-2013\logfrasi.txt"
'        				Set objCreatedFile = objFSO.CreateTextFile(url1, True)
'        				objCreatedFile.WriteLine(querySQL)
'        				objCreatedFile.Close



            ''   Set rsTabella = ConnessioneDB.Execute(QuerySQL)

                'If not(rsTabella.BOF=True And rsTabella.EOF=True) Then
                    ' esiste già non la faccio inserire
					%><!--
                      <BODY onLoad="showText3();">
                                Stai per essere reindirizzato alla pagina precedente </BODY>
                -->
                   <%'else%>
                      <body bgcolor="#FFFFFF">
                <% 'end if %>

 <% end if %>
