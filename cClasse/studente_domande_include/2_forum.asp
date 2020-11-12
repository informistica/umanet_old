<%' carico i messaggi del forum
QuerySQL="SELECT * " &_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and comments<>'InizializzaDB' and Bacheca <>'"&cod&"'" &_ 
" and  DatePosted>=#" &DataClaq  &"#" &_
	 " AND DatePosted<=#" & 1 + cdate(DataClaq2) &"#" &_ 
" ORDER BY ID;"

' " and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND DatePosted<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 
	 


'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni 
' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum0.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL)
'				objCreatedFile.Close
 Set rsTabellaForum = ConnessioneDB1.Execute(QuerySQL)
 appQuery=QuerySQL
 'response.write(QuerySQL)%>
 
 
<!-- Div tendina per i le visualizzazioni di post -->
<a name="ancora_forum" href="#" onClick="Effect.toggle('forum','slide'); return false;"><span style="font-style:normal;" class="sottotitoloquaderno">&nbsp;&nbsp;FORUM</span></a> 
<div id="forum" style="display:none;">
<div class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p>


<table id="zebra_stud" align=center border=1 width="95%"  >
<%If rsTabellaForum.BOF=True And rsTabellaForum.EOF=True Then 

              %>
			  <tr><th align="center">  Non ci sono attività nel forum</th></tr>
			  
<% Else%>
	
			 <% 'conto i post totali 
			  QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and ParentMessage=0 and comments<>'InizializzaDB' " &_
" and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
" AND DatePosted<=#" & 1+ CDate(mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4))  &"#" &_ 
";"

 'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close
				
   Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1) 
	num_post_totali=rsTabella1(0)
	'num_post_totali_punti=0
	num_post_totali_punti=rsTabella1(1)
	if isnull(num_post_totali_punti)  then
	   num_post_totali_punti=0
	end if
	
	 
				
	 QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and ParentMessage <>0 and comments<>'InizializzaDB'" &_ 
" and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
	 " AND DatePosted<=#" & 1+ cdate( mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4))  &"#" &_ 

";"



   Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1) 
	num_messaggi=rsTabella1(0)
	'num_messaggi_punti=0
	num_messaggi_punti=rsTabella1(1)
	if isnull(num_messaggi_punti) then
	   num_messaggi_punti=0
	end if
	
 
	
	 %>
	          	<tr><th colspan=5><center><b>Post(<%=num_post_totali%>) + Commenti(<%=num_messaggi%>) = Punti(<%=num_post_totali_punti+num_messaggi_punti%>)  </b></center> </th></tr>
				<tr><th><b><center>Post</center></b></th><th><b>Messaggio</b></th><th><b>Data/Ora</b></th><th><b>Punti</b></th><th><b>Elimina</b></th></tr>
		 
		<% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   do while not rsTabellaForum.EOF 
		     QuerySQL1="SELECT * "&_
" FROM FORUM_MESSAGES " &_
" WHERE ID=" & rsTabellaForum("ThreadParent") &" and comments<>'InizializzaDB'" &_
 
" ORDER BY ID;"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

Set rsTabella1 = ConnessioneDB1.Execute(QuerySQL1) 
	'num_visualizzazioni=rsTabella1(0) 
			 %>
			<tr><td><a title="Visualizza Post di apertura discussione" href="../../forum/ShowMessage.asp?ID=<%=rsTabellaForum("ThreadParent")%>&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid2%>"><%=rsTabella1("Topic")%></a></td>
            
         
		 
         
            
            
                <td><a title="Visualizza il messaggio nella discussione"    href="../../forum/ShowMessage.asp?ID=<%=rsTabellaForum("ID")%>&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid2%>"><%=rsTabellaForum("Topic")%></a></td> 
                <td><%=rsTabellaForum("DatePosted")%></td> 
                <td><%=rsTabellaForum("Punti")%></td> 
               
			   <td><a onClick="return window.confirm('Vuoi veramente cancellare il messaggio ?');" target="_new" href="../../cancella_messaggio.asp?ID=<%=rsTabellaForum("ID")%>" title="Cancella"><i class=" icon-trash" ></i></a></td>
               
		<%rsTabellaForum.movenext
		loop%>
	<%end if%> 
</table>

</p> 
</div></div>

