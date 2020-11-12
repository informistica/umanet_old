<%
	  'conto i post totali 
			  QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage=0 and comments<>'InizializzaDB' and Id_Social=2 " &_
 " and (DatePosted>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"
' Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\expo2015Server\logForum.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close
				
   Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	num_post_totali=rsTabella1(0)
	'num_post_totali_punti=0
	num_post_totali_punti=rsTabella1(1)
	if isnull(num_post_totali_punti)  then
	   num_post_totali_punti=0
	end if
	
	 
				
	 QuerySQL1="SELECT Count(*) AS Numeropost, sum(Punti) "&_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "'  and ParentMessage <>0 and comments<>'InizializzaDB' and Id_Social=2" &_ 
" and (DatePosted>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104));"


   Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	num_messaggi=rsTabella1(0)
	'num_messaggi_punti=0
	num_messaggi_punti=rsTabella1(1)
	if isnull(num_messaggi_punti) then
	   num_messaggi_punti=0
	end if
%>

<div class="box box-color box-bordered">
                                            <div class="box-title">
                                                <h3>
                                                    <i class="icon-table"></i>
                                                    Diario: Discussioni aperte (<%=num_post_totali%>) + Commenti(<%=num_messaggi%>) = Punti(<%=num_post_totali_punti+num_messaggi_punti%>)
                                                </h3>
                                            </div> 
                                          <div class="box-content nopadding">

<%' carico i messaggi del forum
QuerySQL="SELECT * " &_
" FROM FORUM_MESSAGES " &_
" WHERE CodiceAllievo='" & cod & "' and Id_Classe='" &id_classe & "' and comments<>'InizializzaDB' and Id_Social=2 "&_ 
" and (DatePosted>= CONVERT(DATETIME,'" &DataClaq  &"', 104))" &_
	 " AND (DatePosted<= CONVERT(DATETIME,'" &(1+CDATE(DataClaq2)) &"', 104))"  &_ 
" ORDER BY ID desc;"
'" and  DatePosted>=" &formattaDataCla(DataClaq)  &"" &_
'	 " AND DatePosted<=" & 1 + cdate(formattaDataCla(DataClaq2)) &"" &_ 
'" ORDER BY ID;"

 ' response.write(QuerySQL)
 Set rsTabellaLavagna = ConnessioneDB3.Execute(QuerySQL)
 appQuery=QuerySQL

 'response.write( DataClaq  )
 
 
 QueryCat = "SELECT Descrizione FROM CAT_CAT WHERE Id_Categoria = '"&rsTabellaDiario("Id_Categoria")&"';"
		Set rsTabellaCat = ConnessioneDB.Execute(QueryCat)
		
		categorianome = rsTabellaCat(0)
 
 
 %>
 
 
<!-- Div tendina per i le visualizzazioni di post -->
 



<%If rsTabellaLavagna.BOF=True And rsTabellaLavagna.EOF=True Then %>
    <table class="table table-hover table-nomargin"> 
    <thead>             
			  <tr><th colspan="5" align="center">  Non ci sono attivit&agrave; nel diario</th></tr>
	</thead>		  
<% Else%>
<!--<table class="table table-hover table-nomargin dataTable table-bordered">-->
   <table class="table table-hover table-nomargin"> 
   <thead>
			<tr>
				<th>Post</th>
				<th>Messaggio</th>
                <th class='hidden-480'>Data</th>  
				<th class='hidden-480'>Punti</th>		
              
			</tr>
	</thead>
    <tbody>
		
	
 
	
	  
	          	<tr>
                
				 
		<% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   i=0
		   do while not rsTabellaLavagna.EOF  'and i<10
		   i=i+1
		     QuerySQL1="SELECT * "&_
" FROM FORUM_MESSAGES " &_
" WHERE ID=" & rsTabellaLavagna("ThreadParent") &" and comments<>'InizializzaDB' and Id_Social=2" &_
 
" ORDER BY ID desc;"
'Set objFSO = CreateObject("Scripting.FileSystemObject")
'				url="C:\Inetpub\umanetroot\anno_2012-2013\logForum3.txt"
'				Set objCreatedFile = objFSO.CreateTextFile(url, True)
'				objCreatedFile.WriteLine(QuerySQL1)
'				objCreatedFile.Close

Set rsTabella1 = ConnessioneDB3.Execute(QuerySQL1) 
	'num_visualizzazioni=rsTabella1(0) 
			 %>
			<% if not rsTabella1.eof then%>
            <tr><td><a title="Visualizza Post di apertura discussione" href="../cSocial/ShowMessage.asp?scegli=2&amp;ID=<%=rsTabellaLavagna("ThreadParent")%>&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid2%>&id_categoria<%=rsTabellaDiario("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabella1("Topic")%></a></td>
            
         
		 
         
            
            
                <td><a title="Visualizza il messaggio nella discussione"    href="../cSocial/ShowMessage.asp?scegli=2&amp;ID=<%=rsTabellaLavagna("ID")%>&amp;id_classe=<%=id_classe%>&amp;divid=<%=divid2%>&id_categoria<%=rsTabellaDiario("Id_Categoria")%>&categoria=<%=categorianome%>"><%=rsTabellaLavagna("Topic")%></a></td> 
                <td class='hidden-480'><%=rsTabellaLavagna("DatePosted")%></td> 
                <td class='hidden-480'><%=rsTabellaLavagna("Punti")%></td>                
			   
             </tr>  
		<%
		  end if
		 rsTabellaLavagna.movenext
		loop%>
        
        
        
        
        
        
        
        
       </tbody>
	<%end if%> 
    
   
</table>

 
 

