<% 
if session("Admin")=True and cod=Session("CodAdmin") then
QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo2='" & cod & "'" &_ 
" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"
else
QuerySQL="SELECT AVVISI.*, Allievi.Id_Classe FROM Allievi INNER JOIN AVVISI ON Allievi.CodiceAllievo = AVVISI.CodiceAllievo " &_
" WHERE Avvisi.CodiceAllievo='" & cod & "'" &_ 
" AND Id_Classe='"& id_classe &"' ORDER BY Data desc ;"
end if
'response.write(QuerySQL)
QuerySQL2=QuerySQL

' " and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
' " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 

'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni
appAvvisi= QuerySQL
 Set rsTabellaAvvisi = ConnessioneDB.Execute(QuerySQL)
 
' carico i messaggi della lavagna da lavagna.mdb per gli avvisi sui compiti da svolgere
 QuerySQL="SELECT  *  " &_
" FROM FORUM_MESSAGES " &_
" WHERE Id_Classe ='" & id_classe & "' and comments<>'InizializzaDB' " &_
" and ParentMessage=0 ORDER BY DatePosted desc ;"

' non metto le date tanto li cancello i compiti vecchi
' " and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND DatePosted<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 

'" and  DatePosted>=#" &DataClaq &"#" &_
'" AND DatePosted<=#" & DataClaq2  &"#" &_ 
'ho prelevato solo i moduli e paragrafi per i quali ci sono visualizzazioni 
 Set rsTabellaAvvisi2 = ConnessioneDB2.Execute(QuerySQL)
 'response.write("lllll:"&QuerySQL)
 appAvvisi2= QuerySQL
%>
 
 
 <!-- Tendina per gli avvisi -->
  
<a name="ancora_avvisi" href="#" onClick="Effect.toggle('dAvvisi','appear'); return false;"><span style="font-style:normal;" class="sottotitoloquaderno">&nbsp;&nbsp;AVVISI</span></a> 
<div id="dAvvisi" style="display:none;">
<div class="contenuti" style="background-color:#ffffff;width:auto;padding:10px;"> 
<p> 
<table id="zebra_stud" align=center border=1  width="95%"  >
<thead>
 
 
<%If rsTabellaAvvisi.BOF=True And rsTabellaAvvisi.EOF=True and rsTabellaAvvisi2.BOF=True And rsTabellaAvvisi2.EOF=True Then %>
			  <tr><th align="center"> Non ci sono avvisi ...  </th></tr>
			  
<% 


Else%>
	
			 <% 'conto i post totali 
			'  QuerySQL1="SELECT Count(*)  "&_
'" FROM AVVISI" &_
'" WHERE CodiceAllievo='" & cod & "'" &_
'" and  Data>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND Data<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &_ 
'";"
'   Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
'	num_avvisi_totali=rsTabella1(0)
'	if isnull(num_avvisi_totali) then  num_avvisi_totali=0

'me ne sbatto di far sapere quanti avvisi ci sono non so perchè ma il count (*) sballa
'QuerySQL1="SELECT Count(*) as Tot"&_
'" FROM FORUM_MESSAGES" &_
'" WHERE Id_Classe='" & Id_Classe & "';"
'
'   Set rsTabella1 = ConnessioneDB2.Execute(QuerySQL1) 
'   response.write(QuerySQL1)
  ' " and  DatePosted>=#" & mid(DataClaq,4,2)&"/" &left(DataClaq,2)&"/"& right(DataClaq,4)  &"#" &_
'	 " AND DatePosted<=#" & mid(DataClaq2,4,2)&"/" &left(DataClaq2,2)&"/"& right(DataClaq2,4)  &"#" &";"

  ' num_avvisi_totali_admin=rsTabella1("Tot")
   'if isnull(num_avvisi_totali_admin) then  num_avvisi_totali_admin=0

	'num_post_totali_punti=	
	 %>
    <%'if num_avvisi_totali_admin>0 then ' se ci sono messaggi mostro le torce che illuminano%>
             <tr><th colspan=5><center><img src="../../../img/fuocobacheca.gif" width="12" height="17">&nbsp
<b>Messaggi dal prof. a tutta la classe &nbsp; <img src="../../../img/fuocobacheca.gif" width="12" height="17">
</b></center> </th></tr>  
     <%'else%>
   <!--  <tr><th colspan=5><center><b>Messaggi dal prof. a tutta la classe 
</b></center> </th></tr>  -->
     
     <%'end if%> 
      </thead> 
   
              <% If rsTabellaAvvisi2.BOF=True And rsTabellaAvvisi2.EOF=True then
			  
			   %>
                <tr><th align="center">Nessuno ... </th></tr>
        
        <%else%>
				 <tr><th><b>Oggetto</b></th><th><b>Data</b></th></tr>
         <%end if%>
           
            

          <% k=0 
		     do while not rsTabellaAvvisi2.EOF and k<5 
               k=k+1%>
               
                
               
              <% response.write "<tr class="&classe_riga&"> <TD><A HREF='lavagna/ShowMessage.asp?ID=" & rsTabellaAvvisi2("ID") & "&RCount=" & rsTabellaAvvisi2("ReplyCount")& "&TParent=" & rsTabellaAvvisi2("ID")& "&divid=" & divid2 & "&id_classe=" & id_classe & "'>"  & rsTabellaAvvisi2("Topic") & "</A></FONT></TD>"
               %>
                
               
                <td><%=rsTabellaAvvisi2("DatePosted")%></td> </tr>	   
				<%rsTabellaAvvisi2.movenext
				loop%>
</table>
        <br><br>
        <form method="POST" name="Aggiorna">
      <table id="zebra_stud" align=center border=1 width="95%"  >
         <thead> 
         <%if num_avvisi_totali>0 then ' se ci sono messaggi mostro le torce che illuminano%>
             <tr><th colspan=5><center><img src="../../../img/fuocobacheca.gif" width="12" height="17">
<b>Messaggi personali(<%=num_avvisi_totali%>) <img src="../../../img/fuocobacheca.gif" width="12" height="17">
</b></center> </th></tr>  
     <%else%>
    <tr><th colspan=5><center><b>Messaggi personali(<%=num_avvisi_totali%>) 
</b></center> </th></tr> 
   
     
     <%end if%> 
        
       </thead>  
        
        <% If rsTabellaAvvisi.BOF=True And rsTabellaAvvisi.EOF=True then %>
                <tr><th align="center">Nessuno... </th></tr>
        
        <%else%>
				<% if session("Admin")=true then %>
                 <tr><th><b>Da</b></th><th><b>Testo</b></th><th><b>Data/Ora</b></th><th><b><a onClick="cancella_avviso();" href="#" title="Clicca per eliminare i messaggi selezionati">Elimina</a></b></th></tr>
				<%else%>
                 <tr><th><b>Da</b></th><th><b>Testo</b></th><th><b>Data/Ora</b></th><th><b><a onClick="cancella_avviso();" href="#" title="Clicca per eliminare i messaggi selezionati">Elimina</a></b></th></tr>
                <%end if%>
               
		 
         <%end if%>
		<% 'adesso per ogni messaggio guardo il post (topic) a cui si riferisce
		   i=1
		   do while not rsTabellaAvvisi.EOF 
		    CodiceAllievo=rsTabellaAvvisi("CodiceAllievo")
			 %>
			<tr>
              <td><%=rsTabellaAvvisi("Commentatore")%></td> 
               <td> <%=rsTabellaAvvisi("Azione")%></td>
                <td></b><%=rsTabellaAvvisi("Data")%></td> 
               
			   <td> 
               <input type="checkbox"  name="cbDelete<%=i%>" title="<%=i%>" value="<%=i%>" checked="true"> </td></tr>
               
		<% i=i+1
		   rsTabellaAvvisi.movenext
		loop %>
   
</table> 
     </form>   
        
        
	<%end if%> 
</table>

</p> 
<%
' inserisco avviso personale
  'if (ucase(session("CodiceAllievo"))=ucase(CodiceAllievo)) or (Session("Admin")=true) then %>
<center>  
    <a title="Modifica messaggio" href="#" onClick="Effect.toggle('addMessage','appear'); return false;">
    <span style="background: #c3d6e0;color: #FF0000;font-size: 12px;font-weight:bold;font-style:normal;">
    + Messaggio personale
    </span>
    </a> 
    <div id="addMessage" style="display:none;">
    <div style="width:500px;border:1px solid white;padding:10px;margin : 0 auto 0 auto;"> 
    <form action="../../cMessaggi/inserisci_messaggio_personale.asp?CodiceAllievo=<%=cod%>&amp;DataClaq=<%=DataClaq%>&amp;DataClaq2=<%=DataClaq%>&amp;cbEmail=1" METHOD = "POST">
     
    <br>Messaggio : <br>
    <textarea name="txtMessaggio" cols="60" rows="3"></textarea>
    <br> 
    <p> <input type="checkbox"  name="cbEmail" title="Selezionare per inviare un email allo studente">   Notifica per email &nbsp;&nbsp;&nbsp;<br>
      <p> <input type="submit" value="Inserisci"><br><br><hr style="width:35%"> </center>
     <br>
    <!-- <a href="aggiorna_messaggio.asp> Daglie</a>-->
    </form>
     
<%'end if%>


</div></div> 