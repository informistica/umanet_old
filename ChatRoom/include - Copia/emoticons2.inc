<%

 QuerySQL="Select * from TUTTESMILES where ID_Categoria=1 order by Posizione;"
   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
   rsTabellaS.movefirst
   

if daShowChat2=0 then

	do while not rsTabellaS.eof %>
	 
	  
	 <a href="Javascript:postmessage.AddSmileyIcon('<%=rsTabellaS("Codice")%>');"><img border="0" hspace="10" src="<%=rsTabellaS("Cartella_Cat")&"/"& rsTabellaS("Url")%>"></a> 
  	
	<%    rsTabellaS.movenext
   loop	


 %> 
 <img border="0" hspace="10" src="smilies/on_1.gif" onclick='javascript:postmessage.AddSmileyIcon(":cippa")'> 
 <!--
 <a href="Javascript:postmessage.AddSmileyIcon(':huh?');"><img border="0" hspace="10" src="smilies/on_1.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':s');"><img border="0" hspace="10" src="smilies/on_2.gif"></a>
 <a href="Javascript:postmessage.AddSmileyIcon(':P');"><img border="0" hspace="10" src="smilies/on_3.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon('}:)');"><img border="0" hspace="10" src="smilies/on_4.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':D');"><img border="0" hspace="10" src="smilies/on_5.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon('}:|');"><img border="0" hspace="10" src="smilies/on_6.gif"></a>  
 <a href="Javascript:postmessage.AddSmileyIcon(':)');"><img border="0" hspace="10" src="smilies/on_7.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':oops');"><img border="0" hspace="10" src="smilies/on_8.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(';)');"><img border="0" hspace="10" src="smilies/on_9.gif"></a>  
 <a href="Javascript:postmessage.AddSmileyIcon(':pff');"><img border="0" hspace="10" src="smilies/on_10.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':/');"><img border="0" hspace="10" src="smilies/on_11.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':gg');"><img border="0" hspace="10" src="smilies/on_12.gif"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(':b;');"><img border="0" hspace="10" src="smilies/on_13.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':gg');"><img border="0" hspace="10" src="smilies/on_15.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':nn');"><img border="0" hspace="10" src="smilies/on_16.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':pp');"><img border="0" hspace="10" src="smilies/on_17.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':kk');"><img border="0" hspace="10" src="smilies/on_18.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':yy');"><img border="0" hspace="10" src="smilies/on_19.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':zz');"><img border="0" hspace="10" src="smilies/on_20.gif"></a> 
 <a href="Javascript:postmessage.AddSmileyIcon(':xx');"><img border="0" hspace="10" src="smilies/on_14.gif"></a> 
 -->
 <%else%>
 
<% do while not rsTabellaS.eof %>
	 
	 <a href="Javascript:addsmile('<%=rsTabellaS("Codice")%>');"><img border="0" hspace="10" src="<%=rsTabellaS("Cartella_Cat")&"/"& rsTabellaS("Url")%>"></a> 
	 
	<%    rsTabellaS.movenext
   loop	%>
  
  <!--
  <a href="Javascript:addsmile(':huh?');"><img border="0" hspace="10" src="smilies/on_1.gif"></a> 
 <a href="Javascript:addsmile(':s');"><img border="0" hspace="10" src="smilies/on_2.gif"></a>
 <a href="Javascript:addsmile(':P');"><img border="0" hspace="10" src="smilies/on_3.gif"></a> 
 <a href="Javascript:addsmile('}:)');"><img border="0" hspace="10" src="smilies/on_4.gif"></a> 
 <a href="Javascript:addsmile(':D');"><img border="0" hspace="10" src="smilies/on_5.gif"></a> 
 <a href="Javascript:addsmile('}:|');"><img border="0" hspace="10" src="smilies/on_6.gif"></a>  
 <a href="Javascript:addsmile(':)');"><img border="0" hspace="10" src="smilies/on_7.gif"></a> 
 <a href="Javascript:addsmile(':oops');"><img border="0" hspace="10" src="smilies/on_8.gif"></a> 
 <a href="Javascript:addsmile(';)');"><img border="0" hspace="10" src="smilies/on_9.gif"></a>  
 <a href="Javascript:addsmile(':pff');"><img border="0" hspace="10" src="smilies/on_10.gif"></a> 
 <a href="Javascript:addsmile(':/');"><img border="0" hspace="10" src="smilies/on_11.gif"></a> 
 <a href="Javascript:addsmile(':gg');"><img border="0" hspace="10" src="smilies/on_12.gif"></a> 
 
 <a href="Javascript:addsmile(':b;');"><img border="0" hspace="10" src="smilies/on_13.gif"></a> 
 <a href="Javascript:addsmile(':gg');"><img border="0" hspace="10" src="smilies/on_15.gif"></a> 
 <a href="Javascript:addsmile(':nn');"><img border="0" hspace="10" src="smilies/on_16.gif"></a> 
 <a href="Javascript:addsmile(':pp');"><img border="0" hspace="10" src="smilies/on_17.gif"></a> 
 <a href="Javascript:addsmile(':kk');"><img border="0" hspace="10" src="smilies/on_18.gif"></a> 
 <a href="Javascript:addsmile(':yy');"><img border="0" hspace="10" src="smilies/on_19.gif"></a> 
 <a href="Javascript:addsmile(':zz');"><img border="0" hspace="10" src="smilies/on_20.gif"></a> 
 <a href="Javascript:addsmile(':xx');"><img border="0" hspace="10" src="smilies/on_14.gif"></a> 
 -->
 
 <%end if%>
 
 
 
 
 
           
 
 <br>