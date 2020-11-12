 
<%
 pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
	  if (left(pathEnd1,10)<>"D:\inetpub") then
		 locale=1
	  else
		 locale=0
	  end if 	
	 ' response.write(left(pathEnd1,10))
	'response.write("locale="&locale)
                if locale=1 then
			
				%>
					  <%DataAvviso = giorno & "/" & mese& "/" & anno  %>
					 <% else
					' response.write("online")
					 %>
						 <%DataAvviso = mese & "/" & giorno& "/" & anno  %>
					 <% end if %>    


%>
 