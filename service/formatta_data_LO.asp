 <% 
   function formatta_data_LO(Data)
    DataTest=formatDateTime(Data,2)
    'gira_data=Day(DataTest)&"/"&Month(DataTest)&"/"&Year(DataTest)
  
    if day(DataTest) < 10 then
    giorno="0" & day(DataTest) 
	else
	giorno=day(DataTest)
    end if
	
	if len(year(DataTest) ) = 2 then
	anno="20"& year(DataTest)
	elseif len(year(DataTest) ) =  3 then
	anno="2"& year(DataTest)
	else
	anno=year(DataTest)
	end if
    if month(DataTest) < 10 then
    mese="0" & month(DataTest) 
	else
	mese=month(DataTest)
    end if
	 pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
	  if (left(pathEnd1,10)<>"D:\inetpub") then
		 locale=1
	  else
		 locale=0
	  end if 	
	 ' response.write(left(pathEnd1,10))
	'response.write("locale="&locale)
     if locale=1 then
	     formatta_data_LO = giorno & "/" & mese& "/" & anno   
	 else				 
		 formatta_data_LO = mese & "/" & giorno& "/" & anno  
     end if 
	 
end function%>
