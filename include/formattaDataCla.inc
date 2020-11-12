 <%
function formattaDataCla(DataClass)
  giornoD=DatePart("d",DataClass)
 if len(giornoD)=1 then
    giornoD= "0" & giornoD
 end if
 meseD=DatePart("m",DataClass)
  if len(meseD)=1 then
    meseD= "0" & meseD
 end if
 annoD=DatePart("yyyy",DataClass)
 
 if cint(giornoD)=31 then
   giornoD="01"
   if cint(meseD)=12 then
      meseD="01"
      else
	  meseD= cint(meseD)+ 1 
	  if len(meseD)=1 then
       meseD= "0" & meseD
      end if   
	end if
	annoD=cint(annoD)+1
 end if
 
 
' response.write("Data=" &meseD&"/"&giornoD&"/"&annoD)
 formattaDataCla=meseD&"/"&giornoD&"/"&annoD
end function

function formattaDataCla2(DataClass)
  giornoD=DatePart("d",DataClass)
 if len(giornoD)=1 then
    giornoD= "0" & giornoD
 end if
 meseD=DatePart("m",DataClass)
  if len(meseD)=1 then
    meseD= "0" & meseD
 end if
 annoD=DatePart("yyyy",DataClass)
if (cint(giornoD)=31) or  ( (strcomp(meseD,"02")=0) and (cint(giornoD)=28)) then
   giornoD="01"
   if cint(meseD)=12 then
      meseD="01"
      else
	  meseD= cint(meseD)+ 1 
	  if len(meseD)=1 then
       meseD= "0" & meseD
      end if   
	end if
	annoD=cint(annoD)+1
 end if
 
 
 
' response.write("Data=" &meseD&"/"&giornoD&"/"&annoD)
 formattaDataCla2=meseD&"/"&giornoD+1&"/"&annoD
end function




function invertiDataCla(DataClass)
 giornoD=DatePart("d",DataClass)
 if len(giornoD)=1 then
    giornoD= "0" & giornoD
 end if
 meseD=DatePart("m",DataClass)
  if len(meseD)=1 then
    meseD= "0" & meseD
 end if
 annoD=DatePart("yyyy",DataClass)
' response.write("Data=" &meseD&"/"&giornoD&"/"&annoD)
 invertiDataCla=giornoD&"/"&meseD&"/"&annoD
end function

%>