 
<%
Class TestServer

  'Proprietà
  Private pathEnd1
  Public locale
 

  'Costruttore
  Private Sub Class_Initialize()
	  pathEnd1  =  Server.mappath(Request.ServerVariables("PATH_INFO")) 
	  if (left(pathEnd1,10)="c:\inetpub") then
		 locale=1
	  else
		 locale=0
	  end if 	
  End Sub
  
  
  
  Private Sub Class_Terminate()
    
  End Sub

End Class


%>
 