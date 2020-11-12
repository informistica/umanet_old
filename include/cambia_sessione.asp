 
<%
 id_classe=request.querystring("id_classe")
 divid=request.querystring("divid")
 cartella=request.querystring("cartella")
 
   if (strcomp(id_classe,Session("Id_Classe"))<>0) and (id_classe<>"") then
        
	   Session("Id_Classe")=id_classe
       Session("cartella")=cartella
	   %>
	    
	<!--	<script language="javascript" type="text/javascript"> 
	 	window.alert("Cambio sessione!");
  		</script> 
   -->
 <%  else%>
	 <!--	 <script language="javascript" type="text/javascript"> 
	 	window.alert("Non cambio sessione!"); 
  		</script >-->
     
 <%end if%>
   

