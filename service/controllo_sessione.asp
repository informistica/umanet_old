<%



'response.Redirect "home_manutenzione.html"
' if Session("CodiceAllievo")="" or Session("Id_Classe")="" and 1>2 then  ' il problema è session("id_classe") che rimane ="" mentre 1>2 è un trick per farlo entrare lo stesso, domani ci guardiamo meglio
if Session("CodiceAllievo")="" or Session("Id_Classe")="" then
doc=Request.Cookies("Dati")("DOC")
pageset = Request.Cookies("Dati")("DB")
  if  (strcomp(doc,"1")=0) then %> 

 <script language="javascript" type="text/javascript"> 
    window.alert("Sessione  scaduta, effettua di nuovo il Login!");
     location.href="../../home.asp";
  </script>
<% 

  else%>
  
		  <% if  (strcomp(pageset,"2")=0) then %>
           <script language="javascript" type="text/javascript"> 
            window.alert("Sessione  scaduta, effettua di nuovo il Login!");
             location.href="../../home.asp";
          </script>
          <%end if%>
  <script language="javascript" type="text/javascript"> 
    window.alert("Sessione  scaduta, effettua di nuovo il Login!");
     location.href="../../../../index.html";
  </script>
  
  <%end if
  
  

session.Abandon()
end if%>