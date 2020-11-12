<%if Session("CodiceAllievo")="" or Session("Id_Classe")="" then %> 
 <script language="javascript" type="text/javascript"> 
    window.alert("Sessione  scaduta, effettua nuovamente il Login!");
     location.href="../../home.asp";
  </script>
<% end if%>

