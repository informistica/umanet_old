<%
 QuerySQL="Select * from TUTTESMILES where ID_Categoria=0 order by Posizione, Pos;"
   Set rsTabellaS = ConnessioneDB1.Execute(QuerySQL)   
   rsTabellaS.movefirst

if daShowChat2=0 then  

do while not rsTabellaS.eof %>
	 
	  
	 <a href="Javascript:postmessage.AddSmileyIcon('<%=rsTabellaS("Codice")%>');"><img title='<%=rsTabellaS("Title")%>  '   border="0" width='50' height='50'  src="../img_social/<%=rsTabellaS("Cartella_Cat")&"/"& rsTabellaS("Url")%>"></a> 
  	
	<%    rsTabellaS.movenext
   loop	


 %> 

<!--

 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_00');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_0_MatrixOmino_2.jpg"></a> 

<a href="Javascript:postmessage.AddSmileyIcon(' :;0_01');"> 
<img width='50' height='50'   src="../img_social/connessioni_percezioni/0_1_incredibile_solo.jpg"> </a>
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_02');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_2_vedoilsole.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_03');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_3_NavigaOcchioSi.jpg"></a>
  
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_04');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_4_MondoCoerenzaVerde.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_05');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_5_LampadinaAccesa.jpg"></a>  
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_06');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_6_TestaAccesa.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_07');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_7_vedopioggia.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_08');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_8_NavigaOcchioNo.jpg"></a>  
 
  
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_09');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_9_MondoParadossoRosso.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_10');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_10_LampadinaSpenta.jpg"></a> 
 
 <a href="Javascript:postmessage.AddSmileyIcon(' :;0_11');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_11_TestaSpenta.jpg"></a> 
 -->
 
 
 <%else
 
 do while not rsTabellaS.eof %>
	 
	  
	 <a href="Javascript:addsmile('<%=rsTabellaS("Codice")%>');"><img title='<%=rsTabellaS("Title")%>  '  border="0" width='50' height='50'  src="../img_social/<%=rsTabellaS("Cartella_Cat")&"/"& rsTabellaS("Url")%>"></a> 
  	
	<%    rsTabellaS.movenext
   loop	


 %> 
 
 
 <!--
    <a href="Javascript:addsmile(':;0_00');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_0_MatrixOmino_2.jpg"></a> 

<a href="Javascript:addsmile(':;0_01');"> 
<img width='50' height='50'   src="../img_social/connessioni_percezioni/0_1_incredibile_solo.jpg"> </a>
 
 <a href="Javascript:addsmile(':;0_02');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_2_vedoilsole.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_03');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_3_NavigaOcchioSi.jpg"></a>
  
 <a href="Javascript:addsmile(':;0_04');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_4_MondoCoerenzaVerde.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_05');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_5_LampadinaAccesa.jpg"></a>  
 
 <a href="Javascript:addsmile(':;0_06');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_6_TestaAccesa.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_07');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_7_vedopioggia.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_08');"> 
 <img width='50' height='50' src="../img_social/connessioni_percezioni/0_8_NavigaOcchioNo.jpg"></a>  
 
  
 <a href="Javascript:addsmile(':;0_09');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_9_MondoParadossoRosso.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_10');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_10_LampadinaSpenta.jpg"></a> 
 
 <a href="Javascript:addsmile(':;0_11');"> 
 <img width='50' height='50'    src="../img_social/connessioni_percezioni/0_11_TestaSpenta.jpg"></a> 
 
 -->
 
 <%end if%> 