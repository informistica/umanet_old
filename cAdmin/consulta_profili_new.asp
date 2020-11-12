<%@ Language=VBScript %>
<%
  on error resume next
   CodiceAllievo=Session("CodiceAllievo")
   id_classe=Session("Id_Classe")
  ' classe=Session("Cartella")
    classe=request.QueryString("classe")
   dividA=request.QueryString("dividApro")
   
  
  'Apertura della connessione al database
   Set ConnessioneDB = Server.CreateObject("ADODB.Connection")
    
%>   
  <!-- #include file = "../service/controllo_sessione.asp" -->
   <!-- #include file = "../stringhe_connessione/stringa_connessione.inc" -->
   <!-- #include file = "../var_globali.inc" -->
   
 
 
 <%
' VERIFICHIAMO SE L'UTENTE E' IDENTIFICATO (LOGGATO)

IF Session("Loggato") = True then

QuerySQL="Select * from Allievi where CodiceAllievo='"&CodiceAllievo& "' and Attivo=1;" 
Set rsTabella1 = ConnessioneDB.Execute(QuerySQL)

End IF


Function gira_data()
  	gira_data=Day(date())&"/"&Month(date())&"/"&Year(date())
End Function 
 
 

%>



<html> 
<head>
   
   <title>Profili classe</title>   
   
   <!-- #include file = "../include/header.asp" -->
    
   <script>
   
   
   
function PopUpWindow(w,h) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;

	window.open('../upload_resize/ex2_imgprofilo.asp','../upload_resize/ex2_imgprofilo.asp', 'toolbar=0,location=0,status=0,menubar=0,scrollbars=0,resizable=1,width=800,height=365,top='+wint+',left='+winl)
}
// -->
 

	function uploadImgWindow(form, imgField, thumbField, imgPath, thumbPath, prev, imgWidth, imgHeight, thumbWidth, thumbHeight) {
		var upload = window.open('<%=pageUpload%>?field=' + form + '.' + imgField + '&path=' + imgPath + (prev != '' ? '&prev=' + prev : '') + '&thumbField=' + form + '.' + thumbField + '&thumbPath=' + thumbPath + '&imgWidth=' + imgWidth + '&imgHeight=' + imgHeight + '&thumbWidth=' + thumbWidth + '&thumbHeight=' + thumbHeight, 'upload', 'toolbar=0,location=0,directories=0,status=0,menubar=0,scrollbars=0,resizable=1,width=600,height=200');
		upload.focus();
	}
</script>

 <script language="javascript" type="text/javascript" >
 function validate2() {
	 
	 if (frmDocument.txtCognome.value=="")
	{
	   alert("Non hai inserito il cognome");
	   frmDocument.txtCognome.setfocus();
	   return 0;
	}
 else
 if (frmDocument.txtNome.value=="")
	{
	   alert("Non hai inserito il nome");
	   frmDocument.txtNome.setfocus();
	   return 0;
	}
 else
 if (frmDocument.txtCodiceAllievo.value=="")
	{
	   alert("Non hai inserito il vecchio username ");
	   frmDocument.txtCodiceAllievo.setfocus();
	   return 0;
	}
	else
	if (frmDocument.txtPwdAllievo.value=="")
	{
	   alert("Non hai la vecchia password");
	   frmDocument.txtPwdAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewCodiceAllievo.value=="")
	{
	   alert("Non hai inserito il nuovo username ");
	   frmDocument.txtNewCodiceAllievo.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd.value=="")
	{
	   alert("Non hai inserito la nuova password");
	   frmDocument.txtNewPwd.setfocus();
	   return 0;
	}else
	if (frmDocument.txtNewPwd1.value=="")
	{
	   alert("Non hai confermato la nuova password");
	   frmDocument.txtNewPwd1.setfocus();
	   return 0;
	}else
	{
		
		 document.frmDocument.action = "modifica_pwd.asp?stato=<%=stato%>&cla=<%=cla%>&StringaConnessione=<%=Request.Cookies("Dati")("StrConn")%>&id_classe=<%=id_classe%>&divid=<%=divid%>";  
		
	   
		document.frmDocument.submit();
		
	 
    }
	
}


function validate3() {
	 
	 if (frmDocument1.txtNewEm.value=="")
	{
	   alert("Non hai inserito la nuova email");
	   frmDocument1.txtNewEm.setfocus();
	   return 0;
	}
 else
 if (frmDocument1.txtNewEm1.value=="")
	{
	   alert("Non hai confermato la nuova email");
	   frmDocument1.txtNewEm1.setfocus();
	   return 0;
	}
 else
 if (frmDocument1.txtNewEm1.value !=  frmDocument1.txtNewEm.value)
	{
	   alert("Le due email non corrispondono ");
	   frmDocument1.txtEm1.setfocus();
	   return 0;
	}
	else
	{
		
		 document.frmDocument1.action = "modifica_contatti.asp";  
		
	   
		document.frmDocument1.submit();
		
	 
    }
	
}
 </script>
   
</head>




<body class='theme-<%=session("stile")%>' data-layout-sidebar="fixed" data-layout-topbar="fixed">
	<div id="navigation">
    	 
     
		<!-- #include file = "../include/navigation.asp" -->
       
          
          
	</div>
    
    
    
    
	<div class="container-fluid" id="content">
   
      <!-- #include file = "../include/menu_left.asp" -->
     
     <%
	 id_classe=request.QueryString("id_classe")
  divid=request.QueryString("divid")
  
	'QuerySQL="Select Url_img from Classi where ID_Classe='" & id_classe & "';" 
'	
'	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
'	if not rsTabella.eof then
'		urlimg=rsTabella(0)
'	else
'		urlimg=""
'	end if
'	urlC= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&classe&"/Profili/img" ' vuole il percorso relativo della cartella
'    urlC=Replace(urlC,"\","/")
'	

 
 
  

  QuerySQL="Select Cartella from Classi where ID_Classe='" & id_classe & "';" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	cartella=rsTabella("Cartella")

    QuerySQL="Select * from Allievi where Id_Classe='" & id_classe & "' and Attivo=1 order by Cognome asc;" 
	Set rsTabella = ConnessioneDB.Execute(QuerySQL) 
	classe=rsTabella("Classe")

	QuerySQL1="Select * from Allievi where Id_Classe='" & id_classe & "' and Attivo=1 order by CodiceAllievo asc;" 
	Set rsTabella1 = ConnessioneDB.Execute(QuerySQL1) 
	 %>
   
	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">               
					<div class="pull-left">
						<h1> 	<i class="icon-user"></i> Classe <%=left(classe,1+len (classe)- instr(classe,"$"))%> </h1> 
                        	 
					</div>
					 
				</div>
             
                 
                 
                 
                 
                 
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
			
				      
 
                
               <div class="bs-docs-example">
               <div class="box box-color box-bordered">
                <% if strcomp(urlimg&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è = a ""
    
  	urlimgclasse=urlC&"/"&"profilo_vuoto.png" %>	
     <img class="imground" src="<%=urlimgclasse%>" > <br>
<%else%>
    <% urlimgclasse=urlC&"/"& urlimg ' aggiungo al percorso il nome del file%>
     <img class="imground" src="<%=urlimgclasse%>" >  <br>
<%end if %>
				</div>
                 </div>
                 </div>
                 
                 
                <div class="bs-docs-example"> 
                 
                  <ul id="myTab2" class="nav nav-tabs">
                                    <li class="active"><a href="#profileP" data-toggle="tab">Profili</a></li>   
                                      
									    
                                    <% if session("Admin")=true then %>    
									<li><a href="#profileC" data-toggle="tab">Contatti</a></li>                       
                                    <li><a href="#profileL" data-toggle="tab">Login</a></li>  
									 <li><a href="#profileU" data-toggle="tab">Username</a></li>
                                                                           
       								<%end if%>
                            </ul>
                            <div id="myTabContent2" class="tab-content">
                             
                              <div class="tab-pane fade  in active" id="profileP">
                         
                               
                               
                               
                               
                               
                               
                               
                     <!----Inizio -->           
					<div class="row-fluid">
					<div class="span12">
                    
                    
                    
                    
                    
                    
                    
                    
                   <% i=0
	url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&cartella&"/Profili/thumb" ' vuole il percorso relativo della cartella
    url=Replace(url,"\","/")
   i=0
   do while not rsTabella.eof %>
   
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									  <%=rsTabella("Cognome") & "  " & rsTabella("Nome")%>
								</h3>
							</div>
							<div class="box-content nopadding">
								<form action="#" class="form-horizontal form-bordered">
								 
									<div class="control-group">
										<label for="textfield" class="control-label">Foto</label>
										<div class="controls">
											<div class="fileupload fileupload-new" data-provides="fileupload">
     
                                            <% 'if session("Admin")=true then
											'  url= "../../Materie/"&Session("ID_Materia") &"/"&Session("CartellaAdmin")&"/Profili/thumb"
										  ' else
											   url= "../../DB"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&cartella&"/Profili/thumb" ' vuole il percorso relativo della cartella
										 '  end if
										  url=Replace(url,"\","/")
										 ' response.write(url&"<br>")
										  %>
										 <% if strcomp(rsTabella("Url_img")&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è =   
											urlimg=url&"/"&"profilo_vuoto_thumb.png" %>	
											
										<%else%>
											<% urlimg=url&"/"& rsTabella("Url_img") ' aggiungo al percorso il nome del file%>
											 
										<%end if %> 
												<div class="fileupload-new thumbnail" style="width: 80px; height: 80px; padding-top:20px;"><img class="imground"  src="<%=urlimg%>" /></div>
												<div class="fileupload-preview fileupload-exists thumbnail" style="max-width: 200px; max-height: 150px; line-height: 20px;"></div>
												<div>			 
												</div>
											</div>
										</div>
									</div>
                                    
                                    
									<div class="control-group">
										<label for="textfield"  class="control-label" >Mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;"   class="input-xxlarge"  name="mipiace" value="<%=rsTabella("Mipiace")%>" >
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Non mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xxlarge"  name="nonmipiace" value="<%=rsTabella("Nonmipiace")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label">Descriviti</label>
										<div class="controls">
											 <p><textarea  rows="6" name="S1"   class="input-block-level">
                                             <%=ltrim(response.write(rsTabella("Descriviti")))%>
                                             </textarea></p>
   
										</div>
									</div>
									<!--
									<div class="form-actions">
										<button type="submit" class="btn btn-primary">Salva modifiche</button>
										 
									</div>
                                    -->
								</form>
							</div>
						</div>
                        <%
   rsTabella.movenext
   i=i+1
  loop 
 %>
                        
                        
					</div>
				</div>
                 <!-- >fine form -->   
                 
                            
                               
                               
                               
   
                               
                               
                              </div>
                              
                              
                              
                              
                              
                              
                               <div class="tab-pane fade" id="profileL">
                            
                            
  
     			 <!----Inizio -->           
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Modifica Login
								</h3>
							</div>
							<div class="box-content nopadding">
								<form name="dati" method="post" class="form-horizontal form-bordered">
								 
									 
                                    
                                    
                                  
									<div class="control-group">
										<label for="textfield"  class="control-label">Cognome</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtCognome" value="<%=rsTabella1("Cognome")%>">
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nome</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtNome" value="<%=rsTabella1("Nome")%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield"  class="control-label">Username</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtCodiceAllievo"value="<%=rsTabella1("CodiceAllievo")%>">
										</div>
                                     </div>
                                     
                                     
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Password</label>
										<div class="controls">
                                          <% if session("Admin")=true then%>
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtPwdAllievo" value="<%=rsTabella1("Password")%>">
                                            <%else%>
                                            <input type="password"  class="input-xlarge"  name="txtPwdAllievo" value="<%=rsTabella1("Password")%>">
                                            <%end if%>
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield"  class="control-label">Nuovo Username</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtNewCodiceAllievo" value="" placeholder="Inserisci">
										</div>
                                     </div>
                                     <div class="control-group">
										<label for="textfield"  class="control-label">Nuova Password</label>
										<div class="controls">
											<input type="password"   style="height: auto;" class="input-xlarge"  name="txtNewPwd" value="">
										</div>
                                        <label for="textfield"  style="height: auto;" class="control-label">Conferma Password</label>
										<div class="controls">
											<input type="password"   style="height: auto;" class="input-xlarge"  name="txtNewPwd1" value="">
										</div>
                                     </div>
                                     
                                     
									
                                    <%if (ucase(session("CodiceAllievo"))= ucase(CodiceAllievo)) or (Session("Admin")=true) then %>
                                       
                                  	 <div class="form-actions">
										<button type="button"  class="btn btn-primary" name="B2" onClick="return validate2();">Salva modifiche</button>	 
									</div>
                                    <%end if%>
                                    
									
								</form>  
							</div>
                            
                            
                             <%i=0
					  rsTabella.movefirst()
					  do while not rsTabella.eof%>
 
 
 
 
                    
						<div class="box">
							 
							<div class="box-content nopadding">
								<form action="#" class="form-horizontal form-bordered">
								 
									 
                               
                                    
                                  
									<div class="control-group">
                                     <label for="textfield"  class="control-label">Indirizzo </label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="www.elexpo.net/doc">
                                        </div>
                                       <label for="textfield"  class="control-label">Docente </label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("Cognome") & "  " & rsTabella("Nome")%>">
                                        </div>
										<label for="textfield"  class="control-label">Username</label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("CodiceAllievo")%>">
                                        </div>
                                        <label for="textfield"  class="control-label">Password</label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("Password")%>">
                                        </div>
                                        
                                        
                                     </div>
                                 
									
								</form>
							</div>
						</div>
					
                 <!-- >fine form --> 
                  <%
				   rsTabella.movenext
				   i=i+1
				  loop %>
                            
                            
						</div>
                 <img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa();">     
					</div>
				</div>
                 <!-- >fine form -->                 
 
                              </div>
                               





     <div class="tab-pane fade" id="profileU">
 
     			 <!----Inizio -->           
					<div class="row-fluid">
					<div class="span12">
						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<i class="icon-user"></i>
									Elenco username
								</h3>
							</div>
							<div class="box-content nopadding">
								<form name="dati" method="post" class="form-horizontal form-bordered">
   
									<div class="control-group">
										<label for="textfield"  class="control-label">Cognome</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtCognome" value="<%=rsTabella1("Cognome")%>">
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nome</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtNome" value="<%=rsTabella1("Nome")%>">
										</div>
									</div>
                                    
                                    <div class="control-group">
										<label for="textfield"  class="control-label">Username</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtCodiceAllievo"value="<%=rsTabella1("CodiceAllievo")%>">
										</div>
                                     </div>
        	
								</form>  
							</div>
                            
                            
                             <%i=0
					  rsTabella.movefirst()
					  do while not rsTabella.eof%>
 
						<div class="box">
							 
							<div class="box-content nopadding">
								<form action="#" class="form-horizontal form-bordered">
	
									<div class="control-group">
                                     
                                       <label for="textfield"  class="control-label">Utente </label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("Cognome") & "  " & rsTabella("Nome")%>">
                                        </div>
										<label for="textfield"  class="control-label">Username</label>
										<div class="controls">
                                            <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("CodiceAllievo")%>">
                                        </div>
                                        
                                        
                                     </div>
                                 
									
								</form>
							</div>
						</div>
					
                 <!-- >fine form --> 
                  <%
				   rsTabella.movenext
				   i=i+1
				  loop %>
                            
                            
						</div>
                 <img src="../../img/printer.jpg" title="Stampa questa scheda" onClick="stampa();">     
					</div>
				</div>
                 <!-- >fine form -->                 
 
                              </div>


                               
                               
                               <div class="tab-pane fade" id="profileC">
   
                                	 <!----Inizio -->           
					<div class="row-fluid">
					<div class="span12">
    
                    <%i=0
					  rsTabella.movefirst()
					  do while not rsTabella.eof%>

										 <% if strcomp(rsTabella("Url_img")&"","")=0 then ' evidentemente quando non è indicata un immagine il campo non è =   
											urlimg=url&"/"&"profilo_vuoto_thumb.png" %>	
											
										<%else%>
											<% urlimg=url&"/"& rsTabella("Url_img") ' aggiungo al percorso il nome del file%>
											 
										<%end if %> 

						<div class="box box-color box-bordered">
							<div class="box-title">
								<h3>
									<!--<i class="icon-user"></i>-->
                                     <img src="<%=urlimg%>" title="<%=trim(Cognome)%>&nbsp; <%=trim(Nome)%> " width="38px" height="38px" class="imground">
                                     
									  <%=rsTabella("Cognome") & "  " & rsTabella("Nome")%>
								</h3>
							</div>
							<div class="box-content nopadding">
								<form action="#" class="form-horizontal form-bordered">

									<div class="control-group">
										<label for="textfield"  class="control-label">Email</label>
										<div class="controls">
                                            <input type="hidden" name="txtCodiceAllievo" value="<%=rsTabella("CodiceAllievo")%>">
                                             <% if strcomp(rsTabella("Email")&"","")=0 then %>
                                                <input type="text" style="height: auto;"  class="input-xlarge" placeholder="Nessuna" name="txtEm">
										     <%else%>
                                                <input type="text" style="height: auto;"  class="input-xlarge"  name="txtEm" value="<%=rsTabella("Email")%>">
                                             <%end if%>
                                        </div>
                                     </div>
                                     <% if session("Admin")=true then%>
                                         <div class="control-group">
                                            <label for="textfield" class="control-label">Nuova email</label>
                                            <div class="controls">
                                                <input type="text" style="height: auto;" class="input-xlarge"  name="txtNome" value="">
                                            </div>
                                        </div>
                                        
                                        <div class="control-group">
                                            <label for="textfield"  class="control-label">Conferma email</label>
                                            <div class="controls">
                                                <input type="text" style="height: auto;"  class="input-xlarge"  name="txtCodiceAllievo">
                                            </div>
                                         </div>
                                       
                                         
                                      
                                           
                                         <div class="form-actions">
                                            <button type="button"  class="btn btn-primary" name="B2" onClick="return validate3();">Salva modifiche</button>	 
                                        </div>
                                        <%end if%>
									
								</form>
							</div>
						</div>
					
                 <!-- >fine form --> 
                  <%
				   rsTabella.movenext
				   i=i+1
				  loop 
				 %>             
					</div>
				</div>						   
                               
                               
        
                            </div>
                   </div>
         
       
       
  <%      
QuerySQL="Select count(*) from Allievi where Id_Classe='" & id_classe & "'  and Attivo=1 order by Cognome asc;" 
	Set rsTabella1 = ConnessioneDB.Execute(QuerySQL) 
	

NumStud=rsTabella1(0)
'Numero=cint(NumStud)
'Dim stud()
'Redim stud(Numero)
rsTabella.movefirst

					
%>  
         
        
     <% if session("Admin")=true then %>   
        <div class="accordion" id="accordion2">
          <div class="accordion-group">
										<div class="accordion-heading">
											<a class="accordion-toggle collapsed" data-toggle="collapse" data-parent="#accordion2" href="#collapse4">
												<center>Copia dati studenti</center>
											</a>
										</div>
										<div id="collapse4" class="accordion-body collapse">
											<div class="accordion-inner">
                                            <textarea rows=<%=NumStud%> class="input-block-level">
 <% 
 do while not rsTabella.eof
	 
     'response.write(rsTabella("Cognome")& "  " & left(rsTabella("Nome"),1)&"." & chr(13))
	 response.write(rsTabella("Cognome")& chr(13))
  rsTabella.movenext
 loop
 
 rsTabella.movefirst
  response.write("----------" & chr(13))
  do while not rsTabella.eof
	 
     response.write(rsTabella("Nome")& chr(13)) 
  rsTabella.movenext
 loop
 
 rsTabella.movefirst
  response.write("----------" & chr(13))
  do while not rsTabella.eof
	 
     response.write(rsTabella("Cognome")& "  " & rsTabella("Nome")& " - " & rsTabella("Email")& chr(13))
    
  rsTabella.movenext
 loop
 
 rsTabella.movefirst
  response.write("----------" & chr(13))
  i=1
  do while not rsTabella.eof
	 
     response.write(i&") "&rsTabella("Cognome")& " " & rsTabella("Nome")& chr(13))
      i=i+1
  rsTabella.movenext
 loop

 rsTabella.movefirst
  response.write("email x calendar" & chr(13))
  i=1
  do while not rsTabella.eof
	 
     response.write(rsTabella("Email")& ",")
      i=i+1
  rsTabella.movenext
 loop

  response.write("----------" & chr(13))
  do while not rsTabella1.eof
	 
    response.write(rsTabella1("Cognome")& "," & rsTabella1("Nome")& "," & rsTabella1("CodiceAllievo")& chr(13))
  rsTabella1.movenext
 loop
 

 %>
 </textarea>
                                            </div>
                                         </div>
                                      </div>
                                   </div>
        
      <% end if%>
  
  
   
        
        
 
 
  
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      </div>
			        </div>
			      </div>
			    </div>
			</div>
            
            
		</div> <!--fine main-->
        </div>
        
        <!-- #include file = "../include/colora_pagina_sint.asp" -->
         
  <script>
   function stampa() {
    document.dati.action = "../cFrasi/7_stampa_schede_login.asp?id_classe=<%=id_classe%>";
		//document.dati.action = "../../home.asp"
		document.dati.submit();	
}
   </script>
		 
	</body>

	<!-- InstanceEnd --></html>