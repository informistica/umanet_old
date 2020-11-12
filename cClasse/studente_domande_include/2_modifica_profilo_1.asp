
  <form method="POST" action="../cUtenti/modifica_profilo.asp?CodiceAllievo=<%=cod%>&id_classe=<%=id_classe%>" name="frmDocument1"  class="form-horizontal form-bordered form-validate">

                                 <div class="control-group">

										<label for="textfield"  class="control-label">Cognome</label>
										<div class="controls">
											<input type="text" style="height: auto;"  class="input-xlarge"  name="txtCognome" value="<%=rsTabella("Cognome")%>">
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Nome</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtNome" value="<%=rsTabella("Nome")%>">
										</div>
									</div>
									 <div class="control-group">
                                        <label for="textfield" class="control-label">Gruppo</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtTag" value="<%=rsTabella("Tags")%>">
										</div>
									</div>
									 <div class="control-group">
                                        <label for="textfield" class="control-label">In Quiz</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xlarge"  name="txtInQuiz" value="<%=rsTabella("In_Quiz")%>">
										</div>
									</div>


                                    <% if session("Admin")=true then


											  url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&Session("CartellaAdmin")&"/Profili/thumb"
										   else
											   url= "../../Db"&Session("DB")&"/Materie/"&Session("ID_Materia") &"/"&Session("Cartella")&"/Profili/thumb" ' vuole il percorso relativo della cartella
										   end if
										  url=Replace(url,"\","/")
										  if strcomp(rsTabella("Url_img")&"","")=0 then
										   ' urlimg="https://www.placehold.it/80/EFEFEF/AAAAAA&text=no+image"
										   urlimg="../../img/no-image.jpg"
										  else
										  urlimg=url&"/"& Url_img
										  end if
										 %>

									<div class="control-group">
										<label for="textfield" class="control-label">Foto</label>
										<div class="controls">
                                        <img class="imground"  src="<%=urlimg%>" />
                                        </div>
                                        <div class="controls">
											  <div class="accordion" id="accordion2">
  <div class="accordion-group">
    <div class="accordion-heading">
      <a class="accordion-toggle" data-toggle="collapse" data-parent="#accordion2" href="#collapseOne">
        Aggiungi
      </a>
    </div>
    <div id="collapseOne" class="accordion-body collapse">
      <div class="accordion-inner">
        <iframe src="../upload_resize/ex2_imgprofilo.asp?cartella=<%=Session("Cartella")%>" name="postmessage" id="postmessage" width="100%" height="30%" frameborder="0" SCROLLING="no" border="0" class="iframe">
      </iframe>
      </div>
    </div>
  </div>
</div>

										</div>
									</div>


									<div class="control-group">
										<label for="textfield"  class="control-label" >Mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;"   class="input-xxlarge"  name="txtmipiace" value="<%=rsTabella("Mipiace")%>" >
										</div>
                                     </div>
                                     <div class="control-group">
                                        <label for="textfield" class="control-label">Non mi Piace</label>
										<div class="controls">
											<input type="text" style="height: auto;" class="input-xxlarge"  name="txtnonmipiace" value="<%=rsTabella("Nonmipiace")%>">
										</div>
									</div>
									<div class="control-group">
										<label for="textfield" class="control-label">Descriviti</label>
										<div class="controls">
											 <p><textarea  rows="12" name="S1"  cols="116" class="input-block-level"><%=trim(rsTabella("Descriviti"))%></textarea></p>

										</div>
									</div>

									<div class="form-actions">
										<button type="submit" class="btn btn-primary">Salva modifiche</button>

									</div>
								</form>

								 <% if Session("Admin") = true then %>


		   <div class="box box-bordered">
							<div class="box-title">
								<h3><i class="icon-list-ul"></i> Impostazioni Personalizzate (Utente: <%=CodiceAllievo%>)&nbsp;</h3>
							</div>


    <!-- inizio include-->

<form method="POST" name="votazioni" class="form-horizontal form-striped">

 <% if CIAbilitato = 1 then
	checked1 = "checked"
	checked2 = ""
 else
	checked1 = ""
	checked2 = "checked"

	end if %>


 <div class="control-group">
   <label class="control-label">

      &nbsp; <b>Copia e Incolla</b>
    </label>
    <div class="controls">
        <input id="CIvero" name="CheckCI" <%=checked1%> value="1" type="RADIO">&nbsp;Si&nbsp;&nbsp;
        <input name="CheckCI" <%=checked2%> value="0" type="RADIO">&nbsp;No
    </div>
  </div>

  <div class="control-group">
    <label class="control-label">
       &nbsp; <b>Probababilità 0pt eccezioni ()</b>
     </label>
     <div class="controls">

         <input name="probabilita" id="probabilita" value="<%=Probabilita%>" type="text">&nbsp;
     </div>
   </div>





   <div class="control-group">
   <label class="control-label">

      &nbsp; <b>Aggiorna</b>
    </label>
    <div class="controls">
        <input value="Esegui" name="B1" class="btn" onclick="aggiornaimpostazioni();" type="button">
    </div>
  </div>

   </form>
  </div>


		   <% end if %>
