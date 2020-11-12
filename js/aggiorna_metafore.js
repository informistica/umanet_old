// JavaScript Document
function aggiorna_metafore(tipo){
var cartella, CodiceAllievo,CodiceMetafora,Codice_Test,Modulo,Paragrafo;
 		  cartella=document.getElementById("cartella").value;
		  CodiceAllievo=document.getElementById("CodiceAllievo").value;
		  CodiceMetafora=document.getElementById("CodiceMetafora").value;
		  Codice_Test=document.getElementById("Codice_Test").value;
		  Modulo=document.getElementById("Modulo").value;
		  Paragrafo=document.getElementById("Paragrafo").value;		  
	switch(tipo) {
	  case 0:	  	
		  txtTopolino=document.getElementById("txtTopolino").value;
		  txtFormaggio=document.getElementById("txtFormaggio").value;
		  txtFame=document.getElementById("txtFame").value;
		  txtLabirinto=document.getElementById("txtLabirinto").value;
		  txtStrada=document.getElementById("txtStrada").value;
		  txtStrada_OK=document.getElementById("txtStrada_OK").value;
		  txtStrada_KO=document.getElementById("txtStrada_KO").value;
		  txtTestata=document.getElementById("txtTestata").value;
		  txtDistanza=document.getElementById("txtDistanza").value;
		  txtData=document.getElementById("txtDATA").value;
		  textarea=document.getElementById("textarea").value;
		  segnalata=document.getElementById("cb1").checked;
		  voto=document.getElementById("txtVAl").value;
		  dati2="&txtTopolino="+txtTopolino+"&txtFormaggio="+txtFormaggio+"&txtFame="+txtFame+"&txtLabirinto="+txtLabirinto+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtTestata="+txtTestata+"&txtDistanza="+txtDistanza+"&txtSegnalata="+segnalata+"&S1="+textarea+"&txtData="+txtData+"&txtVAl="+voto;		 
		 break;
	  case 1:
			txtAutista=document.getElementById("txtAutista").value;
			txtDestinazione=document.getElementById("txtDestinazione").value;
			txtCarburante=document.getElementById("txtCarburante").value;
			txtLuogo=document.getElementById("txtLuogo").value;
			txtStrada=document.getElementById("txtStrada").value;
			txtStrada_OK=document.getElementById("txtStrada_OK").value;
			txtStrada_KO=document.getElementById("txtStrada_KO").value;
			txtCespugli=document.getElementById("txtCespugli").value;
			txtLupo=document.getElementById("txtLupo").value;
			txtCestino=document.getElementById("txtCestino").value;
			txtDistanza=document.getElementById("txtDistanza").value;
			txtData=document.getElementById("txtDATA").value;
			textarea=document.getElementById("textarea").value; 
			segnalata=document.getElementById("cb1").checked; 
			voto=document.getElementById("txtVAl").value;
			dati2="&txtAutista="+txtAutista+"&txtDestinazione="+txtDestinazione+"&txtCarburante="+txtCarburante+"&txtLuogo="+txtLuogo+"&txtStrada="+txtStrada+"&txtStrada_OK="+txtStrada_OK+"&txtStrada_KO="+txtStrada_KO+"&txtCespugli="+txtCespugli+"&txtCestino="+txtCestino+"&txtLupo="+txtLupo+"&txtDistanza="+txtDistanza+"&txtSegnalata="+segnalata+"&S1="+textarea+"&txtData="+txtData+"&txtVAL="+voto;
			break;
	  case 2:
			txtSoggettoC=document.getElementById("txtSoggettoC").value;
			txtDomandaC=document.getElementById("txtDomandaC").value;
			txtMotivazioneC=document.getElementById("txtMotivazioneC").value;
			txtDesiderioC=document.getElementById("txtDesiderioC").value;
			txtBisognoC=document.getElementById("txtBisognoC").value;
			txtSoggettoS=document.getElementById("txtSoggettoS").value;
			txtRispostaS=document.getElementById("txtRispostaS").value;
			txtMotivazioneS=document.getElementById("txtMotivazioneS").value;
			txtDesiderioS=document.getElementById("txtDesiderioS").value;
			txtBisognoS=document.getElementById("txtBisognoS").value;
			txtTipoEvento=document.getElementById("txtTipoEvento").value;
			txtTolleranzaC=document.getElementById("txtTolleranzaC").value;
			segnalata=document.getElementById("cb1").checked;
			textarea=document.getElementById("textarea").value;
			txtData=document.getElementById("txtDATA").value;
			voto=document.getElementById("txtVAl").value;
			dati2="&txtSoggettoC="+txtSoggettoC+"&txtDomandaC="+txtDomandaC+"&txtMotivazioneC="+txtMotivazioneC+"&txtDesiderioC="+txtDesiderioC+"&txtBisognoC="+txtBisognoC+"&txtSoggettoS="+txtSoggettoS+"&txtRispostaS="+txtRispostaS+"&txtMotivazioneS="+txtMotivazioneS+"&txtDesiderioS="+txtDesiderioS+"&txtBisognoS="+txtBisognoS+"&txtTipoEvento="+txtTipoEvento+"&txtSegnalata="+segnalata+"&S1="+textarea+"&txtData="+txtData+"&txtVAl="+voto+"&txtTolleranzaC="+txtTolleranzaC;		 
			break;
	} 
	
	dati="cartella="+cartella+"&CodiceAllievo="+CodiceAllievo+"&CodiceMetafora="+CodiceMetafora+"&Codice_Test="+Codice_Test+"&Modulo="+Modulo+"&Paragrafo="+Paragrafo; 
    var url = "7_aggiorna_metafora_ajax.asp?"+dati+dati2;			   
	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
	  if (xhttp.readyState == 4 && xhttp.status == 200) {
		  var testo = xhttp.responseText;		
		  alert(testo);			 
	  }
	};
	xhttp.open("GET", url, true);
	xhttp.send();	
		
}