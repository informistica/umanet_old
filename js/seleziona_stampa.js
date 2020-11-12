
 
function checkTutti() {
	var stringa,stringa2;
	numcb=0;
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		stringa=document.dati.elements[i].name;
		stringa2='cbStampa';
		if (elements[i].type == 'checkbox' && (stringa.search(stringa2) == 0))
		    {
		     elements[i].checked = true;
			 numcb=numcb+1;
			}
		}
	}
	document.dati.txtNUMREC.value=numcb;
}
function uncheckTutti() {
	var stringa,stringa2;
	with (document.dati) {
		for (var i=0; i < elements.length; i++) {
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		stringa=document.dati.elements[i].name;
		stringa2='cbStampa';
		if (elements[i].type == 'checkbox' && (stringa.search(stringa2) == 0))
		//if (elements[i].type == 'checkbox')
		elements[i].checked = false;
		}
	 
	}
	document.dati.txtNUMREC.value=0;
	
}
function aggiorna(nome) {
	 
		with (document.dati) { 
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMREC.value=parseInt(txtNUMREC.value)+1;
		 else
		    txtNUMREC.value=parseInt(txtNUMREC.value)-1;
	    }	
}
function aggiorna2(nome) {
	 
		with (document.dati) { 
		//if (elements[i].type == 'checkbox' && elements[i].name == 'cb')
		// tolgo il controllo sul nome tanto ci sono solo questi nella pagina
		if (elements[nome].checked == true)
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)+1;
			
		 else
		    txtNUMVAL.value=parseInt(txtNUMVAL.value)-1;
	    }	
}
 //assegna la valutazione solo se il record è selezionato per la valutazione
function valutaTutti(voto) {
	var stringa,stringa2;
	var voto=document.dati.txtVoto.value;
	numcb=1;
	 
		for (var i=0; i < document.dati.elements.length; i++) {
			stringa=document.dati.elements[i].name;
			stringa2='txtVAl'+numcb;
			
		if (stringa.search(stringa2) == 0)
		     {
			if (document.dati.elements["cbVal"+numcb].checked == true) document.dati.elements[i].value = voto;
			numcb=numcb+1;
			 
		 	}
		}
}
 
 

 