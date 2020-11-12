
 // JavaScript Document
 function selezionatutti(id) {
	//per modificare tutte le date di un form impostandole uguale al valore della textbox passata per parametro
    //document.dati.date3.value="11/11/1111";
	// document.dati.txtScadenza1.value="19/11/2010";
	
    var el = document.getElementById(id);
    var idtext=0;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=0; i < elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'txtScadenza'+idtext)
		    {
		    elements[i].value = txtScadenza1.value=el.value; 
			idtext=idtext+1;
			}
	 }
	 return true;
    }
 }
 
 /* function selezionauno(id,idtext) {
	//per modificare tutte le date di un form impostandole uguale al valore della textbox passata per parametro
    //document.dati.date3.value="11/11/1111";
	// document.dati.txtScadenza1.value="19/11/2010";
	 
    var el = document.getElementById(id);
    //var idtext=0;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=0; i < elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'txtScadenza'+idtext)
		    {
		    elements[i].value = txtScadenza1.value=el.value; 
			//idtext=idtext+1;
			}
	 }
	 return true;
    }
 }
 */