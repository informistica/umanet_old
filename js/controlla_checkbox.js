// JavaScript Document
 function deselezionatutti() {
	//per modificare tutte le date di un form impostandole uguale al valore della textbox passata per parametro
    //document.dati.date3.value="11/11/1111";
	// document.dati.txtScadenza1.value="19/11/2010";
	
   
      var idtext=1;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=1; i <= elements.length; i++) {
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'inQuiz'+idtext)
		    {
		    elements[i].checked = false; 
			idtext=idtext+1;
			}
	 }
	 return true;
    }
 }
 
 
 
 