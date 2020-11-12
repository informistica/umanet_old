function selezionatutticheckbox() {
	var idtext=1;
    //window.alert(el.value);
    with (document.dati) {
	for (var i=0; i <= elements.length; i++) {
		
		//window.alert(elements[i].name + elements[i].value);
		 if (elements[i].name == 'cb'+idtext)
		    {
		    elements[i].checked = true; 
			idtext=idtext+1;
			}
	 }
    }
	}