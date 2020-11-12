 
 

function onKeyDown() {  
  // current pressed key

  var pressedKey = String.fromCharCode(event.keyCode).toLowerCase();
  if (event.ctrlKey && (pressedKey == "c" || 
                        pressedKey == "v")) {
						alert("Non fare il furbo! Niente copia ed incolla, non ti serve a niente!!");
    // disable key press porcessing
    //event.returnValue = false;
  }
} // onKeyDown
 
/*function disabilita(element)
{ 
  element.oncontextmenu = function()
  { 
    return false; 
  }
} 
*/
function disabilita(elemento)
{ 
  elemento.disabled = true

} 

function abilita(elemento)
{ 
  elemento.disabled = false

} 

function getElement()
{ 
  //abilita(document.getElementById("Btn1"));  
  abilita(document.getElementById("txtS1")); 
  //document.write("  oncontextmenu='return false' ondragstart='return false' onselectstart='return false'");
  
}

function getElement1()
{ 
  abilita(document.getElementById("Btn1"));  
//  disabilita(document.getElementById("txtS1")); 
  //document.write("  oncontextmenu='return false' ondragstart='return false' onselectstart='return false'");
  
}
 
$(document).ready(function()
{
$(document).bind("contextmenu",function(e){
return true;
});

});

$(document).ready(function(){
      $('textarea').bind("cut copy paste drop",function(e) {
          e.preventDefault();
      });
		 $('input').bind("cut copy paste drop",function(e) {
          return true;
      });
	  
	  
    });
  

  
  