var offset = new Date().getTimezoneOffset()/-60;
var now = new Date();
var datanow = new Date(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(),  now.getUTCHours(), now.getUTCMinutes(), now.getUTCSeconds());

var printdata = 0;
if(datanow.getHours()<10){printdata="0"+datanow.getHours();}else{printdata=datanow.getHours();}
if(datanow.getMinutes()<10){printdata+=":0"+datanow.getMinutes();}else{printdata+=":"+datanow.getMinutes();}
if(datanow.getSeconds()<10){printdata+=":0"+datanow.getSeconds();}else{printdata+=":"+datanow.getSeconds();} 
		
	$.ajax({
		method: "POST",
		url: "verificaparametri.asp",
		dataType: "html",
		data: { fuso: offset, h: printdata }
	}) /* .ajax */
	.done(function( ans ) {
		
		if(ans == "notsync"){
			alert("Il tuo orologio non è sincronizzato: sincronizza l'orario del tuo PC con quello di Internet prima di inserire.");
			window.location.href="sincronizzaorologio.asp";
		}
		
	});