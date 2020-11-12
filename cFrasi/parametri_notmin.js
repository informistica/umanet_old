if(l=="/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp"){
frmDocument.S1.value = "";
}
var temp = new Array();
var i = 1;
var caratteri_iniz = 0;

if(l!="/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp"){
caratteri_iniz = frmDocument.txtSpiegazione.value.length;
}
var caratteri_attuali = 0;
var eccessotime = false;

data = Date.now();
temp[0] = parseInt(data/1000);

var t = setInterval(function(){

var tempocontrollo = 1;
var limitecaratteri = 13;
if(i%tempocontrollo == 0){
		if(l=="/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp"){
		caratteri_attuali = frmDocument.S1.value.length;
		}else{
		caratteri_attuali = frmDocument.txtSpiegazione.value.length;
		}
	
	if(caratteri_attuali-caratteri_iniz > limitecaratteri){
		eccessotime = true;
	}else{
		caratteri_iniz = caratteri_attuali;
	}
}

data = Date.now();
temp[i] = parseInt(data/1000);

i++;

}, 1000);

function getParametri(x){
clearInterval(t);

var errore = false;

for(var j=1; j<temp.length && !errore; j++){

	if(temp[j]-temp[j-1] < -1 || temp[j]-temp[j-1] > 4){
	
		if(!(temp[j] == 0 && (temp[j-3] == 57 || temp[j-2] == 58 || temp[j-1] == 59))){
			errore = true;
		}
		
		if(temp[j] == 1 && temp[j-1] < 57){
			errore = true;
		}
		
	}
	
}

/*if(l=="/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp"){
		caratteri_attuali = frmDocument.S1.value.length;
		}else{
		caratteri_attuali = frmDocument.txtSpiegazione.value.length;
		}
	
	if(caratteri_attuali-caratteri_iniz > limitecaratteri){
		eccessotime = true;
	}*/

if(errore){
alert("Hai disabilitato JavaScript! Inserimento annullato.");
window.location.href=window.location.href;
}else if(eccessotime && ci == 0){
alert("Rilevato uno standard input. Inserimento annullato.");
window.location.href=window.location.href;
}else{
	
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
			alert("Il tuo orologio non è sincronizzato: inserimento annullato. Sincronizza l'orario del tuo PC con quello di Internet prima di riprovare.");
			window.location.href=window.location.href;
		}else{
			if(l=="/expo2015Server/UECDL/script/cFrasi/2inserisci_frase.asp"){
				if(x==0){
					//document.getElementById("Btn1").type="submit";
					//document.getElementById("Btn1").click();
					document.getElementById("frmDocument").submit();
				}else{
					invio();
				}
			}else{
				if(x==0){
					//document.getElementById("btnNoImg").type="submit";
					//document.getElementById("btnNoImg").click();
					document.getElementById("frmDocument").submit();
				}else{
					checkImg();
				}
			}
		}
		
	}) /* .done */
	
}

}