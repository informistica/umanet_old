function rimpiazza(testo){
	var pulito = new String(testo);
		pulito = pulito.replace(/&agrave;/g,"à");
		pulito = pulito.replace(/&ograve;/g,"ò");
		pulito = pulito.replace(/&ugrave;/g,"ù");
		pulito = pulito.replace(/&egrave;/g,"è");
		pulito = pulito.replace(/&igrave;/g,"ì");
		pulito = pulito.replace(/&nbsp;/g," ");
		//pulito = pulito.replace(/&/g,"e");
		pulito = pulito.replace(/&#39;/g,"`");
		
		pulito = pulito.replace("'","`");

		return pulito;
}


	function aggiorna_post(){

	   var titolo=document.getElementById("titolopost").value;	 
		var testo= rimpiazza(CKEDITOR.instances.editor1.getData());
		//testo=encodeURI(testo);
	 //  var url = "aggiorna_post_ajax.asp?id="+globalpostid+"&titolo="+titolo+"&testo="+testo;
 	   //alert(testo);
	   var xhttp = new XMLHttpRequest();
	   xhttp.onreadystatechange = function() {
	   if (xhttp.readyState == 4 && xhttp.status == 200) {
		  var testoJSON=JSON.parse(xhttp.responseText);
					stato=testoJSON["stato"];
					messaggio=testoJSON["messaggio"];
					if (stato==0) 
						alert('Errore: '+messaggio);
					document.getElementById("titolo"+globalpostid).innerHTML="<b>"+titolo+"</b>";
					document.getElementById(globalpostid).innerHTML=decodeURIComponent(testo);
					//document.getElementById(globalpostid).innerHTML=decodeURI(testo);
					 $('#chiudi').click();
	  }
	};
	//xhttp.open("GET", url, true);
	//xhttp.send();
	
       
		var url="aggiorna_post_ajax.asp?id="+globalpostid+"&titolo="+titolo;
		//alert($("#editor1").serialize());
		//testo=$("#editor1").serialize();
		testo=encodeURIComponent(testo);
	 	params="testo="+testo;
		xhttp.open('POST', url) 
		xhttp.setRequestHeader('Content-type', 'application/x-www-form-urlencoded')
		xhttp.send(params);
		 
		}


var to;
$().ready(function()
{
    to = setTimeout("TimedOut()",600000); //ogni 10 minuti
});

function TimedOut()
{
    $.post("../service/refresh_session.asp", null,
    function(data)
    { console.log(data);
        if(data == "Session refreshed")
        {
            to = setTimeout("TimedOut()", 600000);
        }
        else
        {
          //  $("#timeout").slideDown('fast'); sessione scaduta
        }
    });
}
