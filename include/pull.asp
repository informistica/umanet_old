<script>
//ATTENTO Dario, non svelare ai tuoi compagni come fare copia ed incolla. In questo modo non li aiuti a crescere ma li aiuti a rubare un valore (punti umanet) al quale non hanno diritto. 
let numCaratteri = 0,
    larghezza;
let l;
let testo;
CKEDITOR.replace('editor1');
CKEDITOR.instances.editor1.on('paste', function(evt) {
    evt.cancel()
});
CKEDITOR.instances.editor1.on('change', function() {
    numCaratteri++;
    testo = String(CKEDITOR.instances.editor1.getData());
    l = testo.length;
    larghezza = $(window).width();
    document.getElementById("progressbar").setAttribute('style', 'width:' + (l / larghezza) * 180 + '%;')
});

function rimpiazza(testo) {
    var pulito = new String(testo);
    pulito = pulito.replace(/&agrave;/g, "à");
    pulito = pulito.replace(/&ograve;/g, "ò");
    pulito = pulito.replace(/&ugrave;/g, "ù");
    pulito = pulito.replace(/&egrave;/g, "è");
    pulito = pulito.replace(/&igrave;/g, "ì");
    pulito = pulito.replace(/&nbsp;/g, " ");
    pulito = pulito.replace(/&#39;/g, "`");
    pulito = pulito.replace("'", "`");
    return pulito
}

function immagineinTesto(){
    var testo = String(CKEDITOR.instances.editor1.getData());
     
    if(testo.search("img") == -1)
        return 0;
    else 
        return 1;
}
function inviaDati(params) {
    var testo = String(CKEDITOR.instances.editor1.getData());
    var xhttp = new XMLHttpRequest();
    var ok=0;
    var url = "2inserisci_frase1.asp?Quesito=<%=Quesito%>&ID_Prefrase=<%=ID_Prefrase%>&prefrase=<%=prefrase%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Img=<%=Img%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>";
    $('#frmDocument').attr('action', url);
    document.getElementById("txtEncode").value = testo;
    


  let immagini;
    if (typeof(frmDocument.txtImg1)==="object")
       immagini=1
    else
        immagini=0

 if (immagini==1){
    if (immagineinTesto()==1)
       ok=1;
       else
       ok=0;
        
    if (ok==0) {
        var stringa1=frmDocument.txtImg1.value;
        var stringa2=frmDocument.txtImg2.value;
        var stringa3=frmDocument.txtImg3.value;
        if(stringa1.search("http") == -1 && stringa2.search("http") == -1 && stringa3.search("http") == -1){
           // alert("Devi inserire almeno un url con protocollo http/https");
            ok=0;
        } else ok=1;
    }
        //alert(testo);
    if (ok==1)
        document.getElementById("frmDocument").submit();
    // alert("invio");
    else 
    alert("Devi inserire almeno un immagine");
    }
    else
        document.getElementById("frmDocument").submit();


}

function caricaDati() {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (xhttp.readyState == 4 && xhttp.status == 200) {
            var testo = xhttp.responseText;
            try {
                CKEDITOR.instances.editor1.setData(decodeURIComponent(testo))
            } catch (e) {
                console.error(e);
                CKEDITOR.instances.editor1.setData(testo)
            }
        }
    };
    var url = "2carica_frase1_ck.asp?Quesito=<%=Quesito%>&ID_Prefrase=<%=ID_Prefrase%>&prefrase=<%=prefrase%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Img=<%=Img%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>";
    xhttp.open('GET', url)
    xhttp.send()
}

function stampaDati() {
    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function() {
        if (xhttp.readyState == 4 && xhttp.status == 200) {
            var testo = xhttp.responseText;
            try {
                CKEDITOR.instances.editor1.setData(decodeURIComponent(testo))
            } catch (e) {
                console.error(e);
                CKEDITOR.instances.editor1.setData(testo)
            }
        }
    };
    var url = "2carica_frase1_ck.asp?Quesito=<%=Quesito%>&ID_Prefrase=<%=ID_Prefrase%>&prefrase=<%=prefrase%>&Cartella=<%=Cartella%>&Nome=<%=Nome%>&Cognome=<%=Cognome%>&CodiceTest=<%=CodiceTest%>&Capitolo=<%=Capitolo%>&Paragrafo=<%=Paragrafo%>&Modulo=<%=Modulo%>&Img=<%=Img%>&CodiceSottopar=<%=CodiceSottopar%>&Sottoparagrafo=<%=Sottoparagrafo%>";
    xhttp.open('GET', url)
    xhttp.send()
}
</script>