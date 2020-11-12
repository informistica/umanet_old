<?php

 
$classe=$_GET["classe"];
$id_classe=$_GET["id_classe"];
$anno=$_GET["anno"];
$anno1="as_".$anno;
$urlreport="../../../grafici/".$anno1."/report&".$classe.".json";
$urlpath="/expo2015Server/UECDL/script/cGrafici/secondaVersione/index.php";
$file=file_get_contents($urlreport);
$elenco=json_decode($file,true);
$periodoInizioM;
$periodoFineM=0;
$periodoInizio=-1;
$periodoFine=-1;
$periodi=dividiPeriodi($elenco["intestazione"]);
$numeroPeriodi=count($periodi);

$pagina = "https://" . $_SERVER['SERVER_NAME'] ;
 

if(isset($_GET["periodoInizio"]))
  if($_GET["periodoInizio"]=="12_09_2013")
    $periodoInizio=0;
  else {
    $periodoInizioM=$_GET["periodoInizio"];
    if(controllaData($periodoInizioM))
      $periodoInizio=trovaData($periodoInizioM,$periodi);
    else
      $periodoInizio=0;
  }
else $periodoInizio=0;

if(isset($_GET["periodoFine"])) {
  $periodoFineM=$_GET["periodoFine"];
  if($periodoFineM=="oggi")
  $periodoFine=$numeroPeriodi-1; 
  else {
  if(controllaData($periodoFineM))
    $periodoFineM=trovaData($periodoFineM,$periodi);
  if($periodoFineM<$periodoInizio) {
    $periodoFine=$periodoInizio-1;
  }else $periodoFine=$periodoFineM;
}
}else $periodoFine=$numeroPeriodi;
$urlcompleto=$pagina.$urlpath."?id_classe=".$id_classe."&classe=".$classe."&anno=".$anno;
 

function controllaData($s) {
  //Data: gg_mm_aaaa
  ////////0123456789
  //Controllo che la lunghezza sia corretta
  if(strlen($s)!=10)
    return false;
  //Controllo che giorno, mese e anno siano numeri
  if(!is_numeric(substr($s,0,2)))
    return false;
  if(!is_numeric(substr($s,3,2)))
    return false;
  if(!is_numeric(substr($s,6,4)))
    return false;
  //Fine dei controlli
  return true;
}
function dividiPeriodi($s) {
  $periodi=explode('&', $s);
  $n=array_filter($periodi,"controllaData");
  return array_values($n);
}
function trovaData($s,$p) {
  for ($i=0; $i < count($p); $i++)
    if($s==$p[$i])
      return $i;
  return count($p);
}
?>
<!DOCTYPE html>
<html lang="it">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@2.9.3/dist/Chart.min.js"></script>
  <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/css/bootstrap.min.css" integrity="sha384-9aIt2nRpC12Uk9gS9baDl411NQApFmC26EwAOH8WgZl5MYYxFfc+NcPb1dKGj7Sk" crossorigin="anonymous">
  <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" integrity="sha384-DfXdz2htPH0lsSSs5nCTpuj/zy4C+OGpamoFVy38MVBnE+IbbVYUew+OrCXaRkfj" crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/popper.js@1.16.0/dist/umd/popper.min.js" integrity="sha384-Q6E9RHvbIyZFJoft+2mJbHaEWldlvI9IOYy5n3zV9zzTtmI3UksdQRVvoxMfooAo" crossorigin="anonymous"></script>
  <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.5.0/js/bootstrap.min.js" integrity="sha384-OgVRvuATP1z7JjHLkuOU7Xw704+h835Lr+6QL9UvYjZE3Ipu6Tp75j7Bh/kR0JKI" crossorigin="anonymous"></script>
  <title>Grafico</title>
</head>

<body>
  <div class="container-fluid" id="contenitore">
    <canvas id="grafico"></canvas>
  </div>
  <div id="periodi">
    <form action="" name="formPeriodi" method="get">
    <select name="periodoInizio" id="periodoInizio">
      <?php
      foreach($periodi as $p)
        echo "<option value=".$p.">".$p."</option>\n";
      ?>
    </select>
    <select name="periodoFine" id="periodoFine">
    <?php
      foreach($periodi as $p)
        echo "<option value=".$p.">".$p."</option>\n";
    ?>
    <option value="oggi">Oggi</option>
    </select>
    <input type="button" class="btn" value="Invia" onclick="aggiorna();">
    </form>
  </div>
  <script src="secondarie.js"></script>
  <script>
     function aggiorna(){       
            let PI=document.getElementById("periodoInizio").value;
            let PF=document.getElementById("periodoFine").value;
            let urlfetch ="<?php echo $urlcompleto ?>"+"&periodoInizio="+PI+"&periodoFine="+PF;
            //alert(urlfetch);
            location.href = urlfetch;
           // document.formPeriodi.action = urlfetch;
		      //	document.formPeriodi.submit();
      }
    let risultati, intestazione;
    let file;
    //Prendo il file con la memoria
    fetch("<?php echo $urlreport; ?>")
    .then(d => d.json())
    .then(d => {
      file = d;
      finito()
    })
    .catch(e => console.error(e));

    function finito() {
      risultati = file.risultati;
      intestazione = file.intestazione;
      let inizio=<?php echo $periodoInizio?>;
      let fine=<?php echo $periodoFine?>;
      let punti = prendiPunti(risultati,inizio,fine);
      let nomi = prendiNomi(risultati);
      resetCanvas(); //Elimino il grafico precedente per evitare che si sovrapponga
      bubbleSort(punti,nomi); //Ordino i nomi e i punti
      creaGrafico(nomi,punti); //Creo il grafico
    }
    //Seleziono le date correnti
    <?php
    echo "$(\"#periodoInizio\").children().eq(".$periodoInizio.").attr(\"selected\",\"selected\");\n";
    if($periodoFineM=="oggi")
    echo "$(\"#periodoFine\").children().eq($(\"#periodoFine\").children().length-1).attr(\"selected\",\"selected\");\n";
    else echo "$(\"#periodoFine\").children().eq(".$periodoFine.").attr(\"selected\",\"selected\");\n";
    ?>
    $("#periodoInizio").change(()=>{
      //Prendo i due select
      let periodoInizio=document.getElementById("periodoInizio");
      let periodoFine=document.getElementById("periodoFine");
      $("#periodoFine").empty(); //Svuoto il select finale
      let indice=periodoInizio.options.selectedIndex; //Prendo l'indice selezionato
      //Riempio il select finale solo con i voti validi
      for(let i=0; i<periodoInizio.options.length; i++)
      if(i>=indice) {
        let option=document.createElement("option");
        option.value=periodoInizio.options[i].value;
        option.innerText=periodoInizio.options[i].value;
        $("#periodoFine").append(option);
      }
      //Aggiungo la voce ultimo
      let option=document.createElement("option");
        option.value="oggi";
        option.innerText="Oggi";
        $("#periodoFine").append(option);
    })
    
  </script>
</body>

</html>