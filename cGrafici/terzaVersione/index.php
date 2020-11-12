<?php

if(isset($_GET["nome"]))
  $nome=$_GET["nome"];
else
  $nome="";

 
  $classe=$_GET["classe"];

  $anno=$_GET["anno"];
  $anno="as_".$anno;

  //$urlreport="../../../grafici/as_1920/report&3Ct$6.json";
  $urlreport="../../../grafici/".$anno."/report&".$classe.".json";

  //$urlreport=$protocollo.$dominio.$homesito."/grafici/".$anno."/report&".$classe.".asp";
  //echo $urlreport;
  //$urlreport="https://www.umanetexpo.net/expo2015Server/UECDL/grafici/as_1920/report&3Ct$6.asp";
  //echo $urlreport;
  //$file=file_get_contents("https://www.umanetexpo.net/expo2015Server/UECDL/grafici/as_1920/report&3Ct$6.asp");
  
  //$file=file_get_contents(realpath($urlreport));
  $file=file_get_contents($urlreport);
  $elenco=json_decode($file,true);

$voti=[];
for ($i=0; $i < count($elenco["risultati"]); $i++)
  if($elenco["risultati"][$i][0]==$nome||trim($elenco["risultati"][$i][0])==$nome)
    for($j=2; $j<count($elenco["risultati"][$i])-2; $j++)
      if(($j-2)%3==0)
        array_push($voti,$elenco["risultati"][$i][$j]);
        
function elencaVoti($v) {
  for ($i=0; $i < count($v); $i++)
    echo $v[$i].",";
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
    
  <?php
  //echo $urlreport;
  //echo "realpath:".realpath($urlreport);
  ?>
      <canvas id="grafico"></canvas>
  </div>
  <script src="secondarie.js"></script>
  <script>
    let risultati, intestazione;
    let file;
    //Prendo il file con la memoria
    fetch("<?php echo $urlreport ?>")
    .then(d => d.json())
    .then(d => {
      file = d;
      finito()
    })
    .catch(e => console.error(e));

    function finito(alunni = []) {
      risultati = file.risultati;
      intestazione = file.intestazione;
      let date = prendiDate(intestazione);
      let voti = [<?php elencaVoti($voti) ?>];
      let nome = "<?php echo $nome ?>";
      resetCanvas(); //Serve per evitare che i grafici si sovrappongano
      console.log(date,voti,nome);
      creaGrafico(date, nome,voti);
    }

  </script>
</body>

</html>