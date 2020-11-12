<?php
//
// miniDesign upload&unzip 2007
//
// Questo piccolo script usa la libreria Php PclZip  
// per estrarre i file da un archivio preventivamente
// caricato sul server.
// Lo script segue la GNU LESSER GENERAL PUBLIC LICENSE
// di PclZip contenuta nel pacchetto.
// Nel caso si intenda redistribuire questo pacchetto
// sarebbe apprezzato un link verso il sito dell'autore.
// https://www.minidesign.it
// Ulteriori indicazioni e istruzioni nel file readme.txt
//
// ******************************************************

require_once('pclzip.lib.php');
$scegli=$_REQUEST['scegli'];
$id_classe=$_REQUEST['id_classe'];
$bacheca=$_REQUEST['bacheca'];
$ID=$_REQUEST['ID'];
$RCount=$_REQUEST['RCount'];
$CodiceAllievo=$_REQUEST['CodiceAllievo'];
$Cartella=$_REQUEST['Cartella'];
$IDPARENT=$_REQUEST['IDPARENT'];
$Materia=$_REQUEST['Materia'];
$Social=$_REQUEST['Social'];
$homesito=$_REQUEST['homesito'];
$homeserver=$_REQUEST['homeserver'];

$indirizzo="../social/ShowMessage.asp?scegli=".$scegli."&id_classe=".$id_classe."&cartella=".$cartella."&bacheca=".$bacheca."&ID=".$ID."&RCount=".$RCount."&Zip=1";


 
//$nome_file= $homeserver."/".$homesito."/"."UECDL"."/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/".$_REQUEST['nome'];
$nome_file= $_REQUEST['nome']; // lo contieno già come variabile di sessione

$archive = new PclZip($nome_file);
echo "Nome=".$nome_file."\n";
//$v_path = $homeserver."/".$homesito."/"."UECDL"."/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/"; //devo settare il path di estrazione
$v_path ="../../Materie/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/"; 

$url=substr($nome_file,0,strlen($nome_file)-4);
//$list = $archive->extract(PCLZIP_OPT_PATH, $v_path,PCLZIP_OPT_EXTRACT_DIR_RESTRICTION, $v_path);
$list = $archive->extract(PCLZIP_OPT_PATH,  $v_path);

 
 $path_parts = pathinfo($_ENV["SCRIPT_FILENAME"]);
// echo $path_parts['dirname']."\n".$path_parts['basename']."\n".$path_parts['extension'];
       

if ($archive->extract() == 0) {
	die("<p>Si è verificato un errore: <br /> ".$archive->errorInfo(true)."</p><p><input type=\"button\" value=\"Riprova\" onClick=\"javascript:history.back(1)\">");
	 }
else { echo "<h1>Operazione eseguita.</h1><p>L'archivio ".$nome_file." è stato aperto con successo.</p><p> Puoi tornare alla <a href='".$indirizzo."'>Discussione</a></p>";

//unlink($nome_file); //rimuovi il file zip decompresso
//header("location: https://www.umanet.net");

	 }
?> 

    