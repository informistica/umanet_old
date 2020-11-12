
<!doctype html>
<html>
<head>

   <title>Unzip</title>

       <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" />
	<!-- Apple devices fullscreen -->
	<meta name="apple-mobile-web-app-capable" content="yes" />
	<!-- Apple devices fullscreen -->
	<meta names="apple-mobile-web-app-status-bar-style" content="black-translucent" />



		<!-- Bootstrap -->
	<link rel="stylesheet" href="../../css/bootstrap2.min.css">
    <!-- jQuery UI -->
    <link rel="stylesheet" href="../../css/plugins/jquery-ui/smoothness/jquery-ui2.css">
     <link rel="stylesheet" href="../../css/style-themes.css">




	<!-- jQuery -->
	<script src="../../js/jquery.min.js"></script>
	<!-- Nice Scroll -->
	<script src="../../js/plugins/nicescroll/jquery.nicescroll.min.js"></script>
	<!-- imagesLoaded -->
	<script src="../../js/plugins/imagesLoaded/jquery.imagesloaded.min.js"></script>

    <!-- jQuery UI -->
	 <script src="../../js/plugins/jquery-ui/megaJQuery.js"></script>

	<!-- Touch enable for jquery UI -->
	<script src="../../js/plugins/touch-punch/jquery.touch-punch.min.js"></script>
	<!-- slimScroll -->
	<script src="../../js/plugins/slimscroll/jquery.slimscroll.min.js"></script>
	<!-- Bootstrap -->
	<script src="../../js/bootstrap.min.js"></script>
	<!-- Bootbox -->
	<script src="../../js/plugins/bootbox/jquery.bootbox.js"></script>

	<!-- Theme framework -->
	<script src="../../js/eakroko.min.js"></script>
	<!-- Theme scripts -->
	<script src="../../js/application.min.js"></script>
	<!-- Just for demonstration -->

	<!-- Favicon -->
	<link rel="shortcut icon" href="../../img/favicon.ico" />
	<!-- Apple devices Homescreen icon -->
	<link rel="apple-touch-icon-precomposed" href="../../img/apple-touch-icon-precomposed.png" />


   <!--   <script type="text/javascript" src="calendar/calendar.js"></script>
<script type="text/javascript" src="calendar/calendar-it.js"></script>
<script type="text/javascript" src="calendar/calendario.js"></script>-->
<link rel="stylesheet" href="https://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" />





</head>



<body class='theme-<%=session("stile")%>'>
	<div id="navigation">



	</div>




	<div class="container-fluid" id="content">

      <!-- #include file = "include/menu_left.asp" -->

	  <div id="main">
	  <div class="container-fluid">
				<div class="page-header">
					<div class="pull-left">
						<h1> <i class="icon-comments"></i> Decompressione archivio</h1>

					</div>
					<div class="pull-right">
                     <!-- se mi interessa devo includere
                         include pull_right.asp-->
                    </div>
				</div>








				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">
				      <div class="box-title">
				        <h3> <i class="icon-reorder"></i>  Esito decompressione
			          </div>
				      <div class="box-content">
				<div class="row-fluid">
				  <div class="span12">
				    <div class="box">



		    <div class="box-content">

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
//$homesito=$_REQUEST['homesito'];
//$homeserver=$_REQUEST['homeserver'];
$homesito="c:/inetpub/umanetroot";
$homeserver="expo2015Server/UECDL";

$indirizzo="../cSocial/ShowMessage.asp?scegli=".$scegli."&id_classe=".$id_classe."&cartella=".$Cartella."&bacheca=".$bacheca."&ID=".$ID."&RCount=".$RCount."&Zip=1";



//$nome_file= $homeserver."/".$homesito."/"."UECDL"."/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/".$_REQUEST['nome'];
$nome_file= $_REQUEST['nome']; // lo contieno già come variabile di sessione
$nome_fileSint= $_REQUEST['nomeSint'];
$nome_fileOrig= $_REQUEST['nomeOrig'];
$archive = new PclZip($nome_file);

$v_path=substr($nome_file,0,-(strlen($nome_fileSint)));
$v_path=$v_path.substr($nome_fileOrig,0,strlen($nome_fileOrig)-4);

//$v_path = $homeserver."/".$homesito."/"."UECDL"."/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/"; //devo settare il path di estrazione
//$v_path ="../../Materie/".$Materia."/".$Cartella."/".$Social."/".$IDPARENT."/".$CodiceAllievo."/";
/*echo "Nome=".$nome_file."\n";
echo "NomeSInt=".$nome_fileSint."\n";
echo "Path?=".$v_path."\n";

*/

$url=substr($nome_file,0,strlen($nome_file)-4);
//$list = $archive->extract(PCLZIP_OPT_PATH, $v_path,PCLZIP_OPT_EXTRACT_DIR_RESTRICTION, $v_path);
$list = $archive->extract(PCLZIP_OPT_PATH,  $v_path);

 //echo $path_parts['dirname']."\n".$path_parts['basename']."\n".$path_parts['extension'];

 //$path_parts = pathinfo($_ENV["SCRIPT_FILENAME"]);


if ($archive->extract() == 0) {
	die("<p>Si &egrave; verificato un errore: <br /> ".$archive->errorInfo(true)."</p><p><input type=\"button\" value=\"Riprova\" onClick=\"javascript:history.back(1)\">");
	 }
else { echo "<h3>L'archivio &egrave; stato aperto con successo.</h3></p>";

//unlink($nome_file); //rimuovi il file zip decompresso


//header("location: https://www.umanet.net");
?>

 <?php
 Echo "<a href=$indirizzo>"?>
   <input class="btn" type="button" name="ApplyMessage" id="ApplyMessage" value="Attendi ... stai per essere reindirizzato alla discussione">

  </a>
 <?php	 }
?>






                      </div>
			        </div>
			      </div>
			    </div>











                      </div>
			        </div>
			      </div>
			    </div>
			</div>


		</div> <!--fine main-->
        </div>

        <!-- #include file = "../include/colora_pagina.asp" -->



	</body>
     <script type="text/javascript">


$(window).load(function () {

	  $('#ApplyMessage').click();


	    event.stopPropagation();

	});

</script>


 </html>
