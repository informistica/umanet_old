<?php

$serverName ="SERVERWIN\SQLEXPRESS";
$usr="utente";
$pwd="123Maurosho";
$db="Copiaditestonline";

$connectionInfo = array("UID" => $usr, "PWD" => $pwd, "Database" => $db);
$conn = sqlsrv_connect($serverName, $connectionInfo) or die (print_r(sqlsrv_errors(), true));
 
?>