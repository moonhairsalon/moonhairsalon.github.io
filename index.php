<?php
$request_uri = trim($_SERVER['REQUEST_URI'], '/');

switch ($request_uri) {
    case 'peace':
        include_once("Peace.html");
        break;
    default:
        include_once("index.html");
        break;
}
?>
