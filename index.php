<?php
$request_uri = trim($_SERVER['REQUEST_URI'], '/');

switch ($request_uri) {
    case 'peace':
        include_once("peace.html");
        break;
    case 'sea':
        include_once("sea.html");
        break;
    case 'namdinh':
        include_once("namdinh.html");
        break;
    case 'denial':
        include_once("denial.html");
        break;
    default:
        include_once("index.html");
        break;
}
?>
