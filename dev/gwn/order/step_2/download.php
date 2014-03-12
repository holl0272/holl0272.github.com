<?php

$file = $_GET['file'];

    //open/save dialog box
    header("Content-Disposition: attachment; filename='$file'");
    //content type
    header('Content-type: application/image');
    //read from server and write to buffer
    readfile('custom_logos/'.$file);


?>