<?php
$filename = $_GET['file'];
$path="custom_logos/".$filename;
if(unlink($path)) {
  echo "Deleted ".$filename;
}
header('location: admin.php');
?>