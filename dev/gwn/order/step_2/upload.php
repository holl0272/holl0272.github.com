<?php
//upload.php
$output_dir = "custom_logos/";

if(isset($_FILES["myfile"]))
{
    //Filter the file types , if you want.
    if ($_FILES["myfile"]["error"] > 0)
    {
      echo "Error: " . $_FILES["myfile"]["error"] . "<br>";
    }
    else
    {
        //move the uploaded file to uploads folder;
        move_uploaded_file($_FILES["myfile"]["tmp_name"],$output_dir. $_FILES["myfile"]["name"]);

     echo "SUCCESS! Uploaded File: ".$_FILES["myfile"]["name"];
     // echo "<a href='http://dev.gamewearnow.com/My Project/custom_logos/".$_FILES["myfile"]["name"]."' download='".$_FILES["myfile"]["name"]."'>".$_FILES["myfile"]["name"]."</a>";
    };
}
?>




<?php

/* These are the variable that tell the subject of the email and where the email will be sent.*/

$emailSubject = 'Custom Logo';
$mailto = 'holl0272@gmail.com';

/* These will gather what the user has typed into the fieled. */

$emailField = $_POST['reply_email'];
$logo = $_FILES["myfile"]["name"];

$headers = "From: $emailField\r\n"; // This takes the email and displays it as who this email is from.
$headers .= "Content-type: text/html\r\n"; // This tells the server to turn the coding into the text.

/* This takes the information and lines it up the way you want it to be sent in the email. */

$body = <<<EOD
<br><hr><br>
Custom Logo: <a href="http://dev.gamewearnow.com/My Project/custom_logos/$logo">$logo</a>
<br><br>
<a href="http://dev.gamewearnow.com/My Project/download.php?file=$logo"><button>DOWNLOAD</button></a>
EOD;

$success = mail($mailto, $emailSubject, $body, $headers); // This tells the server what to send.

?>