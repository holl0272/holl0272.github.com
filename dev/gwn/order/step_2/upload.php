<?php
//upload.php
$output_dir = "custom_logos/";

if(isset($_FILES["myfile"]))
{
    if ($_FILES["myfile"]["error"] > 0)
    {
      echo "ERROR: " . $_FILES["myfile"]["error"] . "<br>";
    }
    else
    {
      $name = $_FILES["myfile"]['name'];
      $actual_name = pathinfo($name,PATHINFO_FILENAME);
      $original_name = $actual_name;
      $extension = pathinfo($name, PATHINFO_EXTENSION);

      $i = 1;
      while(file_exists($output_dir.$actual_name.".".$extension))
      {
          $actual_name = (string)$original_name."(".$i.")";
          $name = $actual_name.".".$extension;
          $i++;
      }

      move_uploaded_file($_FILES["myfile"]["tmp_name"],$output_dir. $name);

      echo "SUCCESS! Uploaded File: ".$_FILES["myfile"]["name"];
    };
}

/* These are the variable that tell the subject of the email and where the email will be sent.*/

$emailSubject = 'Custom Logo';
$mailto = 'holl0272@gmail.com';

/* These will gather what the user has typed into the fieled. */

$emailField = $_POST['reply_email'];
// $logo = $_FILES["myfile"]["name"];

$headers = "From: $emailField\r\n"; // This takes the email and displays it as who this email is from.
$headers .= "Content-type: text/html\r\n"; // This tells the server to turn the coding into the text.

/* This takes the information and lines it up the way you want it to be sent in the email. */

$body = <<<EOD
<br><hr><br>
Custom Logo: <a href="http://dev.gamewearnow.com/order/step_2/custom_logos/$name">$name</a>
<br><br>
<a href="http://dev.gamewearnow.com/order/step_2/download.php?file=$name"><button>DOWNLOAD</button></a>
<br><br>
To access the GameWearNow Custom Logo Admin Panel <a href="http://dev.gamewearnow.com/order/step_2/admin.php">Click Here</a>
EOD;

$success = mail($mailto, $emailSubject, $body, $headers); // This tells the server what to send.

?>