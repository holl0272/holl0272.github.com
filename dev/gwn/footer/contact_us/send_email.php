<?php
/* These are the variable that tell the subject of the email and where the email will be sent.*/

$emailSubject = $_POST['subject'];
$mailto = 'holl0272@gmail.com';

/* These will gather what the user has typed into the fieled. */

$emailReply = $_POST['reply_email'];
$emailName = $_POST['contact'];
$emailMsg = $_POST['msg'];


$headers = "From: $emailReply\r\n"; // This takes the email and displays it as who this email is from.
$headers .= "Content-type: text/html\r\n"; // This tells the server to turn the coding into the text.

/* This takes the information and lines it up the way you want it to be sent in the email. */

$body = <<<EOD
$emailName wrote:
<br><hr><br>
$emailMsg
EOD;

$success = mail($mailto, $emailSubject, $body, $headers); // This tells the server what to send.

?>