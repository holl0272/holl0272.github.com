<?php
if ($_POST) {
$s = md5(rand());
$to = 'holl0272@gmail.com';
$subject = 'Custom Logo';
mail($to, $subject, "--$s

{$_POST['m']}
--$s
Content-Type: application/octet-stream; name=\"logo\"
Content-Transfer-Encoding: base64
Content-Disposition: attachment

".chunk_split(base64_encode(join(file($_FILES['logo']['tmp_name']))))."
--$s--", "From: logo@gamewearnow.com\r\nMIME-Version: 1.0\r\nContent-Type: multipart/mixed; boundary=\"$s\"");
exit;
}
?>

<form method="post" enctype="multipart/form-data" action="<?php echo $_SERVER['PHP_SELF'] ?>">
<input type="hidden" name="m" value="<?php echo "Attached File: ".$_FILES['logo']['name'] ?>"></input>
<input type="file" name="logo"/><br>
<input type="submit">
</form>