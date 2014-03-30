<!doctype html>

<html lang="en">
<head>
  <title>GWN: Custom Logo Panel</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="description" content="GameWearNow - Custom Logo Panel">
  <meta name="author" content="f-robots.com">
  <meta charset="utf-8">

  <link rel="shortcut icon" type="image/png" href="../../images/favicon.gif">
  <link rel="stylesheet" href="../../css/main.css">
  <link rel="stylesheet" href="css/theme.default.css">

  <script src="../../js/jquery-1.10.2.min.js"></script>
  <!-- // <script src="http://code.jquery.com/jquery-1.11.0.min.js"></script> -->
  <script src="js/jquery.tablesorter.js"></script>

  <style type="text/css">

  #heading {
    padding-top: 4%;
  }
  table {
    width: 100%;
    padding: 0 25px;
    font-family: Arial;
  }
  .tablesorter-default .tablesorter-header {
    background-position: center left;
  }
  #container {
    padding: 20px 0;
    margin: -20px auto 0;
  }
  font {
    font-weight: bold;
    font-size: 24px;
  }

  /*VIEW*/
  .view {
    -moz-box-shadow:inset 0px 1px 0px 0px #e6cafc;
    -webkit-box-shadow:inset 0px 1px 0px 0px #e6cafc;
    box-shadow:inset 0px 1px 0px 0px #e6cafc;
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #c579ff), color-stop(1, #a341ee) );
    background:-moz-linear-gradient( center top, #c579ff 5%, #a341ee 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#c579ff', endColorstr='#a341ee');
    background-color:#c579ff;
    -webkit-border-top-left-radius:20px;
    -moz-border-radius-topleft:20px;
    border-top-left-radius:20px;
    -webkit-border-top-right-radius:20px;
    -moz-border-radius-topright:20px;
    border-top-right-radius:20px;
    -webkit-border-bottom-right-radius:20px;
    -moz-border-radius-bottomright:20px;
    border-bottom-right-radius:20px;
    -webkit-border-bottom-left-radius:20px;
    -moz-border-radius-bottomleft:20px;
    border-bottom-left-radius:20px;
    text-indent:0;
    border:1px solid #a946f5;
    display:inline-block;
    color:#ffffff;
    font-family:Arial;
    font-size:15px;
    font-weight:bold;
    font-style:normal;
    height:40px;
    line-height:40px;
    width:86px;
    text-decoration:none;
    text-align:center;
    text-shadow:1px 1px 0px #8628ce;
  }
  .view:hover {
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #a341ee), color-stop(1, #c579ff) );
    background:-moz-linear-gradient( center top, #a341ee 5%, #c579ff 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#a341ee', endColorstr='#c579ff');
    background-color:#a341ee;
  }.view:active {
    position:relative;
    top:1px;
  }

  /*DOWNLOAD*/
  .download {
    -moz-box-shadow:inset 0px 1px 0px 0px #bbdaf7;
    -webkit-box-shadow:inset 0px 1px 0px 0px #bbdaf7;
    box-shadow:inset 0px 1px 0px 0px #bbdaf7;
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #79bbff), color-stop(1, #378de5) );
    background:-moz-linear-gradient( center top, #79bbff 5%, #378de5 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#79bbff', endColorstr='#378de5');
    background-color:#79bbff;
    -webkit-border-top-left-radius:20px;
    -moz-border-radius-topleft:20px;
    border-top-left-radius:20px;
    -webkit-border-top-right-radius:20px;
    -moz-border-radius-topright:20px;
    border-top-right-radius:20px;
    -webkit-border-bottom-right-radius:20px;
    -moz-border-radius-bottomright:20px;
    border-bottom-right-radius:20px;
    -webkit-border-bottom-left-radius:20px;
    -moz-border-radius-bottomleft:20px;
    border-bottom-left-radius:20px;
    text-indent:0;
    border:1px solid #84bbf3;
    display:inline-block;
    color:#ffffff;
    font-family:Arial;
    font-size:15px;
    font-weight:bold;
    font-style:normal;
    height:40px;
    line-height:40px;
    width:86px;
    text-decoration:none;
    text-align:center;
    text-shadow:1px 1px 0px #528ecc;
  }
  .download:hover {
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #378de5), color-stop(1, #79bbff) );
    background:-moz-linear-gradient( center top, #378de5 5%, #79bbff 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#378de5', endColorstr='#79bbff');
    background-color:#378de5;
  }.download:active {
    position:relative;
    top:1px;
  }

  /*DELETE*/
  .delete {
    -moz-box-shadow:inset 0px 1px 0px 0px #f29c93;
    -webkit-box-shadow:inset 0px 1px 0px 0px #f29c93;
    box-shadow:inset 0px 1px 0px 0px #f29c93;
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #fe1a00), color-stop(1, #ce0100) );
    background:-moz-linear-gradient( center top, #fe1a00 5%, #ce0100 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#fe1a00', endColorstr='#ce0100');
    background-color:#fe1a00;
    -webkit-border-top-left-radius:20px;
    -moz-border-radius-topleft:20px;
    border-top-left-radius:20px;
    -webkit-border-top-right-radius:20px;
    -moz-border-radius-topright:20px;
    border-top-right-radius:20px;
    -webkit-border-bottom-right-radius:20px;
    -moz-border-radius-bottomright:20px;
    border-bottom-right-radius:20px;
    -webkit-border-bottom-left-radius:20px;
    -moz-border-radius-bottomleft:20px;
    border-bottom-left-radius:20px;
    text-indent:0;
    border:1px solid #d83526;
    display:inline-block;
    color:#ffffff;
    font-family:Arial;
    font-size:15px;
    font-weight:bold;
    font-style:normal;
    height:40px;
    line-height:40px;
    width:86px;
    text-decoration:none;
    text-align:center;
    text-shadow:1px 1px 0px #b23e35;
  }
  .delete:hover {
    background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #ce0100), color-stop(1, #fe1a00) );
    background:-moz-linear-gradient( center top, #ce0100 5%, #fe1a00 100% );
    filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#ce0100', endColorstr='#fe1a00');
    background-color:#ce0100;
  }.delete:active {
    position:relative;
    top:1px;
  }
  /* These button were generated using CSSButtonGenerator.com */
  </style>

</head>
<body>

<div id="wrapper">

  <div id="header">
    <div id="gwn_logo">
      <a href="../../index.html" title="Home"><image src="../../images/gwn_logo.png" alt="GameWearNow Logo"></a>
    </div>
    <div id="heading">
      <span class="title_txt" id="title">CUSTOM LOGOS</span>
        <br>
      <span class="title_txt" id="sub_title">VIEW, DOWNLOAD OR DELETE LOGOS FROM THE SERVER</span>
    </div>
  </div>

  <div id="container">
    <table id="myTable" class="tablesorter" style="padding: 10px 25px 25px;">
      <thead>
        <tr>
          <th class="sorter-false"></th>
          <th class="sorter-false"></th>
          <th style="padding-left: 25px;">FILE NAME</th>
          <th class="sorter-false"></th>
        </tr>
      </thead>
      <tbody>
        <?
        $dir = opendir('custom_logos/');
        while ($read = readdir($dir)) {
          if ($read!='.' && $read!='..') {
            echo '<tr>';
            echo '<td><a href="custom_logos/'.$read.'" class="view" onClick=\'showPopup(this.href);return(false);\'>View</a></td>';
            echo '<td><a href="custom_logos/'.$read.'" class="download" download>Download</a></td>';
            echo '<td style="vertical-align: middle;"><font>'.$read.'</font></td>';
            echo '<td><a href="delete.php?file='.$read.'" class="delete" onclick="return confirm(\'Are you sure you want to delete this custom logo?\');">Delete</a></td>';
            echo '</tr>';
          }
        }
        closedir($dir);
        ?>
      </tbody>
    </table>
  </div>

<script type="text/javascript">

$(document).ready(function() {
  $("#myTable").tablesorter();
});

function showPopup(url) {
newwindow=window.open(url,'name','left=20,top=20,width=500,height=500,resizable');
if (window.focus) {newwindow.focus()}
}
</script>

</body>
</html>