WebFontConfig = {
  google: { families: [ 'Lato:100,400,900:latin', 'Josefin+Sans:100,400,700,400italic,700italic:latin' ] }
  };
  (function() {
    var wf = document.createElement('script');
    wf.src = ('https:' == document.location.protocol ? 'https' : 'http') +
      '://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js';
    wf.type = 'text/javascript';
    wf.async = 'true';
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(wf, s);
})();

var os;
if(navigator.appVersion.indexOf("Win") != -1) os = "Windows";
if(navigator.appVersion.indexOf("Mac") != -1) os = "Mac";

var isOpera = !!window.opera || navigator.userAgent.indexOf(' OPR/') >= 0;  // Opera 8.0+
var isFirefox = typeof InstallTrigger !== 'undefined';  // Firefox 1.0+
var isSafari = Object.prototype.toString.call(window.HTMLElement).indexOf('Constructor') > 0;   // At least Safari 3+
var isChrome = !!window.chrome && !isOpera;     // Chrome 1+
var isIE = /*@cc_on!@*/false || !!document.documentMode; // At least IE6
if(isOpera) {
  $("#browser-stylesheet").attr("href", "css/opera.css");
};
if(isFirefox) {
  $("#browser-stylesheet").attr("href", "css/firefox.css");
};
if(isSafari) {
  $("#browser-stylesheet").attr("href", "css/safari.css");
};
if(isChrome) {
  $("#browser-stylesheet").attr("href", "css/chrome.css");
};
if(isIE) {
  $("#browser-stylesheet").attr("href", "css/ie.css");
};

if((os == "Windows") && (isChrome)) {
  $("#os-stylesheet").attr("href", "css/windows.css");
};

//PARSE THE URL FOR VAR NAMES AND VALUES
var urlParams;
(window.onpopstate = function () {
    var match,
        pl     = /\+/g,  // Regex for replacing addition symbol with a space
        search = /([^&=]+)=?([^&]*)/g,
        decode = function (s) { return decodeURIComponent(s.replace(pl, " ")); },
        query  = window.location.search.substring(1);

    urlParams = {};
    while (match = search.exec(query))
       urlParams[decode(match[1])] = decode(match[2]);
})();

//URL VARS
var sport = urlParams["sport"];

var pageTitle = (sport).replace(/(^|\s)\S/g, function(match) {
  return match.toUpperCase();
  });
$('#title').html("GWN: "+pageTitle+" Jerserys");

$(document).ready(function(){

if(window.innerWidth < 508){
  $("#size-stylesheet").attr("href", "css/jersey_narrow.css");
  $("#mobile").hide();
  $("#desktop").show();
};

//NAME
$('#urlParams_name').html(name);

//SPORT BOX
$('.box_btn').hide();
$('.mobile_box_btn').hide();
$("#"+sport+"_box").show();
$("#home_box_btn").show();
$("#"+sport+"_box_mobile").show();
$("#home_mobile_box_btn").show();

//DESCRIPTIONS
var classicJerseyProd = $(".Classic_Jersey");
  $(".Classic_Jersey form input[name='sport']").val(sport);
var dazzleMicroProd = $(".Dazzle_Micro_Mesh_Jersey");
  $(".Dazzle_Micro_Mesh_Jersey form input[name='sport']").val(sport);
var fullButtonProd = $(".Full_Button_Mesh_Jersey");
  $(".Full_Button_Mesh_Jersey form input[name='sport']").val(sport);
var gameDayProd = $(".Football_Game_Day_Jersey");
  $(".Football_Game_Day_Jersey form input[name='sport']").val(sport);
var gameDazzleProd = $(".Game_Dazzle_Reversible_Jersey");
  $(".Game_Dazzle_Reversible_Jersey form input[name='sport']").val(sport);
var meshJerseyProd = $(".Mesh_Jersey");
  $(".Mesh_Jersey form input[name='sport']").val(sport);
var meshShortsProd = $(".Mesh_Shorts");
  $(".Mesh_Shorts form input[name='sport']").val(sport);
var mwReversibleProd = $(".Moisture_Wicking_Reversible_Jersey");
  $(".Moisture_Wicking_Reversible_Jersey form input[name='sport']").val(sport);
var mwtShirtProd = $(".Moisture_Wicking_T_Shirt");
  $(".Moisture_Wicking_T_Shirt form input[name='sport']").val(sport);
var reversibleJerseyProd = $(".Reversible_Jersey");
  $(".Reversible_Jersey form input[name='sport']").val(sport);
var three_quarter_sleeveProd = $(".3_4_Sleeve_Jersey");
  $(".3_4_Sleeve_Jersey form input[name='sport']").val(sport);
var twoButtonProd = $(".Two-Button_Jersey");
  $(".Two-Button_Jersey form input[name='sport']").val(sport);
var tShirtProd = $(".T_Shirt");
  $(".T_Shirt form input[name='sport']").val(sport);

$('.jersey').hide();

if(sport == "baseball") {
  fullButtonProd.show();
  meshShortsProd.show();
  mwtShirtProd.show();
  three_quarter_sleeveProd.show();
  twoButtonProd.show();
  tShirtProd.show();
}
else if(sport == "basketball") {
  dazzleMicroProd.show();
  gameDazzleProd.show();
  meshShortsProd.show();
  mwReversibleProd.show();
  mwtShirtProd.show();
  reversibleJerseyProd.show();
  tShirtProd.show();
}
else if(sport == "football") {
  classicJerseyProd.show();
  gameDayProd.show();
  meshJerseyProd.show();
  meshShortsProd.show();
  mwtShirtProd.show();
  tShirtProd.show();
}
else if(sport == "lacrosse") {
  meshJerseyProd.show();
  meshShortsProd.show();
  mwtShirtProd.show();
  tShirtProd.show();
}
else if(sport == "soccer") {
  meshShortsProd.show();
  mwtShirtProd.show();
  tShirtProd.show();
}
else if(sport == "softball"){
  fullButtonProd.show();
  meshShortsProd.show();
  mwtShirtProd.show();
  three_quarter_sleeveProd.show();
  twoButtonProd.show();
  tShirtProd.show();
}
else if(sport == "volleyball") {
  meshShortsProd.show();
  mwtShirtProd.show();
  tShirtProd.show();
};

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/android.css");
};

function adjustStyle(width) {
    if (width < 508) {
      $("#size-stylesheet").attr("href", "css/jersey_narrow.css");
      $("#mobile").show();
      $("#desktop").hide();
    }
    else {
      $("#size-stylesheet").attr("href", "");
      $("#mobile").hide();
      $("#desktop").show();
    };

    if(width <= 970) {
      $('#heading').css({'float':'left','margin-top':'-40px'});
    }
    else  {
      $('#heading').css({'float':'right','margin-top':'0'});
    };
};

$(function() {
    adjustStyle($(this).width());
    $(window).resize(function() {
        adjustStyle($(this).width());
    });
});

function randomColor() {
  var birch = "birch";
  var black = "black";
  var cardinal = "cardinal";
  var charcoal = "charcoal";
  var columbia_blue = "columbia_blue";
  var dark_green = "dark_green";
  var gold = "gold";
  var kelly_green = "kelly_green";
  var light_grey = "light_grey";
  var maroon = "maroon";
  var navy = "navy";
  var navy_gold = "navy_gold";
  var optic_yellow = "optic_yellow";
  var orange = "orange";
  var oxford = "oxford";
  var purple = "purple";
  var royal = "royal";
  var scarlet = "scarlet";
  var white = "white";

  var classic = [];
  classic.push(cardinal, gold, navy, oxford, scarlet, white);
  var classicColor = Math.floor(Math.random()*classic.length);
  var classic_img_source = "../images/products/classic_"+classic[classicColor]+"_solid.gif";
  $('.classic_image').attr('src', classic_img_source);

  var dazzle = [];
  dazzle.push(black, columbia_blue, maroon, navy, scarlet);
  var dazzleColor = Math.floor(Math.random()*dazzle.length);
  var dazzle_img_source = "../images/products/dazzle_"+dazzle[dazzleColor]+".gif";
  $('.dazzle_image').attr('src', dazzle_img_source);

  var fullbutton = [];
  fullbutton.push(black, navy, scarlet);
  var fullbuttonColor = Math.floor(Math.random()*fullbutton.length);
  var full_button_img_source = "../images/products/fullbutton_"+fullbutton[fullbuttonColor]+"_solid.gif";
  $('.full_button_image').attr('src', full_button_img_source);

  var gameday = [];
  gameday.push(black, maroon, navy, purple, scarlet);
  var gamedayColor = Math.floor(Math.random()*gameday.length);
  var game_day_img_source = "../images/products/gameday_"+gameday[gamedayColor]+"_solid.gif";
  $('.game_day_image').attr('src', game_day_img_source);

  var gamedazzle = [];
  gamedazzle.push(black, maroon, navy, scarlet);
  var gamedazzleColor = Math.floor(Math.random()*gamedazzle.length);
  var game_dazzle_img_source = "../images/products/gamedazzle_"+gamedazzle[gamedazzleColor]+".gif";
  $('.game_dazzle_image').attr('src', game_dazzle_img_source);

  var mesh = [];
  mesh.push(black, gold, navy, purple, scarlet, white);
  var meshColor = Math.floor(Math.random()*mesh.length);
  var mesh_img_source = "../images/products/mesh_"+mesh[meshColor]+"_solid.gif";
  $('.mesh_image').attr('src', mesh_img_source);

  var mwrev = [];
  mwrev.push(black, navy, purple, scarlet);
  var mwrevColor = Math.floor(Math.random()*mwrev.length);
  var mwrev_img_source = "../images/products/mw_rev_"+mwrev[mwrevColor]+".gif";
  $('.mw_rev_image').attr('src', mwrev_img_source);

  var mwt = [];
  mwt.push(black, charcoal, optic_yellow, scarlet, white);
  var mwtColor = Math.floor(Math.random()*mwt.length);
  var mwt_img_source = "../images/products/mwt_"+mwt[mwtColor]+"_solid.gif";
  $('.mwt_image').attr('src', mwt_img_source);

  var rev = [];
  rev.push(black, kelly_green, maroon, navy, navy_gold, purple, scarlet);
  var revColor = Math.floor(Math.random()*rev.length);
  var rev_img_source = "../images/products/rev_"+rev[revColor]+".gif";
  $('.rev_image').attr('src', rev_img_source);

  var threequarter = [];
  threequarter.push(black, gold, navy, scarlet);
  var threequarterColor = Math.floor(Math.random()*threequarter.length);
  var threequarter_img_source = "../images/products/three_quarter_sleeve_"+threequarter[threequarterColor]+".gif";
  $('.three_quarter_image').attr('src', threequarter_img_source);

  var twobutton = [];
  twobutton.push(birch, black, navy, purple, scarlet);
  var twobuttonColor = Math.floor(Math.random()*twobutton.length);
  var two_button_img_source = "../images/products/twobutton_"+twobutton[twobuttonColor]+"_solid.gif";
  $('.two_button_image').attr('src', two_button_img_source);

  var tshirt = [];
  tshirt.push(black, cardinal, dark_green, gold, kelly_green, navy, purple, scarlet);
  var tshirtColor = Math.floor(Math.random()*tshirt.length);
  var tshirt_img_source = "../images/products/tshirt_"+tshirt[tshirtColor]+"_solid.gif";
  $('.tshirt_image').attr('src', tshirt_img_source);

  var shorts = [];
  shorts.push(black, navy, scarlet);
  var shortsColor = Math.floor(Math.random()*shorts.length);
  var shorts_img_source = "../images/products/shorts_"+shorts[shortsColor]+"_solid.gif";
  $('.shorts_image').attr('src', shorts_img_source);
}

randomColor();

});
