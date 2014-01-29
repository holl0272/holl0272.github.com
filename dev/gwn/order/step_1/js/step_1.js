if(window.innerWidth <= 800 && window.innerHeight <= 600) {
   $("#init-stylesheet").attr("href", "../../css/narrow.css");
   $('#wrapper').hide();
 };

$(document).ready(function(){

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "../../css/android.css");
};

var isOpera = !!window.opera || navigator.userAgent.indexOf(' OPR/') >= 0;  // Opera 8.0+
var isFirefox = typeof InstallTrigger !== 'undefined';  // Firefox 1.0+
var isSafari = Object.prototype.toString.call(window.HTMLElement).indexOf('Constructor') > 0;   // At least Safari 3+
var isChrome = !!window.chrome && !isOpera;     // Chrome 1+
var isIE = /*@cc_on!@*/false || !!document.documentMode; // At least IE6
if(isOpera) {
  $("#browser-stylesheet").attr("href", "../../css/opera.css");
};
if(isFirefox) {
  $("#browser-stylesheet").attr("href", "../../css/firefox.css");
};
if(isSafari) {
  $("#browser-stylesheet").attr("href", "../../css/safari.css");
};
if(isChrome) {
  $("#browser-stylesheet").attr("href", "../../css/chrome.css");
};
if(isIE) {
  $("#browser-stylesheet").attr("href", "../../css/ie.css");
};

   $('#wrapper').show();

function adjustStyle(width) {
  width = parseInt(width);
    if (width < 508) {
      $("#size-stylesheet").attr("href", "../../css/narrow.css");
    }
    else {
      $("#size-stylesheet").attr("href", "");
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

//NAME
var name = urlParams["name"];
$('#urlParams_name').html(name);

//DESCRIPTIONS
var gameDazzle = "Game Dazzle<br>Reversible Jersey";
var dazzleMicro = "Dazzle-Micro<br>Mesh Jersey";
var reversibleJersey = "Reversible Jersey";
var tShirt = "T-Shirt";
var meshShorts = "Mesh Shorts";

$('.prod_description').hide();

if(name == gameDazzle) {
  $('#gameDazzle').show();
}
else if(name == dazzleMicro) {
  $('#dazzleMicro').show();
}
else if(name == reversibleJersey) {
  $('#reversibleJersey').show();
}
else if(name == tShirt) {
  $('#tShirt').show();
}
else if(name == meshShorts) {
  $('#meshShorts').show();
};

//COST
var cost = (urlParams["price"] / 100).toFixed(2);
var cost_IV = ((urlParams["price"] / 100)*.95).toFixed(2);
var cost_XII = ((urlParams["price"] / 100)*.90).toFixed(2);
var cost_XXXIV = ((urlParams["price"] / 100)*.85).toFixed(2);

if(urlParams["name"] == "T-Shirt") {
  cost_IV = (((urlParams["price"] / 100)*.95)+.01).toFixed(2);
  cost_XXXIV = (((urlParams["price"] / 100)*.85)+.01).toFixed(2);
};

$('#urlParams_price_1').html(cost);
$('#urlParams_price_6').html(cost_IV);
$('#urlParams_price_12').html(cost_XII);
$('#urlParams_price_36').html(cost_XXXIV);

//MISC SCRIPTS
$('.notApplicable').prop('disabled', true);

});