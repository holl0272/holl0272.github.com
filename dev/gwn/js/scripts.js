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

$(document).ready(function(){
  var device = navigator.userAgent.toLowerCase();
  var isAndroid = device.indexOf("android") > -1;
  if(isAndroid) {
    $("#device-stylesheet").attr("href", "css/android.css");
};

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

function adjustStyle(width) {
    if (width < 508) {
      $("#size-stylesheet").attr("href", "css/narrow.css");
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

$('.resize').each(function( index, element ) {

  $("<div id='hidden-resizer' style='font-size:22px;' />").hide().appendTo(document.body);

  var size;
  var desired_width = 140;
  var htmlSpan = $(this).html();
  var resizer = $("#hidden-resizer");
  resizer.html(htmlSpan);

  while(resizer.width() > desired_width) {
  size = parseInt(resizer.css("font-size"));
  resizer.css("font-size", size - 1);
  };

  $(this).css("font-size", size).html(resizer.html());

  $('#hidden-resizer').remove();
  });

});


