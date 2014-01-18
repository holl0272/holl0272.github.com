$(document).ready(function(){

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/privacy_policy_android.css");
};

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

function adjustStyle(width) {
  width = parseInt(width);
    if (width < 508) {
      $("#size-stylesheet").attr("href", "css/privacy_policy_narrow.css");
    }
    else {
      $("#size-stylesheet").attr("href", "css/wide.css");
    };

    if(width <= 970) {
      $('#heading').css({'float':'left','margin-top':'-40px'});
    }
    else  {
      $('#heading').css({'float':'right','margin-top':'0'});
    };
}

$(function() {
    adjustStyle($(this).width());
    $(window).resize(function() {
        adjustStyle($(this).width());
    });
});

$('#footer').show();

$(".not_selected").hover(
  function() {
    $('#current_page a').css('color','#cccdce');
  }, function() {
    $('#current_page a').css('color','#e8d606');
  }
);

});
