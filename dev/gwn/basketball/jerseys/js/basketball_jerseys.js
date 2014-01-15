$(document).ready(function(){

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/android.css");
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
    if (width < 480) {
      $("#size-stylesheet").attr("href", "css/narrow.css");
    }
      $("#size-stylesheet").attr("href", "css/wide.css");
};

$(function() {
    adjustStyle($(this).width());
    $(window).resize(function() {
        adjustStyle($(this).width());
    });
});

$('#footer').show();

$(".select_btn").hover(
  function() {
    $(this).parent().find('.price').css('margin-top','11px');
    $(this).parent().parent().next().css('padding-left','1px');
    $(this).parent().parent().find('img').css('width','140px').css('border','2px solid #e8d606');
  }, function() {
    $(this).parent().find('.price').css('margin-top','7px');
    $(this).parent().parent().next().css('padding-left','0px');
    $(this).parent().parent().find('img').css('width','145px').css('border','none');
  }
);

});
