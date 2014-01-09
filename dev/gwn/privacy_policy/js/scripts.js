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
if(isFirefox) {
  $("#browser-stylesheet").attr("href", "css/firefox.css");
};

function adjustStyle(width) {
  width = parseInt(width);
    if (width < 480) {
        $("#size-stylesheet").attr("href", "css/narrow.css");
    }
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
    $('#current_page a').css('color','white'));
  }, function() {
    $('#current_page a').css('color','yellow'));
  }
);


});
