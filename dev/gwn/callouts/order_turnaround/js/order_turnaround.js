if(window.innerWidth <= 800 && window.innerHeight <= 600) {
 $("#init-stylesheet").attr("href", "css/order_turnaround_narrow.css");
 $('#wrapper').hide();
};

$(document).ready(function(){

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/order_turnaround_android.css");
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
      $("#size-stylesheet").attr("href", "css/order_turnaround_narrow.css");
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

var shipDate = new Date();
var turnaround = 2;
shipDate.setDate(shipDate.getDate() + turnaround);

var weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
var day = weekdays[shipDate.getDay()];
  if(day == "Saturday") {
    shipDate.setDate(shipDate.getDate() + 2);
    day = weekdays[shipDate.getDay()];
  }
  else if(day == "Sunday") {
    shipDate.setDate(shipDate.getDate() + 1);
    day = weekdays[shipDate.getDay()];
  };
var date = shipDate.getDate();
var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
var month = months[shipDate.getMonth()];
var year = shipDate.getFullYear();

var displayDate = day +', ' + month + ' ' + date + ', '+ year;

$('#shippingDate').html(displayDate);

});

$(window).load(function() {
   $('#wrapper').show();
 });
