$(document).ready(function(){

var ua = navigator.userAgent.toLowerCase();
var isAndroid = ua.indexOf("android") > -1; //&& ua.indexOf("mobile");
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/android.css");
};

function adjustStyle(width) {
  width = parseInt(width);
    if (width < 824) {
        $("#size-stylesheet").attr("href", "css/narrow.css");
    } else {
       $("#size-stylesheet").attr("href", "css/wide.css");
    }
}

$(function() {
    adjustStyle($(this).width());
    $(window).resize(function() {
        adjustStyle($(this).width());
    });
});

$('#footer').show();

});
