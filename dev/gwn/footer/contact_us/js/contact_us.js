if(window.innerWidth <= 800 && window.innerHeight <= 600) {
 $("#email_form").hide();
 $("#init-stylesheet").attr("href", "css/contact_us_narrow.css");
 $('#wrapper').hide();
};

$(document).ready(function(){

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/contact_us_android.css");
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

   $('#wrapper').show();

if(window.innerWidth > 800 && window.innerHeight > 600) {
 $("#email_form").show();
};

function adjustStyle(width) {
  width = parseInt(width);
    if (width < 408) {
      $("#email_form").hide();
    }
    else {
      $("#email_form").show();
    };

    if (width <= 630) {
      $("#size-stylesheet").attr("href", "css/contact_us_narrow.css");
    }
    else {
      $("#size-stylesheet").attr("href", "");
    };

    if(width <= 970) {
      $('#heading').css({'float':'left','margin-top':'-40px'});
      $('#email_form td').css('padding', '0 10px 5px 0');
      $('.collapse').hide();
      $('.collapsed').show();
        if($('.collapsed select option').filter(':selected').text() == "Other Question") {
          $('.other_field_collapsed').show();
          $('#other_field').hide();
        }
        else  {
          $('.other_field_collapsed').hide();
        };
    }
    else  {
      $('#heading').css({'float':'right','margin-top':'0'});
      $('#email_form td').css('padding', '5px 10px 5px 0');
      $('.collapse').show();
      $('.collapsed').hide();
        if($('.collapse select option').filter(':selected').text() == "Other Question") {
          $('#other_field').show();
          $('.other_field_collapsed').hide();
        }
        else  {
          $('#other_field').hide();
        };
    };
};

$(function() {
    adjustStyle($(this).width());
    $(window).resize(function() {
        adjustStyle($(this).width());
    });
});

$('.collapse select').change(function() {
  var collapseOption = $('.collapse select option').filter(':selected').text();
  $('.collapsed select>option[value="' + collapseOption + '"]').prop('selected', true);
  if(collapseOption == "Other Question") {
    $('#other_field').show();
  }
  else {
    $('#other_field').hide();
  };
});

$('.collapsed select').change(function() {
  var collapsedOption = $('.collapsed select option').filter(':selected').text();
  $('.collapse select>option[value="' + collapsedOption + '"]').prop('selected', true);
  if(collapsedOption == "Other Question") {
    $('.other_field_collapsed').show();
  }
  else {
    $('.other_field_collapsed').hide();
  };
});

$(".not_selected").hover(
  function() {
    $('#current_page a').css('color','#cccdce');
  }, function() {
    $('#current_page a').css('color','#e8d606');
  }
);

});
