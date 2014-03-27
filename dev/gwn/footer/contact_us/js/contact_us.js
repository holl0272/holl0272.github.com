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

if(window.innerWidth > 800 && window.innerHeight > 600) {
 $("#email_form").show();
};

function adjustStyle(width) {
  width = parseInt(width);
    if (width <= 630) {
      $("#size-stylesheet").attr("href", "css/contact_us_narrow.css");
    }
    else {
      $("#size-stylesheet").attr("href", "");
    };

    if (width < 850) {
      $("#email_form").hide();
      $("#email_form_narrow").show();
      if($('.collapsed select option').filter(':selected').text() == "Other Question") {
        $("#email_form_narrow select>option[value='Other Question']").prop('selected', true);
        $('.other_field_narrow').show();
      }
      else  {
        $('.other_field_narrow').hide();
      };
    }
    else {
      $("#email_form").show();
      $("#email_form_narrow").hide();
    };

    if(width <= 970) {
      $('#heading').css({'float':'left','margin-top':'-40px'});
      $('#email_form td').css('padding', '0 10px 5px 0');
      $('.collapse').hide();
      $('.collapsed').show();
      var collapsedOption = $('.collapsed select option').filter(':selected').text();
      $('#email_form_narrow select>option[value="' + collapsedOption + '"]').prop('selected', true);
        if($('.collapsed select option').filter(':selected').text() == "Other Question") {
          // $("#email_form_narrow>option[value='Other Question']").prop('selected', true);
          $('.other_field_collapsed').show();
          $('.other_field_narrow').show();
          $('#other_field').hide();
        }
        else  {
          $('.other_field_collapsed').hide();
          $('.other_field_narrow').hide();
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

$('#email_form_narrow select').change(function() {
  var collapsedOption = $('#email_form_narrow select option').filter(':selected').text();
  $('#email_form_narrow>option[value="' + collapsedOption + '"]').prop('selected', true);
  if(collapsedOption == "Other Question") {
    $('.other_field_narrow').show();
  }
  else {
    $('.other_field_narrow').hide();
  };
});

$('select').change(function() {
  var selectedOption = $(this).find('option').filter(':selected').text();
  if(selectedOption != "Other Question") {
    $('.other_field').val('');
    $(this).blur();
    $('.msg').filter(':visible').focus();
  }
  else{
    $(this).blur();
    $('.other_field').filter(':visible').focus();
  }
  $('select').each(function() {
    $(this).find('option').each(function() {
      $(this).removeAttr('selected');
    });
    $(this).find("option[value='"+selectedOption+"']").attr('selected', 'selected');
  });
});

$('.contact').on('blur', function() {
  var custName = $(this).val();
  $('.contact').val(custName);
});
$('.reply_email').on('blur', function() {
  var replyAddress = $(this).val();
  $('.reply_email').val(replyAddress);
});
$('.other_field').on('blur', function() {
  var otherField = $(this).val();
  $('.other_field').val(otherField);
});
$('.msg').on('blur', function() {
  var message = $(this).val();
  $('.msg').val(message);
});

function resetForm() {
  $('.contact').val("");
  $('.reply_email').val("");
  $('.other_field').val("");
  $('.msg').val("");
  $('#other_field').hide();
  $('.other_field_narrow').hide();
  $(".email_label").css('color', '#666666');
  $(".category_label").css('color', '#666666');
  $("select option[value='default']").attr('selected', 'selected');
};

function validateForm() {
  if(($('select').val() == "") && (($(".reply_email").filter(":visible").val().indexOf("@") < 1) && ($(".reply_email").filter(":visible").val().indexOf(".") < 1) && ($(".reply_email").filter(":visible").val().length < 7))) {
    $('.category_label').css('color', 'red');
    $("input[name='reply_email']").val('').focus();
    $(".email_label").css('color', 'red');
    $("input[name='reply_email']").attr('placeholder','Please Provide a Valid Email');
  }
  else if(($('select').val() != "") && (($(".reply_email").filter(":visible").val().indexOf("@") < 1) && ($(".reply_email").filter(":visible").val().indexOf(".") < 1) && ($(".reply_email").filter(":visible").val().length < 7))) {
    $(".category_label").css('color', '#666666');
    $("input[name='reply_email']").val('').focus();
    $(".email_label").css('color', 'red');
    $("input[name='reply_email']").attr('placeholder','Please Provide a Valid Email');
  }
  else if(($('select').val() == "") && (($(".reply_email").filter(":visible").val().indexOf("@") != -1) && ($(".reply_email").filter(":visible").val().indexOf(".") != -1) && ($(".reply_email").filter(":visible").val().length >= 7))) {
    $(".email_label").css('color', '#666666');
    $('.category_label').css('color', 'red');
  };
};

$('#submit_contact').on('click', function(){
  if((($(".reply_email").filter(":visible").val().indexOf("@") != -1) && ($(".reply_email").filter(":visible").val().indexOf(".") != -1) && ($(".reply_email").filter(":visible").val().length >= 7)) && ($('select').val() != "")){
    var contact = $('.contact').val();
    var reply_email = $('.reply_email').val();
    var subject = $('select').val();
    var msg = $('.msg').val();
    if(subject == "Other Question") {
      subject = $('.other_field').val();
    }
    var success = "Thank you for contacting us!  Your message has been sent and we will respond to you shortly.";
    var dataString = 'contact='+ contact + '&reply_email=' + reply_email + '&subject=' + subject + '&msg=' + msg;
    $.ajax({
      type: "POST",
      url: "send_email.php",
      data: dataString,
      success: function() {
        resetForm();
        $('#success_msg').html(success);
        $('.inner').slideToggle().delay(5000).slideToggle();
      }
    });
    return false;
  }
  else {
    validateForm();
  };
});

$('#submit_contact_narrow').on('click', function(){
  if((($(".reply_email").filter(":visible").val().indexOf("@") != -1) && ($(".reply_email").filter(":visible").val().indexOf(".") != -1) && ($(".reply_email").filter(":visible").val().length >= 7)) && ($('select').val() != "")){
    var contact = $('.contact').val();
    var reply_email = $('.reply_email').val();
    var subject = $('select').val();
    var msg = $('.msg').val();
    if(subject == "Other Question") {
      subject = $('.other_field').val();
    }
    var success_narrow = "Thank you for contacting us!<br><br>Your message has been sent<br>and we will respond to you shortly.<br><br>";
    var dataString = 'contact='+ contact + '&reply_email=' + reply_email + '&subject=' + subject + '&msg=' + msg;
    $.ajax({
      type: "POST",
      url: "send_email.php",
      data: dataString,
      success: function() {
        resetForm();
        $('#success_msg').html(success_narrow);
        $('.inner').slideToggle().delay(5000).slideToggle();
      }
    });
    return false;
  }
  else {
    validateForm();
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
