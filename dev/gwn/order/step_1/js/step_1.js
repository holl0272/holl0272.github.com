
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
var name = urlParams["name"];
var sport = urlParams["sport"];
var img = urlParams["img"];
var price = urlParams["price"];
var product_id = urlParams["id"];

$(document).ready(function(){

if(window.innerWidth < 508){
  $("#size-stylesheet").attr("href", "css/jersey_narrow.css");
  $('.mobile').show();
}

var device = navigator.userAgent.toLowerCase();
var isAndroid = device.indexOf("android") > -1;
if(isAndroid) {
  $("#device-stylesheet").attr("href", "css/step_1_android.css");
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
      $("#size-stylesheet").attr("href", "css/step_1_narrow.css");
      $('.mobile').show();
    }
    else {
      $("#size-stylesheet").attr("href", "");
      $('.mobile').hide();
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

//NAME
$('#urlParams_name').html(name);

//SPORT BOX
$('.sport_box').hide();
$('.sport_box_mobile').hide();
$("#"+sport+"_box").show();
$("#"+sport+"_box_mobile").show();

//DESCRIPTIONS
var classicJersey = "Classic Jersey";
var dazzleMicro = "Dazzle-Micro Mesh Jersey";
var fullButton = "Full-Button Mesh Jersey";
var gameDay = "Football Game Day Jersey";
var gameDazzle = "Game Dazzle Reversible Jersey";
var meshJersey = "Mesh Jersey";
var meshShorts = "Mesh Shorts";
var mwReversible = "Moisture Wicking Reversible Jersey";
var mwtShirt = "Moisture Wicking T-Shirt";
var reversibleJersey = "Reversible Jersey";
var three_quarter_sleeve = "3/4 Sleeve Jersey";
var twoButton = "Two Button Jersey";
var tShirt = "T-Shirt";

function shortsDisplay() {
  if(name == meshShorts) {
    $('.shorts_show').show();
    $('.shorts_hide').hide();
    //$("#image_container a > img").unwrap();
  };
};

$('.description').hide();
$('.color_option').hide();

if(name == classicJersey) {
  $('#classicJersey').show();
  $('#reversible').hide();
    $('#cardinal_solid_option').show();
    $('#gold_solid_option').show();
    $('#navy_solid_option').show();
    $('#oxford_solid_option').show();
    $('#scarlet_solid_option').show();
    $('#white_solid_option').show();
}
else if(name == dazzleMicro) {
  $('#dazzleMicro').show();
    $('#black_option').show();
    $('#columbia_blue_option').show();
    $('#maroon_option').show();
    $('#navy_option').show();
    $('#scarlet_option').show();
}
else if(name == fullButton) {
  $('#fullButton').show();
  $('#reversible').hide();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
    $('#scarlet_solid_option').show();
}
else if(name == gameDay) {
  $('#gameDay').show();
  $('#reversible').hide();
    $('#black_solid_option').show();
    $('#maroon_solid_option').show();
    $('#navy_solid_option').show();
    $('#purple_solid_option').show();
    $('#scarlet_solid_option').show();
}
else if(name == gameDazzle) {
  $('#gameDazzle').show();
    $('#black_option').show();
    $('#maroon_option').show();
    $('#navy_option').show();
    $('#scarlet_option').show();
}
else if(name == meshJersey) {
  $('#meshJersey').show();
  $('#reversible').hide();
    $('#black_solid_option').show();
    $('#gold_solid_option').show();
    $('#navy_solid_option').show();
    $('#purple_solid_option').show();
    $('#scarlet_solid_option').show();
    $('#white_solid_option').show();
}
else if(name == meshShorts) {
  $('#meshShorts').show();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
    $('#scarlet_solid_option').show();
  shortsDisplay();
}
else if(name == mwReversible) {
  $('#mwReversible').show();
    $('#black_option').show();
    $('#navy_option').show();
    $('#purple_option').show();
    $('#scarlet_option').show();
}
else if(name == mwtShirt) {
  $('#mwtShirt').show();
  $('#reversible').hide();
    $('#black_solid_option').show();
    $('#charcoal_solid_option').show();
    $('#optic_yellow_solid_option').show();
    $('#scarlet_solid_option').show();
    $('#white_solid_option').show();
}
else if(name == reversibleJersey) {
  $('#reversibleJersey').show();
    $('#black_option').show();
    $('#kelly_green_option').show();
    $('#maroon_option').show();
    $('#navy_option').show();
    $('#navy_gold_option').show();
    $('#purple_option').show();
    $('#scarlet_option').show();
}
else if(name == three_quarter_sleeve) {
  $('#three_quarter_sleeve').show();
  $('#reversible').hide();
    $('#black_option').show();
    $('#gold_option').show();
    $('#navy_option').show();
    $('#scarlet_option').show();
}
else if(name == twoButton) {
  $('#twoButton').show();
  $('#reversible').hide();
    $('#birch_solid_option').show();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
    $('#purple_solid_option').show();
    $('#scarlet_solid_option').show();
}
else if(name == tShirt) {
  $('#tShirt').show();
  $('#reversible').hide();
    $('#black_solid_option').show();
    $('#cardinal_solid_option').show();
    $('#dark_green_solid_option').show();
    $('#kelly_green_solid_option').show();
    $('#gold_solid_option').show();
    $('#navy_solid_option').show();
    $('#purple_solid_option').show();
    $('#scarlet_solid_option').show();
    $('#white_solid_option').show();
};
//set value of step_1_rev_product
if($('#reversible').css('display') == "none") {
  $('#step_1_rev_product').val('no')
}
else{
  $('#step_1_rev_product').val('yes')
  if(name == meshShorts) {
    $('#step_1_rev_product').val('no');
  };
};

//COST
var cost = (price / 100);
var cost_IV = ((price / 100)*.95);
var cost_XII = ((price / 100)*.90);
var cost_XXXIV = ((price / 100)*.85);

if(name == "T-Shirt") {
  cost_IV = (((price / 100)*.95)+.01);
  cost_XXXIV = (((price / 100)*.85)+.01);
};

$('#urlParams_price_1').html((cost).toFixed(2));
$('#urlParams_price_6').html((cost_IV).toFixed(2));
$('#urlParams_price_12').html((cost_XII).toFixed(2));
$('#urlParams_price_36').html((cost_XXXIV).toFixed(2));

$('#price_per_jersey').html((cost).toFixed(2));

//COLOR SWATCHES

//init color selection
if($(':radio').filter(':checked').length == 0) {
var firstRadioSelect = $('input[type=radio]').filter(":visible").first();
var firstColorSwatch = firstRadioSelect.val();
var enAttrPos_JerseyColor = firstRadioSelect.data('color');
    firstRadioSelect.prop('checked', true);
    //populates this first color input in the submit form
    $('#step_1_color').val(firstColorSwatch);
    $('#step_1_enAttrPos_JerseyColor').val(enAttrPos_JerseyColor);
};

//remove margin on first child
$('#color_select').filter(":visible").first().css('margin-left',0);
//chage image on radio select
$('input[type=radio]').change(function() {
  if (this.checked) {
    $('#color_select').find('input[type=radio]').not(this).prop('checked', false);
  }
  //populates the color input in the submit form
  var checkedColor = $('input[type=radio]').filter(":checked").val();
  var checkedColor_enAttrPos_JerseyColor = $('input[type=radio]').filter(":checked").data('color');
  $('#step_1_color').val(checkedColor);
  $('#step_1_enAttrPos_JerseyColor').val(checkedColor_enAttrPos_JerseyColor);
  imageDisplay();
});

$('.color_square').on('click', function() {
  var square_id = $(this).prop('id');
  $('#step_1_color').val(square_id);
  $("input[value="+square_id+"]").prop('checked', true);
  var square_id_enAttrPos_JerseyColor = $("input[value="+square_id+"]").data('color');
  $('#step_1_enAttrPos_JerseyColor').val(square_id_enAttrPos_JerseyColor);
  $('#color_select').find('input[type=radio]').not("input[value="+square_id+"]").prop('checked', false);
  imageDisplay();
});

//IMAGE
//init image selection
imageDisplay();
//error handeling
$('.product_img').error(function(){
  $(this).attr('src', '../../images/products/no_preview.gif').parent().css('cursor', 'default');
  var thisImg = $(this).attr('id');
  $("#image_container > a #"+thisImg).unwrap();
});
$('.lightbox_img').error(function() {
  $(this).parent().attr('href','../../images/products/large/no_preview.gif');
});
//captures the image color based on the radio selection
function imageColor() {
  var selectedColor = "";
  var selected = $("#color_select input[type='radio']:checked");
  if (selected.length > 0) {
    selectedColor = selected.val();
  }
  return selectedColor;
};

//concatenate the image source
function imageDisplay() {
  var color = imageColor();
  var captionColor  = color;
      captionColor = color.replace(/_/g, " ").replace('solid', '');
  var captionProduct = name.replace('<br>', '&nbsp;');
  var caption = (captionColor+" "+captionProduct).replace(/(^|\s)\S/g, function(match) {
    return match.toUpperCase();
    });
  var img_source = "../../images/products/"+img+color+".gif";
  var lightbox_img = "../../images/products/large/"+img+color+".gif";
  var lightbox_img_back = "../../images/products/back/large/"+img+color+".gif";
  $('#product_img_front').attr('src', img_source);
  $('#product_img_front_large').attr('src', lightbox_img);
  $('#product_img_back_large').attr('src', lightbox_img_back);
  $('#product_img_front').parent().attr('href', lightbox_img).attr('data-lightbox', img+color).attr('title', caption);
  $('#product_img_back').attr('href', lightbox_img_back).attr('data-lightbox', img+color).attr('title', caption+' (Back)');
}

//TOGGLE ANIMATED CALCULATION GRAPHIC
function calculating() {
  $('#next_step').hide();
  $('#save_clear_btns').hide();
  $('#calculated').hide();
  $('#calculating').show();
};
function calculated() {
  $('#calculating').hide();
  $('#calculated').show();
  nextStep();
};
function re_calculate() {
  var qty = parseInt($('#order_qty').val());
  calculateCost(qty);
};
function reversableCosts() {
  $('.reversable > span').each(function() {
    var multiplier = $(this).text() * 2;
    $(this).text(multiplier);
  });
};

/*
    .:| ORPER OPTIONS |:.
*/

//HOW MANY JERSEYS DO YOU WANT TO ORDER?
var typingTimer;
var doneTypingInterval = 2000;
//on keyup, start the countdown
$('#order_qty').keyup(function(){
  calculating();
  $('#sub_selections table tr').remove();
  typingTimer = setTimeout(doneTyping, doneTypingInterval);
});
//on keydown, clear the countdown
$('#order_qty').keydown(function(){
  clearTimeout(typingTimer);
});

$('#order_qty').on('click', function(){
  $(this).prop('placeholder', "");
});
//user is "finished typing"
function doneTyping() {
  var qty = parseInt($('#order_qty').val());
  calculateCost(qty);
  calculated();
  $('#order_qty').blur();
  if($('.resize_select').length == 1){
    $('.resize_select').removeClass('resize_select').addClass('size_options');
  }
  if(name != meshShorts) {
    buildRows(qty);
  }
  else {
    buildRows(1);
  };
  //populates the qty input in the submit form
  $('#step_1_print_qty').val(qty);
};

//re-run row build if returning from order_step_2
if($('#order_qty').val() != ""){
  doneTyping();
}
else {
  $.removeCookie('returnJSON', { path: '/' });
  $('#cookieTbl').remove();
}

function calculateCost(qty) {
  var price_per;
  if((qty >= 6) && (qty <= 11)) {
    price_per = cost_IV;
  }
  else if((qty >= 12) && (qty <= 35)) {
    price_per = cost_XII;
  }
  else if(qty >= 36) {
    price_per = cost_XXXIV;
  }
  else{
    price_per = cost;
  };

    if(reversibleOnly() == 2) {
      price_per += addNumbers() * 2;
      price_per += addNameOnBack() * 2;
      price_per += teamNameDesign() * 2;
      reversableCosts();
    }
    else {
      price_per += addNumbers();
      price_per += addNameOnBack();
      price_per += teamNameDesign();
    };

    if((customLogo() != 0) && (qty > 0)) {
      if(customLogo() == 35) {
        var per_jersey = customLogo()/qty;
        price_per += per_jersey;
        $('#custom_logo_cost font').html(qty);
        $('#custom_logo_cost span').html((per_jersey).toFixed(2));
      };
      if(customLogo() == 24) {
        $('#custom_logo_cost').hide();
        $('#custom_logo_cost_waived').show();
      };
    };

    var xxl_size_mult = +($(".xxl_price").filter(":visible").text());
    var xxxl_size_mult = +($(".xxxl_price").filter(":visible").text());
    var xxl_jersey_price = (price_per + xxl_size_mult).toFixed(2);
    var xxxl_jersey_price = (price_per + xxxl_size_mult).toFixed(2);

  $('#price_per_jersey').html((price_per).toFixed(2));
  $('#jersey_price').val((price_per).toFixed(2));
  $('#xxl_jersey').val(xxl_jersey_price);
  $('#xxxl_jersey').val(xxxl_jersey_price);

  priceEachJersey();
};


function buildRows(qty) {
  var header = "<tr class='transparent'><td>#</td><td></td><td style='min-width:80px'>Size</td><td style='min-width:50px'>Price</td><td class='numbers_input'></td><td class='numbers_input'>Num</td><td class='names_input' style='min-width:115px'></td><td class='names_input' style='min-width:140px'>Name</td><td></td><td style='min-width:55px'></td><td class='hide'>Qty</td></tr>";
  var shortsheader = "<tr class='transparent'><td>#</td><td></td><td style='min-width:80px'>Size</td><td style='min-width:50px'>Price</td><td></td><td style='min-width:55px'>Qty</td></tr>";
  var row_number = "<td class='row_number'><font></font></td>";
  var row_number_mobile = "<td style='padding-right: 5px;' class='row_number_mobile'><font></font></td>";
    var sizeSelect = $('.description').filter(":visible").find('.size_select').parent().html();
    var resizeSelect = $('.description').filter(":visible").find('.size_select').removeClass('size_select').addClass('resize_select').parent().html();
  var product_size = "<td class='size_label'>Size</td><td class='jersey_size'>"+sizeSelect+"</td>";
  var product_size_mobile = "<td style='padding-right: 5px;' class='size_label'>Size</td><td style='text-align: center;' class='jersey_size'>"+sizeSelect+"</td>";
  var jersey_price = "<td class='jersey_price'></td>";
  var jersey_price_mobile = "<td style='padding-right: 25px;' class='jersey_price'></td>";
    var numberInput = "<input type='text' class='number_input' style='width: 25px;' maxlength='2'>";
    var newnumberInput = "<input type='text' class='newnumber_input' style='width: 25px;' maxlength='2'>";
  var product_number = "<td class='numbers_input number_label'>Number</td><td class='numbers_input number_reset'>"+numberInput+"</td>";
  var product_number_mobile = "<td colspan='2' class='numbers_input number_label'>Number</td><td colspan='2' style='padding-left: 10px;' class='numbers_input number_reset'>"+numberInput+"</td>";
    var nameInput = "<input type='text' class='name_input'>";
    var newnameInput = "<input type='text' class='newname_input'>";
  var name_on_jersey = "<td class='names_input name_label'>Name on Jersey</td><td class='names_input name_reset'>"+nameInput+"</td>";
  var name_on_jersey_mobile = "<td colspan='2' class='names_input name_label'>Name on Jersey</td><td colspan='2' style='padding-left: 10px;' class='names_input name_reset'>"+nameInput+"</td>";
  var product_qty = "<td class='qty_label'>Quantity</td><td><input type='hidden' class='row_qty' value='1'><font style='padding-right: 10px;'></font>";
  var product_qty_mobile = "<td style='min-width:139px'>QTY: <input type='hidden' class='row_qty' value='1'><font style='padding-right: 10px;'></font>";
  var qty_btns = "<span class='btns'><span class='plus_one' style='font-weight: bold; padding: 0 5px; cursor: pointer;'> + </span><span class='less_one' style='font-weight: bold; padding-left:5px; cursor: pointer;'> - </span></td><span>";
  var qty_btns_mobile = "<span class='btns'><span class='mobile_plus_one' style='font-weight: bold; padding: 0 5px; cursor: pointer;'> + </span><span class='mobile_less_one' style='font-weight: bold; padding-left:5px; cursor: pointer;'> - </span></td><span>";
  var raw_qty = "<td class='hide'></td>";

  if(name != meshShorts) {
    $('#sub_selections table').append(header);
      var jersey_row_desktop = "<tr class='desktop_row'>"+row_number+product_size+jersey_price+product_number+name_on_jersey+product_qty+qty_btns+raw_qty+"</tr>";
      var jersey_row_mobile = "<tr class='mobile mobile_row'>"+row_number_mobile+product_size_mobile+jersey_price_mobile+product_qty_mobile+qty_btns_mobile+raw_qty+"</tr><tr class='mobile mobile_row_number' style='padding-top: 5px;'><td style='padding-right: 20px;'>&nbsp;</td>"+product_number_mobile+"</tr><tr class='mobile mobile_row_name' style='padding-top: 5px;'><td style='padding-right: 20px;'>&nbsp;</td>"+name_on_jersey_mobile+"</tr>";
      var jersey_row = jersey_row_desktop + jersey_row_mobile;
  }
  else {
    $('#sub_selections table').append(shortsheader);
    var jersey_row = "<tr>"+row_number+product_size+jersey_price+product_qty+"</tr>";
  };
  //builds rows X qty input
  for (var i = 1; i <= qty; i++) {
    $('#sub_selections table').append(jersey_row);
  };
  //leading 0 for single digits
  function numberRows() {
    if(name != meshShorts) {
      $(".row_number").each(function(i) {
        var n = ++i;
        var row_number = ("0" + n).slice(-2);
        $(this).find("font").text(row_number);
      });
      $(".row_number_mobile").each(function() {
        var row_number_mobile = $(this).parent().prev().find('td:first').text()
        $(this).find("font").text(row_number_mobile);
      });
    }
  };
  numberRows();

  $('.mobile_row select').on('change', function() {
    var mobile_size = $(this).find('option').filter(':selected').val()
    $(this).closest('tr').prev().find('option').removeAttr('selected').filter(function(index) { return $(this).text() === mobile_size; }).attr('selected', true).change();
  });

  // $('.mobile_row .jersey_size').on('click', function() {
  //   $(this).html(sizeSelect);
  // });

  function numberCells() {
    //hides "number" cells if add numbers if init selected is NO
    if($('#print_numbers_select').val() == "no") {
      $('.numbers_input').hide();
    }
    else {
      $('.numbers_input').show();
    };
  };
  numberCells();

  function nameCells() {
    //hides "name" cells if add numbers if init selected is NO
    if($('#print_name_on_back').val() == "no") {
      $('.names_input').hide();
    }
    else {
      $('.names_input').show();
    };
  };
  nameCells();

  $(".row_qty").each(function() {
    if(name == meshShorts) {
      $(this).val(parseInt($('#order_qty').val()));
    };
    var qty_txt = $(this).val();
    $(this).next("font").text(qty_txt).closest('td').next('td').html(qty_txt);
  });

  function togglePlusLess() {
    $('#sub_selections table tr td .btns').show();
    if($("#sub_selections table tr:gt(0):not('.mobile')").filter(":visible").length > 1) {
      $("#sub_selections table tr:not('.mobile')").filter(":visible").find('.row_qty').next().addClass('count');
      $("#sub_selections table tr:not('.mobile')").filter(":visible").last().find('.row_qty').next().removeClass('count');
      var sum = 0;
        $('.count').each(function() {
          sum += Number($(this).text());
        });
      var lastRowQty = qty - sum;
        if(lastRowQty <= 0){
          lastRowQty = 1
        };
      $("#sub_selections table tr:not('.mobile') td").filter(":visible").last().find('.row_qty').val(lastRowQty).next().text(lastRowQty);

      $("#sub_selections table tr:not('.mobile')").filter(":visible").last().find('.btns').hide();
    }
    else if($("#sub_selections table tr:gt(0):not('.mobile')").filter(":visible").length == 1) {
      $("#sub_selections table tr:not('.mobile') td").filter(":visible").find('.row_qty').val(qty).next().text(qty);
    };
  };
  if(name != meshShorts) {
    togglePlusLess();
  };

  $('.plus_one').on('click', function() {
    var qty_plus = Number($(this).parent().prev().prev().val());
    var increase = qty_plus + 1;
    var rowTwoQty = Number($("#sub_selections table tr:not('.mobile')").filter(":visible").last().find('.row_qty').val());
    var decreaseRowTwoQty = rowTwoQty - 1;

    if(increase <= qty){
      $(this).parent().prev().text(increase);
      $(this).parent().prev().prev().val(increase).closest('td').next('td').html(increase);
      if(($("#sub_selections table tr:not('.mobile')").filter(":visible").length == 2) && (rowTwoQty > 1)) {
        $("#sub_selections table tr:not('.mobile')").filter(":visible").last().find('.row_qty').val(decreaseRowTwoQty).next().text(decreaseRowTwoQty);
      }
      else {
        $("#sub_selections table tr:not('.mobile')").filter(":visible").last().hide();
      };
      togglePlusLess();
    };
  });

  $('.less_one').on('click', function() {
    var qty_less = Number($(this).parent().prev().prev().val());
    var decrease = qty_less - 1;
    if(decrease > 0){
      $(this).parent().prev().text(decrease);
      $(this).parent().prev().prev().val(decrease).closest('td').next('td').html(decrease);
      $("#sub_selections table tr:not('.mobile')").filter(":hidden").first().show();
      togglePlusLess();
    };
  });

  function mobile_togglePlusLess() {
    $('#sub_selections table tr.mobile_row td .btns').show();
    if($("#sub_selections table tr.mobile_row").filter(":visible").length > 1) {
      $("#sub_selections table tr:not('.desktop_row')").filter(":visible").find('.row_qty').next().addClass('count');
      $("#sub_selections table tr:not('.desktop_row')").filter(":visible").last().find('.row_qty').next().removeClass('count');
      var sum = 0;
        $('.count').each(function() {
          sum += Number($(this).text());
        });
      var lastRowQty = qty - sum;
      if(lastRowQty <= 0){
        lastRowQty = 1
      };
      $("#sub_selections table tr:not('.desktop_row') td").filter(":visible").last().find('.row_qty').val(lastRowQty).next().text(lastRowQty);
      if(lastRowQty == 1){
        $("#sub_selections table tr.mobile_row").filter(":visible").last().find('.btns').hide();
      };
    }
    else if($("#sub_selections table tr.mobile_row").filter(":visible").length == 1) {
      $("#sub_selections table tr:not('.desktop_row') td").filter(":visible").find('.row_qty').val(qty).next().text(qty);
    };
  };
  if(name != meshShorts) {
    mobile_togglePlusLess();
  };

  $('.mobile_plus_one').on('click', function() {
    var qty_plus = Number($(this).parent().prev().prev().val());
    var increase = qty_plus + 1;
    var rowTwoQty = Number($("#sub_selections table tr.mobile_row").filter(":visible").last().find('.row_qty').val());
    var decreaseRowTwoQty = rowTwoQty - 1;

    if(increase <= qty){
      $(this).parent().prev().text(increase);
      $(this).parent().prev().prev().val(increase).closest('td').next('td').html(increase);
      if(($("#sub_selections table tr.mobile_row").filter(":visible").length == 2) && (rowTwoQty > 1)) {
        $("#sub_selections table tr.mobile_row").filter(":visible").last().find('.row_qty').val(decreaseRowTwoQty).next().text(decreaseRowTwoQty);
      }
      else {
        $("#sub_selections table tr.mobile_row").filter(":visible").last().prev().hide();
        $("#sub_selections table tr.mobile_row").filter(":visible").last().hide().next().hide().next().hide()
      };
      mobile_togglePlusLess();
    };
    // $("#sub_selections table tr.mobile_row").filter(":visible").last().find('.btns').hide();
  });

  $('.mobile_less_one').on('click', function() {
    var qty_less = Number($(this).parent().prev().prev().val());
    var decrease = qty_less - 1;
    if(decrease > 0){
      $(this).parent().prev().text(decrease);
      $(this).parent().prev().prev().val(decrease).closest('td').next('td').html(decrease);
      $("#sub_selections table tr.mobile_row").filter(":hidden").first().prev().show();
      $("#sub_selections table tr.mobile_row").filter(":hidden").first().show().next().show().next().show();
      mobile_togglePlusLess();
    };
  });

  $('.size_select').change(function() {
    var size = $(this).val();
    if(size == "XXL"){
      $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
    }
    else if(size == "XXXL"){
      $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
    }
    else {
      $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
    };
    $(this).closest('td').html("<a style='color:#cccdce;'><font class='set_size'>"+size+"</font></a>");
    $('.set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
      $(this).html(resizeSelect);
      $(this).closest('td').next('td').empty();
      $('.resize_select').change(function() {
        var resize = $(this).val();
        if(resize == "XXL"){
          $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
        }
        else if(resize == "XXXL"){
          $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
        }
        else {
          $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
        };
        $(this).closest('td').html("<a style='color:#cccdce;'><font class='set_size'>"+resize+"</font></a>");
        $('.set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')});
      });
    });
  });

  $('.mobile_row select').change(function() {
    alert('change1');
    var size = $(this).val();
    if(size == "XXL"){
      $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
    }
    else if(size == "XXXL"){
      $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
    }
    else {
      $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
    };
    $(this).closest('td').html("<a style='color:#cccdce;'><font class='set_size'>"+size+"</font></a>").closest('tr').prev().find('.jersey_size').html(size);
    // $('.set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
    $('.mobile_row .set_size').closest('td').on('click', function(e) {
      $(this).html(resizeSelect);
      alert('change2')
      e.preventDefault;
      $(this).closest('td').next('td').empty();
        $('.mobile_row .resize_select').change(function() {
            alert('change3');
        var resize = $(this).val();
        if(resize == "XXL"){
          $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
        }
        else if(resize == "XXXL"){
          $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
        }
        else {
          $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
        };
        $(this).closest('td').html("<a style='color:#cccdce;'><font class='set_size'>"+resize+"</font></a>").closest('tr').prev().find('.jersey_size').text(resize);
        // $('.set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')});
      });
    });
  });


  // $('.number_input').on('blur', function() {
  //   var number = $(this).val();
  //   $(this).closest('td').html("<font class='set_size'>"+number+"</font>").on('click', function() {
  //     $(this).html(newnumberInput);
  //     $('.newnumber_input').on('blur', function() {
  //       var newnumber = $(this).val();
  //       $(this).closest('td').html("<font class='set_size'>"+newnumber+"</font>");
  //     });
  //   });
  // });

var numberTimer;
var doneTypingNumber = 1000;
//on keyup, start the countdown
$('.number_input').keyup(function(){
  $(this).attr('id', 'temp');
  numberTimer = setTimeout(doneTyping, doneTypingNumber);
});
//on keydown, clear the countdown
$('.number_input').keydown(function(){
  clearTimeout(numberTimer);
});
//user is "finished typing"
function doneTyping() {
  var number = $('#temp').val();
  $('#temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+number+"</font></a>");
    $('.set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
    $(this).html(newnumberInput).find('input').focus();
    var numberTimer;
    var doneTypingNumber = 1000;
    //on keyup, start the countdown
    $('.newnumber_input').keyup(function(){
      $(this).attr('id', 'temp');
      numberTimer = setTimeout(doneTypingNum, doneTypingNumber);
    });
    //on keydown, clear the countdown
    $('.newnumber_input').keydown(function(){
      clearTimeout(numberTimer);
    });
    //user is "finished typing"
    function doneTypingNum() {
      var newnumber = $('#temp').val();
      $('#temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+newnumber+"</font></a>");
      $('.set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')});
    };
  });
};

var mobile_numberTimer;
var mobile_doneTypingNumber = 1000;
//on keyup, start the countdown
$('.mobile_row_number .number_input').keyup(function(){
  $(this).attr('id', 'mobile_temp');
  mobile_numberTimer = setTimeout(mobile_doneTyping, mobile_doneTypingNumber);
});
//on keydown, clear the countdown
$('.mobile_row_number .number_input').keydown(function(){
  clearTimeout(mobile_numberTimer);
});
//user is "finished typing"
function mobile_doneTyping() {
  var mobile_number = $('#mobile_temp').val();
  $('#mobile_temp').closest('tr').prev().prev().find('.number_reset').text(mobile_number);
  $('#mobile_temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+mobile_number+"</font></a>")
    $('.mobile_row_number .set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
    $(this).html(newnumberInput).find('input').focus();
    var mobile_numberTimer;
    var mobile_doneTypingNumber = 1000;
    //on keyup, start the countdown
    $('.newnumber_input').keyup(function(){
      $(this).attr('id', 'temp');
      mobile_numberTimer = setTimeout(mobile_doneTypingNum, mobile_doneTypingNumber);
    });
    //on keydown, clear the countdown
    $('.newnumber_input').keydown(function(){
      clearTimeout(mobile_numberTimer);
    });
    //user is "finished typing"
    function mobile_doneTypingNum() {
      var mobile_newnumber = $('#temp').val();
      $('#temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+mobile_newnumber+"</font></a>").closest('tr').prev().prev().find('.number_reset').text(mobile_newnumber);
      $('.mobile_row_number .set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')});
    };
  });
};

var nameTimer;
var doneTypingName = 2000;
//on keyup, start the countdown
$('.name_input').keyup(function(){
  $(this).attr('id', 'temp_name');
  nameTimer = setTimeout(doneTypingNam, doneTypingName);
});
//on keydown, clear the countdown
$('.name_input').keydown(function(){
  clearTimeout(nameTimer);
});
//user is "finished typing"
function doneTypingNam() {
  var name = $('#temp_name').val();
  $('#temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+name+"</font></a>");
    $('.set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
    $(this).html(newnameInput).find('input').focus();
    var nameTimer;
    var doneTypingName = 2000;
    //on keyup, start the countdown
    $('.newname_input').keyup(function(){
      $(this).attr('id', 'temp_name');
      nameTimer = setTimeout(doneReTypingName, doneTypingName);
    });
    //on keydown, clear the countdown
    $('.newname_input').keydown(function(){
      clearTimeout(nameTimer);
    });
    //user is "finished typing"
    function doneReTypingName() {
      var newname = $('#temp_name').val();
      $('#temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+newname+"</font></a>");
      $('.set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td')
    };
  });
};

var mobile_nameTimer;
var mobile_doneTypingName = 2000;
//on keyup, start the countdown
$('.mobile_row_name .name_input').keyup(function(){
  $(this).attr('id', 'mobile_temp_name');
  mobile_nameTimer = setTimeout(mobile_doneTypingNam, mobile_doneTypingName);
});
//on keydown, clear the countdown
$('.name_input').keydown(function(){
  clearTimeout(mobile_nameTimer);
});
//user is "finished typing"
function mobile_doneTypingNam() {
  var mobile_name = $('#mobile_temp_name').val();
  $('#mobile_temp_name').closest('tr').prev().prev().prev().find('.name_reset').text(mobile_name);
  $('#mobile_temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+mobile_name+"</font></a>");
    $('.mobile_row_name .set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
    $(this).html(newnameInput).find('input').focus();
    var mobile_nameTimer;
    var mobile_doneTypingName = 2000;
    //on keyup, start the countdown
    $('.mobile_row_name .newname_input').keyup(function(){
      $(this).attr('id', 'temp_name');
      mobile_nameTimer = setTimeout(mobile_doneReTypingName, mobile_doneTypingName);
    });
    //on keydown, clear the countdown
    $('.newname_input').keydown(function(){
      clearTimeout(nameTimer);
    });
    //user is "finished typing"
    function mobile_doneReTypingName() {
      var mobile_newname = $('#temp_name').val();
      $('#temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+mobile_newname+"</font></a>").closest('tr').prev().prev().prev().find('.name_reset').text(mobile_newname);
      $('.mobile_row_name .set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td')
    };
  });
};

};

function priceEachJersey(){
  $('#jersey_details tr .set_size').each(function(){
    if($(this).text() == "XXL"){
      $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
    }
    else if($(this).text() == "XXXL"){
      $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
    }
    else {
      $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
    };
  });
};

function returnPriceEachJersey(){
  $('.jersey_size font').each(function(){
   if($(this).text() == "XXL"){
      $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
    }
    else if($(this).text() == "XXXL"){
      $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
    }
    else {
      $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
    };
  });
};

//DO YOU WANT TO PRINT NUMBERS ON THE JERSEYS?
function addNumbers() {
  var numbers_one_side = 2;  //add $2
  var numbers_both_sides = 4; //add $4
  if($('#print_numbers_select').val() == "yes") {
    if(($('#numbers_front_back').val() == "front") || ($('#numbers_front_back').val() == "back"))  {
      $('#numbers_front_back_cost').show();
      $('#numbers_front_back_cost span').html(numbers_one_side);
      return numbers_one_side;
    }
    else if($('#numbers_front_back').val() == "front_back") {
      $('#numbers_front_back_cost span').html(numbers_both_sides);
      return numbers_both_sides;
    }
  }
  else {
    $('#numbers_front_back_cost').hide();
    return 0;
  }
};

$('#print_numbers_select').on('change', function() {
  if($(this).val() == "yes"){
    $('.numbers_input').show();
    $('#print_numbers_yes').show();
    // $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
    //$('#custom_logo option').prop('disabled', true);
    // populates print_number input and init number_placement in the submit form
    $('#step_1_print_numbers').val('yes');
    $('#step_1_number_placement').val('back');
    if($("#team_name_design :selected").val() == "letters_graphic") {
      $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
      $("#team_name_design option[value='letters']").prop("selected",true);
      $('#step_1_team_name').val('letters');
    }
    else {
      $("#team_name_design option[value='letters_graphic']").prop("disabled",false);
    };
    // if($("#custom_logo :selected").val() == "yes") {
    //   $("#custom_logo option[value='no']").prop("selected",true);
    //   $('#step_1_logo').val('no');
    // };
  }
  else {
    $('.numbers_input').hide();
    $('#print_numbers_yes').hide();
    $("#numbers_front_back option:eq(0)").prop('selected', true);
    $('#step_1_print_numbers').val('no');
    $('#step_1_number_placement').val('');
    $("#team_name_design option[value='letters_graphic']").prop("disabled",false);
    $(".number_reset").html("<input type='text' class='number_input' style='width: 25px;' maxlength='2'>");
  };
  re_calculate();
});

if($('#print_numbers_select').val() == "yes") {
  $('#print_numbers_yes').show();
  $('#numbers_front_back_cost').show();
};

//WANT TO PRINT NUMBERS ON FRONT AND BACK?
$('#numbers_front_back').on('change', function() {
  if($(this).val() == "front"){
    $('#step_1_number_placement').val('front');
    $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
    $('#custom_logo option').prop('disabled', true);
    if($("#team_name_design :selected").val() == "letters_graphic") {
      $("#team_name_design option[value='letters']").prop("selected",true);
      $('#step_1_team_name').val('letters');
    };
    if($("#custom_logo :selected").val() == "yes") {
      $("#custom_logo option[value='no']").prop("selected",true);
      $('#step_1_logo').val('no');
    };
    // $("#print_name_on_back option").prop('disabled',false);
  }
  else if($(this).val() == "back"){
    $('#step_1_number_placement').val('back');
    $("#team_name_design option[value='letters_graphic']").prop("disabled",false);
    $('#custom_logo option').prop('disabled', false);
    // $("#print_name_on_back option").prop('disabled',false);
    $('#custom_logo option').prop('disabled', false);
  }
  else if($(this).val() == "front_back"){
    // $("#print_name_on_back option[value='no']").prop("selected",true);
    //$('#step_1_print_names').val('no');
    // $('#print_name_on_back option').prop('disabled',true);
    $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
    $('#step_1_number_placement').val('front_back');
    $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
    $('#custom_logo option').prop('disabled', true);
    if($("#team_name_design :selected").val() == "letters_graphic") {
      $("#team_name_design option[value='letters']").prop("selected",true);
      $('#step_1_team_name').val('letters');
    };
    if($("#custom_logo :selected").val() == "yes") {
      $("#custom_logo option[value='no']").prop("selected",true);
      $('#step_1_logo').val('no');
    };
  };
  re_calculate();
});

//DO YOU WANT TO PRINT NAMES ON THE BACK OF JERSEYS?
function addNameOnBack() {
  var name_on_back = 4;  //add $4
  if($('#print_name_on_back').val() == "yes") {
    $('#print_name_on_back_cost').show();
    $('#print_name_on_back_cost span').html(name_on_back);
    return name_on_back;
  }
  else {
    $('#print_name_on_back_cost').hide();
    return 0;
  }
};
$('#print_name_on_back').on('change', function() {
  if($(this).val() == "yes"){
    $('.names_input').show();
    $('#step_1_print_names').val('yes');
  }
  else {
    $('.names_input').hide();
    $('#step_1_print_names').val('no');
    $(".name_reset").html("<input type='text' class='name_input' style='width: 150px;'>");
  };
  re_calculate();
});

//PRINT ON BOTH SIDES OF THE JERSEY? (Reversible jerseys only)
function reversibleOnly() {
  var reversible_only = 2;  //multiply add_ons 2X
  if($('#reversible_only').val() == "yes") {
    $('#reversible_only_note').hide();
    $('#reversible_only_cost').show();
    return reversible_only;
  }
  else {
    $('#reversible_only_cost').hide();
    $('#reversible_only_note').show();
    return 0;
  };
};

$('#reversible_only').on('change', function() {
  if($(this).val() == "yes"){
    $('#step_1_rev').val('yes');
  }
  else {
    $('#step_1_rev').val('no');
  };
  re_calculate();
});

//WHAT DO YOU WANT FOR YOUR TEAM NAME DESIGN?
function teamNameDesign() {
  var team_name_design = 4;  //add $4
  if($('#team_name_design').val() != "none") {
    $('#team_name_design_cost').show();
    $('#team_name_design_cost span').html(team_name_design);
    return team_name_design;
  }
  else {
    $('#team_name_design_cost').hide();
    return 0;
  }
};

$('#team_name_design').on('change', function() {
  var designOption = $(this).val();
  $('#step_1_team_name').val(designOption);
  if(designOption != "none" ) {
    $('#custom_logo').val('no');
    $('#step_1_logo').val('no');
    $('#custom_logo_line').hide();
  }
  else {
    $('#custom_logo_line').show();
  }
  re_calculate();
});

if($('#team_name_design').val() != "none") {
  $('#custom_logo_line').hide();
};

//DO YOU WANT TO SUPPLY YOUR OWN TEAM LOGO?
function customLogo() {
  var qty = Number($('#order_qty').val());
  var custom_logo = 24;
  if($('#custom_logo').val() == "yes") {
    $('#custom_logo_note').hide();
    $('#custom_logo_cost').show();
    if(qty <= 24) {
      custom_logo = 35;
      $('#custom_logo_cost_waived').hide();
    };
    return custom_logo;
  }
  else {
    $('#custom_logo_cost').hide();
    $('#custom_logo_cost_waived').hide();
    $('#custom_logo_note').show();
    return 0;
  }
};

$('#custom_logo').on('change', function() {
  var qty = Number($('#order_qty').val());
  if($(this).val() == "yes"){
    $('#step_1_logo').val('yes');
    $('#name_design').hide();
  }
  else {
    $('#name_design').show();
    $('#step_1_logo').val('no');
   // $('#step_1_print_names').val('no');
  };
  if($('#order_qty').val() == ""){
    $('#order_qty').val(1);
    doneTyping();
  }
  else {
    re_calculate();
  };
});

if($('#custom_logo').val() == "yes") {
  $('#name_design').hide();
};

//POPULATE TABLE WITH RETURN JSON DATA FROM STEP 2
if($.cookie('returnJSON')){
  //$('#jersey_details').hide();
  $('.next_btn').attr('id','continue');
  //disable select option on GO-BACK
  ////
  $('.side_select').removeClass('side_select');
  $('#jersey_details').empty();

  var data = JSON.parse($.cookie('returnJSON'));
  var options = {
    source: data,
  };

  var detailsTable = $("<table id='cookieTbl'></table>");

  detailsTable.jsonTable({
    json : ['#', 'Size', 'Price', 'Num', 'Name', 'Qty']
  });

  detailsTable.jsonTableUpdate(options);
  // $("#json_data").append(detailsTable);

function handle(table){
  var header = "<tr class='transparent' id='json_header'><td>#</td><td></td><td style='min-width:80px'>Size</td><td style='min-width:50px'>Price</td><td class='numbers_input'></td><td class='numbers_input'>Num</td><td class='names_input' style='min-width:115px'></td><td class='names_input' style='max-width:140px'>Name</td><td></td><td></td><td style='min-width:55px'>Qty</td><td class='hide'>Qty</td></tr>"
  $('#jersey_details').append(header);

  table.find('tr').each(function(){
    var jersey = $(this).find('td:eq(0)').text();
    var size = $(this).find('td:eq(1)').text();
    var price = $(this).find('td:eq(2)').text();
    var number = $(this).find('td:eq(3)').text();
    var numberInput = "<input type='text' class='number_input return_number' style='width: 25px;' maxlength='2'>";
    var nameInput = "<input type='text' class='name_input return_name' style='width: 150px;'>";
    var name = $(this).find('td:eq(4)').text();
    var qty = $(this).find('td:eq(5)').text();

    var row_number = "<td class='row_number'><font>"+jersey+"</font></td>";
    var product_size = "<td>Size</td><td class='jersey_size'><a class='float'><font class='return_size'>"+size+"</font></a></td>";
    var jersey_price = "<td class='jersey_price'>"+price+"</td>";
    if(number != "") {
      var product_number = "<td class='numbers_input'>Number</td><td class='numbers_input number_reset'><a class='float'><font class='return_number'>"+number+"</font></a></td>";
    }
    else {
      var product_number = "<td class='numbers_input'>Number</td><td class='numbers_input number_reset'>"+numberInput+"</td>";
    };
    if(name != "") {
     var name_on_jersey = "<td class='names_input'>Name On Jersey</td><td class='names_input name_reset'><a class='float'><font style='text-transform: uppercase;' class='return_name'>"+name+"<font></a></td><td></td>";
    }
    else {
     var name_on_jersey = "<td class='names_input'>Name On Jersey</td><td class='names_input name_reset'>"+nameInput+"</td>";
    };
    var product_qty = "<td>Quantity</td><td><input type='hidden' class='row_qty' value='1'><font style='padding-right: 10px;'>"+qty+"</font>";
    var qty_btns = "<span class='btns'><span class='plus_one' style='font-weight: bold; padding: 0 5px; cursor: pointer;'> + </span><span class='less_one' style='font-weight: bold; padding-left:5px; cursor: pointer;'> - </span></td><span>";
      if(name == meshShorts) {
        var jersey_row = "<tr>"+row_number+product_size+jersey_price+product_number+name_on_jersey+product_qty+qty_btns+"</tr>";
      }
      else {
        var jersey_row = "<tr>"+row_number+product_size+jersey_price+product_number+name_on_jersey+product_qty+"</tr>";
      };

    $('#jersey_details').append(jersey_row);
  });
  if($('#print_numbers_select').val() == "no") {
    $('.numbers_input').hide();
  }
  else {
    $('.numbers_input').show()
  };
  if($('#print_name_on_back').val() == "no") {
    $('.names_input').hide();
  }
  else {
    $('.names_input').show()
  };

  $('#json_header').next('tr').remove();
  //$('#json_data').remove();

  //edit size on return from step_2
  $('.return_size').on('click', function() {
    $('.description').filter(":visible").find('.resize_select').removeClass('resize_select');
    var sizeSelect = $('.description').filter(":visible").find('.size_select').parent().html();
    var resizeSelect = $('.description').filter(":visible").find('.size_select').removeClass('size_options').addClass('resize_select').parent().html();
    $(this).closest('td').html(sizeSelect);
    $('.size_select').change(function() {
      var returnSize = $(this).val();
      if(returnSize == "XXL"){
        $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
      }
      else if(returnSize == "XXXL"){
        $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
      }
      else {
        $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
      };
      $(this).closest('td').html("<a style='color:#cccdce;'><font class='return_set_size'>"+returnSize+"</font></a>");
      $('.return_set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
        $(this).html(resizeSelect);
        $(this).closest('td').next('td').empty();
        $('.resize_select').change(function() {
          var resize = $(this).val();
          if(resize == "XXL"){
            $(this).closest('td').next('td').html('$'+$('#xxl_jersey').val());
          }
          else if(resize == "XXXL"){
            $(this).closest('td').next('td').html('$'+$('#xxxl_jersey').val());
          }
          else {
            $(this).closest('td').next('td').html('$'+$('#jersey_price').val());
          };
          $(this).closest('td').html("<a style='color:#cccdce;'><font class='return_set_size'>"+resize+"</font></a>");
          $('.return_set_size').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')})
        });
      });
    });
  });

  //edit number on return from step_2
  $('.return_number').on('click', function() {
    var numberInput = "<input type='text' class='number_input' style='width: 25px;' maxlength='2'>";
    var newnumberInput = "<input type='text' class='newnumber_input' style='width: 25px;' maxlength='2'>";
    $(this).closest('td').html(numberInput).on('click', function() {
      $(this).find('input').focus();
      var numberTimer;
      var doneTypingNumber = 1000;
      //on keyup, start the countdown
      $('.number_input').keyup(function(){
        $(this).attr('id', 'temp');
        numberTimer = setTimeout(doneTyping, doneTypingNumber);
      });
      //on keydown, clear the countdown
      $('.number_input').keydown(function(){
        clearTimeout(numberTimer);
      });
      //user is "finished typing"
      function doneTyping() {
        var number = $('#temp').val();
        $('#temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+number+"</font></a>");
        $('.set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).closest('td').on('click', function() {
          $(this).html(newnumberInput).find('input').focus();
          var numberTimer;
          var doneTypingNumber = 1000;
          //on keyup, start the countdown
          $('.newnumber_input').keyup(function(){
            $(this).attr('id', 'temp');
            numberTimer = setTimeout(doneTypingNum, doneTypingNumber);
          });
          //on keydown, clear the countdown
          $('.newnumber_input').keydown(function(){
            clearTimeout(numberTimer);
          });
          //user is "finished typing"
          function doneTypingNum() {
            var newnumber = $('#temp').val();
            $('#temp').closest('td').html("<a style='color:#cccdce;'><font class='set_number'>"+newnumber+"</font></a>");
            $('.set_number').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')})
          };
        });
      };
    });
  });

  //edit name on return from step_2
  $('.return_name').on('click', function() {
    var nameInput = "<input type='text' class='name_input'>";
    var newnameInput = "<input type='text' class='newname_input'>";
    $(this).closest('td').html(nameInput).on('click', function() {
      $(this).find('input').focus();
      var nameTimer;
      var doneTypingName = 2000;
      //on keyup, start the countdown
      $('.name_input').keyup(function(){
        $(this).attr('id', 'temp_name');
        nameTimer = setTimeout(doneTypingNam, doneTypingName);
      });
      //on keydown, clear the countdown
      $('.name_input').keydown(function(){
        clearTimeout(nameTimer);
      });
      //user is "finished typing"
      function doneTypingNam() {
        var name = $('#temp_name').val();
        $('#temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+name+"</font></a>");
        $('.set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')}).on('click', function() {
          $(this).html(newnameInput).find('input').focus();
          var nameTimer;
          var doneTypingName = 2000;
          //on keyup, start the countdown
          $('.newname_input').keyup(function(){
            $(this).attr('id', 'temp_name');
            nameTimer = setTimeout(doneReTypingName, doneTypingName);
          });
          //on keydown, clear the countdown
          $('.newname_input').keydown(function(){
            clearTimeout(nameTimer);
          });
          //user is "finished typing"
          function doneReTypingName() {
            var newname = $('#temp_name').val();
            $('#temp_name').closest('td').html("<a style='color:#cccdce;'><font class='set_name'>"+newname+"</font></a>");
            $('.set_name').parent().mouseover(function(){$(this).removeAttr('style').addClass('float')})
          };
        });
      };
    });
  });

  function togglePlusLess() {
    var qty = Number($('#order_qty').val());
    $('#sub_selections table tr td .btns').show();
    if($('#sub_selections table tr').filter(":visible").length > 1) {
      $('#sub_selections table tr').filter(":visible").find('.row_qty').next().addClass('count');
      $('#sub_selections table tr').filter(":visible").last().find('.row_qty').next().removeClass('count');
      var sum = 0;
        $('.count').each(function() {
          sum += Number($(this).text());
        });
      var lastRowQty = qty - sum;
      $('#sub_selections table tr td').filter(":visible").last().find('.row_qty').val(lastRowQty).next().text(lastRowQty);
    }
    else if($('#sub_selections table tr').filter(":visible").length == 1) {
      $('#sub_selections table tr td').filter(":visible").find('.row_qty').val(qty).next().text(qty);
    };
  };
  if(name != meshShorts) {
    togglePlusLess();
  };

  $('.plus_one').on('click', function() {
    var qty = Number($('#order_qty').val());
    var qty_plus = Number($(this).parent().prev().prev().val());
    var increase = qty_plus + 1;
    var rowTwoQty = Number($('#sub_selections table tr').filter(":visible").last().find('.row_qty').val());
    var decreaseRowTwoQty = rowTwoQty - 1;

    if(increase <= qty){
      $(this).parent().prev().text(increase);
      $(this).parent().prev().prev().val(increase).closest('td').next('td').html(increase);
      if(($('#sub_selections table tr').filter(":visible").length == 2) && (rowTwoQty > 1)) {
        $('#sub_selections table tr').filter(":visible").last().find('.row_qty').val(decreaseRowTwoQty).next().text(decreaseRowTwoQty);
      }
      else {
        $('#sub_selections table tr').filter(":visible").last().hide();
      };
      togglePlusLess();
    };
  });

  $('.less_one').on('click', function() {
    var qty_less = Number($(this).parent().prev().prev().val());
    var decrease = qty_less - 1;
    if(decrease > 0){
      $(this).parent().prev().text(decrease);
      $(this).parent().prev().prev().val(decrease).closest('td').next('td').html(decrease);
      $('#sub_selections table tr').filter(":hidden").first().show();
      togglePlusLess();
    };
  });

  }

  handle(detailsTable)
  $('select').change(function(){
    returnPriceEachJersey();
  });
};

//NEXT STEP
function nextStep() {
  var qty = Number($('#order_qty').val());
  if((qty == 0) || (isNaN(qty))) {
    $('#next_step').hide();
    $('#save_clear_btns').hide();
  }
  else{
    $('#next_step').show();
    $('#save_clear_btns').show();
  }
};
nextStep();

//SAVE BUTTON
$('.save_btn').on('click', function(e) {
  e.preventDefault();
});

//RESET BUTTON
$('.reset_btn').on('click', function(e) {
  $('select').each(function() {
    var selectID = $(this).attr('id');
    var firstOption = $("#"+selectID+" option:first").val();
    $("#"+selectID+" option[value="+firstOption+"]").attr('selected', 'selected');
  });
  $('#order_qty').val('').attr('placeholder','0');
  nextStep();
  re_calculate();
  $('#print_numbers_select').change();
  $('#name_design').show();
  $('#custom_logo_line').show();
  $('#sub_selections table tr').remove();
  $('input[type=radio]').prop('checked', false).filter(":visible").first().prop('checked', true);
  imageDisplay();
  e.preventDefault();
});

//CANCEL BUTTON
$('.cancel_btn').on('click', function() {
  var href = "../../sports/jersey.html?sport="+sport;
  window.location = href;
});

//CLEAR ALL JERSEY ROW DATA
$('.clear_btn').on('click', function(e) {
  var qty = parseInt($('#order_qty').val());
  $('#sub_selections table tr').remove();
  buildRows(qty)
  e.preventDefault();
});

//NEXT STEP BUTTON
$('.next_btn').on('click', function(e) {
  var msg;
  if (name != meshShorts) {
   if($('.mobile').css('display') == 'none') {
    var verify = "Please verify the jersey details you have entered are acurate - click <span class='mock_btn'>NEXT STEP</span> to continue"
    var missing = "The jersey details are incomplete - please review the section above for missing information"
    }
    else {
    var verify = "Please verify the jersey details<br>you have entered are acurate<br>Click <span class='mock_btn'>NEXT STEP</span> to continue"
    var missing = "The jersey details are incomplete<br>Please review the section above<br>for missing information"
    };
  }
  else {
  var verify = "Please verify the size and quantity you have entered are acurate - click <span class='mock_btn'>NEXT STEP</span> to continue"
  var missing = "The quantity of shorts or their size selection is incomplete - please review the section above for missing information"
  }
  var emptyInputs = $('#jersey_details').find('input[type=text]:empty').filter(":visible").length;
  var emptySelects = $('#jersey_details').find('select').filter(":visible").length;
  if((emptyInputs > 0) || (emptySelects > 0)) {
    msg = missing;
  }
  else{
    msg = verify;
  };
  if($(this).attr('id') != 'continue') {
    $('#next_step_msg').html(msg);
    $('.inner').slideToggle().delay(5000).slideToggle();
    if(msg == verify) {
      $(this).attr('id','continue');
    };
    e.preventDefault();
  }
  else{
    captureValues();
  };
});

if($('#numbers_front_back').val() == "back") {
  $("#team_name_design option[value='letters_graphic']").prop("disabled",false);
}
else {
  $("#team_name_design option[value='letters_graphic']").prop("disabled",true);
};

});

//CAPTURE VALUES AND SUBMIT FORM TO STEP 2
function captureValues() {
  $('.btns').closest('td').find('span').remove();
  $('tr.mobile').hide();
  $.removeCookie('returnJSON', { path: '/' });
  var detailsToJSON = $('#jersey_details').tableToJSON();
  var data = JSON.stringify(detailsToJSON);
  $('<input type="hidden" name="json"/>').val(data).appendTo('#step_1_form');

  //url vars
  $('#step_1_url').val(window.location.href)
  $('#step_1_sport').val(sport);
  $('#step_1_name').val(name);
  $('#step_1_img').val(img);
  $('#step_1_price').val($('#price_per_jersey').text());
  $('#step_1_product_id').val(product_id);
  //SELECTION VALUES
    //quantity populates on doneTyping()
    //rev init no regardless of product but toggles on #reversible_only change
    //color checked populates on radio change
    //print_numbers YES/NO toggles on #print_numbers_select change
    //number_placement dictated by #numbers_front_back select
    //print_names YES/NO toggles on #print_name_on_back change
    //team_name options toggle on #team_name_design change#
    //logo option toggles on #custom_logo change
    if($('#step_1_logo').val() == "no") {
      $.removeCookie('reply', { path: '/' });
      $.removeCookie('logoName', { path: '/' });
      $.removeCookie('logoPath', { path: '/' });
    };
alert(data);
  // $('#step_1_form').submit();
};