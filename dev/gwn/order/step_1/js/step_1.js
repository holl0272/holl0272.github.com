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

function shortsDisplay() {
  if(name == meshShorts) {
    $('.shorts_show').show();
    $('.shorts_hide').hide();
  };
};

$('.description').hide();
$('.color_option').hide();

if(name == gameDazzle) {
  $('#gameDazzle').show();
    $('#black_option').show();
    $('#maroon_option').show();
    $('#navy_option').show();
    $('#scarlet_option').show();
}
else if(name == dazzleMicro) {
  $('#dazzleMicro').show();
    $('#black_option').show();
    $('#columbia_blue_option').show();
    $('#maroon_option').show();
    $('#navy_option').show();
    $('#scarlet_option').show();
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
}
else if(name == meshShorts) {
  $('#meshShorts').show();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
    $('#scarlet_solid_option').show();
  shortsDisplay();
};

//COST
var cost = (urlParams["price"] / 100);
var cost_IV = ((urlParams["price"] / 100)*.95);
var cost_XII = ((urlParams["price"] / 100)*.90);
var cost_XXXIV = ((urlParams["price"] / 100)*.85);

if(urlParams["name"] == "T-Shirt") {
  cost_IV = (((urlParams["price"] / 100)*.95)+.01);
  cost_XXXIV = (((urlParams["price"] / 100)*.85)+.01);
};

$('#urlParams_price_1').html((cost).toFixed(2));
$('#urlParams_price_6').html((cost_IV).toFixed(2));
$('#urlParams_price_12').html((cost_XII).toFixed(2));
$('#urlParams_price_36').html((cost_XXXIV).toFixed(2));

$('#price_per_jersey').html((cost).toFixed(2));

//COLOR SWATCHES

//init color selection
$('input[type=radio]').filter(":visible").first().prop('checked', true);
//remove margin on first child
$('#color_select').filter(":visible").first().css('margin-left',0);
//chage image on radio select
$('input[type=radio]').change(function() {
  if (this.checked) {
    $('#color_select').find('input[type=radio]').not(this).prop('checked', false);
  }
  imageDisplay();
});

$('.color_square').on('click', function() {
  var square_id = $(this).prop('id');
  $("input[value="+square_id+"]").prop('checked', true);
  $('#color_select').find('input[type=radio]').not("input[value="+square_id+"]").prop('checked', false);
  imageDisplay();
});

//IMAGE
//init image selection
imageDisplay();
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
  var img = urlParams["img"]
  var color = imageColor();
  var img_source = "../../images/products/"+img+color+".gif";
  $('#product_img').attr('src', img_source);
}

//TOGGLE ANIMATED CALCULATION GRAPHIC
function calculating() {
  $('#calculated').hide();
  $('#calculating').show();
};
function calculated() {
  $('#calculating').hide();
  $('#calculated').show();
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
var doneTypingInterval = 3000;
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
  buildRows(qty);
};

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

    if((customLogo() == 35) && (qty > 0)) {
      var per_jersey = customLogo()/qty;
      price_per += per_jersey;
      $('#custom_logo_cost font').html(qty);
      $('#custom_logo_cost span').html((per_jersey).toFixed(2));
    };

  $('#price_per_jersey').html((price_per).toFixed(2));
};

function buildRows(qty) {
  var row_number = "<td class='row_number'><font></font></td>";
  var product_size = "<td>Size<select><option value='m' selected>M</option><option value='l'>L</option><option value='xl'>XL</option><option value='xxl'>XXL</option><option value='xxXl'>XXXL</option></select></td>";
  var product_number = "<td>Number</td><td><input type='text' class='input_num'></td>";
  var name_on_jersey = "<td>Name On Jersey</td><td><input type='text' class='input_num'></td>";
  var product_qty = "<td>Quantity</td><td><input type='hidden' class='row_qty' value='1'><font></font>";
    var qty_btns = "<span class='btns'><button class='plus_one'> + </button><button class='less_one'> - </button></td><span>";

  if(name != meshShorts) {
    var jersey_row = "<tr>"+row_number+product_size+product_number+name_on_jersey+product_qty+qty_btns+"</tr>";
  }
  else {
    var jersey_row = "<tr>"+row_number+product_size+product_qty+qty_btns+"</tr>";
  };

  //builds rows X quty input
  for (var i = 1; i <= qty; i++) {
     $('#sub_selections table').append(jersey_row);
  };
  //leading 0 for single digits
  function numberRows() {
    $(".row_number").each(function(i) {
      var n = ++i;
      var row_number = ("0" + n).slice(-2);
      $(this).find("font").text(row_number);
    });
  };
  numberRows()

  $(".row_qty").each(function() {
    var qty_txt = $(this).val();
    $(this).next("font").text(qty_txt);
  });

  function togglePlusLess() {
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

  togglePlusLess();

  $('.plus_one').on('click', function() {
    var qty_plus = Number($(this).parent().prev().prev().val());
    var increase = qty_plus + 1;
    var rowTwoQty = Number($('#sub_selections table tr').filter(":visible").last().find('.row_qty').val());
    var decreaseRowTwoQty = rowTwoQty - 1;

    if(increase <= qty){
      $(this).parent().prev().text(increase);
      $(this).parent().prev().prev().val(increase);
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
      $(this).parent().prev().prev().val(decrease);
      $('#sub_selections table tr').filter(":hidden").first().show();
      togglePlusLess();
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
    $('#print_numbers_yes').show();
  }
  else {
    $('#print_numbers_yes').hide();
    $("#numbers_front_back option:eq(0)").prop('selected', true);
  };
  re_calculate();
});

//WANT TO PRINT NUMBERS ON FRONT AND BACK?
$('#numbers_front_back').on('change', function() {
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
  re_calculate();
});

//DO YOU WANT TO SUPPLY YOUR OWN TEAM LOGO?
function customLogo() {
  var custom_logo = 35;
  if($('#custom_logo').val() == "yes") {
    $('#custom_logo_note').hide();
    $('#custom_logo_cost').show();
    return custom_logo;
  }
  else {
    $('#custom_logo_cost').hide();
    $('#custom_logo_note').show();
    return 0;
  }
};

$('#custom_logo').on('change', function() {
  if($('#order_qty').val() == ""){
    $('#order_qty').val(1);
    doneTyping();
  }
  else {
    re_calculate();
  };
});

//MISC SCRIPTS
$('.notApplicable').prop('disabled', true);

});