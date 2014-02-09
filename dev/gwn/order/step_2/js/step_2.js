if(window.innerWidth <= 800 && window.innerHeight <= 600) {
   $("#init-stylesheet").attr("href", "../../css/narrow.css");
   $('#wrapper').hide();
 };

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
var url = urlParams["url"];
var name = urlParams["name"];
var sport = urlParams["sport"];
var img = urlParams["img"];
var price = urlParams["price"];
var color = urlParams["color"];
var qty = urlParams["qty"];
var rev = urlParams["rev"];
var print_numbers = urlParams["print_numbers"];
var number_placement = urlParams["number_placement"];
var print_names = urlParams["print_names"];
var team_name = urlParams["team_name"];
var logo = urlParams["logo"];

// var hashCount = 0
// function backOneStep() {
//   var index = -(hashCount + 1);
//   alert(index);
//   window.history.back(index);
// };

$(document).ready(function(){

  $('.form_rev').html(rev);
  $('.form_team_name').html(team_name);
  $('.form_logo').html(logo);
  $('.form_print_names').html(print_names);
  $('.form_print_numbers').html(print_numbers);

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
/*
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
*/
//NAME
$('#urlParams_name').html(name);

//SPORT BOX
$('.sport_box').hide();
$('.sport_box_mobile').hide();
$("#"+sport+"_box").show();
$("#"+sport+"_box_mobile").show();

//STEP ONE RESULTS DISPLAY ROWS
//color row
if((print_names == 'no') && (print_numbers == 'no') && (team_name == 'none') && (logo == 'no')) {
  $('#resulting_rows').hide();
  $('#no_results').show();
};
if((print_names == 'no') && (print_numbers == 'no') && (team_name == 'none') && (logo == 'yes')) {
  $('.non_logo_span').hide();
};
//rev row
if(rev == 'no') {
  $('#rev_row').hide();
}
else {
  if(color != "navy_gold") {
    var upperCaseFirstColor = color;
        upperCaseFirstColor = color.replace(/_/g, " ").replace('solid', '');
        upperCaseFirstColor = upperCaseFirstColor.toLowerCase().replace(/\b[a-z]/g, function(letter) {
            return letter.toUpperCase();
        });
    var colorOne = "<option value='colorOne' selected=''>"+upperCaseFirstColor+" Side</option>";
    $('#color_1_select').prepend(colorOne);
  }
  else{
    var colorOne = "<option value='colorOne' selected=''>Navy Side</option>";
    $('#color_1_select').prepend(colorOne);
    var colorTwo = "<option value='colorOne' selected=''>Gold Side</option>";
    $('#color_2_select').find(':first-child').remove()
    $('#color_2_select').prepend(colorTwo);
  };
};
//font row
$('#both_spans').hide();
if((print_names == 'no') && ((team_name == 'none') || (team_name == 'letters_graphic'))) {
  $('#font_row').hide();
}
if((print_names == 'yes') && (team_name == 'letters')) {
  $('#both_spans').show();
};
if(print_names == 'no') {
  $('#player_names_span').hide();
};
if((team_name == 'none') || (team_name == 'letters_graphic')) {
  $('#team_name_span').hide();
};
//team name and placement rows
if(team_name == 'none') {
  $('#team_name_row').hide();
  $('#placement_row').hide();
};
//graphic row
if(team_name != 'letters_graphic') {
  $('#graphic_row').hide();
};
if(logo == 'yes') {
  $('#placement_row').show();
  $('#graphic_row').hide();
  $('#team_lettering_row').hide();
}
else {
  $('.logo_row').hide();
}
//team name lettering row
if((team_name == 'none') || (team_name == 'letters_graphic')) {
  $('#team_lettering_row').hide();
};
// player names lettering row
if((print_names == 'no')) {
  $('#print_names_row').hide();
};

// if(((team_name != 'none') || (print_names != 'none')) && (logo == 'yes')) {
//   $('.logo_span').hide();
//   $('.non_logo_span').show();
//   $('.logo_row').show();
// };

//DESCRIPTIONS
var classicJersey = "Classic Jersey";
var dazzleMicro = "Dazzle-Micro<br>Mesh Jersey";
var fullButton = "Full-Button<br>Mesh Jersey";
var gameDazzle = "Game Dazzle<br>Reversible Jersey";
var meshJersey = "Mesh Jersey";
var meshShorts = "Mesh Shorts";
var reversibleJersey = "Reversible Jersey";
var twoButton = "Two Button<br>Jersey";
var tShirt = "T-Shirt";

function shortsDisplay() {
  if(name == meshShorts) {
    $('.shorts_show').show();
    $('.shorts_hide').hide();
  };
};

$('.description').hide();
$('.color_option').hide();

if(name == classicJersey) {
  $('#classicJersey').show();
  $('#rev_row').hide();
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
  $('#rev_row').hide();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
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
  $('#rev_row').hide();
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
else if(name == twoButton) {
  $('#twoButton').show();
  $('#rev_row').hide();
    $('#birch_solid_option').show();
    $('#black_solid_option').show();
    $('#navy_solid_option').show();
    $('#purple_solid_option').show();
    $('#scarlet_solid_option').show();
}
else if(name == tShirt) {
  $('#tShirt').show();
  $('#rev_row').hide();
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
$(":radio[value="+color+"]").prop('checked', true);

//remove margin on first child
$('#color_select').filter(":visible").first().css('margin-left',0);
//chage image on radio select
$('input[type=radio]').change(function() {
  if (this.checked) {
    $('#color_select').find('input[type=radio]').not(this).prop('checked', false);
  }
  //populates the color input in the submit form
  var checkedColor = $('input[type=radio]').filter(":checked").val();
  $('#step_1_color').val(checkedColor);
  imageDisplay();
});

$('.color_square').on('click', function() {
  var square_id = $(this).prop('id');
  $('#step_1_color').val(square_id);
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
  var color = imageColor();
  var img_source = "../../images/products/"+img+color+".gif";
  $('#product_img').attr('src', img_source);
}

//TOGGLE ANIMATED CALCULATION GRAPHIC
function calculating() {
  $('#next_step').hide();
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

$('select').change(function() {
  if($(this).val() != "default") {
    $(this).prev().find('img').attr('src', 'images/check.png');
  }
  else {
    $(this).prev().find('img').attr('src', 'images/info.png');
  }
});

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
  //populates the qty input in the submit form
  $('#step_1_print_qty').val(qty);
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
  var product_size = "<td>Size<select style='margin-left: 10px;''><option value='m' selected>M</option><option value='l'>L</option><option value='xl'>XL</option><option value='xxl'>XXL</option><option value='xxXl'>XXXL</option></select></td>";
  var product_number = "<td class='numbers_input'>Number</td><td class='numbers_input'><input type='text' class='input_num' style='width: 25px;'></td>";
  var name_on_jersey = "<td>Name On Jersey</td><td><input type='text' class='input_num' style='width: 150px;'></td>";
  var product_qty = "<td>Quantity</td><td><input type='hidden' class='row_qty' value='1'><font style='padding-right: 10px;'></font>";
    var qty_btns = "<span class='btns'><span class='plus_one' style='font-weight: bold; padding: 0 5px; cursor: pointer;'> + </span><span class='less_one' style='font-weight: bold; float:right; padding-left:5px; cursor: pointer;'> - </span></td><span>";

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
  numberRows();

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
    $('.numbers_input').show();
    $('#print_numbers_yes').show();
    //populates print_number input and init number_placement in the submit form
    $('#step_1_print_numbers').val('yes');
    $('#step_1_number_placement').val('front');
  }
  else {
    $('.numbers_input').hide();
    $('#print_numbers_yes').hide();
    $("#numbers_front_back option:eq(0)").prop('selected', true);
    $('#step_1_print_numbers').val('no');
    $('#step_1_number_placement').val('');
  };
  re_calculate();
});

//WANT TO PRINT NUMBERS ON FRONT AND BACK?
$('#numbers_front_back').on('change', function() {
  if($(this).val() == "front"){
    $('#step_1_number_placement').val('front');
  }
  else if($(this).val() == "back"){
    $('#step_1_number_placement').val('back');
  }
  else if($(this).val() == "front_back"){
    $('#step_1_number_placement').val('front_back');
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
    $('#step_1_print_names').val('yes');
  }
  else {
    $('#step_1_print_names').val('no');
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
  if($(this).val() == "yes"){
    $('#step_1_logo').val('yes');
  }
  else {
    $('#step_1_print_names').val('no');
  };
  if($('#order_qty').val() == ""){
    $('#order_qty').val(1);
    doneTyping();
  }
  else {
    re_calculate();
  };
});

//NEXT STEP
function nextStep() {
  var qty = Number($('#order_qty').val())
  if(qty == 0) {
    $('#next_step').hide();
  }
  else{
    $('#next_step').show();
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
  $('#team_name_input').val('').attr('placeholder','EAGLES');
  e.preventDefault();
});

$('#team_name_input').on('click', function(){
  $(this).attr('placeholder','');
});

$('#team_name_input').on('blur', function(){
  if($(this).val() == "") {
    $(this).attr('placeholder','EAGLES');
  };
});

//CANCEL BUTTON
$('.cancel_btn').on('click', function() {
  var href = "../../sports/"+sport+"/jerseys/"+sport+"_jerseys.html";
  window.location = href;
});

//FINALIZE BUTTON
$('.finalize_btn').on('click', function(e) {
  e.preventDefault();
});

});


//CAPTURE VALUES AND SUBMIT FORM TO STEP 2
function captureValues() {
//url vars
$('#step_1_sport').val(sport);
$('#step_1_name').val(name);
$('#step_1_img').val(img);
$('#step_1_price').val(price);
//SELECTION VALUES
  //quantity populates on doneTyping()
  //rev init no regardless of product but toggles on #reversible_only change
  //color checked populates on radio change
  //print_numbers YES/NO toggles on #print_numbers_select change
  //number_placement dictated by #numbers_front_back select
  //print_names YES/NO toggles on #print_name_on_back change
  //team_name options toggle on #team_name_design change#
  //logo option toggles on #custom_logo change
$('#step_1_form').submit();
};