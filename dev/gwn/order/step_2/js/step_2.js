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
var rev_prod = urlParams["rev_product"];
var rev = urlParams["rev"];
var print_numbers = urlParams["print_numbers"];
var number_placement = urlParams["number_placement"];
var print_names = urlParams["print_names"];
var team_name_design = urlParams["team_name"];
var logo = urlParams["logo"];
var json_source = urlParams["json"];

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

  //LIGHTBOX IMAGES BASED ON SPORT
  $("."+sport+"_stock").attr('data-lightbox', 'graphics');
  var graphicCount = $("[data-lightbox='graphics']").length - 1;
  var upperCaseSport = sport;
      upperCaseSport = sport.replace(/_/g, " ").replace('solid', '');
      upperCaseSport = upperCaseSport.toLowerCase().replace(/\b[a-z]/g, function(letter) {
          return letter.toUpperCase();
      });
  $("#stock_graphics").attr('href', "images/info/"+sport+".png").attr('title', graphicCount+" Stock "+upperCaseSport+" Graphics");

  //STEP ONE RESULTS DISPLAY ROWS
  //color row
  if((print_names == 'no') && (print_numbers == 'no') && (team_name_design == 'none') && (logo == 'no')) {
    $('#resulting_rows').hide();
    $('#no_results').show();
  };
  if((print_names == 'no') && (print_numbers == 'no') && (team_name_design == 'none') && (logo == 'yes')) {
    $('.non_logo_span').hide();
  };
  //rev row
  if(rev_prod == 'no'){
    $('#rev_prod_side_row').hide()
    $('#rev_row').hide();
  }
  else {
    if(rev == 'no') {
      $('#rev_row').hide();
      $('#rev_prod_side_row').show();
      if(color != "navy_gold") {
        var upperCaseFirstSideColor = color;
            upperCaseFirstSideColor = color.replace(/_/g, " ").replace('solid', '');
            upperCaseFirstSideColor = upperCaseFirstSideColor.toLowerCase().replace(/\b[a-z]/g, function(letter) {
                return letter.toUpperCase();
            });
        var colorSideOne = "<option value="+upperCaseFirstSideColor.toLowerCase()+">"+upperCaseFirstSideColor+" Side</option>";
        var colorSideWhite = "<option value='white'>White Side</option>";
        $('#side_select').append(colorSideOne);
        $('#side_select').append(colorSideWhite);
      }
      else{
        var colorSideOne = "<option value='navy'>Navy Side</option>";
        var colorSideTwo = "<option value='gold'>Gold Side</option>";
        $('#side_select').append(colorSideOne);
        $('#side_select').append(colorSideTwo);
      };
    }
    else {
      $('#rev_prod_side_row').hide();
      $('#rev_row').show();
      if(color != "navy_gold") {
        var upperCaseFirstColor = color;
            upperCaseFirstColor = color.replace(/_/g, " ").replace('solid', '');
            upperCaseFirstColor = upperCaseFirstColor.toLowerCase().replace(/\b[a-z]/g, function(letter) {
                return letter.toUpperCase();
            });
        var colorOne = "<option value='default' selected>"+upperCaseFirstColor+" Side</option>";
        $('#color_1_select').prepend(colorOne);
        $('#color_1_select option:eq(1)').remove();
      }
      else{
        var colorOne = "<option value='default' selected>Navy Side</option>";
        $('#color_1_select').prepend(colorOne);
        $('#color_1_select option:eq(1)').remove();
        var colorTwo = "<option value='default' selected>Gold Side</option>";
        $('#color_2_select').find(':first-child').remove()
        $('#color_2_select').prepend(colorTwo);
      };
    };
  };
  //font row
  $('#both_spans').hide();
  if((print_names == 'no') && ((team_name_design == 'none') || (team_name_design == 'letters_graphic'))) {
    $('#font_row').hide();
  }
  if((print_names == 'yes') && (team_name_design == 'letters')) {
    $('#both_spans').show();
  };
  if(print_names == 'no') {
    $('#player_names_span').hide();
  };
  if((team_name_design == 'none') || (team_name_design == 'letters_graphic')) {
    $('#team_name_span').hide();
  };
  function disableLetterStyles() {
    $("#player_name_style_select [value='default']").prop('selected', true);
    $('#player_name_style_select option').prop('disabled', true);
    $('#player_name_style_select option:selected').prop('disabled', false);
    $("#player_name_style_select").change();
    $("#team_name_style_select [value='default']").prop('selected', true);
    $('#team_name_style_select option').prop('disabled', true);
    $('#team_name_style_select option:selected').prop('disabled', false);
    $("#team_name_style_select").change();
  };
  function enableLetterStyle() {
    $('#player_name_style_select option').prop('disabled', false);
    $('#team_name_style_select option').prop('disabled', false);
  };

  if($('#font_select').val() == 'default'){
    disableLetterStyles();
  }
  else {
    enableLetterStyle();
  };
  $('#font_select').on('change', function(){
    if($('#font_select').val() == 'default'){
      disableLetterStyles();
    }
    else {
      enableLetterStyle();
    };
  })
  //team name and placement rows
  if(team_name_design == 'none') {
    $('#team_name_row').hide();
    $('#placement_row').hide();
  };
  if((number_placement == "front") || (number_placement == "front_back")) {
    $('#placement_row').hide();
  };
  //graphic row
  if(team_name_design != 'letters_graphic') {
    $('#graphic_row').hide();
  }
  else {
    $("."+sport+"_stock").each(function(){
      var graphicId = $(this).attr('id');
      var graphicValue  = graphicId;
          graphicValue = graphicValue.replace(/_/g, " ");
          graphicValue = graphicValue.replace(/(^|\s)\S/g, function(match) {
        return match.toUpperCase();
        });
      var graphicOption = "<option value="+graphicId+">"+graphicValue+"</option>"
      $('#graphic_select').append(graphicOption)
    });
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
  if((team_name_design == 'none') || (team_name_design == 'letters_graphic')) {
    $('#team_lettering_row').hide();
  };
  // player names lettering row
  if((print_names == 'no')) {
    $('#print_names_row').hide();
  };

  //IMAGE
  //init color selection
  $(":radio[value="+color+"]").prop('checked', true);
  //init image selection
  imageDisplay();
  //error handeling
  $('.product_img').error(function(){
    $(this).attr('src', '../../images/products/no_preview.gif').parent().css('cursor', 'default');
    $(this).next().remove();
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
    var img_source_front = "../../images/products/"+img+color+".gif";
    var img_source_back = "../../images/products/back/"+img+color+".gif";
    $('#product_img_front').attr('src', img_source_front);
    $('#product_img_back').attr('src', img_source_back);
    var lightbox_img_front = "../../images/products/large/"+img+color+".gif";
    var lightbox_img_back = "../../images/products/back/large/"+img+color+".gif";
    $('#product_img_front').parent().attr('href', lightbox_img_front).attr('title', caption+' (Front)');
    $('#product_img_back').parent().attr('href', lightbox_img_back).attr('title', caption+' (Back)');
    $('#product_img_front_large').attr('src', lightbox_img_front);
    $('#product_img_back_large').attr('src', lightbox_img_back);
    //lightbox image for jersey side info icon
    var lightbox_img_reversable = "../../images/products/large/"+img+color+".gif";
    $('#jersey_side_lightbox').attr('href', lightbox_img_reversable);
  };

  //ELEMENT OVERLAYS
  //color row
  function printColor(color) {
    var numbersColor;
    var numbersColorOverlay;
    $('.number_element').remove();
    $('#numbersColorOverlay').remove();
    if(color != "default") {
      numbersColor = "<image src='images/elements/numbers/"+color+".png' class='product_img_element number_element'>";
      if(number_placement == 'front'){
        $('#front_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/numbers/large/"+color+".png' id='numbersColorOverlay' class='front_element'>";
      }
      else if(number_placement == 'back'){
        $('#back_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/numbers/large/"+color+".png' id='numbersColorOverlay' class='back_element'>";
      }
      else if(number_placement == 'front_back'){
        $('#front_elements').append(numbersColor);
        $('#back_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/numbers/large/"+color+".png' id='numbersColorOverlay' class='front_back_element'>";
      };
      if($('#placement_select').val() != "chest") {
        graphicColor(color);
      }
      else {
        placementColor(color);
      };
      playerLetteringColor(color);
      teamLetteringColor(color);
    }
    else {
      numbersColor = "<image src='images/elements/default.png' class='product_img_element number_element'>";
      $('#front_elements').append(numbersColor);
      $('#back_elements').append(numbersColor);
    };
    $(numbersColorOverlay).insertAfter(".lb-image");
  };
  $('#color_1_select').on('change', function() {
    var color = $('#color_1_select').val();
    if(rev_prod == 'no'){
      printColor(color);
    }
    else {
      $('#side_select').change();
    };
  });

  function otherSide(side) {
    var otherSide = "left"
    if(side == "left") {
      otherSide = "right"
    };
    return otherSide;
  };

  //reversable product printing on one side
  function printOneRevColor(color, side) {
    var numbersColor;
    var numbersColorOverlay;
    if(rev == "no") {
      $('#front_elements').empty();
      $('#back_elements').empty();
    }
    //$(".back_element").remove();
    $('#numbersColorOverlay').remove();
    if(color != "default") {
      numbersColor = "<image src='images/elements/reversable/numbers/"+side+"/"+color+".png' class='product_img_element number_element'>";
      if(number_placement == 'front'){
        $('#front_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersColorOverlay' class='front_element'>";
      }
      else if(number_placement == 'back'){
        $('#back_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersColorOverlay' class='back_element'>";
      }
      else if(number_placement == 'front_back'){
        $('#front_elements').append(numbersColor);
        $('#back_elements').append(numbersColor);
        numbersColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersColorOverlay' class='front_back_element'>";
      };
      if($('#placement_select').val() != "chest") {
        graphicColor(color);
      }
      else {
        placementColor(color);
      };
      playerLetteringRevOneColor(color, side)
      teamLetteringColor(color);
    }
    else {
      numbersColor = "<image src='images/elements/default.png' class='product_img_element number_element'>";
      $('.number_element').remove();
    };
    $(numbersColorOverlay).insertAfter(".lb-image");
  };
  $('#side_select').on('change', function() {
    var color = $('#color_1_select').val();
    if(rev == 'yes') {
      var side = "left";
      printOneRevColor(color, side);
    }
    else {
      if($(this).val() != "default") {
        if($('#side_select option:eq(1)').prop('selected') == true) {
          var side = "left";
        }
        else if($('#side_select option:eq(2)').prop('selected') == true){
          var side = "right"
        }
        printOneRevColor(color, side);
      };
    };
  });

  //reversable product printing on both side
  function printTwoRevColor(color, side) {
    var numbersColor;
    var numbersRevColorOverlay;
    $('#rev_number_element').remove();
    $('#numbersRevColorOverlay').remove();
    if(color != "default") {
      numbersColor = "<image src='images/elements/reversable/numbers/"+side+"/"+color+".png' class='product_img_element rev_number_element'>";
      if(number_placement == 'front'){
        $('#front_elements').append(numbersColor);
        numbersRevColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersRevColorOverlay' class='front_element'>";
      }
      else if(number_placement == 'back'){
        $('#back_elements').append(numbersColor);
        numbersRevColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersRevColorOverlay' class='back_element'>";
      }
      else if(number_placement == 'front_back'){
        $('#front_elements').append(numbersColor);
        $('#back_elements').append(numbersColor);
        numbersRevColorOverlay = "<image src='images/elements/reversable/numbers/"+side+"/large/"+color+".png' id='numbersRevColorOverlay' class='front_back_element'>";
      };
      if($('#placement_select').val() != "chest") {
        graphicColor(color);
      }
      else {
        placementColor(color);
      };
      playerLetteringRevOneColor(color, side)
      teamLetteringColor(color);
    }
    else {
      $('.rev_number_element').remove();
    };
    $(numbersRevColorOverlay).insertAfter(".lb-image");
  };
  $('#color_2_select').on('change', function() {
    var color = $('#color_2_select').val();
    var side = "right";
      printTwoRevColor(color, side);
  });

  //player lettering style row
  function playerLetteringColor(color) {
    var playerLetteringColor;
    var playerLetteringColorOverlay;
    var font = $('#font_select').val();
    var letteringStyle = $('#player_name_style_select').val();
    $('#player_name_element').remove();
    $('#playerLetteringColorOverlay').remove();
    if(letteringStyle != "default") {
      playerLetteringColor = "<image src='images/elements/player_lettering/"+font+"_"+letteringStyle+"_"+color+".png' id='player_name_element' class='product_img_element'>";
      playerLetteringColorOverlay = "<image src='images/elements/player_lettering/large/"+font+"_"+letteringStyle+"_"+color+".png' id='playerLetteringColorOverlay' class='back_element'>";
    }
    else {
      playerLetteringColor = "<image src='images/elements/default.png' id='player_name_element' class='product_img_element'>";
    };
    $('#back_elements').append(playerLetteringColor);
    $(playerLetteringColorOverlay).insertAfter(".lb-container");
  };
  $('#player_name_style_select').on('change', function() {
    if(rev_prod == "no") {
      var color = $('#color_1_select').val();
      playerLetteringColor(color);
    }
    else {
      if(rev == "yes") {
        var color1 = $('#color_1_select').val();
        var color2 = $('#color_2_select').val();
        playerLetteringRevTwoColor(color1, color2);
      }
      else {
        var color = $('#color_1_select').val();
        if($(this).val() != "default") {
          if($('#side_select option:eq(1)').prop('selected') == true) {
            var side = "left";
          }
          else if($('#side_select option:eq(2)').prop('selected') == true){
            var side = "right"
          }
          playerLetteringRevOneColor(color, side);
        };
      };
    };
  });

  //reversable product player lettering on one side
  function playerLetteringRevOneColor(color, side) {
    var oposite = otherSide(side);
    var playerLetteringColor;
    var playerLetteringColorOverlay;
    var font = $('#font_select').val();
    var letteringStyle = $('#player_name_style_select').val();
    $("#player_name_"+side+"_element").remove();
    $("#playerLettering"+side+"ColorOverlay").remove();
    if(rev != "yes") {
      $("#playerLettering"+oposite+"ColorOverlay").remove();
    }
    if(letteringStyle != "default") {
      playerLetteringColor = "<image src='images/elements/reversable/player_lettering/"+side+"/"+font+"_"+letteringStyle+"_"+color+".png' id='player_name_"+side+"_element' class='product_img_element'>";
      playerLetteringColorOverlay = "<image src='images/elements/reversable/player_lettering/"+side+"/large/"+font+"_"+letteringStyle+"_"+color+".png' id='playerLettering"+side+"ColorOverlay' class='back_element'>";
    }
    else {
      playerLetteringColor = "<image src='images/elements/default.png' id='player_name_element' class='product_img_element'>";
    };
    $('#back_elements').append(playerLetteringColor);
    $(playerLetteringColorOverlay).insertAfter(".lb-container");
  };

  //reversable product player lettering on both side
  function playerLetteringRevTwoColor(color1, color2) {
    var playerLetteringColorOne;
    var playerLetteringColorTwo;
    var playerLetteringColorOneOverlay;
    var playerLetteringColorTwoOverlay;
    var font = $('#font_select').val();
    var letteringStyle = $('#player_name_style_select').val();
    $('#player_name_left_element').remove();
    $('#player_name_right_element').remove();
    $('#playerLetteringleftColorOverlay').remove();
    $('#playerLetteringrightColorOverlay').remove();
    if(letteringStyle != "default") {
      playerLetteringColorOne = "<image src='images/elements/reversable/player_lettering/left/"+font+"_"+letteringStyle+"_"+color1+".png' id='player_name_left_element' class='product_img_element'>";
      playerLetteringColorTwo = "<image src='images/elements/reversable/player_lettering/right/"+font+"_"+letteringStyle+"_"+color2+".png' id='player_name_right_element' class='product_img_element'>";
      playerLetteringColorOneOverlay = "<image src='images/elements/reversable/player_lettering/left/large/"+font+"_"+letteringStyle+"_"+color1+".png' id='playerLetteringleftColorOverlay' class='back_element'>";
      playerLetteringColorTwoOverlay = "<image src='images/elements/reversable/player_lettering/right/large/"+font+"_"+letteringStyle+"_"+color2+".png' id='playerLetteringrightColorOverlay' class='back_element'>";
    }
    else {
      playerLetteringColorOne = "<image src='images/elements/default.png' id='player_name_element' class='product_img_element'>";
      playerLetteringColorTwo = "<image src='images/elements/default.png' id='player_name_element' class='product_img_element'>";
    };
    $('#back_elements').append(playerLetteringColorOne);
    $('#back_elements').append(playerLetteringColorTwo);
    $(playerLetteringColorOneOverlay).insertAfter(".lb-container");
    $(playerLetteringColorTwoOverlay).insertAfter(".lb-container");
  };
  // $('#player_name_style_select').on('change', function() {
  //   if(rev == "yes") {
  //     var color1 = $('#color_1_select').val();
  //     var color2 = $('#color_2_select').val();
  //     playerLetteringRevTwoColor(color1, color2);
  //   }
  //   else {
  //     var color = $('#color_1_select').val();
  //     if($(this).val() != "default") {
  //       if($('#side_select option:eq(1)').prop('selected') == true) {
  //         var side = "left";
  //       }
  //       else if($('#side_select option:eq(2)').prop('selected') == true){
  //         var side = "right"
  //       }
  //       playerLetteringRevOneColor(color, side)
  //     };
  //   }
  // });

  //team lettering style row
  function teamLetteringColor(color) {
    var teamLetteringColor;
    var teamLetteringColorOverlay;
    var font = $('#font_select').val();
    var letteringStyle = $('#team_name_style_select').val();
    $('#team_name_element').remove();
    $('#teamLetteringColorOverlay').remove();
    if(letteringStyle != "default") {
      teamLetteringColor = "<image src='images/elements/team_lettering/"+font+"_"+letteringStyle+"_"+color+".png' id='team_name_element' class='product_img_element'>";
      teamLetteringColorOverlay = "<image src='images/elements/team_lettering/large/"+font+"_"+letteringStyle+"_"+color+".png' id='teamLetteringColorOverlay' class='front_element'>";
    }
    else {
      teamLetteringColor = "<image src='images/elements/default.png' id='team_name_element' class='product_img_element'>";
    };
    $('#front_elements').append(teamLetteringColor);
    $(teamLetteringColorOverlay).insertAfter(".lb-container");
  };
  $('#team_name_style_select').on('change', function() {
    var color = $('#color_1_select').val();
    teamLetteringColor(color);
  });

  //graphic row
  function graphicColor(color) {
    var graphicColor;
    var graphicColorOverlay;
    var graphic = $('#graphic_select').val();
    $('#front_graphic_element').remove();
    $('#graphicColorOverlay').remove();
    if(graphic != "default") {
      graphicColor = "<image src='images/elements/graphics/"+graphic+"_"+color+".png' id='front_graphic_element' class='product_img_element'>";
      graphicColorOverlay = "<image src='images/elements/graphics/large/"+graphic+"_"+color+".png' id='graphicColorOverlay' class='front_element'>";
    }
    else {
      graphicColor = "<image src='images/elements/default.png' id='front_graphic_element' class='product_img_element'>";
    }
    $('#front_elements').append(graphicColor);
    $(graphicColorOverlay).insertAfter(".lb-container");
  };
  $('#graphic_select').on('change', function() {
    var color = $('#color_1_select').val();
    $("#placement_select [value='front']").prop('selected', true);
    graphicColor(color);
  });

  //position row
  function placementColor(color) {
    if(team_name_design == 'letters') {
      var teamLetteringColor;
      var teamLetteringColorOverlay;
      var font = $('#font_select').val();
      var letteringStyle = $('#team_name_style_select').val();
      $('#team_name_element').remove();
      $('#teamLetteringColorOverlay').remove();
      if(letteringStyle != "default") {
        teamLetteringColor = "<image src='images/elements/placement/team_lettering/"+font+"_"+letteringStyle+"_"+color+".png' id='team_name_element' class='product_img_element'>";
        teamLetteringColorOverlay = "<image src='images/elements/placement/team_lettering/large/"+font+"_"+letteringStyle+"_"+color+".png' id='teamLetteringColorOverlay' class='front_element'>";
      }
      else {
        teamLetteringColor = "<image src='images/elements/default.png' id='team_name_element' class='product_img_element'>";
      }
      $('#front_elements').append(teamLetteringColor);
      $(teamLetteringColorOverlay).insertAfter(".lb-container");
    }
    else {
      var graphicColor;
      var graphicColorOverlay;
      var graphic = $('#graphic_select').val();
      $('#front_graphic_element').remove();
      $('#graphicColorOverlay').remove();
      if(graphic != "default") {
        graphicColor = "<image src='images/elements/placement/graphics/"+graphic+"_"+color+".png' id='front_graphic_element' class='product_img_element'>";
        graphicColorOverlay = "<image src='images/elements/placement/graphics/large/"+graphic+"_"+color+".png' id='graphicColorOverlay' class='front_element'>";
      }
      else {
        graphicColor = "<image src='images/elements/default.png' id='front_graphic_element' class='product_img_element'>";
      }
      $('#front_elements').append(graphicColor);
      $(graphicColorOverlay).insertAfter(".lb-container");
    };
  };

  $('#placement_select').on('change', function() {
    if($(this).val() != "chest"){
      var color = $('#color_1_select').val();
      if(team_name_design == 'letters') {
        teamLetteringColor(color)
      }
      else {
        graphicColor(color);
      };
    }
    else {
      var color = $('#color_1_select').val();
      placementColor(color);
    };
  });

  //element error handeling
  $('select').change(function() {
    $('.product_img_element').error(function(){
      $(this).attr('src', 'images/elements/default.png');
    });
  });

  //TOGGLE INFO AND CHECKMARK ICON
  $('select').change(function() {
    if($(this).val() != "default") {
      $(this).prev().find('img').attr('src', 'images/check.png');
    }
    else {
      $(this).prev().find('img').attr('src', 'images/info.png');
    }
  });
  //team name input
  $('#team_name_input').on('click', function(){
    $(this).attr('placeholder','');
  });
  $('#team_name_input').on('blur', function() {
    if($(this).val() != ""){
      $('#team_name_info_icon').attr('src', 'images/check.png');
    }
    else {
      $(this).attr('placeholder','EAGLES');
      $('#team_name_info_icon').attr('src', 'images/info.png');
    };
  });
  $('#team_name_info_icon').on('click', function(){
    $(this).next().attr('placeholder', 'Up to X Characters');
  });

  //GO BACK
  $('.return_to_step_1').click(function(){
    $.cookie('returnJSON', json_source, { path: '/' });
    window.history.back(-1);
    return false;
  });

  //SAVE BUTTON
  $('.save_btn').on('click', function(e) {
    e.preventDefault();
  });

  //RESET BUTTON
  $('.reset_btn').on('click', function(e) {
    $('#front_elements').empty();
    $('#back_elements').empty();
    $('.number_element').remove();
    $('#numbersColorOverlay').remove();
    $('select').each(function() {
      var selectID = $(this).attr('id');
      var firstOption = $("#"+selectID+" option:first").val();
      $("#"+selectID+" option[value="+firstOption+"]").attr('selected', 'selected');
    });
    $('#team_name_input').val('').attr('placeholder','EAGLES');
    $('#team_name_info_icon').attr('src', 'images/info.png');
    $('select').change();
    e.preventDefault();
  });

  //CANCEL BUTTON
  $('.cancel_btn').on('click', function() {
    var href = "../../sports/"+sport+"/jerseys/"+sport+"_jerseys.html";
    window.location = href;
  });

  //FINALIZE BUTTON
  $('.finalize_btn').on('click', function(e) {
    var msg;
    var blank = "There are no design options avalible for this order - please click <span class='mock_btn'>BACK TO YOUR ORDER OPTIONS</span> to continue"
    var verify = "Please verify the jersey details you have entered are acurate - click <span class='mock_btn'>FINALIZE ORDER</span> to continue"
    var missing = "Not all design options have been selected - please review the section above for missing information"
    var infoIcon = $(".info_btn[src*='info']").filter(":visible").length;
    var blankIcon = $(".info_btn").filter(":visible").length
    if(infoIcon > 0) {
      msg = missing;
    }
    else if(blankIcon == 0) {
      msg = blank;
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

//TOGGLES THE LIGHTBOX OVERLAY ELEMENTS
$('#front').on('click', function() {
  $('.back_element').hide();
  $('.front_element').show();
  setTimeout(function(){
    $('.front_element').css('opacity', 1);
    $('.front_back_element').css('opacity', 1);
  },1000);
});
$('#back').on('click', function() {
  $('.front_element').hide();
  $('.back_element').show();
  setTimeout(function(){
    $('.back_element').css('opacity', 1);
    $('.front_back_element').css('opacity', 1);
  },1000);
});
//close outside the lb
$('#lightbox').on('click', function() {
    $('.front_element').css('opacity', 0);
    $('.back_element').css('opacity', 0);
    $('.front_back_element').css('opacity', 0);
});
//close on lb X
$('a.lb-close').on('click', function() {
    $('.front_element').css('opacity', 0);
    $('.back_element').css('opacity', 0);
    $('.front_back_element').css('opacity', 0);
});

});

//CAPTURE VALUES AND SUBMIT FORM TO STEP 3
function captureValues() {
//url vars
  //$('#step_2_url').val(url);
    var upperCaseSport = sport;
      upperCaseSport = upperCaseSport.toLowerCase().replace(/\b[a-z]/g, function(letter) {
      return letter.toUpperCase();
    });
  $('#step_2_sport').val(upperCaseSport);
  $('#step_2_name').val(name);
  $('#step_2_qty').val(qty);
  $('#step_2_price').val('$'+price);
  $('#step_2_product_sub_total').val('$'+(price*qty).toFixed(2));
    var upperCaseColor = color;
      upperCaseColor = upperCaseColor.replace(/_/g, " ").replace('solid', '');
      upperCaseColor = upperCaseColor.replace(/(^|\s)\S/g, function(match) {
      return match.toUpperCase();
    });
  $('#step_2_color').val(upperCaseColor);
  if(rev_prod == "no") {
    $('#step_2_rev_product').val("No");
    $('#step_2_rev_side').remove();
  }
  else {
    $('#step_2_rev_product').val("Yes");
  };
  if(print_numbers == "no") {
    $('#step_2_print_numbers').val("No");
    $('#step_2_number_placement').remove();
  }
  else {
    $('#step_2_print_numbers').val("Yes");
    if((number_placement == "front") || (number_placement == "front_back")) {
      $('#step_2_name_design_placement').remove();
    };
  };
    var upperCaseNumPlacement = number_placement;
      upperCaseNumPlacement = upperCaseNumPlacement.replace(/_/g, " & ");
      upperCaseNumPlacement = upperCaseNumPlacement.replace(/(^|\s)\S/g, function(match) {
      return match.toUpperCase();
    });
  $('#step_2_number_placement').val(upperCaseNumPlacement);
  if(print_names == "no") {
    $('#step_2_print_names').val("No");
    $('#step_2_player_name_font').remove();
    $('#step_2_player_name_style').remove();
  }
  else {
    $('#step_2_print_names').val("Yes");
  };
  if(team_name_design == "none") {
    $('#step_2_team_name_design').val("None");
    $('#step_2_team_name').remove();
    $('#step_2_name_design_option').remove();
    $('#step_2_name_lettering_style').remove();
    $('#step_2_name_design_placement').remove();
  }
  else {
    $('#step_2_logo').remove();
    var upperCaseTeamNameDesign = team_name_design;
      upperCaseTeamNameDesign = upperCaseTeamNameDesign.replace(/_/g, " & ");
      upperCaseTeamNameDesign = upperCaseTeamNameDesign.replace(/(^|\s)\S/g, function(match) {
      return match.toUpperCase();
    });
    $('#step_2_team_name_design').val(upperCaseTeamNameDesign);
    if(team_name_design == "letters") {
      $('#step_2_name_design_option').remove();
    }
    else{
      $('#step_2_name_lettering_style').remove();
    };
  };
    var upperCaseLogo = logo;
      upperCaseLogo = upperCaseLogo.toLowerCase().replace(/\b[a-z]/g, function(letter) {
      return letter.toUpperCase();
    });
  $('#step_2_logo').val(upperCaseLogo);
  //selection values
  var print_side;
    if(rev == "yes"){
      print_side = "Both";
      $('#step_2_print_color').remove();
      var sideOne = $("#color_1_select [value='default']").text();
      var sideOnePrintColor = $("#color_1_select option:selected").text();
      var sideTwo = $("#color_2_select [value='default']").text();
      var sideTwoPrintColor = $("#color_2_select option:selected").text();
    }
    else {
      print_side = $("#side_select option:selected").text();
      $('#step_2_side_color_one').remove();
      $('#step_2_side_color_two').remove();
    };
  $('#step_2_rev_side').val(print_side);
  $('#step_2_print_color').val($("#color_1_select option:selected").text());
  $('#step_2_side_color_one').val(sideOne+" / "+sideOnePrintColor);
  $('#step_2_side_color_two').val(sideTwo+" / "+sideTwoPrintColor);
  $('#step_2_player_name_font').val($("#font_select option:selected").text());
  $('#step_2_player_name_style').val($('#player_name_style_select option:selected').text());
  $('#step_2_team_name').val($("#team_name_input").val());
  $('#step_2_name_design_option').val($("#graphic_select option:selected").text());
  $('#step_2_name_lettering_style').val($("#team_name_style_select option:selected").text());
  $('#step_2_name_design_placement').val($("#placement_select option:selected").text());

         var data = JSON.parse(json_source);
         var options = {
                source: data,
            };

          var detailsTable = $("<br><table></table>");

          detailsTable.jsonTable({
              head : ['Jersey', 'Size', 'Price', 'Number', 'Name', 'Qty'],
              json : ['Jersey', 'Size', 'Price', 'Number', 'Name', 'Qty']
          });

          detailsTable.jsonTableUpdate(options);

          $("#json_table").append(detailsTable);

  $('#step_2_form').submit(function(){
    return false;
  });

  $("#form_results").show();
};