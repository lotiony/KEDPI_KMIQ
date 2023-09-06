
//tab motion
$(document).on("click", ".tab li", function(e) { 	
    //tab button action
	$(this).parents('.tab').find('li').removeClass('on');
    $(this).addClass('on');
    
    //tab list active
    var idx = $(".tab li").index($(this));
    $('.tab_list').removeClass('on');
    $('.tab_list').eq(idx).addClass('on');
   
   
});

$(document).on("click", "#pagination li", function(e) { 	
    //tab button action
	$(this).parents('#pagination').find('li').removeClass('on');
    $(this).addClass('on');
    
    //tab list active
    var idx = $("#pagination li").index($(this));
    $('.tab_list').removeClass('on');
    $('.tab_list').eq(idx).addClass('on');
   
   
});


