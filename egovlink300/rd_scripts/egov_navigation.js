jQuery(document).ready(function($) {
    // BEGIN: Set up dropdown navigation menu --------------------------------
    /* prepend menu icon */
    //$('#nav-wrap').prepend('<div id="menu-icon" onmouseover="expandNavigation();">Main Menu</div>');
    //$('#nav-wrap').prepend('<div id="menu-icon">Main Menu</div>');

    /* toggle nav */
    $('#menu-icon').click(function() {
        $('#navmenu').show();
    });
    $('#menu-icon').mouseover(function() {
        $('#navmenu').show();
    });

    $('#navmenu').mouseleave(function() {
        $('#navmenu').hide();
    });
    // END: Set up dropdown navigation menu ---------------------------------

    // BEGIN: Set up Search -------------------------------------------------
    //$('#searchBox').css('display', 'none');
    $('#classesSearchBox, .classesSearchBox').css('display', 'none');

    $('#searchBoxText').click(function() {
     	$('#classesSearchBox').show();
            $('#txtsearchphrase').focus();
    });
    $('#searchBoxText').mouseover(function() {
     	$('#classesSearchBox').show();
            $('#txtsearchphrase').focus();
    });

    $('#classesSearchBox').mouseleave(function() {
     	$('#classesSearchBox').hide();
    });

    $('.searchBoxText').click(function() {
        $('.classesSearchBox').show();
            $('.txtsearchphrase').focus();
    });
    $('.searchBoxText').mouseover(function() {
        $('.classesSearchBox').show();
            $('.txtsearchphrase').focus();
    });

    $('.searchBox').mouseleave(function() {
     	$('.classesSearchBox').hide();
    });

    var runSearch = function() {
        var lcl_searchphrase = '';

        if ($('#txtsearchphrase')) {
            lcl_searchphrase = $('#txtsearchphrase').val();
        }
        if ($('.txtsearchphrase')) {
		//alert(lcl_searchphrase);
	    if (lcl_searchphrase + '' == 'undefined')
    	    {
	        lcl_searchphrase = '';
    	    }
	    //alert(lcl_searchphrase);


		//alert($('#iframenav').find('.txtsearchphrase').val());
            lcl_searchphrase += $('#iframenav').find('.txtsearchphrase').val();
		//alert(lcl_searchphrase);
        }

        //if (lcl_searchphrase == '') {
            //$('#txtsearchphrase').focus();
            //inlineMsg(document.getElementById('searchButton').id, '<strong>Required Field Missing: </strong> Please enter a value to search.', 10, 'searchButton');
            //return false;
        //} else {
            //clearMsg('searchButton');
	    searchURL = 'rd_classes/class_list.aspx?keywordSearch=' + lcl_searchphrase;
	    if (window.location.pathname.split("/").length - 1 > 2)
    	    {
		searchURL = '../' + searchURL;
    	    }
	    //alert(window.location.pathname);
            location.href = searchURL;
        //}
    };

    $('#searchButton, .searchButton').click(runSearch);
    $('.txtsearchphrase').keypress(function(e) {
    	if(e.which == 13) {
		runSearch();
    	}
    });

    //$('#txtsearchphrase, .txtsearchphrase').change(function() {
        //clearMsg('searchButton');
    //});
    // END: Set up Search ---------------------------------------------------

    // BEGIN: Set up submenu navigation menu --------------------------------
    //$('#submenu_lists').css('display', 'none');
    //
    var showQL = function()
    {
        //expand_submenu('QUICKLINKS');
	//hide submenu_categories_list
	$("#submenu_categories_list").hide();
	//submenu_lists show
	$("#submenu_lists").show();
	//submenu_quicklinks_list show
	$("#submenu_quicklinks_list").show();
    };
    var showCat = function() {
        //expand_submenu('CATEGORIES');
	$("#submenu_quicklinks_list").hide();
	//submenu_lists show
	$("#submenu_lists").show();
	//submenu_quicklinks_list show
	$("#submenu_categories_list").show();
    };

    $('#submenu_quicklinks').click(showQL).mouseover(showQL);

    /* -- Classes/Events -- */
    $('#submenu_categories').click(showCat).mouseover(showCat);

    $("#submenu_lists").mouseleave(function() {
	$("#submenu_lists").hide();
    });

    //$('#submenu_search').click(function() {
    //expand_submenu('SEARCH');
    //});
    // END: Set up submenu navigation menu ----------------------------------
});

function expandSearchBox() {

}
function expandiframeSearchBox() {
    if ($('.classesSearchBox').css('display') != 'block') {
        $('.classesSearchBox').slideDown('slow', function() {
            $('.txtsearchphrase').focus();
        });
    } else {
        $('.classesSearchBox').slideUp('slow');
    }
}

function expand_submenu(iSubMenuOption) {
    var lcl_list_option_selected = '';
    var lcl_list_option_deselected1 = '';
    //var lcl_list_option_deselected2 = '';
    var lcl_list_show = '';
    var lcl_list_hide1 = '';
    //var lcl_list_hide2 = '';

    if (iSubMenuOption == 'CATEGORIES') {
        lcl_list_option_selected = 'submenu_categories';
        lcl_list_option_deselected1 = 'submenu_quicklinks';
        //lcl_list_option_deselected2 = 'submenu_search';
        lcl_list_show = 'submenu_categories_list';
        lcl_list_hide1 = 'submenu_quicklinks_list';
        //lcl_list_hide2 = 'submenu_search_box';
    //} else if (iSubMenuOption == 'SEARCH') {
    //    lcl_list_option_selected = 'submenu_search';
    //    lcl_list_option_deselected1 = 'submenu_quicklinks';
    //    lcl_list_option_deselected2 = 'submenu_categories';
    //    lcl_list_show = 'submenu_search_box';
    //    lcl_list_hide1 = 'submenu_quicklinks_list';
    //    lcl_list_hide2 = 'submenu_categories_list';
    } else {
        lcl_list_option_selected = 'submenu_quicklinks';
        lcl_list_option_deselected1 = 'submenu_categories';
        //lcl_list_option_deselected2 = 'submenu_search';
        lcl_list_show = 'submenu_quicklinks_list';
        lcl_list_hide1 = 'submenu_categories_list';
        //lcl_list_hide2 = 'submenu_search_box';
    }

    if ($('#' + lcl_list_show).css('display') == 'block') {
        $('#submenu_lists').hide( function() {
            $('#submenu_quicklinks').removeClass('submenu_option_selected');
            $('#submenu_categories').removeClass('submenu_option_selected');
            //$('#submenu_search').removeClass('submenu_option_selected');

            $('#submenu_quicklinks_list').css('display', 'none');
            $('#submenu_categories_list').css('display', 'none');
            //$('#submenu_search_box').css('display', 'none');
        });
    } else {
        if ($('#submenu_lists').css('display') == 'block') {
            $('#submenu_quicklinks_list').css('display', 'none');
            $('#submenu_categories_list').css('display', 'none');
            
            $('#submenu_lists').hide(function() {
                $('#' + lcl_list_option_deselected1).removeClass('submenu_option_selected');
                //$('#' + lcl_list_option_deselected2).removeClass('submenu_option_selected');
                $('#' + lcl_list_option_selected).addClass('submenu_option_selected');

                $('#' + lcl_list_hide1).css('display', 'none');
                //$('#' + lcl_list_hide2).css('display', 'none');
                $('#' + lcl_list_show).css('display', 'block');

                $('#submenu_lists').show();
                //$('#submenu_lists').slideDown('slow',function() {
                //    if(lcl_field_focus != '') {
                //        $('#' + lcl_field_focus).focus();
                //    }
                //});
            });
        } else {
            $('#submenu_quicklinks').addClass('submenu_quicklinks_selected');
            $('#' + lcl_list_option_deselected1).removeClass('submenu_option_selected');
            //$('#' + lcl_list_option_deselected2).removeClass('submenu_option_selected');
            $('#' + lcl_list_option_selected).addClass('submenu_option_selected');

            $('#' + lcl_list_hide1).css('display', 'none');
            //$('#' + lcl_list_hide2).css('display', 'none');
            $('#' + lcl_list_show).css('display', 'block');
            
            $('#submenu_lists').show();
            //$('#submenu_lists').slideDown('slow', function() {
            //    if(lcl_field_focus != "") {
            //        $('#' + lcl_field_focus).focus();
            //    }
            //});
            
        }
    }
}

function HideThings()
	{
		// Steve Loar 2/21/2006 - To hide form selects that block the dropdown menu

		// events/calendar.asp
		var formnames = document.getElementsByName("frmDate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmDate["month"]);
			if(bexists)
			{
				document.frmDate.month.style.visibility="hidden";
			}
			bexists = eval(document.frmDate["year"]);
			if(bexists)
			{
				document.frmDate.year.style.visibility="hidden";
			}
		}
		// recreation/facility_availability.asp
		formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="hidden";
			}
		}
}

function UnhideThings()
	{
		// Steve Loar 2/21/2006 - To unhide form selects that block the dropdown menu

		// events/calendar.asp
		var formnames = document.getElementsByName("frmDate");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmDate["month"]);
			if(bexists)
			{
				document.frmDate.month.style.visibility="visible";
			}
			bexists = eval(document.frmDate["year"]);
			if(bexists)
			{
				document.frmDate.year.style.visibility="visible";
			}
		}
		// recreation/facility_availability.asp
		formnames = document.getElementsByName("frmcal");
		if (formnames.length == 1)
		{
			var bexists = eval(document.frmcal["selfacility"]);
			if(bexists)
			{
				document.frmcal.selfacility.style.visibility="visible";
			}
		}
}
