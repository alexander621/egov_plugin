jQuery(document).ready(function($) {

    // BEGIN: Set up dropdown navigation menu --------------------------------
    /* prepend menu icon */
    //    $('#nav-wrap').prepend('<div id="menu-icon">Main Menu</div>');

    /* toggle nav - class categories */
    $('#mobile_subcategorymenu_option').click(function() {
        if ($('#mobile_subcategorylist').html() == '') {
            $('#mobile_subcategorylist').html($('#subcategorymenu_new').html());
            location.href = '#mobile_bookmark';
        } else {
            $('#mobile_subcategorylist').slideToggle('slow', function() {
                location.href = '#mobile_bookmark';
            });
        }
    });
});