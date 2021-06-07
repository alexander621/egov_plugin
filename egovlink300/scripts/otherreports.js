
var othermap;
var otherbounds;
var infoWindow;
var infoWindowContent;
var markets;
function initalizeOtherReports() {
    otherbounds = new google.maps.LatLngBounds();

    var mapOptions = {
        mapTypeId: 'roadmap'
    };
                    
    // Display a map on the page
    othermap = new google.maps.Map(document.getElementById("map_canvas"), mapOptions);
    othermap.setTilt(45);
        
    // Display multiple markers on a map
    infoWindow = new google.maps.InfoWindow()
    var marker, i;
    
    // Loop through our array of markers & place each one on the map  
    for( i = 0; i < markers.length; i++ ) {
        var position = new google.maps.LatLng(markers[i][1], markers[i][2]);
        otherbounds.extend(position);
        marker = new google.maps.Marker({
            position: position,
            map: othermap,
            title: markers[i][0]
        });
        
        // Allow each marker to have an info window    
        google.maps.event.addListener(marker, 'click', (function(marker, i) {
            return function() {
	        var latLng = marker.getPosition(); // returns LatLng object
		othermap.setCenter(latLng); // setCenter takes a LatLng object

		othermap.setZoom(18);
                infoWindow.setContent(infoWindowContent[i][0]);
                infoWindow.open(othermap, marker);
            }
        })(marker, i));


        // Automatically center the map fitting all markers on the screen
        othermap.fitBounds(otherbounds);
    }


    // Override our map zoom level once our fitBounds function runs (Make sure it only runs once)
    var boundsListener = google.maps.event.addListener((othermap), 'bounds_changed', function(event) {
        //this.setZoom(14);
        //google.maps.event.removeListener(boundsListener);
    });
    
}
	function hideModal()
	{
		infoWindow.close();
		document.getElementById("cover").style.display = "none";
	}

    window.addEventListener('click', function(e){   
  	if (document.getElementById('map_wrapper').contains(e.target) || document.getElementById('otherrptsbtn').contains(e.target) || e.target.tagName == "BUTTON" ){
    		// Clicked in box
  	} else{
    		// Clicked outside the box
		infoWindow.close();
		document.getElementById("cover").style.display = "none";
 	}
    });

function upvote( id )
{
	//alert(id);
	
	jQuery.ajax({
		url:    'https://www.egovlink.com/eclink/upvote.asp',
		type: "GET",
		data: {
			id : id
		},
   		success: function(result) {
			if (result != "success")
			{
				alert("Sorry, there was an error with your request.  Please try again later.");
			}
			else
			{
				//jQuery('#map_wrapper').html("<div>We've recorded your report in our system.</div>");
				infoWindow.close();
				jQuery("#mapstuff").hide();
				jQuery("#successMessage").show();
			}
                }
    	});      
}

function showMapModal()
{
	jQuery('#successMessage').hide();
	jQuery('#mapstuff').show();
	jQuery('#cover').show();
	othermap.fitBounds(otherbounds);
}
