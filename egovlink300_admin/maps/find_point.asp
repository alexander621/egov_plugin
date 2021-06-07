<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="content-type" content="text/html; charset=UTF-8"/>
    <title>Google Maps API Example - Geocoding API</title>
    <script src="http://maps.google.com/maps?file=api&amp;v=2.x&amp;key=ABQIAAAAStOU3ZcVxvfIkpDt0vutZRQternwZke6s-i5RcvQUBYaa6sGNBRU9VUuUsvdIq7TKwzuH6PixiPYcw" type="text/javascript"></script>
    <script type="text/javascript">
    //<![CDATA[

    var map = null;
    var geocoder = null;

    function load() {
      if (GBrowserIsCompatible()) {
        map = new GMap2(document.getElementById("map"));
        map.setCenter(new GLatLng(37.4419, -122.1419), 13);
        geocoder = new GClientGeocoder();
      }
    }

    function showAddress(address) 
	{
      if (geocoder) 
	  {
        pointy = new geocoder.getLatLng(
          address,
          function(point) {
            if (!point) 
			{
              alert(address + " not found");
            } 
			else 
			{
				//alert(point);
              map.setCenter(point, 13);
              var marker = new GMarker(point);
              map.addOverlay(marker);
              marker.openInfoWindowHtml(address + ' ' + point);
            }
          }
        );
		//document.getElementById("lat").value = geocoder.lat();
		//document.getElementById("lng").value = geocoder.lng();
      }
    }
    //]]>
    </script>
  </head>

  <body onload="load()" onunload="GUnload()">
    <form action="#" onsubmit="showAddress(this.address.value); return false">
      <p>
        <input type="text" size="60" name="address" value="1600 Amphitheatre Parkway, Mountain View, CA" />
        <input type="submit" value="Go!" />
      </p>
      <div id="map" style="width: 500px; height: 300px"></div>
	  <p>
	  Lat: <input type="text" id="lat" /><br />
	  Lng: <input type="text" id="lng" />
	  </p>
    </form>
  </body>
</html>

