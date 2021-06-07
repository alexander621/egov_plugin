<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN"  "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <meta http-equiv="content-type" content="text/html; charset=utf-8"/>
    <title>EGOVLINK MAP EXAMPLE</title>
    <script src="http://maps.google.com/maps?file=api&amp;v=2&amp;key=ABQIAAAAStOU3ZcVxvfIkpDt0vutZRQternwZke6s-i5RcvQUBYaa6sGNBRU9VUuUsvdIq7TKwzuH6PixiPYcw"
      type="text/javascript"></script>
	  <style>
		div.info {
			text-align:left;
			FONT-SIZE: 11px;
			FONT-FAMILY: Verdana,Tahoma,Arial
			}
		table.excel {}
		td.excelheader {font-weight:bold;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldata {background-color:white;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
		td.excelheaderleft {font-weight:bold;border-left: solid #000000 1px;background-color:#eeeeee;border-top: solid #000000 1px;border-right: solid #000000 1px;font-family: verdana,sans-serif; font-size: 10px;border-bottom: solid #000000 1px;}
		td.exceldataleft {background-color:white;border-left: solid #c0c0c0 1px;border-right: solid #c0c0c0 1px;border-bottom: solid #c0c0c0 1px;font-family: verdana,sans-serif; font-size: 10px;height:12px;}
	  </style>
    
  </head>


  <body  >
	<center>
	    <div id="map" style="border: solid 1px #000000;width: 795px; height: 595px"></div>
	</center>

	 <script type="text/javascript">

			var map = new GMap2(document.getElementById("map"));
			map.addControl(new GSmallMapControl());
			map.addControl(new GMapTypeControl());
			gmarkers = new Array(); 
			badaddresses = new Array();
			var i = 0;
			var ba = 0;
			var geocoder = new GClientGeocoder();

			// ADD MARKERS FOR TICKETS FROM DATABASE
			<% fnAddTroubleTickets() %>

			function showAddress(address,sMsg) {
			  geocoder.getLatLng(
				address,
				function(point) {
				  if (!point) {
					// CREATE AN ARRAY OF ADDRESSED UNABLE TO BE PLOTTED
					badaddresses[ba]=i;
					ba++;
					i++;
				  } else {

					// SET MAP CENTER POINT AT THIS LOCATION
					map.setCenter(point, 13);
					
					// ADD MARKER TO MAP
					var marker = new GMarker(point);
					map.addOverlay(marker);
					
					// SET CLICK EVENT TO DISPLAY MORE INFORMATION
					GEvent.addListener(marker, "click", function() {
					marker.openInfoWindowHtml(sMsg);
					});
					
					//INCREMENT MARKER ARRAY TO REFERENCE OFF MAP LINKS
					gmarkers[i]=marker;
					i++;				

				  }
				}
			  );
			}

			function checkplot(iIndex){
				
				// IF FALSE POINT WAS NOT PLOTTED
				if (isbadaddress(iIndex) != true) {
					alert('Error: This address could not be plotted!');
				}
				else
				{
					// POINT WAS PLOTTED. MOVE TO MAP TO SHOW MARKER
					GEvent.trigger(gmarkers[iIndex], 'click');
				}
			}

			function isbadaddress(iIndex){
				// LOOP THRU LIST OF BAD OF ADDRESS MAKE THIS POINT WAS PLOTTED
				for (var l=0;l<badaddresses.length;l++){
					if (iIndex == badaddresses[l]) {
						// NOT PLOTTED
						return false;
					}
				}
				// WAS PLOTTED
				return true;
			}


	 </script>


	 <P>
	 <center>
			<%fnGetTroubleTickets()%>
	 </center>
	 </p>

  </body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED FUNCTIONS
'--------------------------------------------------------------------------------------------------


'--------------------------------------------------------------------------------------------------
' FUNCTION FNGETTROUBLETICKETS()
'--------------------------------------------------------------------------------------------------
Function fnGetTroubleTickets()
	
	sSQL = replace(session("MAP_QUERY"),"SELECT","Select TOP 50 *,")
	Set oTickets = Server.CreateObject("ADODB.Recordset")
	oTickets.Open sSQL, Application("DSN"), 3, 1
	iMarkerCount = 0

	If NOT oTickets.EOF Then
		
		' START TABLE
		response.write "<table cellpadding=2 cellspacing=0 class=excel>"
		response.write "<tr><td class=excelheaderleft>Marker</td><td class=excelheader>Tracking Number</td><td class=excelheader>Ticket Name</td><td class=excelheader>Problem Address</td></tr>"

		Do While NOT oTickets.EOF 

			If TRIM(oTickets("useraddress")) <> "" AND NOT ISNULL(oTickets("useraddress")) Then
				response.write "<tr><td class=exceldataleft><a href=# onClick=""checkplot("& iMarkerCount &");"" >" & iMarkerCount & "</a></td><td class=exceldata>" & oTickets("action_autoid") & replace(FormatDateTime(oTickets("submit_date"),4),":","") & "</td><td class=exceldata>" & oTickets("action_formTitle") & "</td><td class=exceldata>" &  oTickets("useraddress") & "</td></tr>"
				iMarkerCount = iMarkerCount + 1
			End If

			oTickets.MoveNext
		Loop

		response.write "</table>"

	End If

End Function


'--------------------------------------------------------------------------------------------------
' FUNCTION FNADDTROUBLETICKETS()
'--------------------------------------------------------------------------------------------------
Function fnAddTroubleTickets()
	
	sSQL = replace(session("MAP_QUERY"),"SELECT","Select TOP 50 *,")
	Set oTickets = Server.CreateObject("ADODB.Recordset")
	oTickets.Open sSQL, Application("DSN"), 3, 1
	iMarkerCount = 0

	If NOT oTickets.EOF Then
			
		Do While NOT oTickets.EOF 
			
			If TRIM(oTickets("useraddress")) <> "" AND NOT ISNULL(oTickets("useraddress")) Then
				iTrackingNumber = oTickets("action_autoid") & replace(FormatDateTime(oTickets("submit_date"),4),":","")
				sMsg = "<div class=info>"
				sMsg = sMsg & "<b>Marker: </b>" & iMarkerCount & "<BR>"
				sMsg = sMsg & "<b>Tracking Number: </b>" & iTrackingNumber & "<BR>"
				sMsg = sMsg & "<b>Ticket Name: </b>" & replace(oTickets("action_formTitle"),vbcrlf,"") & "<br>"
				sMsg = sMsg & "<b>Problem Address: </b>" & oTickets("useraddress") & "<br>"
				sMsg = sMsg & "</div>"
				iMarkerCount = iMarkerCount + 1
				
				response.write "showAddress('" & oTickets("useraddress") & " " & oTickets("usercity") & ", OH " & oTickets("userzip") & ",USA','" & sMsg & "');" & vbcrlf
			End If

			oTickets.MoveNext
		Loop



	End If

	

End Function
%>

