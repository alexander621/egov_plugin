<%@ Page Language="C#" AutoEventWireup="true" CodeFile="mockup_response.aspx.cs" Inherits="action_response" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1"><meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Action Line Respond</title>

    <!-- Bootstrap -->
    <link href="mockup_css/bootstrap.min.css" rel="stylesheet" type="text/css" />
    <link href="mockup_css/nav.css" rel="stylesheet" type="text/css" />
    <link href="mockup_css/bootstrap-dropdown-menu-sliding.css" rel="stylesheet" type="text/css" />
    <link href="mockup_css/Site.css" rel="stylesheet" type="text/css" />
    <link href="mockup_css/grid.css" rel="stylesheet" type="text/css" />
    <link href="mockup_css/datepicker3.css" rel="stylesheet" type="text/css" />
    

    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
      <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->

        <!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
        <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.0/jquery.min.js"></script>
        <!-- Include all compiled plugins (below), or include individual files as needed -->
        <script src="mockup_js/bootstrap.min.js" type="text/javascript"></script>
	<script src="mockup_js/bootstrap-datepicker.js" type="text/javascript"></script>
	<script type="text/javascript" src="../scripts/getdates.js"></script>
	
	<style>
		table
		{
			font-size:14px;
		}
		.al-table .col-sm-3
		{
			font-weight:bold;
		}
	</style>
	
</head>
<body>
<div class="container">
<h3>Pothole</h3>

<div class="al-table">
	<div class="row">
		<div class="col-sm-3">Tracking Number</div>
		<div class="col-sm-9">10838190953</div>
	</div>
	<div class="row">
		<div class="col-sm-3">Date Time Received:</div>
		<div class="col-sm-9">3/21/2017 2:39:47 PM ( Eastern Time (US & Canada))</div>
	</div>
	<div class="row">
		<div class="col-sm-3">Created By</div>
		<div class="col-sm-9">Admin ECLink (Admin Employee)</div>
	</div>
	<div class="row">
		<div class="col-sm-3">Completed Date</div>
		<div class="col-sm-9"></div>
	</div>
	<div class="row">
		<div class="col-sm-3">Submitted By IP Address</div>
		<div class="col-sm-9">74.87.250.138</div>
	</div>
</div>


<div class="tabbable">
    <ul class="nav nav-tabs">
        <li class="active"><a href="#detailspane" data-toggle="tab">Details</a></li>
        <li><a href="#updatepane" data-toggle="tab">Update</a></li>
        <li><a href="#sharepane" data-toggle="tab">Share</a></li>
    </ul>
    <div class="tab-content">
        <div id="detailspane" class="tab-pane active">
	<fieldset>
		<legend>Contact Information</legend>
		<table class="table">
  			<tr><td>Name</td><td>Christina Schappacher</td></tr>
  			<tr><td>Last Name</td><td>test</td></tr>
  			<tr><td>Business Name</td><td>adsfadsf</td></tr>
  			<tr><td>Email</td><td>asdfadsf@dsfasdf.com</td></tr>
  			<tr><td>Daytime Phone</td><td></td></tr>
  			<tr><td>Fax</td><td></td></tr>
  			<tr><td>Address</td><td></td></tr>
  			<tr><td>Preferred Contact Method</td><td>No Response Necessary</td></tr>
		</table>
	</fieldset>
<br />
	<fieldset>
		<legend>Location</legend>
		<!--p>
		Latitude: 39.28534
		<br />
		Longitude: -84.48997
		</p-->
<div id="map-canvas"></div><style>
#map-canvas {
height: 200px;
width: 200px;
margin: 0px;
padding: 0px
}
</style>
<script>
var map;
function initialize() {
var myLatlng = new google.maps.LatLng(39.28534, -84.48997);
var mapOptions = {
zoom: 17,
center: myLatlng}
map = new google.maps.Map(document.getElementById('map-canvas'), mapOptions);
var marker = new google.maps.Marker({ 
position: myLatlng, 
map: map 
}); 
}
</script>
<script src="https://maps.googleapis.com/maps/api/js?v=3.exp&sensor=false&callback=initialize&key=AIzaSyBe75DmQ_4W2L0-LK089mcCx9IMoGztYdg" type="text/javascript"></script>
<strong>Address</strong><br />
9957 ADAMS AV<br />
Loveland, OH 45140
<br />

          <table class="table">
            <tr>
                <td>County</td>
                <td>Hamilton</td>
            </tr>
            <tr>
                <td>Parcel ID</td>
                <td>062000920063</td>
            </tr>
            <tr>
                <td>Listed Owner</td>
                <td>HUTCHENS TERRY L</td>
            </tr>
            <tr>
                <td>Legal Description</td>
                <td>ADAMS ST 25 X 100- LOT 92 A & J KNECHTS SUB OFTWIGH TWEE</td>
            </tr>
            <tr>
                <td>Comments</td>
                <td></td>
            </tr>
          </table>
	</fieldset>
	<br />
	<fieldset>
		<legend>Submitted Info</legend>
		<b>Nature of your inquiry...</b><br />terst</p><p><b>What type of business?</b><br />Retail<br /><br />
	</fieldset>
	<br />
	<fieldset>
		<legend>Internal Info</legend>
		There are no values entered.
	</fieldset>
	<br />
	<fieldset>
		<legend>Attachments</legend>
		<div class="table-responsive">
			<table class="table table-striped">
				<thead>
  					<tr><th>Date Added - Added By - Name - Action</th></tr>
				</thead>
				<tbody>
  					<tr><td colspan="3"><i>No Attachments added.</i></td></tr>
				</tbody>
			</table>
		</div>
		<input type="button" id="addAttach" onclick="$('#frmAddAttachment').toggleClass('hidden');$('#addAttach').toggleClass('hidden');$('#hideAttach').toggleClass('hidden');" class="btn btn-primary btn-sm" value="Add Attachment" />
		<input type="button" id="hideAttach" onclick="$('#frmAddAttachment').toggleClass('hidden');$('#addAttach').toggleClass('hidden');$('#hideAttach').toggleClass('hidden');" class="hidden btn btn-danger btn-sm" value="Cancel" />
		<form name="frmAddAttachment" id="frmAddAttachment" action="attachment_save.asp" method="POST" enctype="multipart/form-data" class="hidden">
  			<input type="hidden" name="status" id="status" value="SUBMITTED" />
  			<input type="hidden" name="irequestid" id="irequestid" value="1083962" />
  			<input type="hidden" name="screentype" id="screentype" value="E" />
  			<input type="hidden" name="attachmentFormName" id="attachmentFormName" value="frmAddAttachment" />
			<input type="hidden" name="attachmentIsSecure" id="attachmentIsSecure" value="off" checked="checked" />
			File
      			<input type="file" name="filAttachment" id="filAttachment" class="form-control" onchange="validateAttachment();" />
			Description
			<textarea name="attachmentdesc" class="form-control"></textarea>
      			<input type="submit" name="saveAttachmentButton" id="saveAttachmentButton" value="Save" class="btn btn-primary btn-sm" onclick="validateAttachment();" />
		</form>
		</fieldset>
		<br />
		<fieldset>
			<legend>Request Activity Log</legend>
			<div class="table-responsive">
				<table class="table table-striped">
					<thead>
      						<tr><th>User Name - Status <em>(Sub-Status)</em> - Edit Date</th></tr>
					</thead>
					<tbody>
  						<tr>
      						<td>Admin ECLink - SUBMITTED - 3/21/2017 2:39:47 PM</td>
  						</tr>  <tr>
      						<td>&nbsp;&nbsp;&nbsp;Internal Note: <em>This request was submitted by Admin ECLink.</em></td>
  						</tr>
  						<tr><td>Admin ECLink (Admin Employee) - SUBMITTED - 3/21/2017 2:39:47 PM</td></tr>
					</tbody>
				</table>
			</div>
		</fieldset>





	</div>
        <div id="updatepane" class="tab-pane">
		<fieldset>
  			<legend>Update Action Request</legend>
			<form name="frmUpdate" id="frmUpdate" action="action_respond.asp" method="post">
				<input type="hidden" name="TrackID" id="TrackID" value="1083962" />
				<input type="hidden" name="currentStatus" id="currentStatus" value="SUBMITTED" />
				<input type="hidden" name="currentSubStatus" id="currentSubStatus" value="" />
				<input type="hidden" name="prevAssignedemployeeid" id="prevAssignedemployeeid" value="5883" />
				<input type="hidden" name="currentDepartmentID" id="currentDepartmentID" value="3371" />
				<input type="hidden" name="currentDueDate" id="currentDueDate" value="" />
                		Assigned Employee
                    		<select name="assignedemployeeid" id="assignedemployeeid" class="form-control">
					<option value="5903,David Boyer">David Boyer</option>
					<option value="5901,Tim Boyle">Tim Boyle</option>
					<option value="12966,NKU Dev">NKU Dev</option>
					<option value="5904,Steve Loar">Steve Loar</option>
					<option value="5902,Jeff Nelson">Jeff Nelson</option>
					<option value="5884,Christina Schappacher">Christina Schappacher</option>
					<option value="5883,Peter Selden" selected="selected">Peter Selden</option>
                    		</select>
				<br />
                		Status
                    		<select name="selStatus" id="selStatus" onchange="changeSubStatus();" class="form-control">
                      			<option value="SUBMITTED" selected="selected">SUBMITTED</option>
                      			<option value="INPROGRESS">INPROGRESS</option>
                      			<option value="WAITING">WAITING</option>
                   			<option value="RESOLVED">RESOLVED</option>
                   			<option value="DISMISSED">DISMISSED</option>
                    		</select>
				<br />
                		Sub-Status
                    		<select name="selSubStatus" id="selSubStatus" class="form-control">
                      			<option value="0"></option>
                    		</select>
				<br />
				Due Date
				<input type="text" name="due_date" id="due_date" value="" readonly class="form-control"  />
				<script>
                            	$("#due_date").datepicker({ format: "m/d/yyyy", startDate: $("#due_date").val(), date: $("#toDate").val(), autoclose:true, clearBtn:true});
                        	</script>
				
				<br />
				Department
				<select name="deptid" id="deptid" onchange="clearMsg('deptid');" class="form-control">
					<option value="3370">Administrators</option>
					<option value="3371" selected="selected">City Employees</option>
				</select>
				<br />
                    		Internal Note
                    		<textarea name="internal_comment" id="internal_comment" rows="5" cols="80" class="form-control"></textarea>
				<br />
                          	Note to Citizen<!--input type="button" value="Add a Link" class="button" onclick="doPicker('frmUpdate.external_comment','Y','Y','Y','Y');" /-->
				<div class="checkbox"><label><input type="checkbox" class="big" name="sendemail" id="sendemail" value="yes" /> Send email to Citizen?</label></div>

                      		<textarea name="external_comment" id="external_comment" rows="5" cols="80" class="form-control"></textarea>
				<br />
          			<input type="button" name="sAction" class="btn btn-primary btn-sm" style="float:left;" value="UPDATE REQUEST" onclick="updateRequest()" />
				<input type="button" name="sAction" class="btn btn-primary btn-sm" style="float:right;" value="DELETE REQUEST" onclick="deleteconfirm(1083962);" />
			</form>
		</fieldset>
	</div>
        <div id="sharepane" class="tab-pane">
		<fieldset>
			<legend>Send Email Notification</legend>
			<form name="frmNotify" action="action_respond.asp" method="post">
  				<input type="hidden" name="TrackID" id="TrackID" value="1083962" />
  				<input type="hidden" name="prevnotifyuserid" id="prevnotifyuserid" value="" />
      				Notify User
          			<select name="notifyuserid" id="notifyuserid" onchange="clearMsg('notifyuserid')" class="form-control">
            				<option value=""></option>
  					<option value="5903">David Boyer</option>
  					<option value="5901">Tim Boyle</option>
  					<option value="5904">Steve Loar</option>
  					<option value="5902">Jeff Nelson</option>
  					<option value="5884">Christina Schappacher</option>
  					<option value="5883">Peter Selden</option>
          			</select>
				<br />
      				Notify Department
          			<select name="notifydeptid" id="notifydeptid" onchange="clearMsg('notifyuserid')" class="form-control">
            				<option value=""></option>
  					<option value="3370">Administrators</option>
  					<option value="3371">City Employees</option>
          			</select>
				<br />
          			Additional Comments
          			<textarea name="notify_additional_comments" id="notify_additional_comments" rows="5" cols="80" class="form-control"></textarea>
				<br />
          			<input type="submit" name="sAction" id="sAction" class="btn btn-primary btn-large" value="SEND NOTIFICATION" onclick="return checkDepartmentInactive();" /> 
			</form>
		</fieldset>
	</div>
    </div>
</div>
<br />
<br />
<br />
<br />
<br />
<br />
</div>
</body>
</html>
