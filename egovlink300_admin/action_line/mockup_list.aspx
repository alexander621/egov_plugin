<%@ Page Language="C#" AutoEventWireup="true" CodeFile="mockup_list.aspx.cs" Inherits="list" %>
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1"><meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Action Line List</title>

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
	
	
</head>
<body>
<div class="container" style="position:relative;">
<div  id="searchparams" class="hidden" style="margin-left:-15px;padding: 0 15px 0 15px;position:absolute;z-index:100;background:white;width:100%;height:100%;">
	<br />
	<fieldset>
		<legend style="background:white;">Available Filters <a class="btn btn-sm btn-primary" onclick="javascript:$('#searchparams').toggleClass('hidden');">< Back</a></legend>
		<form class="searchparamsfrm">
		<div class="row">
			<div class="col-md-9">

Assigned To: 
<select name="selectAssignedto" id="selectAssignedto" class="form-control">
  <option value="all">Anyone</option>
  <option value="5903" >David Boyer</option>
  <option value="5901" >Tim Boyle</option>
  <option value="12966" >NKU Dev</option>
  <option value="5882" >Admin ECLink</option>
  <option value="5904" >Steve Loar</option>
  <option value="5902" >Jeff Nelson</option>
  <option value="5884" >Christina Schappacher</option>
  <option value="5883" >Peter Selden</option>
</select>

                      Order By:
                      <select name="orderBy" id="orderBy" class="form-control">
  <option value="assigned_Name">Assigned To</option>
  <option value="action_Formid">Form</option>
  <option value="submit_date" selected="selected">Date Descending</option>
  <option value="deptId">Department</option>
  <option value="streetname">Issue/Problem Location Street Name</option>
  <option value="submittedby">Submitted By</option>
  <option value="status">Status</option>
                      </select>

                      Status:
		      <div>
<div class="checkbox"><label><input class="big" type="checkbox" name="statusSUBMITTED" id="statusSUBMITTED" value="yes" checked="checked" /> Submitted</label></div>
<div class="checkbox"><label><input class="big" type="checkbox" name="statusINPROGRESS" id="statusINPROGRESS" value="yes" checked="checked" /> In Progress</label></div>
<div class="checkbox"><label><input class="big" type="checkbox" name="statusWAITING" id="statusWAITING" value="yes" checked="checked" /> Waiting</label></div>
<div class="checkbox"><label><input class="big" type="checkbox" name="statusRESOLVED" id="statusRESOLVED" value="yes" checked="checked" /> Resolved</label></div>
<div class="checkbox"><label><input class="big" type="checkbox" name="statusDISMISSED" id="statusDISMISSED" value="yes" checked="checked" /> Dismissed</label></div>
          <input type="hidden" name="substatus_hidden" id="substatus_hidden" value="" />
          <input type="hidden" name="show_hide_substatus" id="show_hide_substatus" value="HIDE" />
	  </div>
                      Categories and Forms:
                      <select name="selectFormId" id="selectFormId" onChange="toggleregAddy(this.value);" class="form-control">
                        <option value="">All Categories</option>
  <option value="C1541" >----Category: Citizen Comments and Concerns</option>
  <option value="10836" >Crime Stoppers</option>
  <option value="10851" >General Complaint Form</option>
  <option value="10798" >General Suggestion</option>
  <option value="10835" >Request to Add Calendar Item</option>
  <option value="C1542" >----Category: Requests for Information</option>
  <option value="10839" >Building Permits</option>
  <option value="10834" >Burning Permit</option>
  <option value="10846" >Business License Inquiry</option>
  <option value="10831" >Mailing Lists </option>
  <option value="10830" >New Resident Packet</option>
  <option value="10842" >Other - Request for Information</option>
  <option value="10833" >Sewer General </option>
  <option value="10832" >Tax Questions</option>
  <option value="10840" >Water General</option>
  <option value="C1543" >----Category: Repairs and Requests for Service</option>
  <option value="10802" >Appliance Collection </option>
  <option value="10854" >Code Enforcement Request</option>
  <option value="10803" >Curb and Gutter </option>
  <option value="10809" >Downtown Items</option>
  <option value="10805" >Dumpster Billing </option>
  <option value="10806" >Dumpster Service </option>
  <option value="10807" >Landscaping</option>
  <option value="10808" >Line of Sight Obstruction </option>
  <option value="10801" >Mowing </option>
  <option value="10810" >Parks and Recreation Survey</option>
  <option value="10811" >Pothole </option>
  <option value="10918" >Property Maintenance Code Violations </option>
  <option value="10853" >Public Works Service Request</option>
  <option value="10812" >Recycling </option>
  <option value="10814" >Sewer Leak </option>
  <option value="10815" >Sewer Odor </option>
  <option value="10816" >Sewer Spill </option>
  <option value="10817" >Sidewalk Repair </option>
  <option value="10845" >Snow - Plowing/Removal</option>
  <option value="10818" >Storm Drainage/Erosion </option>
  <option value="10819" >Street Improvements </option>
  <option value="10820" >Street Light Out </option>
  <option value="10813" >Street Name Signs</option>
  <option value="10822" >Street Sweeping </option>
  <option value="10823" >Traffic Signals / Signs</option>
  <option value="10821" >Tree Service</option>
  <option value="10824" >Utility Cuts </option>
  <option value="10825" >Utility Landscape </option>
  <option value="10826" >Utility Patching </option>
  <option value="10827" >Water & Sewer Locates (Request to Dig)</option>
  <option value="10852" >Water Connect/Disconnect</option>
  <option value="10828" >Water Leak </option>
  <option value="10829" >Yard Waste</option>
  <option value="C1544" >----Category: Nuisance/Code Violations</option>
  <option value="10841" >Brush</option>
  <option value="10799" >Building Code Violations </option>
  <option value="10849" >Grass complaint</option>
  <option value="10800" >Junk Cars </option>
  <option value="10837" >Lot Mowing</option>
  <option value="10843" >Other - Nuisance/Code Violations</option>
  <option value="10838" >Parking in Yard</option>
  <option value="C1545" >----Category: Animal Control</option>
  <option value="10804" >Dead animal removal </option>
  <option value="10844" >Other - Animal Control</option>
  <option value="10797" >Stray Animals</option>
  <option value="C1546" >----Category: Licenses</option>
  <option value="10847" >Dog Registration Annual Application</option>
  <option value="10848" >Pool Registration</option>
  <option value="C1547" >----Category: Other</option>
  <option value="10850" >IT Request</option>
                      </select>
        Department: 
                      <select name="selectDeptId" id="selectDeptId" class="form-control">
                        <option value="all">All Departments</option>
  <option  value="3370">Administrators</option>
  <option  value="3371">City Employees</option>
                      </select>
	      <input type="hidden" name="reporttype" id="reporttype" value="List" />

<br />
                        Show <select name="selectDateType" id="selectDateType" class="form-control" style="width:initial;display:inline-block;">
                          <option value="active" selected>Active Requests</option>
                          <option value="submit">Submit Date</option>
                        </select>
                        <nobr>From: 
                        <input type="text" name="fromDate" id="fromDate" value="3/21/2016" size="6" maxlength="10" class="form-control" style="width:initial;display:inline-block;" />
                        To:
                        <input type="text" name="toDate" id="toDate" value="3/20/2017" size="6" maxlength="10" class="form-control" style="width:initial;display:inline-block;" />
			</nobr>
			<script>
                            $("#fromDate").datepicker({ format: "m/d/yyyy", date: $("#fromDate").val(), autoclose: true, clearBtn: true })
                            .on('changeDate', function (dateEvent) {
                                start_date = $("#fromDate").val();
                                $('#toDate').datepicker('setStartDate', start_date);

                                if (Date.parse($('#toDate').val()) < Date.parse(start_date))
                                    $('#toDate').val(start_date)
                                    $('#toDate').datepicker('setDate', start_date);
                            });

                            $("#toDate").datepicker({ format: "m/d/yyyy", startDate: $("#fromDate").val(), date: $("#toDate").val(), autoclose:true, clearBtn:true});
                        </script>
			
<select onChange="getDates(this.value, 'Date');" class="calendarinput form-control" name="Date" id="fromToDateSelection" style="width:initial;display:inline-block;">
  <option value="0">Or Select Date Range from Dropdown...</option>
  <option value="16">Today</option>
  <option value="17">Yesterday</option>
  <option value="18">Tomorrow</option>
  <option value="11">This Week</option>
  <option value="12">Last Week</option>
  <option value="14">Next Week</option>
  <option value="1">This Month</option>
  <option value="2">Last Month</option>
  <option value="13">Next Month</option>
  <option value="3">This Quarter</option>
  <option value="4">Last Quarter</option>
  <option value="15">Next Quarter</option>
  <option value="6">Year to Date</option>
  <option value="5">Last Year</option>
  <option value="7">All Dates to Date</option>
</select>
<br />
<br />

                      <h4>Submitted By:</h4>
                      Contact First Name: <input type="text" name="selectUserFName" id="selectUserFName" placeholder="All" size="12" class="form-control" />
                      Contact Last Name: <input type="text" name="selectUserLName" id="selectUserLName" placeholder="All" size="12" class="form-control" />
                      Contact Street Name:&nbsp; <input type="text" name="selectContactStreet" id="selectContactStreet" placeholder="All" class="form-control" />
		      <br />

                      <h4>Issue/Problem Location:</h4>
                      Street Number: <input type="text" name="selectIssueStreetNumber" id="selectIssueStreetNumber" value="" size="10" maxlength="150" class="form-control" />
                      Street Name: <input type="text" name="selectIssueStreet" id="selectIssueStreet" value="" size="30" maxlength="300" class="form-control" />

                      County:
                      <input type="text" name="selectCounty" id="selectCounty" value="" size="30" maxlength="50" class="form-control" />
		<br />
                      Business Name:
                      <input type="text" name="selectBusinessName" id="selectBusinessName" Placeholder="All" class="form-control" />

                      Tracking Number:
                      <input type="text" name="selectTicket" id="selectTicket" value="" size="15" class="form-control" />

                      Display Open Over X Days
                      <input type="text" name="pastDays" id="pastDays" value="" size="2" onchange="clearMsg('pastDays');" class="form-control" />
                      <input type="hidden" name="searchDaysType" id="searchDaysType" value="OPEN" />

<br />
                      <input type="button" name="searchButton" id="searchButton" value="Show Results" class="btn btn-lg btn-primary" style="float:right;" onclick="clearScreenMsg();submitForm();" />
		      <div style="clear:right;"></div>
		      <br />
		      <br />
		      <br />
		      <br />
		      </div>
			<div class="col-md-3">
			
<input type="hidden" name="userReportName" id="userReportName" value="ActionLine - User Saved Search Options" size="10" maxlength="200" />
<fieldset class="fieldset">
  <legend>Default Search Options</legend>
  <table border="0" cellspacing="0" cellpadding="0" style="padding-top:5px">
    <tr valign="top">
        <td align="center" nowrap="nowrap">
            Currently using<br />
            <span id="customSearchDisplay" class="redText">System Options</span>
        </td>
        <td align="right"><input type="button" value="Run Default Search" class="btn btn-sm btn-primary" onclick="location.href='action_line_list.asp?init=Y';" /></td>
    </tr>
    <tr><td colspan="2">&nbsp;</td></tr>
    <tr><td colspan="2"><input type="button" name="useMyDefaults" id="useMyDefaults" class="btn btn-sm btn-primary" style="width:300px" value="Set Current Search Options as Default" onclick="updateCustomReport('8189','USER','Y');" /></td></tr>
    <tr><td colspan="2" style="padding-top:5px"><input type="button" name="useSystemDefaults" id="useSystemDefaults" class="btn btn-sm btn-primary" style="width:300px" value="Set System Search Options as Default" onclick="updateCustomReport('8189','SYSTEM','N');" /></td></tr>
  </table>
</fieldset>
</div>
		</div>
	</div>
		</form>
	</fieldset>
	<div style="float:left;"><h3>Action Line Items</h3></div>
	<div style="float:right;"><a class="btn btn-lg btn-primary" onclick="javascript:$('#searchparams').toggleClass('hidden');">Refine</a></div>
	<div id="list-group">
	<div class="row">
		<div class="col-xs-12">
			<div class="grid-row">
				<div class="container">
        				<div style="float:left;" class="wo-row col-xs-4">
            					<a class="btn btn-sm btn-default woduedate btn-prevm" href="#" onclick="return false;" style="">Submitted<br />3/21/2017 </a>
				
            					<div id="wostatus7528div" style="display:inline-block">
    							<a class="btn btn-sm btn-prevm btn-status" href="#statuschangemodal" data-id="7528" data-val="11" data-toggle="modal" style="background:cyan; color:black;"><div class="overlay"></div>INPROGRESS</a>
            					</div>
        				</div>
        				<div class="wo-row col-xs-4 wo-row-name" style="float:left;">
						10838190953 Pothole
						<br />
						<span style="font-weight:normal">From: Christina Schappacher</span>
					</div>
        				<div class="wo-row col-xs-4" style="float:left;text-align:right;">
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Assigned<br />Peter Selden</a>
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Department<br />City Employees</a>
        				</div>
				</div>
			</div>
		</div>
	</div>
	<div class="row">
		<div class="col-xs-12">
			<div class="grid-row">
				<div class="container">
        				<div style="float:left;" class="wo-row col-xs-4">
            					<a class="btn btn-sm btn-default woduedate btn-prevm" href="#" onclick="return false;" style="">Submitted<br />3/17/2017 </a>
				
            					<div id="wostatus7528div" style="display:inline-block">
    							<a class="btn btn-sm btn-prevm btn-status" href="#statuschangemodal" data-id="7528" data-val="11" data-toggle="modal" style="background:orange; color:white;"><div class="overlay"></div>WAITING</a>
            					</div>
        				</div>
        				<div class="wo-row col-xs-4 wo-row-name" style="float:left;">
						10793231041 Brush
						<br />
						<span style="font-weight:normal">From: Jane Doe</span>
					</div>
        				<div class="wo-row col-xs-4" style="float:left;text-align:right;">
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Assigned<br />Joe Felix</a>
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Department<br />City Employees</a>
        				</div>
				</div>
			</div>
		</div>
	</div>
	<div class="row">
		<div class="col-xs-12">
			<div class="grid-row">
				<div class="container">
        				<div style="float:left;" class="wo-row col-xs-4">
            					<a class="btn btn-sm btn-default woduedate btn-prevm" href="#" onclick="return false;" style="">Submitted<br />3/10/2017 </a>
				
            					<div id="wostatus7528div" style="display:inline-block">
    							<a class="btn btn-sm btn-prevm btn-status" href="#statuschangemodal" data-id="7528" data-val="11" data-toggle="modal" style="background:red; color:white;"><div class="overlay"></div>SUBMITTED</a>
            					</div>
        				</div>
        				<div class="wo-row col-xs-4 wo-row-name" style="float:left;">
						10838160947 Appliance Collection
						<br />
						<span style="font-weight:normal">From: Terry Foster</span>
					</div>
        				<div class="wo-row col-xs-4" style="float:left;text-align:right;">
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Assigned<br />Jerry Felix</a>
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Department<br />City Employees</a>
        				</div>
				</div>
			</div>
		</div>
	</div>
	<div class="row">
		<div class="col-xs-12">
			<div class="grid-row">
				<div class="container">
        				<div style="float:left;" class="wo-row col-xs-4">
            					<a class="btn btn-sm btn-default woduedate btn-prevm" href="#" onclick="return false;" style="">Submitted<br />1/10/2017 </a>
				
            					<div id="wostatus7528div" style="display:inline-block">
    							<a class="btn btn-sm btn-prevm btn-status" href="#statuschangemodal" data-id="7528" data-val="11" data-toggle="modal" style="background:Green; color:white;"><div class="overlay"></div>RESOLVED</a>
            					</div>
        				</div>
        				<div class="wo-row col-xs-4 wo-row-name" style="float:left;">
						1136231451 Property Maintenance Code Violations
						<br />
						<span style="font-weight:normal">From: n/a</span>
					</div>
        				<div class="wo-row col-xs-4" style="float:left;text-align:right;">
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Assigned<br />Jerry Felix</a>
            					<a id="woU7528" data-id="7528" data-val="32" class="btn btn-sm btn-default btn-prevm" href="#userassignment" data-toggle="modal" style="">Department<br />City Employees</a>
        				</div>
				</div>
			</div>
		</div>
	</div>
	</div> <!--end list group -->
</div>
<script>
$(function () {
	$(".wo-row").click(function (e) {
        if (e.target == e.currentTarget) {
            window.location = "mockup_response.aspx";
        }
    });
});
</script>
</body>
</html>
