<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: CLIENT_TEMPLATE_PAGE.ASP
' AUTHOR: JOHN STULLENBERGER
' CREATED: 01/17/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  
'
' MODIFICATION HISTORY
' 1.0   01/17/06   JOHN STULLENBERGER - INITIAL VERSION
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
%>


<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->


<html>
<head>
	<title>E-Gov Administration Console</title>
	
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

</head>


<body>

 
<%DrawTabs tabRecreation,1%>


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
		<p>
			Name: <input type="text" name="classname" value="Drawing Landscapes with Ease" size="50" maxlength="50" />
			&nbsp; <strong>This is a Single Class/Event</strong> <!--&nbsp; <input type="button" class="button" name="gotoparent" value="Edit Parent" onclick="javascript:location.href='class_editseries.asp'; " />-->
		</p>
		<p>
			Status: <strong>Active</strong>  &nbsp; &nbsp; <input type="button" class="button" name="cancel" value="Cancel Class/Event" />
			&nbsp; If cancelled, the reason is displayed here.
		</p>
		<fieldset class="edit"><legend><strong> Categories </strong></legend>
		<p>
			<!--<div id="topbuttons">-->
				<input type="button" class="assignbuttons" name="categories" value="Assign Categories" />
			<!--</div> -->
			Fitness, Adult Learning 
			
		</p>
		</fieldset>
		<fieldset class="edit"><legend><strong> General Information </strong></legend>
		<p>
			Description:<br /><textarea name="classdescription" id="classdescription">Outdoor Workshop for Teens and Adults
					Bring any drawing art supplies for experimentation 
					(sketch pad, pencils, crayons & erasers included)</textarea>
		</p>
		<p>
			Image: (show image here) URL: <input type="text" value=""><input type="button" value="...">
		</p>
		<p>
			Search Keywords:<br />
			<input type="text" name="searchkeywords" value="art drawing" size="100" maxlength="1000" />
		</p>
		<p>
			Minimum Age: <input type="text" name="minage" value="13" size="3" maxlength="2" /> &nbsp; 
			Maximum Age: <input type="text" name="maxage" value="" size="3" maxlength="2" />  
		</p>
		<p>
			Location: <select name="locationid" >
						<option value="0">None</option>
						<option value="1">Annex Building</option>
						<option value="2" selected="selected">Terwilliger's Lodge</option>
						<option value="6">Weller Park</option>
					</select>
		</p>
		<p>
			Point of Contact: <select name="pocid" >
						<option value="1" selected="selected">City Hall</option>
						<option value="2">Pat Stroz</option>
						<option value="3">Matthew Vanderhorst</option>
						<option value="2">Amber Morris</option>
					</select>
		</p>
		<p>
			External URL: <input type="text" name="externalurl" value="" size="50" maxlength="255" />
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Waivers </strong></legend>
			<p>
				<input type="button" class="assignbuttons" name="waivers" value="Assign Waivers" />
				General Terms 
			</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Instructors </strong></legend>
			<p>
				<input type="button" class="assignbuttons" name="instructors" value="Assign Instructors" />
				None 
			</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Critical Dates </strong></legend>
		<p>
			Starts: <input type="text" class="datefield" name="startdate" value="5/14/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
			Ends: <input type="text" class="datefield" name="enddate" value="5/14/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
		</p>
		<p>
			Publication Starts: <input type="text" class="datefield" name="publishstartdate" value="4/12/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
			Publication Ends: <input type="text" class="datefield" name="publishenddate" value="5/14/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
		</p>
		<p>
			Registration Starts: <input type="text" class="datefield" name="registrationstartdate" value="4/12/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
			Registration Ends: <input type="text" class="datefield" name="registrationenddate" value="5/14/2006" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
		</p>
		<p>
			Send Evaluation: <input type="text" class="datefield" name="evaluationdate" value="" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
			Alternate Date: <input type="text" class="datefield" name="alternatedate" value="" />&nbsp;<img src="../images/calendar.gif" height="16" width="16" border="0" />
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Purchasing </strong></legend>
		<p>
			Requires: <select name="optionid" >
						<option value="2">Tickets</option>
						<option value="1" selected="selected">Registration</option>
						<option value="3">Open Attendance</option>
						<option value="4">Information Only</option>
					</select>
		</p>
		<p>
			<table id="pricingtable">
				<caption>Pricing:</caption>
				<tr><td><input type="checkbox" name="pricetypeid" value="1" /> Resident </td><td><input type="text" name="amount1" value="" /></td></tr>
				<tr><td><input type="checkbox" name="pricetypeid" value="2" /> NonResident </td><td><input type="text" name="amount2" value="" /></td></tr>
				<tr><td><input type="checkbox" name="pricetypeid" value="3" /> Member </td><td><input type="text" name="amount3" value="" /></td></tr>
				<tr><td><input type="checkbox" name="pricetypeid" value="4" /> NonMember </td><td><input type="text" name="amount4" value="" /></td></tr>
				<tr><td><input type="checkbox" name="pricetypeid" value="5" checked="checked" /> Everyone </td><td><input type="text" name="amount5" value="30.00" /></td></tr>
			</table>
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Occurs </strong></legend>
		<p>
			Days of the Week: 
				<input type="checkbox" name="dayofweek" value="1" /> Sunday
				<input type="checkbox" name="dayofweek" value="2" /> Monday
				<input type="checkbox" name="dayofweek" value="3" /> Tuesday
				<input type="checkbox" name="dayofweek" value="4" /> Wednesday
				<input type="checkbox" name="dayofweek" value="5" /> Thursday
				<input type="checkbox" name="dayofweek" value="6" /> Friday
				<input type="checkbox" name="dayofweek" value="7" checked="checked" /> Saturday
		</p>
		<p>
				<table id="seriestime" border="0" cellpadding="0" cellspacing="0">
				<caption>Time:</caption>
				<tr><th>Start</th><th>End</th><th>Min</th><th>Max</th><th>Waitlist<br />Max</th><th>&nbsp;</th></tr>
				<tr>
					<td><input type="text" name="starttime1" value="1:00PM" size="8" maxlength="7" /></td>
					<td><input type="text" name="endtime1" value="4:00PM" size="8" maxlength="7" /></td>
					<td><input type="text" name="min1" value="" size="4" maxlength="5" /></td>
					<td><input type="text" name="max1" value="" size="4" maxlength="5" /></td>
					<td><input type="text" name="maxwaitlist1" value="" size="4" maxlength="5" /></td>
					<td><input type="button" class="button" name="remove" value="Remove" /></td>
				</tr>
				<tr>
					<td><input type="text" name="starttime2" value="" size="8" maxlength="7" /></td>
					<td><input type="text" name="endtime2" value="" size="8" maxlength="7" /></td>
					<td><input type="text" name="min2" value="" size="4" maxlength="5" /></td>
					<td><input type="text" name="max2" value="" size="4" maxlength="5" /></td>
					<td><input type="text" name="maxwaitlist2" value="" size="4" maxlength="5" /></td>
					<td><input type="button" class="button" name="add" value="Add" /></td>
				</tr>
				</table>
		</p>
		</fieldset>

		<p>
			<input type="button" class="button" name="update" value="Update" />
			<input type="button" class="button" name="copy" value="Copy to New Class/Event" />
			
		<p>

		<!--<fieldset class="edit"><legend><strong> Child Classes/Events </strong></legend>
		<p>
			<input type="button" name="child" value="Create New Child" id="newchild" /><br />
			<table id="children" border="1" cellpadding="0" cellspacing="0">
				<tr><th>Class/Event Name</th><th>Starts</th><th>Ends</th><th>&nbsp;</th></tr>
				<tr class="alt_row"><td>Summer Music and Mixer - May 14th</td><td>5/14/2006</td><td>5/14/2006</td><td><input type="button" name="editchild" value="Edit" /></td></tr>
				<tr><td>Summer Music and Mixer - June 11th</td><td>6/11/2006</td><td>6/11/2006</td><td><input type="button" name="editchild" value="Edit" /></td></tr>
			</table>
		</p>
		</fieldset>-->

	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>



<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------
%>


