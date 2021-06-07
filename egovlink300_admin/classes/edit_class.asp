<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: edit_class.asp
' AUTHOR: Steve Loar
' CREATED: 04/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the editing of classes and events
'
' MODIFICATION HISTORY
' 1.0	 04/19/2006	Steve Loar - INITIAL VERSION
' 1.1	 10/11/2006	Steve Loar - Security, Header and nav changed
' 2.0	 03/09/2007	Steve Loar - Total make over for Menlo Park Project
' 2.1	 11/29/2007	Steve Loar - Added yui tabs
' 2.2	 02/15/2008	Steve Loar - Early Registration added
' 2.3	 03/13/2008	Steve Loar - Early Registration changed to allow multiple classes to be picked.
' 2.4  12/30/2008 David Boyer - Added "DisplayRosterPublic" checkbox for Craig, CO custom registration fields.
' 2.5  06/17/2009	David Boyer	- Added "Show Terms" checkbox
' 2.6  11/17/2009 David Boyer - Added team registration options for t-shirts/pants (input options)
' 2.7	 12/02/2009	Steve Loar - Option to only allow purchases on admin but display on public
' 2.8	 12/07/2009	Steve Loar - Changes to link to rentals for Menlo Park Project
' 2.9	 10/10/2011	Steve Loar - Added Gender Restriction
' 3.0  02/02/2012 David Boyer - Modified the "Team Registration" section to add more choices
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iMaxTimeRows, iMin, iMax, lcl_scripts, lcl_orghasfeature_class_supervisors
Dim lcl_orghasfeature_gl_accounts, lcl_orghasfeature_discounts
Dim lcl_orghasfeature_custom_registration_craigco, lcl_onload
Dim iGenderRestrictionId, bHasGenderRestrictions
Dim sSql, oRs, iClassId, bIsSeriesParent

sLevel       = "../"  'Override of value from common.asp
iMaxTimeRows = 0
iMin         = 0
iMax         = 0
lcl_scripts  = ""

iClassId        = request("classid")
bIsSeriesParent = False

'Check if page is online and user has permissions in one call not two
PageDisplayCheck "manage classes", sLevel	' In common.asp

'Check for org features
lcl_orghasfeature_class_supervisors           = orghasfeature("class supervisors")
lcl_orghasfeature_gl_accounts                 = orghasfeature("gl accounts")
lcl_orghasfeature_discounts                   = orghasfeature("discounts")
lcl_orghasfeature_custom_registration_craigco = orghasfeature("custom_registration_CraigCO")
bHasGenderRestrictions                        = orghasfeature("gender restriction") 
lcl_orghasfeature_rentals                     = orghasfeature("rentals")
lcl_orghasfeature_classes_have_rentals        = orghasfeature("classes have rentals")

'Set up BODY onload
lcl_onload = ""
lcl_onload = lcl_onload & "setMaxLength();"

If lcl_orghasfeature_custom_registration_craigco Then 
	lcl_onload = lcl_onload & "enableDisableTeamRosterFields();"
End If 

If request("s") = "u" Then
	sLoadMsg = "displayScreenMsg('Your Changes Were Successfully Saved');"
End If 
If request("s") = "n" Then
	sLoadMsg = "displayScreenMsg('This Class/Event Was Successfully Created');"
End If

blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")
%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
 	<title>E-Gov Administration Console {Edit Class/Event}</title>

  	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />

	<style type="text/css">
		/*margin and padding on body element
		  can introduce errors in determining
		  element position and are not recommended;
		  we turn them off as a foundation for YUI
		  CSS treatments. */
		body {
			margin:0;
			padding:0;
			}
	</style>

	<link rel="stylesheet" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />

  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script src="../scripts/ajaxLib.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>
	<script src="../scripts/setfocus.js"></script>
	<script src="../scripts/isvaliddate.js"></script>
	<script src="../scripts/textareamaxlength.js"></script>

	<script>
	<!--
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("#screenMsg").html("*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;");
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() 
		{
			$("#screenMsg").html("");
		}

		function showDateSelection( iRowCount )
		{
			// check the start date and end date and display a message instead if they are missing
			if ($("#startdate").val() == '' || $("#enddate").val() == '')
			{
				if ($("#startdate").val() == '')
				{
					alert('The Start Date for this class is blank, so this reservation cannot be made.');
					return;
				}
				if ($("#enddate").val() == '')
				{
					alert('The End Date for this class is blank, so this reservation cannot be made.');
					return;
				}
			}
			// All is there so display the date selection page
			var iTimeId = $("#rentaltimeid" + iRowCount).val();
			var iRentalId = $("#rentalid" + iRowCount).val();
			location.href='../rentals/classdateselection.asp?rentalid=' + iRentalId + '&timeid=' + iTimeId;
		}

		function ClassChange() 
		{
			// Try to get a drop down of names
			//alert($("earlyregistrationclassseasonid").options[$("earlyregistrationclassseasonid").selectedIndex].value);
			doAjax('getseasonclasses.asp', 'classseasonid=' + $("#earlyregistrationclassseasonid").val(), 'UpdateClasses', 'get', '0');
		}

		function UpdateClasses( sResult )
		{
			$("#earlyclass").html(sResult);
		}

		function selectAll( selectBox, selectAll ) 
		{
			// have we been passed an ID
			if (typeof selectBox == "string") {
				selectBox = document.getElementById(selectBox);
			}
			// is the select box a multiple select box?
			if (selectBox.type == "select-multiple") {
				for (var i = 0; i < selectBox.options.length; i++) {
					selectBox.options[i].selected = selectAll;
				}
			}
		}

		function ValidatePrice( oRs )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oRs.value = removeSpaces(oRs.value);
			//Remove commas that would cause problems in validation
			oRs.value = removeCommas(oRs.value);

			// Validate the format of the price
			if (oRs.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oRs.value);
				if ( Ok )
				{
					oRs.value = format_number(Number(oRs.value),2);
				}
				else 
				{
					oRs.value = "";
					alert("Prices must be numbers in currency format or blank.\nPlease correct to continue.");
					tabView.set('activeIndex',5);
					setfocus(oRs);
					return false;
				}
			}
		}

		function NewTimeRow()
		{
			document.ClassForm.maxtimeid.value = parseInt(document.ClassForm.maxtimeid.value) + 1;
			document.ClassForm.maxtimedayid.value = parseInt(document.ClassForm.maxtimedayid.value) + 1;
			var tbl = document.getElementById("seriestime");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.ClassForm.maxtimeid.value);
			var row = tbl.insertRow(lastRow);

			var cellZero = row.insertCell(0);
			cellZero.className = 'ref';

			var e1 = document.createElement('input');
			e1.type = 'hidden';
			e1.id = 'timeid' + newRow;
			e1.name = 'timeid' + newRow;
			e1.value = '0';
			cellZero.appendChild(e1);

			var e2 = document.createElement('input');
			e2.type = 'text';
			e2.id = 'activity' + newRow;
			e2.name = 'activity' + newRow;
			e2.value = '';
			e2.size = 10;
			e2.maxLength = 10;
			cellZero.appendChild(e2);

			var cellOne = row.insertCell(1);
			cellOne.align = 'center';
			//instructor pick here
			var e01 = document.createElement('select');
			e01.name = 'instructorid' + newRow;
			e01.id = 'instructorid' + newRow;
			cellOne.appendChild(e01);
			var slength = $("#instructorid0 option").length;
			//alert(slength);
			var op;
			var newText; 
			for ( var s=0; s < slength; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.getElementById("instructorid0").options[s].text );
				op.appendChild( newText );
				op.setAttribute( 'value', document.getElementById("instructorid0").options[s].value );
				e01.appendChild(op);
			}

			var cellTwo = row.insertCell(2);
			cellTwo.align = 'center';
			var e3 = document.createElement('input');
			e3.type = 'text';
			e3.name = 'min' + newRow;
			e3.value = '';
			e3.size = 4;
			e3.maxLength = 5;
			cellTwo.appendChild(e3);

			var cell3 = row.insertCell(3);
			cell3.align = 'center';
			var e4 = document.createElement('input');
			e4.type = 'text';
			e4.name = 'max' + newRow;
			e4.value = '';
			e4.size = 4;
			e4.maxLength = 5;
			cell3.appendChild(e4);

			var cell4 = row.insertCell(4);
			cell4.align = 'center';
			newText = document.createTextNode( '0' );
			cell4.appendChild( newText );

			var cell5 = row.insertCell(5);
			cell5.align = 'center';
			var e5 = document.createElement('input');
			e5.type = 'text';
			e5.name = 'waitlistmax' + newRow;
			e5.value = '';
			e5.size = 4;
			e5.maxLength = 5;
			cell5.appendChild(e5);

			var cell15 = row.insertCell(6);
			cell15.align = 'center';

			var cell6 = row.insertCell(7);
			cell6.className = 'firstday';
			var e6 = document.createElement('input');
			e6.type = 'hidden';
			e6.name = 'timedayid' + newRow;
			e6.value = '0';
			cell6.appendChild(e6);

			var e7 = document.createElement('input');
			e7.type = 'checkbox';
			e7.name = 'su' + newRow;
			cell6.appendChild(e7);

			var cell7 = row.insertCell(8);
			var e8 = document.createElement('input');
			e8.type = 'checkbox';
			e8.name = 'mo' + newRow;
			cell7.appendChild(e8);

			var cell8 = row.insertCell(9);
			var e9 = document.createElement('input');
			e9.type = 'checkbox';
			e9.name = 'tu' + newRow;
			cell8.appendChild(e9);

			var cell9 = row.insertCell(10);
			var e10 = document.createElement('input');
			e10.type = 'checkbox';
			e10.name = 'we' + newRow;
			cell9.appendChild(e10);

			var cell10 = row.insertCell(11);
			var e11 = document.createElement('input');
			e11.type = 'checkbox';
			e11.name = 'th' + newRow;
			cell10.appendChild(e11);

			var cell11 = row.insertCell(12);
			var e12 = document.createElement('input');
			e12.type = 'checkbox';
			e12.name = 'fr' + newRow;
			cell11.appendChild(e12);

			var cell12 = row.insertCell(13);
			var e13 = document.createElement('input');
			e13.type = 'checkbox';
			e13.name = 'sa' + newRow;
			cell12.appendChild(e13);

			var cell13 = row.insertCell(14);
			cell13.align = 'center';
			var e14 = document.createElement('input');
			e14.type = 'text';
			e14.id = 'starttime' + newRow;
			e14.name = 'starttime' + newRow;
			e14.value = '';
			e14.size = 8;
			e14.maxLength = 7;
			cell13.appendChild(e14);

			var cell14 = row.insertCell(15);
			cell14.align = 'center';
			var e15 = document.createElement('input');
			e15.type = 'text';
			e15.id = 'endtime' + newRow;
			e15.name = 'endtime' + newRow;
			e15.value = '';
			e15.size = 8;
			e15.maxLength = 7;
			cell14.appendChild(e15);
		}
		
		function AddInstructor()
		{
			if (document.ClassForm.firstname.value == "")
			{
				alert("Please enter a first name for the new instructor.");
				document.ClassForm.firstname.focus();
				return;
			}
			if (document.ClassForm.lastname.value == "")
			{
				alert("Please enter a last name for the new instructor.");
				document.ClassForm.lastname.focus();
				return;
			}
			// Fire off Ajax routine
			doAjax('addinstructor.asp', 'firstname=' + document.ClassForm.firstname.value + '&lastname=' + document.ClassForm.lastname.value, 'IncludeInstructor', 'get', '0');
		}

		function IncludeInstructor( sNewInstructor )
		{
			// Process the Ajax CallBack by adding the new instructor to the list
			var InstructorPick = document.getElementById( 'instructorid' );
			var newOption = document.createElement( 'OPTION' );
			var newText = document.createTextNode( document.ClassForm.firstname.value + ' ' + document.ClassForm.lastname.value );
			newOption.appendChild( newText );
			newOption.setAttribute( 'value', sNewInstructor );
			InstructorPick.appendChild(newOption);
			
			// Add to the time rows
			var op;
			for (var s=0; s <= parseInt(document.ClassForm.maxtimeid.value); s++)
			{
				if ($("#instructorid"+s).length > 0)
				{
					op = document.createElement('OPTION');
					newText = document.createTextNode( document.ClassForm.firstname.value + ' ' + document.ClassForm.lastname.value );
					op.appendChild( newText );
					op.setAttribute( 'value', sNewInstructor );
					document.getElementById("instructorid"+s).appendChild(op);
				}
			}
			document.ClassForm.firstname.value = "";
			document.ClassForm.lastname.value = "";
			alert('Instructor Added');
		}

		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=ClassForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function doPicker(sFormField) 
		{
		  w = (screen.width - 350)/2;
		  h = (screen.height - 350)/2;
		  eval('window.open("imagepicker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=400,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function insertAtURL (textEl, text) 
		{
			if (textEl.createTextRange && textEl.caretPos) 
			{
				var caretPos = textEl.caretPos;
				caretPos.text =
				caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				text + ' ' : text;
			}
			else
				textEl.value  = text;

			$("#" + textEl.name + "pic").attr("src",text);
		}

		function openWin1(url, name) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 350)/2;
			popupWin = eval('window.open(url, name,"resizable,width=800,height=400,left=' + 100 + ',top=' + h + '")');
		}

		function openWin2(url, name) 
		{
			var w = (screen.width - 350)/2;
			var h = (screen.height - 550)/2;
			popupWin = eval('window.open(url, name,"resizable,width=820,height=600,scrollbars=1,left=' + 80 + ',top=' + h + '")');
		}

		function openWin(url) {
			var w = 900;
			var h = 400;
			var l = (screen.availWidth/2) - (w/2);
			var t = (screen.availHeight/2) - (h/2);
	 		popupWin = eval('window.open(url, "maintain_options", "resizable,scrollbars=1,width='+w+',height='+h+',left='+l+',top='+t+'")');
		}

		function CopyClass(iClassid)
		{
			if (confirm('Copy this to a new Class/Event?'))
			{
				//location.href='class_copyclass.asp?classid=' + iClassid + '&seasonid=' + something + '&bring=' + bring;
				document.copyForm.submit();
			}
		}

		function CopyToChild(iClassid)
		{
			if (confirm('Create a new Individual for this Series?'))
			{
				location.href='class_copyaschild.asp?classid=' + iClassid;
			}
		}

		function AddSingle( )
		{
			if (confirm('Add \'' + document.SingleForm.classid.options[document.SingleForm.classid.selectedIndex].text + '\' to the Series?'))
			{
				document.SingleForm.submit();
			}
		}

  function enableDisableTeamRosterFields() {
    lcl_disabled_options     = false;
    //lcl_disabled_label_color = "#000000";
    lcl_disabled_maintlinks  = false;

    if($("#displayrosterpublic").is(':checked') != true) {
       lcl_disabled_options     = true;
       //lcl_disabled_label_color = "#c0c0c0";
    } else {
       setupDefaultAccessoryOptions();
    }

    //Determine if default values are to be assigned
    if(lcl_disabled_options != true) {
       assignDefaultAccessoryOptions("TSHIRT");
       assignDefaultAccessoryOptions("PANTS");
    }

    document.getElementById("teamreg_TSHIRT_enabled").disabled      = lcl_disabled_options;
    document.getElementById("teamreg_PANTS_enabled").disabled       = lcl_disabled_options;
    document.getElementById("teamreg_GRADE_enabled").disabled       = lcl_disabled_options;
    document.getElementById("teamreg_COACH_enabled").disabled       = lcl_disabled_options;

    document.getElementById("teamreg_TSHIRT_inputtype").disabled    = lcl_disabled_options;
    document.getElementById("teamreg_PANTS_inputtype").disabled     = lcl_disabled_options;
    document.getElementById("teamreg_GRADE_inputtype").disabled     = lcl_disabled_options;

    document.getElementById("teamreg_TSHIRT_input_button").disabled = lcl_disabled_options;
    document.getElementById("teamreg_PANTS_input_button").disabled  = lcl_disabled_options;
    document.getElementById("teamreg_GRADE_input_button").disabled  = lcl_disabled_options;

    enableDisableMaintainButton("TSHIRT");
    enableDisableMaintainButton("PANTS");
    enableDisableMaintainButton("GRADE");

    //Now check the values in the dropdown list(s) and see if the maintain link is to be shown
    //if(lcl_disabled_options != true) {
    //   enableDisableMaintainButton("TSHIRT");
    //   enableDisableMaintainButton("PANTS");
    //}
  }

  function enableDisableMaintainButton(p_type) {
    var lcl_disabled_options = true;

    if($("#displayrosterpublic").is(':checked') == true) {
       //If the value in the dropdown list = "TEXT" then disable the "maintain options" button
       if($("#teamreg_" + p_type + "_inputtype").val() == "LOV") {
          lcl_disabled_options = false;
       }

    }

    document.getElementById("teamreg_" + p_type + "_input_button").disabled = lcl_disabled_options;

    //If the value in the dropdown list = "LOV" then check to see if default values need to be assigned.
    if(lcl_disabled_options != true) {
       assignDefaultAccessoryOptions(p_type);
    }
  }

  function setupDefaultAccessoryOptions() {
    var sParameter  = 'orgid='    + encodeURIComponent("<%=session("orgid")%>");
        sParameter += '&classid=' + encodeURIComponent("<%=iClassID%>");
       	sParameter += '&isAjaxRoutine=Y';

    doAjax('setupDefaultAccessoryOptions.asp', sParameter, '', 'post', '0');
  }

		function assignDefaultAccessoryOptions(p_type)	{
    if($("#displayrosterpublic").is(':checked') == true) {
       //if($("teamreg_" + p_type + "_enabled").checked == true) {
       var lcl_fieldvalue = document.getElementById('teamreg_' + p_type + '_enabled').value;

       if(lcl_fieldvalue != 'DISABLED') {
          if($("#teamreg_" + p_type + "_inputtype").val() != "TEXT") {

           		var sParameter = 'orgid='          + encodeURIComponent("<%=session("orgid")%>");
       		    sParameter    += '&classid='       + encodeURIComponent($("#classid").val());
       		    sParameter    += '&accessorytype=' + encodeURIComponent($("#teamreg_" + p_type + "_accessorytype").val());
       		    sParameter    += '&isAjaxRoutine=Y';

       		    doAjax('assignDefaultAccessoryOptions.asp', sParameter, '', 'post', '0');
          }
       }
    }
		}

		function ValidateForm()
		{
			var rege;
			var Ok;
			var iRow;
			var x;
			var regOK;
			var TimeRequired;

			// check the class name
			if (document.ClassForm.classname.value == "")
			{
				alert('Please enter a name.');
				document.ClassForm.classname.focus();
				return;
			}
			if (document.ClassForm.classname.value.length > 255)
			{
				alert('The classname is limited to 255 characters.\nPlease shorten this name.');
				document.ClassForm.classname.focus();
				return;
			}
			// check the class description
			if (document.ClassForm.classdescription.value == "")
			{
				tabView.set('activeIndex',0);
				alert('Please enter a description.');
				document.ClassForm.classdescription.focus();
				return;
			}

			var bHasAge = false; 

			// check the minimum age
			if (document.ClassForm.minage.value.length > 0)
			{
				bHasAge = true;
				rege = /^\d{0,2}\.?\d?$/;
				Ok = rege.test(document.ClassForm.minage.value);

				if (! Ok)
				{
					tabView.set('activeIndex',0);
					alert("The minimum age must be a number less than 100 with at most one decimal place, or be blank.");
					document.ClassForm.minage.focus();
					return;
				}
				else
				{
					if (parseFloat(document.ClassForm.minage.value) > 99.9)
					{
						tabView.set('activeIndex',0);
						alert("The minimum age must be a number less than 100 with at most one decimal place, or be blank.");
						document.ClassForm.minage.focus();
						return;
					}
				}
			}
			// check the maximum age
			if (document.ClassForm.maxage.value.length > 0)
			{
				bHasAge = true;
				rege = /^\d{0,2}\.?\d?$/;
				Ok = rege.test(document.ClassForm.maxage.value);

				if (! Ok)
				{
					tabView.set('activeIndex',0);
					alert("The maximum age must be a number less than 100 with at most one decimal place, or be blank.");
					document.ClassForm.maxage.focus();
					return;
				}
				else
				{
					if (parseFloat(document.ClassForm.maxage.value) > 99.9)
					{
						tabView.set('activeIndex',0);
						alert("The maximum age must be a number less than 100 with at most one decimal place, or be blank.");
						document.ClassForm.maxage.focus();
						return;
					}
				}
			}

			if (bHasAge)
			{
				// check the agecomparedate
				if (document.ClassForm.agecomparedate.value == "")
				{
					tabView.set('activeIndex',0);
					alert("Please enter an age comparison date");
					document.ClassForm.agecomparedate.focus();
					return;
				}
				else
				{
					//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
					//Ok = rege.test(document.ClassForm.agecomparedate.value);
					//if (! Ok)
					if (! isValidDate(document.ClassForm.agecomparedate.value))
					{
						tabView.set('activeIndex',0);
						alert("The age comparison date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
						document.ClassForm.agecomparedate.focus();
						return;
					}
				}
			}

			// check the length of the search key words
			if (document.ClassForm.searchkeywords.value.length > 1024)
			{
				tabView.set('activeIndex',0);
				alert("The maximum length of search keywords is 1024 characters. \nPlease make this smaller.");
				document.ClassForm.searchkeywords.focus();
				return;
			}

			// check the startdate
			if (document.ClassForm.startdate.value == "")
			{
				tabView.set('activeIndex',4);
				alert("Please enter a start date");
				document.ClassForm.startdate.focus();
				return;
			}
			else
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.startdate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.startdate.value))
				{
					tabView.set('activeIndex',4);
					alert("Start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.startdate.focus();
					return;
				}
			}

			// check the enddate
			if (document.ClassForm.enddate.value == "")
			{
				tabView.set('activeIndex',4);
				alert("Please enter an end date");
				document.ClassForm.enddate.focus();
				return;
			}
			else
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.enddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.enddate.value))
				{
					tabView.set('activeIndex',4);
					alert("End date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.enddate.focus();
					return;
				}
			}

			// check publish start date
			if (document.ClassForm.publishstartdate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.publishstartdate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.publishstartdate.value))
				{
					tabView.set('activeIndex',4);
					alert("Publication start date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.publishstartdate.focus();
					return;
				}
			}

			// check publish end date
			if (document.ClassForm.publishenddate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.publishenddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.publishenddate.value))
				{
					tabView.set('activeIndex',4);
					alert("Publication end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.publishenddate.focus();
					return;
				}
			}

			// check evaluation date
			if (document.ClassForm.evaluationdate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.evaluationdate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.evaluationdate.value))
				{
					tabView.set('activeIndex',4);
					alert("Evaluation date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.evaluationdate.focus();
					return;
				}
			}
			
			// check registration end date
			if (document.ClassForm.registrationenddate.value != "")
			{
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.ClassForm.registrationenddate.value);
				//if (! Ok)
				if (! isValidDate(document.ClassForm.registrationenddate.value))
				{
					tabView.set('activeIndex',4);
					alert("Registration end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.registrationenddate.focus();
					return;
				}
			}

			// Check pricing on the ones they checked
			for (var p = parseInt(document.ClassForm.minpricetypeid.value); p <= parseInt(document.ClassForm.maxpricetypeid.value); p++)
			{
				//alert($("pricetypeid" + p).value);
				// Does it exist
				if ($("#pricetypeid" + p).length > 0)
				{
					//alert($("pricetypeid" + p).checked);
					//Is is checked
					if($("#pricetypeid" + p).is(':checked') == true)
					{
						// Remove any extra spaces
						$("#amount" + p).val(removeSpaces($("#amount" + p).val()));
						//Remove commas that would cause problems in validation
						$("#amount" + p).val(removeCommas($("#amount" + p).val()));

						// Is the price formated correctly
						rege = /^\d+\.\d{2}$/;
						Ok = rege.test($("#amount" + p).val());
						if (! Ok)
						{
							tabView.set('activeIndex',5);
							alert("Selected prices cannot be blank and must be in currency format.");
							$("#amount" + p).focus();
							return;
						}
						//Is the instructor % formatted correctly
						rege = /^\d+$/;
						Ok = rege.test($("#instructorpercent" + p).val());
						if (! Ok)
						{
							tabView.set('activeIndex',5);
							alert("Selected instructor percentages cannot be blank and must be in the range of 0-100.");
							$("#instructorpercent" + p).focus();
							return;
						}
						if (parseInt($("#instructorpercent" + p).val()) > 100)
						{
							tabView.set('activeIndex',5);
							alert("Selected instructor percentages cannot be blank and must be in the range of 0-100.");
							$("#instructorpercent" + p).focus();
							return;
						}
						// check that there is a registration start date
						if (document.getElementById("registrationstartdate" + p).getAttribute("type") != 'hidden')
						{
							//alert($("registrationstartdate" + p).value);
							//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
							//Ok = rege.test($("registrationstartdate" + p).value);
							//if (! Ok)
							if (! isValidDate($("#registrationstartdate" + p).val()))
							{
								tabView.set('activeIndex',5);
								alert("Selected registration start dates cannot be blank and should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
								$("#registrationstartdate" + p).focus();
								return;
							}
						}
					}
				}
			}

			// Check the time rows.  If there is an activity and a time is entered, check that they are formatted right
			var timerege;
			var timeOk;
			for (var t = 0; t <= parseInt(document.ClassForm.maxtimeid.value); t++)
			{
				if (parseInt($("#timeid" + t).val()) != 0)
				{
					// If not marked for delete
					if ($("#delete" + t).is(':checked') == false)
					{
						if ($("#activity" + t).val() != '')
						{
							if ($("#starttime" + t).val() != '')
							{
								timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
								timeOk = timerege.test($("#starttime" + t).val());

								if (! timeOk)
								{
									tabView.set('activeIndex',6);
									alert("The start time must be formatted as HH:MM(AM|PM).");
									$("#starttime" + t).focus();
									return;
								}
							}
							if ($("#endtime" + t).val() != '')
							{
								timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
								timeOk = timerege.test($("#endtime" + t).val());

								if (! timeOk)
								{
									tabView.set('activeIndex',6);
									alert("The end time must be formatted as HH:MM(AM|PM).");
									$("#endtime" + t).focus();
									return;
								}
							}
						}
						else
						{
							// The activity number is blank.
							tabView.set('activeIndex',6);
							alert("The activity number cannot be blank for existing activities.");
							$("#activity" + t).focus();
							return;
						}
					}
				}
				else
				{
					// handle new activities
					if ($("#activity" + t).val() != '')
					{
						if ($("#starttime" + t).val() != '')
						{
							timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
							timeOk = timerege.test($("#starttime" + t).val());

							if (! timeOk)
							{
								tabView.set('activeIndex',6);
								alert("The start time must be formatted as HH:MM(AM|PM).");
								$("#starttime" + t).focus();
								return;
							}
						}
						if ($("#endtime" + t).val() != '')
						{
							timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
							timeOk = timerege.test($("#endtime" + t).val());

							if (! timeOk)
							{
								tabView.set('activeIndex',6);
								alert("The end time must be formatted as HH:MM(AM|PM).");
								$("#endtime" + t).focus();
								return;
							}
						}
					}
				}
			}

			// Check the Early Registration tab
			if($("#allowearlyregistration").is(':checked') == true)
			{
				// Check the Early Registration date
				if ($("#earlyregistrationdate").val() == '')
				{
					tabView.set('activeIndex',7);
					alert("You have selected to allow Early Registration without giving an Early Registration Start Date.\nPlease correct this and try saving again.");
					$("#earlyregistrationdate").focus();
					return;
				}
				else
				{
					// Check the date format
					if (! isValidDate($("#earlyregistrationdate").val()))
					{
						tabView.set('activeIndex',7);
						alert("The Early Registration Start Date must be a valid date (mm/dd/yyyy).\nPlease correct this and try saving again.");
						$("#earlyregistrationdate").focus();
						return;
					}
				}
				// Check that a class was picked
				if (document.getElementById("earlyregistrationclassid").type == 'hidden')
				{
					tabView.set('activeIndex',7);
					alert("You have selected to allow Early Registration without selecting a class.\nPlease correct this and try saving again.");
					$("#earlyregistrationclassseasonid").focus();
					return;
				}
			}

			//alert('OK');
			//return;

			// all is ok, so submit and save the changes
			document.ClassForm.submit();
		}

		function SetUpPage()
		{
			<%=lcl_onload%>
			<%=sLoadMsg%>
			$("#classname").focus();
		}

	//-->
	</script>
</head>
<body class="yui-skin-sam" onload="SetUpPage();">
 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">

	<!--BEGIN: PAGE TITLE-->
	<p>
		<span id="screenMsg"></span>
		<font size="+1"><strong>Edit Classes and Events</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

	<p>
		<input type="button" class="button" name="backBtn" value="<< Return to Class/Event List" onclick="location.href='class_list.asp'" /> &nbsp;
		<input type="button" class="button" name="update1" id="update1" value="Save Changes" onclick="ValidateForm();" /> &nbsp;
		<input type="button" class="button" name="rosterBtn" value="Rosters &amp; Registration" onclick="location.href='class_offerings.asp?classid=<%=iClassId%>'" />
	</p>
<%
	sSql = "SELECT "
	sSql = sSql & " classname, "
	sSql = sSql & " ISNULL(classseasonid,0) AS classseasonid, "
	sSql = sSql & " classdescription, "
	sSql = sSql & " ISNULL(startdate,'') AS startdate, "
	sSql = sSql & " C.publiccanonlyview, "
	sSql = sSql & " ISNULL(enddate,'') AS enddate, "
	sSql = sSql & " ISNULL(publishstartdate,'') AS publishstartdate, "
	sSql = sSql & " ISNULL(publishenddate,'') AS publishenddate, "
	sSql = sSql & " ISNULL(registrationstartdate,'') AS registrationstartdate, "
	sSql = sSql & " ISNULL(registrationenddate,'') AS registrationenddate, "
	sSql = sSql & " ISNULL(promotiondate,'') AS promotiondate, "
	sSql = sSql & " ISNULL(evaluationdate,'') AS evaluationdate, "
	sSql = sSql & " ISNULL(locationid,0) AS locationid, "
	sSql = sSql & " ISNULL(alternatedate,'') AS alternatedate, "
	sSql = sSql & " searchkeywords, "
	sSql = sSql & " externalurl, "
	sSql = sSql & " externallinktext, "
	sSql = sSql & " noenddate, "
	sSql = sSql & " isparent, "
	sSql = sSql & " C.classtypeid, "
	sSql = sSql & " ISNULL(parentclassid,0) AS parentclassid, "
	sSql = sSql & " statusid, "
	sSql = sSql & " cancelreason, "
	sSql = sSql & " ISNULL(imgurl,'') AS imgurl, "
	sSql = sSql & " ISNULL(imgalttag,'') AS imgalttag, "
	sSql = sSql & " optionid, "
	sSql = sSql & " ISNULL(pricediscountid,0) AS pricediscountid, "
	sSql = sSql & " ISNULL(minage,0) AS minage, "
	sSql = sSql & " ISNULL(maxage,99) AS maxage, "
	sSql = sSql & " mingrade, "
	sSql = sSql & " maxgrade, "
	sSql = sSql & " ISNULL(pocid,0) AS pocid, "
	sSql = sSql & " T.classtypename, "
	sSql = sSql & " ISNULL(membershipid,0) AS membershipid, "
	sSql = sSql & " ISNULL(supervisorid,0) AS supervisorid, "
	sSql = sSql & " notes, "
	sSql = sSql & " allowearlyregistration, "
	sSql = sSql & " earlyregistrationdate, "
	sSql = sSql & " ISNULL(earlyregistrationclassseasonid,0) AS earlyregistrationclassseasonid, "
	sSql = sSql & " ISNULL(earlyregistrationclassid,0) AS earlyregistrationclassid, "
	sSql = sSql & " ISNULL(minageprecisionid, 0) AS minageprecisionid, "
	sSql = sSql & " ISNULL(maxageprecisionid, 0) AS maxageprecisionid, "
	sSql = sSql & " ISNULL(agecomparedate,'') AS agecomparedate, "
	sSql = sSql & " displayrosterpublic, "
	'sSql = sSql & " ISNULL(teamreg_tshirt_enabled,1) AS teamreg_tshirt_enabled, "
	'sSql = sSql & " ISNULL(teamreg_pants_enabled,0) AS teamreg_pants_enabled, "
	sSql = sSql & " teamreg_tshirt_enabled, "
	sSql = sSql & " teamreg_pants_enabled, "
	sSql = sSql & " teamreg_grade_enabled, "
	sSql = sSql & " teamreg_coach_enabled, "
	sSql = sSql & " teamreg_pants_inputtype, "
	sSql = sSql & " teamreg_tshirt_inputtype, "
	sSql = sSql & " teamreg_grade_inputtype, "
	sSql = sSql & " showTerms, "
	sSql = sSql & " ISNULL(genderrestrictionid,0) AS genderrestrictionid,norefunds "
	sSql = sSql & " FROM egov_class C, "
	sSql = sSql &      " egov_class_type T "
	sSql = sSql & " WHERE T.classtypeid = C.classtypeid "
	sSql = sSql & " AND C.classid = " & iClassId
	sSql = sSql & " AND C.orgid = "   & session("orgid")

	set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
'------------------------------------------------------------------------------
	If Not oRs.EOF Then 
'------------------------------------------------------------------------------
		'Setup variables -------------------------------------------------------
		lcl_minage = ""
		lcl_maxage = ""
		lcl_agecomparedate = ""
		lcl_selected_displayrosterpublic = ""

		lcl_classname_label = ""
		lcl_gotoparent_button = ""
		cCopyText = ""

		If clng(oRs("genderrestrictionid")) > clng(0) Then 
			iGenderRestrictionId = oRs("genderrestrictionid")
		Else
			iGenderRestrictionId = GetGenderNotRequiredId( )
		End If 

		'Determine if this is an "individual" or "series" class
		If clng(oRs("parentclassid")) > 0 Then 
			lcl_classname_label   = "&nbsp;Individual"
			lcl_gotoparent_button = "&nbsp;<input type=""button"" class=""button"" name=""gotoparent"" id=""gotoparent"" value=""Edit Series"" onclick=""javascript:location.href='edit_class.asp?classid=" & oRs("parentclassid") & "';"" />"
			cCopyText             = "Copy to New Series Individual"
		Else 
			If clng(oRs("classtypeid")) = 1 Then 
				cCopyText       = "Copy to New Series"
				bIsSeriesParent = True
			Else 
				cCopyText = "Copy to New Class/Event"
			End If 
		End If 

		response.write "<form name=""ClassForm"" id=""ClassForm"" accept-charset=""UTF-8"" method=""post"" action=""update_class.asp"">"
		response.write "<input type=""hidden"" name=""classid"" id=""classid"" value=""" & iClassId & """ />"
		response.write "<p>"
		response.write "Name: <input type=""text"" name=""classname"" id=""classname"" value=""" & Replace(oRs("classname"),"""","&quot;") & """ size=""50"" maxlength=""50"" />"
		response.write "&nbsp; <strong>This is a " & oRs("classtypename") & lcl_classname_label & " Class/Event </strong>"
		response.write lcl_gotoparent_button
		response.write "</p>"
		response.write "<p>"
		response.write "<div id=""cancelreason"">" & oRs("cancelreason") & "</div>"
		response.write "Status: <strong>" & GetStatusName(oRs("statusid")) & "</strong>&nbsp;&nbsp;"

		If clng(oRs("statusid")) = 1 Then 
			response.write "<input type=""button"" class=""button"" name=""cancel"" id=""cancel"" value=""Cancel Class/Event"" onclick=""location.href='class_cancel.asp?classid=" & iClassId & "';"" />&nbsp;"
		Else 
			response.write "<input type=""button"" class=""button"" name=""activate"" id=""activate"" value=""Activate Class/Event"" onclick=""location.href='class_changestatus.asp?classid=" & iClassId & "&statusid=1';"" />&nbsp;"
		End If 

		response.write "</p>"
		response.write "<p>"
		response.write "Season: "

		ShowClassSeasonFilterPicks oRs("classseasonid") 'In class_global_functions.asp

		response.write "</p>"
%>
	<div id="demo" class="yui-navset widetabnav">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Information</em></a></li>
			<li><a href="#tab2"><em>Categories</em></a></li>
			<li><a href="#tab3"><em>Waivers</em></a></li>
			<li><a href="#tab4"><em>Instructors</em></a></li>
			<li><a href="#tab5"><em>Dates</em></a></li>
			<li><a href="#tab6"><em>Purchasing</em></a></li>
			<li><a href="#tab7"><em>Occurs</em></a></li>
			<li><a href="#tab8"><em>Early Registration</em></a></li>
<%			If lcl_orghasfeature_rentals And lcl_orghasfeature_classes_have_rentals Then	%>
				<li><a href="#tab9"><em>Rentals</em></a></li>
<%			End If								%>
		 </ul>            
	<div class="yui-content widetabs">

  		<div id="tab1"> <!-- General Information -->
		   <p>
				Description:<br />
				<textarea name="classdescription" id="classdescription" maxlength="6000" wrap="soft"><%=oRs("classdescription")%></textarea>
 			</p>
 			<p>
				Image:
				<input type="<%if blnHasWP then %>hidden<%else%>text<%end if%>" name="imgurl" class="imageurl" id="imgurl" value="<%=oRs("imgurl")%>" size="50" maxlength="255" style="display:block" />
				<img src="<%=oRs("imgurl")%>" id="imgurlpic" name="imgurlpic" border="0" alt="<%=oRs("imgalttag")%>" align="middle" width="180" height="180"  onerror="this.src = '../images/placeholder.png';" />
				<% if blnHasWP then %>
					<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'imgurl');" />
				<% else%>
				<input type="button" class="button" value="Browse..." onclick="javascript:doPicker('ClassForm.imgurl');" />
				&nbsp; &nbsp; <input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
				<% end if %>
				<br /><span id="imgalttag">Image Alt Tag:</span> <input type="text" name="imgalttag" value="<%=oRs("imgalttag")%>" size="50" maxlength="255" />
				<br /><span id="imgalttabdesc"> The Image Alt Tag is a description used for ADA compliance.</span>
			</p>
			<p>
				Search Keywords:<br />
				<textarea name="searchkeywords" id="searchkeywords" maxlength="1024" wrap="soft"><%=oRs("searchkeywords")%></textarea>
			</p>
      <%
		      response.write "<p>"

        'Minimum Age -----------------------------------------------------------
		      response.write "Minimum Age: "

        If clng(oRs("minage")) <> 0 Then 
			        lcl_minage = oRs("minage")
        End If 

        response.write "<input type=""text"" name=""minage"" id=""minage"" value=""" & lcl_minage & """ size=""4"" maxlength=""4"" /> &nbsp;"

        ShowAgeCheckPrecision oRs("minageprecisionid"), "minageprecisionid"

        response.write "&nbsp;&nbsp;"

        'Maximum Age -----------------------------------------------------------
        response.write "Maximum Age: "

        If clng(oRs("maxage")) <> 99 Then 
           lcl_maxage = oRs("maxage")
        End If 

        response.write "<input type=""text"" name=""maxage"" id=""maxage"" value=""" & lcl_maxage & """ size=""4"" maxlength=""4"" /> &nbsp;"

        ShowAgeCheckPrecision oRs("maxageprecisionid"), "maxageprecisionid"

        response.write "</p>"
        response.write "<p>"

        'Age Compare to Date ---------------------------------------------------
        If oRs("agecomparedate") <> "1/1/1900" Then 
           lcl_agecomparedate = oRs("agecomparedate")
        End If 

        response.write "Check registrant age against this date: "
        response.write "<input type=""text"" maxlength=""10"" class=""datefield"" name=""agecomparedate"" id=""agecomparedate"" value=""" & lcl_agecomparedate & """ />&nbsp;"
        response.write "<span class=""calendarimg"" style=""cursor:hand;"">"
        response.write "<img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('agecomparedate');"" />"
        response.write "</span>"
        response.write "</p>"

		' Gender Restriction Picks
		If bHasGenderRestrictions Then		
			response.write vbcrlf & "<p>"
			response.write vbcrlf & "Gender Restriction: " 
			ShowGenderRestrictions iGenderRestrictionId 
			response.write vbcrlf & "</p>"
		Else								
			response.write vbcrlf & "<input type=""hidden"" name=""genderrestrictionid"" value=""" & GetGenderNotRequiredId( ) & """ />"
		End If								

        response.write "</p>"
        response.write "<div class=""classeditrightbuttons"">"
        response.write "<input type=""button"" class=""assignbuttons"" name=""locationmgr"" id=""locationmgr"" value=""Manage Locations"" onclick=""openWin2('location_mgmt.asp','_blank')"" />"
        response.write "</div>"

        'Location --------------------------------------------------------------
        response.write "Location: "

        ShowLocationPicks oRs("locationid")

        response.write "</p>"
        response.write "<p>"
        response.write "Point of Contact: "

        ShowPOCPicks oRs("pocid"), session("orgid")

        response.write "</p>"

        'Supervisor ------------------------------------------------------------
        If lcl_orghasfeature_class_supervisors Then 
           response.write "<p>"
           response.write "Supervisor: "

           ShowSupervisorPicks oRs("supervisorid")  'In class_global_functions.asp

           response.write "</p>"
        Else 
           response.write "<input type=""hidden"" name=""supervisorid"" id=""supervisorid"" value=""0"" />"
        End If 

       'External URL and Text -------------------------------------------------
        response.write vbcrlf & "<p>"
        response.write "External URL: "
        response.write "<input type=""text"" name=""externalurl"" id=""externalurl"" value=""" & oRs("externalurl") & """ size=""50"" maxlength=""255"" />"
        response.write "<br />"
        response.write "External URL Text: "
        response.write "<input type=""text"" name=""externallinktext"" id=""externallinktext"" value=""" & oRs("externallinktext") & """ size=""50"" maxlength=""255"" />"
        response.write "</p>"

       'Receipt Notes ---------------------------------------------------------
        response.write "<p>"
        response.write "   Receipt Notes:<br />"
        response.write "   <textarea name=""notes"" id=""receiptnotes"" maxlength=""1024"" wrap=""soft"">" & oRs("notes") & "</textarea>"
        response.write "</p>"

       'Display Team Roster Options to Public (Craig, CO custom request) ------
        If lcl_orghasfeature_custom_registration_craigco Then 
           if oRs("displayrosterpublic") then
              lcl_selected_displayrosterpublic = " checked=""checked"""
              'lcl_checked_tshirt               = ""
              'lcl_checked_pants                = ""

             'Check to see if the tshirt/pants checkbox is enabled/disabled
              'if oRs("teamreg_tshirt_enabled") then
              '   lcl_checked_tshirt = " checked=""checked"""
              'end if

              'if oRs("teamreg_pants_enabled") then
              '   lcl_checked_pants = " checked=""checked"""
              'end if

             'Determine which value to select in the tshirt enabled dropdown list
              'lcl_teamreg_tshirt_enabled = ""

              'if oRs("teamreg_tshirt_enabled") <> "" then
              '   lcl_teamreg_tshirt_enabled = ucase(oRs("teamreg_tshirt_enabled"))
              'end if

              'if lcl_teamreg_tshirt_enabled = "INTERNAL ONLY" then
              '   lcl_selected_tshirt_enabled_both         = ""
              '   lcl_selected_tshirt_enabled_internalonly = " selected=""selected"""
              '   lcl_selected_tshirt_enabled_disabled     = ""
              'elseif lcl_teamreg_tshirt_enabled = "DISABLED" then
              '   lcl_selected_tshirt_enabled_both         = ""
              '   lcl_selected_tshirt_enabled_internalonly = ""
              '   lcl_selected_tshirt_enabled_disabled     = " selected=""selected"""
              'else
              '   lcl_selected_tshirt_enabled_both         = " selected=""selected"""
              '   lcl_selected_tshirt_enabled_internalonly = ""
              '   lcl_selected_tshirt_enabled_disabled     = ""
              'end if

             'Determine which value to select in the pants enabled dropdown list
              'lcl_teamreg_pants_enabled = ""

              'if oRs("teamreg_pants_enabled") <> "" then
              '   lcl_teamreg_pants_enabled = ucase(oRs("teamreg_pants_enabled"))
              'end if

              'if lcl_teamreg_pants_enabled = "INTERNAL ONLY" then
              '   lcl_selected_pants_enabled_both         = ""
              '   lcl_selected_pants_enabled_internalonly = " selected=""selected"""
              '   lcl_selected_pants_enabled_disabled     = ""
              'elseif lcl_teamreg_pants_enabled = "DISABLED" then
              '   lcl_selected_pants_enabled_both         = ""
              '   lcl_selected_pants_enabled_internalonly = ""
              '   lcl_selected_pants_enabled_disabled     = " selected=""selected"""
              'else
              '   lcl_selected_pants_enabled_both         = " selected=""selected"""
              '   lcl_selected_pants_enabled_internalonly = ""
              '   lcl_selected_pants_enabled_disabled     = ""
              'end if

             'Determine which value to select in the tshirt dropdown list
              'if oRs("teamreg_tshirt_inputtype") = "TEXT" then
              '   lcl_selected_tshirt_type_lov  = ""
              '   lcl_selected_tshirt_type_text = " selected=""selected"""
              'else
              '   lcl_selected_tshirt_type_lov  = " selected=""selected"""
              '   lcl_selected_tshirt_type_text = ""
              'end if

             'Determine which value to select in the pants dropdown list
              'if oRs("teamreg_pants_inputtype") = "TEXT" then
              '   lcl_selected_pants_type_lov  = ""
              '   lcl_selected_pants_type_text = " selected=""selected"""
              'else
              '   lcl_selected_pants_type_lov  = " selected=""selected"""
              '   lcl_selected_pants_type_text = ""
              'end if

             'Determine which value to select in the pants dropdown list
              'if oRs("teamreg_grade_inputtype") = "TEXT" then
              '   lcl_selected_grade_type_lov  = ""
              '   lcl_selected_grade_type_text = " selected=""selected"""
              'else
              '   lcl_selected_grade_type_lov  = " selected=""selected"""
              '   lcl_selected_grade_type_text = ""
              'end if

              'Determine if the default, dropdown list, options need to be set.
              '  NOTE: ONLY if NO values have been assigned to the dropdown list(s) then
              '        assign the default options to this field for this class.
              'setupDefaultAccessoryOptions session("orgid"), iClassId, "TSHIRT", "Y"
              'setupDefaultAccessoryOptions session("orgid"), iClassId, "TPANTS", "Y"

           end if

          'Check for an "edit display" for the label
           lcl_label_tshirt = "T-Shirt"
           lcl_label_pants  = "Pants"
           lcl_label_grade  = "Grade"
           lcl_label_coach  = "Coach"

           if orgHasDisplay(session("orgid"),"class_teamregistration_tshirt_label") then
              lcl_label_tshirt = getOrgDisplay(session("orgid"),"class_teamregistration_tshirt_label")
           end if

'BEGIN: SCRIPT TO USE WHEN CHANGING COLUMNS ON EGOV_CLASS ---------------------
'sSql = "SELECT classid, teamreg_tshirt_enabled, teamreg_pants_enabled "
'sSql = sSql & " FROM egov_class "
'sSql = sSql & " order by classid "

'	Set oGetExistingInfo = Server.CreateObject("ADODB.Recordset")
'	oGetExistingInfo.Open sSql, Application("DSN"), 0, 1

' if not oGetExistingInfo.eof then
'    do while not oGetExistingInfo.eof

'       if oGetExistingInfo("teamreg_tshirt_enabled") then
'          lcl_current_tshirt_value = "1"
'          lcl_new_tshirt_value     = "BOTH"
'       else
'          lcl_current_tshirt_value = "0"
'          lcl_new_tshirt_value     = "DISABLED"
'       end if

'       if oGetExistingInfo("teamreg_pants_enabled") then
'          lcl_current_pants_value = "1"
'          lcl_new_pants_value     = "BOTH"
'       else
'          lcl_current_pants_value = "0"
'          lcl_new_pants_value     = "DISABLED"
'       end if

'       sSqlu = "UPDATE egov_class SET "
'       sSqlu = sSqlu & " teamreg_tshirt_enabled_new = '"     & lcl_new_tshirt_value     & "', "
'       sSqlu = sSqlu & " teamreg_pants_enabled_new = '"      & lcl_new_pants_value      & "', "
'       sSqlu = sSqlu & " teamreg_tshirt_enabled_original = " & lcl_current_tshirt_value & ", "
'       sSqlu = sSqlu & " teamreg_pants_enabled_original = "  & lcl_current_pants_value
'       sSqlu = sSqlu & " WHERE classid = " & oGetExistingInfo("classid")
'response.write sSqlu & "<br /><br />" & vbcrlf
'      	set oUpdateExistingInfo = Server.CreateObject("ADODB.Recordset")
'     	 oUpdateExistingInfo.Open sSqlu, Application("DSN"), 0, 1

'       set oUpdateExistingInfo = nothing

'       oGetExistingInfo.movenext
'    loop
' end if

 'oGetExistingInfo.close
 'set oGetExistingInfo = nothing
'END: SCRIPT TO USE WHEN CHANGING COLUMNS ON EGOV_CLASS -----------------------

           response.write "<p>" & vbcrlf
           response.write "<fieldset class=""fieldset"">" & vbcrlf
           response.write "  <legend>Team Registration Options:&nbsp;</legend>" & vbcrlf
           response.write "  <p>" & vbcrlf
           response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"" style=""width:600px"">" & vbcrlf
           response.write "    <tr>" & vbcrlf
           response.write "        <td>" & vbcrlf
           response.write "            <input type=""checkbox"" name=""displayrosterpublic"" id=""displayrosterpublic"" value=""on"" onclick=""enableDisableTeamRosterFields();""" & lcl_selected_displayrosterpublic & " />" & vbcrlf
           response.write "        </td>" & vbcrlf
           response.write "        <td colspan=""2"">Display Team Roster Options to Public</td>" & vbcrlf
           response.write "    </tr>" & vbcrlf
           response.write "    <tr>" & vbcrlf
           response.write "        <td colspan=""3"">&nbsp;</td>" & vbcrlf
           response.write "    </tr>" & vbcrlf

           response.write "    <tr>" & vbcrlf
           response.write "        <td>&nbsp;</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"" align=""right"">" & lcl_label_tshirt & " Input Options:</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"">" & vbcrlf
                                       lcl_fieldtype                = "TSHIRT"
                                       lcl_show_input_fields_tshirt = true

                                       setupRosterOptions iClassID, _
                                                          lcl_fieldtype, _
                                                          lcl_label_tshirt, _
                                                          oRs("teamreg_tshirt_enabled"), _
                                                          lcl_show_input_fields_tshirt, _
                                                          oRs("teamreg_tshirt_inputtype")
           response.write "        </td>" & vbcrlf
           response.write "    </tr>" & vbcrlf
           response.write "    <tr>" & vbcrlf
           response.write "        <td>&nbsp;</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"" align=""right"">Pants Input Options:</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"">" & vbcrlf
                                       lcl_fieldtype               = "PANTS"
                                       lcl_show_input_fields_pants = true

                                       setupRosterOptions iClassID, _
                                                          lcl_fieldtype, _
                                                          lcl_label_pants, _
                                                          oRs("teamreg_pants_enabled"), _
                                                          lcl_show_input_fields_pants, _
                                                          oRs("teamreg_pants_inputtype")
           response.write "        </td>" & vbcrlf
           response.write "    </tr>" & vbcrlf

           response.write "    <tr>" & vbcrlf
           response.write "        <td>&nbsp;</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"" align=""right"">Grade Input Options:</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"">" & vbcrlf
                                       lcl_fieldtype               = "GRADE"
                                       lcl_show_input_fields_grade = true

                                       setupRosterOptions iClassID, _
                                                          lcl_fieldtype, _
                                                          lcl_label_grade, _
                                                          oRs("teamreg_grade_enabled"), _
                                                          lcl_show_input_fields_grade, _
                                                          oRs("teamreg_grade_inputtype")
           response.write "        </td>" & vbcrlf
           response.write "    </tr>" & vbcrlf
           response.write "    <tr>" & vbcrlf
           response.write "        <td>&nbsp;</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"" align=""right"">Coach Input Options:</td>" & vbcrlf
           response.write "        <td class=""noWrapTD"">" & vbcrlf
                                       lcl_fieldtype               = "COACH"
                                       lcl_show_input_fields_coach = false
                                       lcl_inputtype_coach         = ""

                                       setupRosterOptions iClassID, _
                                                          lcl_fieldtype, _
                                                          lcl_label_coach, _
                                                          oRs("teamreg_coach_enabled"), _
                                                          lcl_show_input_fields_coach, _
                                                          lcl_inputtype_coach
           response.write "        </td>" & vbcrlf
           response.write "    </tr>" & vbcrlf
           'response.write "    <tr>" & vbcrlf
           'response.write "        <td>&nbsp;</td>" & vbcrlf
           'response.write "        <td align=""right"">" & lcl_label_tshirt & " Input Options:</td>" & vbcrlf
           'response.write "        <td>" & vbcrlf
           'response.write "            <input type=""hidden"" name=""teamreg_TSHIRT_accessorytype"" id=""teamreg_TSHIRT_accessorytype"" value=""TSHIRT"" />" & vbcrlf
           'response.write "            <input type=""checkbox"" name=""teamreg_TSHIRT_enabled"" id=""teamreg_TSHIRT_enabled"" value=""on"" onclick=""assignDefaultAccessoryOptions('TSHIRT')""" & lcl_checked_tshirt & " />&nbsp;<span id=""teamreg_TSHIRT_enabled_label"">Enabled</span>&nbsp;" & vbcrlf
           'response.write "            <select name=""teamreg_TSHIRT_enabled"" id=""teamreg_TSHIRT_enabled"">" & vbcrlf
           'response.write "              <option value=""BOTH"""          & lcl_selected_tshirt_enabled_both         & ">Both</option>" & vbcrlf
           'response.write "              <option value=""INTERNAL ONLY""" & lcl_selected_tshirt_enabled_internalonly & ">Internal Only</option>" & vbcrlf
           'response.write "              <option value=""DISABLED"""      & lcl_selected_tshirt_enabled_disabled     & ">Disabled</option>" & vbcrlf
           'response.write "            </select>" & vbcrlf
           'response.write "            <select name=""teamreg_TSHIRT_inputtype"" id=""teamreg_TSHIRT_inputtype"" onchange=""enableDisableMaintainButton('TSHIRT');"">" & vbcrlf
           'response.write "              <option value=""LOV"""  & lcl_selected_tshirt_type_lov  & ">Drop Down List</option>" & vbcrlf
           'response.write "              <option value=""TEXT""" & lcl_selected_tshirt_type_text & ">Input Text Field</option>" & vbcrlf
           'response.write "            </select>&nbsp;" & vbcrlf
           'response.write "            <input type=""button"" name=""teamreg_TSHIRT_input_button"" id=""teamreg_TSHIRT_input_button"" value=""Maintain " & lcl_label_tshirt & " Options"" class=""button"" onclick=""openWin('class_accessoryoptions_list.asp?classid=" & iClassId & "&atype=TSHIRT');"" />" & vbcrlf
           'response.write "        </td>" & vbcrlf
           'response.write "    </tr>" & vbcrlf
           'response.write "    <tr>" & vbcrlf
           'response.write "        <td>&nbsp;</td>" & vbcrlf
           'response.write "        <td align=""right"">Pants Input Options:</td>" & vbcrlf
           'response.write "        <td>" & vbcrlf
           'response.write "            <input type=""hidden"" name=""teamreg_PANTS_accessorytype"" id=""teamreg_PANTS_accessorytype"" value=""PANTS"" />" & vbcrlf
           'response.write "            <input type=""checkbox"" name=""teamreg_PANTS_enabled"" id=""teamreg_PANTS_enabled"" value=""on"" onclick=""assignDefaultAccessoryOptions('PANTS')""" & lcl_checked_pants & ">&nbsp;<span id=""teamreg_PANTS_enabled_label"">Enabled</span>&nbsp;" & vbcrlf
           'response.write "            <select name=""teamreg_PANTS_enabled"" id=""teamreg_PANTS_enabled"">" & vbcrlf
           'response.write "              <option value=""BOTH"""          & lcl_selected_pants_enabled_both         & ">Both</option>" & vbcrlf
           'response.write "              <option value=""INTERNAL ONLY""" & lcl_selected_pants_enabled_internalonly & ">Internal Only</option>" & vbcrlf
           'response.write "              <option value=""DISABLED"""      & lcl_selected_pants_enabled_disabled     & ">Disabled</option>" & vbcrlf
           'response.write "            </select>" & vbcrlf
           'response.write "            <select name=""teamreg_PANTS_inputtype"" id=""teamreg_PANTS_inputtype"" onchange=""enableDisableMaintainButton('PANTS');"">" & vbcrlf
           'response.write "              <option value=""LOV"""  & lcl_selected_pants_type_lov  & ">Drop Down List</option>" & vbcrlf
           'response.write "              <option value=""TEXT""" & lcl_selected_pants_type_text & ">Input Text Field</option>" & vbcrlf
           'response.write "            </select>&nbsp;" & vbcrlf
           'response.write "            <input type=""button"" name=""teamreg_PANTS_input_button"" id=""teamreg_PANTS_input_button"" value=""Maintain Pants Options"" class=""button"" onclick=""openWin('class_accessoryoptions_list.asp?classid=" & iClassId & "&atype=PANTS');"" />" & vbcrlf
           'response.write "            <span id=""teamreg_PANTS_input_button"">[Maintain PANTS Options]</span>"
           'response.write "        </td>" & vbcrlf
           'response.write "    </tr>" & vbcrlf
           response.write "  </table>" & vbcrlf
           response.write "  </p>" & vbcrlf
           response.write "</fieldset>" & vbcrlf
           response.write "</p>" & vbcrlf
        End If 
      %>
			</div>

			<div id="tab2"> <!-- Categories -->
				<p class="classeditdescription">
					Select the catagories you wish this class/event to show up under on the public pages. 
				</p>
				<div style="display: block;">
					<div>
						<div class="classeditrightbuttons">
							<input type="button" class="assignbuttons" name="categorymgr" value="Manage Categories" onclick="openWin2('category_mgmt.asp','_blank')" /> 
						</div>
						<% ShowCategories iClassId, session("orgid") %>
					</div>
				</div>
			</div>

<%
			 'Determine if the "Show Terms" checkbox is checked.
			  If oRs("showTerms") Then 
				 lcl_selected_showTerms = " checked=""checked"""
			  Else 
				 lcl_selected_showTerms = ""
			  End If 
%>

			<div id="tab3"> <!-- Waivers -->
  			<div>
				<div class="classeditrightbuttons">
					<input type="button" class="assignbuttons" name="waivermgr" value="Manage Waivers" onclick="openWin2('class_waivers.asp','_blank')" /><br />

<%
				'Check to see if any terms exist
				 If checkWaiversExist(session("orgid"),"TERM") > 0 Then 
					response.write "<input type=""checkbox"" name=""showTerms"" id=""showTerms"" value=""on""" & lcl_selected_showTerms & " /> Show Terms"
				 Else 
					response.write "<input type=""hidden"" name=""showTerms"" id=""showTerms"" value=""on"" />"
				 End If 
%>
				</div>
					<% ShowWaiverPicks iClassId %>
				<br /><input type="button" class="assignbuttons" name="NoRss" value="Clear Selection" onclick="selectAll(document.getElementById('waiverid'),false)" />
			</div>
				
  			<div id="waivernote">
				Note: To add new waivers, click on Manage Waivers, and create the waiver.
				The new waiver will not appear in this list until after you save your changes to this page.
  			</div>
			</div>

			<div id="tab4"> <!-- Instructors -->
				<p class="classeditdescription">
					The instructors selected here are only for display purposes. 
				</p>
				<div>
					<div class="classeditrightbuttons">
						<fieldset class="edit"><legend><strong> New Instructor </strong></legend>
							<table id="newinstructor" cellspacing="0" cellpadding="0" border="0">
								<tr>
									<th align="center">First Name</th>
									<th align="center">Last Name</th>
									<th>&nbsp;</th>
								</tr>
								<tr>
								   	<td align="center" valign="top"><input type="text" id="instructorfirstname" name="firstname" value="" size="27" maxlength="25" /></td>
									   <td align="center" valign="top"><input type="text" id="instructorlastname" name="lastname" value="" size="27" maxlength="25" /></td>
									   <td align="center" valign="bottom"><input type="button" class="assignbuttons" name="addinstructors" value="Add Instructor" onclick="AddInstructor();" /></td>
								</tr>
								<tr>
   									<td colspan="3">Remember to add details for any instructors you add after you save this class</td>
								</tr>
							</table>
						</fieldset>
					</div>
					<% ShowInstructorPicks iClassId %>
					<br />(select for display only)
					<br /><input type="button" class="assignbuttons" name="NoRss" value="Clear Selection" onclick="selectAll(document.getElementById('instructorid'),false)" />
				</div>
			</div>

			<div id="tab5"> <!-- Critical Dates -->
				<p class="classeditdescription">
					These dates tell the public when the class starts and ends. They control when the class/event is 
					visible to the public and when they can register. The registration start date here is just for display 
					purposes. The actual registration start is on the Purchasing tab and is tied to the pricing types offered.
				</p>
			  <p>
				 <table id="criticaldates" border="0" cellpadding="1" cellspacing="3">
					  <tr>
      					<td align="right">Class/Event Starts:</td>
						<td><input type="text" maxlength="10" class="datefield" id="startdate" name="startdate" value="<%	If oRs("startdate") <> "1/1/1900" Then 
																												response.write oRs("startdate")
																											End If %>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
					    </td>																							
      					<td align="right">Class/Event Ends:</td><td><input type="text" maxlength="10" class="datefield" id="enddate" name="enddate" value="<%	If oRs("enddate") <> "1/1/1900" Then 
																																						response.write oRs("enddate")
																																					End If %>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
      					</td>
   					</tr>
					<tr>
						<td align="right">Publication Starts:</td>
						<td>
							<input type="text" maxlength="10" class="datefield" name="publishstartdate" value="<%	If oRs("publishstartdate") <> "1/1/1900" Then 
																														response.write oRs("publishstartdate")
																													End If %>" />
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishstartdate');" /></span>
						</td>
						<td align="right">Publication Ends:</td>
						<td>
							<input type="text" maxlength="10" class="datefield" name="publishenddate" value="<%	If oRs("publishenddate") <> "1/1/1900" Then 
																													response.write oRs("publishenddate")
																												End If %>" />
							<a href="#criticaldates"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishenddate');" /></a>
						</td>
					</tr>
   					<tr>
			      		<td align="right">Registration Starts:<br />(for display only)</td>
						<td><input type="text" maxlength="10" class="datefield" name="registrationstartdate" value="<%	If oRs("registrationstartdate") <> "1/1/1900" Then 
																															response.write oRs("registrationstartdate")
																														End If %>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationstartdate');" /></span>
						</td>
      					<td align="right">Registration Ends:</td>
						<td><input type="text" maxlength="10" class="datefield" name="registrationenddate" value="<%	If oRs("registrationenddate") <> "1/1/1900" Then 
																															response.write oRs("registrationenddate")
																														End If %>" />&nbsp;
							<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationenddate');" /></span>
        				</td>
					</tr>
					<tr>
						<td align="right">Send Evaluation:</td>
						<td><input type="text" maxlength="10" class="datefield" name="evaluationdate" value="<% If oRs("evaluationdate") <> "1/1/1900" Then 
																													response.write oRs("evaluationdate")
																												End If %>" />&nbsp;
						<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('evaluationdate');" /></span>
						</td>
      					<td>&nbsp;</td>
   					</tr>
				</table>
			  </p>
			</div>

			<div id="tab6"> <!-- Purchasing -->
				<p class="classeditdescription">
					On this tab you set up the pricing types that are offered for this class/event. Depending on what you offer,
					pricing can vary based on residency, or membership. Various fees can be also added seperately to the total
					price charged. The registration dates entered here control when the public can start registation based on 
					how they match the pricing types you have selected.
				</p>
				<p>
					<strong>Pricing:</strong><br /><br />
					<table id="pricingtable" border="0" cellpadding="0" cellspacing="0">
						<tr>
							<th>Type</th>
           <%
							If lcl_orghasfeature_gl_accounts Then 
								response.write "<th>Account</th>"
							End If 
           %>
			      			<th>Price</th>
           <%
							If ClassCanNeedMemberships() Then  'In class_global_functions.asp
								response.write "<th>Membership</th>"
							Else 
								response.write "<th>&nbsp;</th>"
							End If 
           %>
  				    		<th>Instructor %</th>
							<th>Registration Starts</th>
						</tr>

  					<%	GetPricing iClassId, Session("OrgID"), oRs("membershipid"), oRs("classseasonid"), iMin, iMax %>

					</table>
<%
  					response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMin & """ />"
					response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMax & """ />"
%>
				</p>
				<p>
					Requires: <% ShowRegistrationPicks oRs("optionid") %>
				</p>
				<p>
					<input type="checkbox" id="publiccanonlyview" name="publiccanonlyview" 
<%						If oRs("publiccanonlyview") Then 
							response.write " checked=""checked"" "
						End If	%> /> Allow public to view but not purchase
				</p>
				<% if session("orgid") = "60" then%>
				<p>
					<input type="checkbox" id="norefunds" name="norefunds" 
<%						If oRs("norefunds") Then 
							response.write " checked=""checked"" "
						End If	%> /> Only Allow Administrative Refunds
				</p>

				<%end if %>
		<%		If lcl_orghasfeature_discounts Then  %> 
				<p>
					Discount: <% 
					If lcl_orghasfeature_discounts Then 
						ShowPriceDiscountPicks oRs("pricediscountid"), Session("OrgID")
					End If %> 
				</p>
		<%		Else %>
					<input type="hidden" name="pricediscountid" value="0" />
		<%		End If %>
			</div>

			<div id="tab7"> <!-- Occurs -->
				<p class="classeditdescription">This tab is where the individual activity occurrences are set up.
					Every class/event needs at least one activity occurrence to tell when it happens.
					Input your activity number which is like a course catalog number. Then select your instructor, 
					if there is one, and your enrollment sizes. You will also need to input the days and times that this meets. 
					If the activity meets more than once on the same day, then add a new activity row, input the same activity
					number and then just skip to the days where you select the day and input the second meeting time.
				</p>
				<p>
				<input type="button" class="button" value="Add Row" id="addref" onClick="NewTimeRow()" />
				<table id="seriestime" border="0" cellpadding="0" cellspacing="0">
					<tr id="occurstabletitle"><th align="center" colspan="7">Activities and Enrollment</th><th align="center" colspan="10" class="firstday">Days and Times This Meets</th></tr>
					<tr>
						<th>Activity #</th><th>Instructor</th><th>Min</th><th>Max</th><th>Enrld</th><th>Waitlist<br />Max</th><th>Can-<br />celed</th>
						<th class="firstday">Su</th><th>M</th><th>T</th><th>W</th><th>Th</th><th>F</th><th>S</th><th>Start<br />Time</th><th>End<br />Time</th><th>Delete<br />Row*</th>
					</tr>
<%					
					iMaxTimeRows = ShowClassTimes( iClassId ) 
%>
					</table>

					<strong>* Activities cannot be completely deleted when there are people on the roster.</strong>
					<input type="hidden" name="maxtimeid" value="<%=iMaxTimeRows%>" />
					<input type="hidden" name="maxtimedayid" value="<%=iMaxTimeRows%>" />

				</p>
			</div>

			<div id="tab8"> <!-- Early Registration -->
				<p>
					<input type="checkbox" id="allowearlyregistration" name="allowearlyregistration" <%If oRs("allowearlyregistration") Then response.write " checked=""checked"" " %> />
					&nbsp; Allow Early Registration For This Class/Event
				</p>
				<p>
					Early Registration Start Date: &nbsp; <input type="text" id="earlyregistrationdate" name="earlyregistrationdate" value="<%=oRs("earlyregistrationdate")%>" size="10" maxlength="10" />
					&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('earlyregistrationdate');" /></span>
				</p>
				<p>
					<strong>Select The Class/Event That Gets Early Registration:</strong><br /><br />
					<table cellpadding="5" cellspacing="0" border="0">
						<tr>
							<td align="right" valign="top" class="earlyreglabel">Season: &nbsp;</td>
							<td valign="top"><select id="earlyregistrationclassseasonid" name="earlyregistrationclassseasonid" onchange="javascript:ClassChange();">
											<% iEarlySeasonId = ShowEarlyRegistrationClassSeasons( oRs("earlyregistrationclassseasonid") ) %>
											</select> &nbsp; &nbsp;
							</td>
						</tr>
						<tr>
							<td align="right" valign="top" class="earlyreglabel">Class/Event: &nbsp</td>
							<td><span id="earlyclass"><% ShowEarlRegistrationClasses iEarlySeasonId, iClassId %></span>
								<br />* multiple class selection allowed
								<br /><input type="button" class="assignbuttons" name="NoEarlyClasses" value="Clear Selection" onclick="selectAll(document.getElementById('earlyregistrationclassid'),false)" />
							</td>
						</tr>
					</table>
				</p>
			</div>

<%			If lcl_orghasfeature_rentals And lcl_orghasfeature_classes_have_rentals Then	%>

			<div id="tab9"> <!-- Rentals -->
				<p class="classeditdescription">This tab allows classes to be scheduled into the rentals. 
					For this to work the class must have a location, a start date, and an end date.
					There must also be activities set up that have an activity number, selected days of the week, 
					a start time, and an end time.
				</p>
				<p><strong>Location: </strong><%= GetLocationName( oRs("locationid") ) %></p>
				<p><strong>Classes Start: </strong><% If oRs("startdate") <> "1/1/1900" Then response.write oRs("startdate") %></p>
				<p><strong>Classes End: </strong><% If oRs("enddate") <> "1/1/1900" Then response.write oRs("enddate") %></p>
				<table id="rentalschedule" cellspacing="0" cellpadding="2" border="0">
					<tr><th>Activity #</th><th>Days</th><th>Start<br />Time</th><th>End<br />Time</th><th>Rental</th><th>&nbsp;</th></tr>
<%					iRentalTimeRows = ShowActivitiesAndRentals( iClassId, oRs("locationid") )		%>
				</table>
				<input type="hidden" name="maxrentaltimerows" value="<%=iRentalTimeRows%>" />
			</div>

<%			Else		%>
				<input type="hidden" name="maxrentaltimerows" value="0" />
<%			End If		%>

		</div>
	</div>

		<p>
			<input type="button" class="button" name="update2" value="Save Changes" onclick="ValidateForm();" />
			<!--<input type="button" class="button" name="copy" value="<%=cCopyText%>" onclick="CopyClass('<%=iClassId%>');" />-->
		</p>
	</form>

	<fieldset class="edit"><legend><strong> Copy To A New Class </strong></legend>
		<form name="copyForm" accept-charset="UTF-8" method="post" action="class_copyclass.asp">
			<input type="hidden" name="classid" value="<%=iClassId%>" />
			<p>
				To Season: <% ShowClassSeasonFilterPicks 0  ' In class_global_functions.asp %> &nbsp; 
				<input type="checkbox" name="copyattendees" /> Bring Attendees as Waitlist &nbsp; 
				<input type="button" class="button" name="copy" value="<%=cCopyText%>" onclick="CopyClass('<%=iClassId%>');" />
			</p>
		</form>
	</fieldset>

<% 
		If bIsSeriesParent Then %>
		<fieldset class="edit"><legend><strong><a name="children"> Individual Classes/Events </a></strong></legend>
			<p id="newchild">
				<input type="button" id="childbutton" class="button" name="child" value="Create New Individual" onclick="CopyToChild('<%=iClassId%>');" />
			</p>
			<p id="childrenlist">
				<table id="children" class="style-alternate" border="1" cellpadding="0" cellspacing="0">
					<tr><th id="classname">Class/Event Name</th><th>Status</th><th>Starts</th><th>Ends</th><th>&nbsp;</th></tr>
					<% ShowSeriesChildren iClassId  %>
				</table>
			</p>
			<p id="addsingle">
				<form name="SingleForm" method="post" accept-charset="UTF-8" action="class_addsingletoseries.asp">
					<input type="hidden" name="parentclassid" value="<%=iClassId%>" />
					<% ShowAvailableClasses %> &nbsp; 
					<input type="button" class="button" name="addexisting" value="Add Existing Single" onclick="AddSingle();" />
				</form>
			</p>
		</fieldset>
<% 
		End If 

'------------------------------------------------------------------------------
	 Else 
		response.write "<p>No information could be found for this Class/Event.</p>"
	 End If 
'------------------------------------------------------------------------------

	oRs.Close
	Set oRs = Nothing 
%>
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<%

'------------------------------------------------------------------------------
' integer ShowClassTimes( iClassId )
'------------------------------------------------------------------------------
Function ShowClassTimes( ByVal iClassId )
	Dim sSql, oRs, sOldActivityId, iRowNo, iActivityCount

	sSql = "SELECT T.timeid, T.activityno, T.min, T.max, T.waitlistmax, ISNULL(T.instructorid,0) AS instructorid, T.enrollmentsize, T.iscanceled, "
	sSql = sSql & " D.timedayid, sunday, monday, tuesday, wednesday, thursday, friday, saturday, D.starttime, D.endtime "
	sSql = sSql & " FROM egov_class_time T, egov_class_time_days D " 
	sSql = sSql & " WHERE T.timeid = D.timeid AND T.classid = " & iClassId
	sSql = sSql & " ORDER BY T.timeid, D.timedayid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		sOldActivityId = ":a:"
		iRowNo = -1
		iActivityCount = 0

		Do While Not oRs.EOF
			iRowNo = iRowNo + 1
			'Build the table of times that were originally created
			If sOldActivityId <> oRs("activityno") Then 
				iActivityCount = iActivityCount + 1

				If iActivityCount Mod 2 = 0 Then 
					sRowClass = " class=""altrow"" "
				Else 
					sRowClass = ""
				End If 
			End If 

			response.write vbcrlf & "<tr" & sRowClass & ">"
			response.write "<td class=""ref"">"
			response.write "<input type=""hidden"" name=""timeid" & iRowNo & """ id=""timeid" & iRowNo & """  value=""" & oRs("timeid") & """ />"
			response.write "<input type=""hidden"" name=""timedayid" & iRowNo & """ value=""" & oRs("timedayid") & """ />"

			If sOldActivityId <> oRs("activityno") Then 
				response.write "<input type=""text"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""" & oRs("activityno") & """ size=""10"" maxlength=""10"" />" 
			Else 
				'response.write oRs("activityno") & " <input type=""hidden"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""" & oRs("activityno") & """ />"
				response.write "&nbsp; <input type=""hidden"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""skip"" />"
			End If 

			response.write "</td>" 
			response.write "<td align=""center"">" 

			If sOldActivityId <> oRs("activityno") Then 
				ShowInitialInstructorPicks iRowNo, oRs("instructorid")
			Else 
				response.write "&nbsp;"
			End If 

			response.write "</td>"

			If sOldActivityId <> oRs("activityno") Then 
				If oRs("iscanceled") Then 
					lcl_checked_cancelled = " checked=""checked"""
				Else 
					lcl_checked_cancelled = ""
				End If 

				response.write "<td align=""center""><input type=""text"" name=""min" & iRowNo & """ value=""" & oRs("min") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center""><input type=""text"" name=""max" & iRowNo & """ value=""" & oRs("max") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center"">" & oRs("enrollmentsize") & "<input type=""hidden"" id=""enrollmentsize" & iRowNo & """ name=""enrollmentsize" & iRowNo & """ value=""" & oRs("enrollmentsize") & """ /></td>"
				response.write "<td align=""center""><input type=""text"" name=""waitlistmax" & iRowNo & """ value=""" & oRs("waitlistmax") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center""><input type=""checkbox""  id=""iscanceled" & iRowNo & """ name=""iscanceled" & iRowNo & """" & lcl_checked_cancelled & " />"
			Else 
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
			End If 

			setupWeekdayCheckbox True, True,  "su", iRowNo, oRs("sunday")
			setupWeekdayCheckbox True, False, "mo", iRowNo, oRs("monday")
			setupWeekdayCheckbox True, False, "tu", iRowNo, oRs("tuesday")
			setupWeekdayCheckbox True, False, "we", iRowNo, oRs("wednesday")
			setupWeekdayCheckbox True, False, "th", iRowNo, oRs("thursday")
			setupWeekdayCheckbox True, False, "fr", iRowNo, oRs("friday")
			setupWeekdayCheckbox True, False, "sa", iRowNo, oRs("saturday")

			response.write "<td align=""center""><input type=""text"" id=""starttime" & iRowNo & """ name=""starttime" & iRowNo & """ value=""" & oRs("starttime") & """ size=""8"" maxlength=""7"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""endtime" & iRowNo & """ name=""endtime" & iRowNo & """ value=""" & oRs("endtime") & """ size=""8"" maxlength=""7"" /></td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""delete" & iRowNo & """ name=""delete" & iRowNo & """ /></td>"
			response.write "</tr>"

			sOldActivityId = oRs("activityno")
			oRs.MoveNext
		Loop 
	Else 
		'They do not have any time rows
		response.write vbcrlf & "<tr>"
		response.write "<td class=""ref"">"
		response.write "<input type=""hidden"" id=""timeid0"" name=""timeid0"" value=""0"" />"
		response.write "<input type=""text"" id=""activity0"" name=""activity0"" value="""" size=""10"" maxlength=""10"" />"
		response.write "<input type=""hidden"" id=""timedayid0"" name=""timedayid0"" value=""0"" />"
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowInitialInstructorPicks 0, 0 
		response.write "</td>"
		response.write "<td align=""center""><input type=""text"" name=""min0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center""><input type=""text"" name=""max0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center""><input type=""text"" name=""waitlistmax0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""iscanceled0"" name=""iscanceled0"" /></td>"

		setupWeekdayCheckbox True, True,  "su", "0", ""
		setupWeekdayCheckbox True, False, "mo", "0", ""
		setupWeekdayCheckbox True, False, "tu", "0", ""
		setupWeekdayCheckbox True, False, "we", "0", ""
		setupWeekdayCheckbox True, False, "th", "0", ""
		setupWeekdayCheckbox True, False, "fr", "0", ""
		setupWeekdayCheckbox True, False, "sa", "0", ""

		response.write "<td align=""center""><input type=""text"" id=""starttime0"" name=""starttime0"" value="""" size=""8"" maxlength=""7"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""endtime0"" name=""endtime0"" value="""" size=""8"" maxlength=""7"" /></td>"
		response.write "</tr>"

		iRowNo = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

	ShowClassTimes = iRowNo

End Function 


'------------------------------------------------------------------------------
' void GetPricing iClassId, iOrgID, iMembershipId, iClassSeasonId, iMin, iMax
'------------------------------------------------------------------------------
Sub GetPricing( iClassId, iOrgID, iMembershipId, iClassSeasonId, ByRef iMin, ByRef iMax )
	Dim sSql, oRs

	iMax = 0
	iMin = 10000

	sSql = "SELECT pricetypeid, pricetypename, ismember, ISNULL(instructorpercent,0) AS instructorpercent, "
	sSql = sSql & " needsregistrationstartdate, isfee, isdropin, basepricetypeid "
	sSql = sSql & " FROM egov_price_types "
	sSql = sSql & " WHERE isactiveforclasses = 1 AND orgid = " & iOrgId
	sSql = sSql & " ORDER BY displayorder "

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If CLng(oRs("pricetypeid")) > CLng(iMax) Then 
			iMax = oRs("pricetypeid")
		End If 
		If CLng(oRs("pricetypeid")) < CLng(iMin) Then 
			iMin = oRs("pricetypeid")
		End If 
		response.write vbcrlf & "<tr>"
		response.write "<td class=""type"">"
		response.write "<input type=""checkbox"" name=""pricetypeid"" id=""pricetypeid" & oRs("pricetypeid") & """ value=""" & oRs("pricetypeid") & """"
		If ClassHasPriceType( iClassId, clng(oRs("pricetypeid")) ) Then
			response.write " checked=""checked"" "
		End If 
		response.write " />&nbsp;" & oRs("pricetypename") 
		If oRs("isfee") Then
			response.write " (fee)"
		ElseIf oRs("isdropin") Then
			response.write " (one time)"
		ElseIf Not IsNull(oRs("basepricetypeid")) Then
			response.write " (+)" 
		End If 

		If lcl_orghasfeature_gl_accounts Then 
			response.write "</td>"
			iAccountId = GetClassPricetypeAccount( iClassId, clng(oRs("pricetypeid")) )
			response.write "<td align=""center"">"
			ShowAccountPicks iAccountId, oRs("pricetypeid")  ' In common.asp
		Else
			response.write vbcrlf & "<input type=""hidden"" name=""accountid" & oRs("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		response.write "<td align=""center""><input type=""text"" name=""amount" & oRs("pricetypeid") & """ id=""amount" & oRs("pricetypeid") & """ value=""" & GetPriceAmount( oRs("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" /></td>"

		response.write "<td align=""center"">"
		If ClassCanNeedMemberships() Then
			If oRs("ismember") Then
				'Show the membership picks for the one that requires membership
				ShowClassMembershipPicks iMembershipId, oRs("pricetypeid")  ' In class_global_functions.asp
			Else
				response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oRs("pricetypeid") & """ value=""0"" />"
			End If 
		Else 
			response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oRs("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		response.write "<td align=""center""><input type=""text"" name=""instructorpercent" & oRs("pricetypeid") & """ id=""instructorpercent" & oRs("pricetypeid") & """ value=""" & GetClassPricetypeInstructorPercent( iClassId, oRs("pricetypeid"), oRs("instructorpercent") ) & """ size=""3"" maxlength=""3"" /></td>"

		If oRs("needsregistrationstartdate") Then 
			response.write "<td align=""center"">"

			lcl_registration_start_date = GetSeasonalStartDate( iClassSeasonId, oRs("pricetypeid") )

			If ClassHasPriceType( iClassId, clng(oRs("pricetypeid")) ) Then 
				lcl_registration_start_date = GetClassPricetypeRegistrationStart( iClassId, oRs("pricetypeid"), lcl_registration_start_date )
				'     else
				'if a start date for the price type does NOT exist or is NULL then get the default season date
				'        if lcl_registration_start_date = "" OR isnull(lcl_registration_start_date) then
				'           lcl_registration_start_date = GetDefaultSeasonDate( iClassSeasonId, "registrationstartdate" )
				'        end if
			End If 

			response.write "<input type=""text"" maxlength=""10"" class=""datefield"" name=""registrationstartdate" & oRs("pricetypeid") & """ id=""registrationstartdate" & oRs("pricetypeid") & """ value=""" & lcl_registration_start_date & """ />&nbsp;"
			response.write "<span class=""calendarimg"" style=""cursor:hand;"">"
			response.write "<img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('registrationstartdate" & oRs("pricetypeid") & "');"" />"
			response.write "</span>" 
			response.write "</td>"
		Else 
			response.write "<td>&nbsp;<input type=""hidden"" name=""registrationstartdate" & oRs("pricetypeid") & """ id=""registrationstartdate" & oRs("pricetypeid") & """ value="""" /></td>"
		End If 

		response.write "</tr>"	
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassHasPriceType( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ClassHasPriceType( ByVal iClassId, ByVal iPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(pricetypeid) AS hits "
	sSql = sSql & " FROM egov_class_pricetype_price "
	sSql = sSql & " WHERE pricetypeid = " & iPriceTypeId
	sSql = sSql & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If clng(oRs("hits")) > clng(0) Then 
		ClassHasPriceType = True
	Else 
		ClassHasPriceType = False
	End If 

	oRs.Close
	Set oRs = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' integer GetClassPricetypeAccount( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeAccount( ByVal iClassId, ByVal iPriceTypeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(accountid,0) AS accountid "
	sSql = sSql & " FROM egov_class_pricetype_price "
	sSql = sSql & " WHERE pricetypeid = " & iPriceTypeId
	sSql = sSql & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		GetClassPricetypeAccount = oRs("accountid")
	Else 
		GetClassPricetypeAccount = 0
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' integer GetClassPricetypeInstructorPercent( iClassId, iPriceTypeId, iDefaultPercent )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeInstructorPercent( ByVal iClassId, ByVal iPriceTypeId, ByVal iDefaultPercent )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(instructorpercent,0) AS instructorpercent "
	sSql = sSql & "FROM egov_class_pricetype_price "
	sSql = sSql & "WHERE pricetypeid = " & iPriceTypeId & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetClassPricetypeInstructorPercent = oRs("instructorpercent")
	Else 
		GetClassPricetypeInstructorPercent = iDefaultPercent
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetClassPricetypeRegistrationStart( iClassId, iPriceTypeId, sDefaultDate )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeRegistrationStart( ByVal iClassId, ByVal iPriceTypeId, ByVal sDefaultDate )
	Dim sSql, oRs

	sSql = "SELECT registrationstartdate "
	sSql = sSql & " FROM egov_class_pricetype_price "
	sSql = sSql & " WHERE pricetypeid = " & iPriceTypeId
	sSql = sSql & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		  GetClassPricetypeRegistrationStart = oRs("registrationstartdate")
	Else 
		  GetClassPricetypeRegistrationStart = sDefaultDate
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string GetPriceAmount( iPriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetPriceAmount( ByVal iPriceTypeId, ByVal iClassId )
	Dim sSql, oRs, sAmount

	sSql = "SELECT amount FROM egov_class_pricetype_price "
	sSql = sSql & "WHERE pricetypeid = " & iPriceTypeId & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		sAmount = FormatNumber(oRs("amount"),2,,,0) ' no commas in this
		GetPriceAmount = sAmount
	Else 
		GetPriceAmount = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' string GetDefaultSeasonDate( iClassSeasonId, sDateField )
'--------------------------------------------------------------------------------------------------
Function GetDefaultSeasonDate( ByVal iClassSeasonId, ByVal sDateField )
	Dim sSql, oRs

	sSql = "SELECT " & sDateField & " AS keydatefield FROM egov_class_seasons "
	sSql = sSql & "WHERE classseasonid = " & iClassSeasonId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If IsNull(oRs(sDateField)) Then 
			GetDefaultSeasonDate = ""
		Else 
			GetDefaultSeasonDate = DateValue(oRs("keydatefield"))
		End If 
	Else 
		GetDefaultSeasonDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void GetPricing1( iClassId, iOrgID, iMembershipId
'--------------------------------------------------------------------------------------------------
Sub GetPricing1( ByVal iClassId, ByVal iOrgID, ByVal iMembershipId )
	Dim sSql, oRs, iRow

	iRow = 0
	sSql = "SELECT pricetypeid, pricetypename, ismember FROM egov_price_types "
	sSql = sSql & "WHERE orgid = " & iOrgId & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRow = iRow + 1
		response.write vbcrlf & "<tr><td><input type=""hidden"" name=""pricetypeid" & iRow & """ value=""" & oRs("pricetypeid") & """ />" & oRs("pricetypename") & "</td>"
		response.write "<td><input type=""text"" name=""amount" & iRow & """ value=""" & GetPriceAmount( oRs("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" /></td><td>"
		If oRs("ismember") Then
			' Show the membership picks for the one that requires membership
			response.write "Membership: &nbsp; "
			ShowMembershipPicks iMembershipId
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td></tr>"	
		oRs.MoveNext
	Loop 
	response.write vbcrlf & "<input type=""hidden"" name=""pricetypeidcount"" value=""" & iRow & """ />"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowSeriesChildren iParentClassId
'--------------------------------------------------------------------------------------------------
Sub ShowSeriesChildren( ByVal iParentClassId )
	Dim sSql, oRs, iRowCOunt

	sSql = "SELECT C.classid, C.classname, C.startdate, C.enddate, S.statusname FROM egov_class C, egov_class_status S "
	sSql = sSql & " WHERE C.parentclassid = " & iParentClassId & " AND C.statusid = S.statusid ORDER BY C.startdate, C.classname"
	iRowCOunt = 1

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCOunt = iRowCOunt + 1
		If ( iRowCOunt Mod 2 ) = 0 Then 
			response.write vbcrlf & "<tr class=""alt_row"">"
		Else 
			response.write vbcrlf & "<tr>"
		End If 
		response.write vbcrlf & "<td>" & oRs("classname") & "</td><td align=""center"">" & oRs("statusname") & "</td><td align=""center"">" & oRs("startdate") & "</td><td align=""center"">" & oRs("enddate") & "</td>"
		response.write "<td align=""center""><input type=""button"" class=""button"" name=""editchild"" value=""Edit"" onclick=""javascript:location.href='edit_class.asp?classid=" & oRs("classid") & "';"" /></td></tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowLocationPicks iLocationId 
'--------------------------------------------------------------------------------------------------
Sub ShowLocationPicks( ByVal iLocationId )
	Dim sSql, oRs

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & Session("orgid") & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""locationid"">"
	
	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("locationid") & """ "
			If CLng(oRs("locationid")) = CLng(iLocationId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
	Else
		response.write vbcrlf & vbtab & "<option value=""0"">Unknown Location</option>"
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPOCPicks iPocId, iOrgId
'--------------------------------------------------------------------------------------------------
Sub ShowPOCPicks( ByVal iPocId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT pocid, name FROM egov_class_pointofcontact WHERE orgid = " & iOrgId & " ORDER BY name"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""pocid"" >"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("pocid") & """ "
				If CLng(oRs("pocid")) = CLng(iPocId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oRs("name") & "</option>"
			oRs.MoveNext
		Loop 
		response.write "</select>"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowInstructors iClassId 
'--------------------------------------------------------------------------------------------------
Sub ShowInstructors( ByVal iClassId )
	Dim sSql, oRs, iRow

	iRow = 0
	sSql = "SELECT L.lastname, L.firstname FROM egov_class_instructor L, egov_class_to_instructor C"
	sSql = sSql & " WHERE L.instructorid = C.instructorid"
	sSql = sSql & " AND C.classid = " & iClassId & " ORDER BY L.lastname, L.firstname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oRs("firstname") & " " & oRs("lastname")
			oRs.MoveNext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No Instructor Assigned"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowWaivers iClassId
'--------------------------------------------------------------------------------------------------
Sub ShowWaivers( ByVal iClassId )
	Dim sSql, oRs, iRow

	iRow = 0
	sSql = "SELECT waivername FROM egov_class_waivers W, egov_class_to_waivers C"
	sSql = sSql & " WHERE W.waiverid = C.waiverid"
	sSql = sSql & " AND C.classid = " & iClassId & " ORDER BY waivername"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		Do While Not oRs.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oRs("waivername") 
			oRs.MoveNext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No 3rd Party Waivers Assigned"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowRegistrationPicks iOptionId
'--------------------------------------------------------------------------------------------------
Sub ShowRegistrationPicks( ByVal iOptionId )
	Dim sSql, oRs

	sSql = "SELECT optionid, optionname, optiondescription, requirestime "
	sSql = sSql & "FROM egov_registration_option ORDER BY optionid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 1, 1

	If Not oRs.EOF Then
		response.write vbcrlf & "<select name=""optionid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select an Option</option>"
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("optionid") & """ "
				If CLng(oRs("optionid")) = CLng(iOptionId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oRs("optionname") & " &ndash; " & oRs("optiondescription") & "</option>"
			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"

		' Put in the hidden fields to check for time requirements
		oRs.MoveFirst
		Do While Not oRs.EOF
			response.write vbcrlf & "<input type=""hidden"" name=""requirestime" & oRs("optionid") & """ value=""" & oRs("requirestime") & """ />"
			oRs.MoveNext
		Loop 

	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' string CheckDayOfWeek( iClassId, iDayOfWeek )
'--------------------------------------------------------------------------------------------------
Function CheckDayOfWeek( ByVal iClassId, ByVal iDayOfWeek )
	Dim sSql, oRs

	sSql = "SELECT rowid FROM egov_class_dayofweek "
	sSql = sSql & "WHERE dayofweek = " & iDayOfWeek & " AND classid = " & iClassId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		CheckDayOfWeek = " checked=""checked"" "
	Else 
		CheckDayOfWeek = ""
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' string CheckClassInCategory( iClassId, iCategoryId )
'--------------------------------------------------------------------------------------------------
Function CheckClassInCategory( ByVal iClassId, ByVal iCategoryId )
	Dim sSql, oCat

	sSql = "SELECT rowid FROM egov_class_category_to_class "
	sSql = sSql & "WHERE categoryid = " & iCategoryId & " AND classid = " & iClassId

	Set oCat = Server.CreateObject("ADODB.Recordset")
	oCat.Open sSql, Application("DSN"), 0, 1

	If Not oCat.EOF Then
		CheckClassInCategory = " checked=""checked"" "
	Else 
		CheckClassInCategory = ""
	End If 

	oCat.Close
	Set oCat = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowCategories iClassId, iOrgId
'--------------------------------------------------------------------------------------------------
Sub ShowCategories( ByVal iClassId, ByVal iOrgId )
	Dim sSql, oRs, iRow

	iRow = 0
	sSql = "SELECT categoryid, categorytitle FROM egov_class_categories "
	sSql = sSql & "WHERE isroot = 0 AND orgid = " & iOrgId & " ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<input type=""checkbox"" name=""categoryid"" value=""" & oRs("categoryid") & """ "
		response.write CheckClassInCategory( iClassId, oRs("categoryid") )
		response.write " /> " & oRs("categorytitle") & " <br />"
		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowClassCategories iClassId
'--------------------------------------------------------------------------------------------------
Sub ShowClassCategories( ByVal iClassId )
	Dim sSql, oRs, iRow

	iRow = 0
	sSql = "SELECT categorytitle FROM egov_class_categories C, egov_class_category_to_class G "
	sSql = sSql & "WHERE C.categoryid = G.categoryid AND classid = " & iClassId & " ORDER BY categorytitle"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If iRow > 0 Then
			response.write ", "
		End If 
		response.write oRs("categorytitle") 
		oRs.MoveNext
		iRow = iRow + 1
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPriceDiscountPicks iPriceDiscountId, iOrgId
'--------------------------------------------------------------------------------------------------
Sub ShowPriceDiscountPicks( ByVal iPriceDiscountId, ByVal iOrgId )
	Dim sSql, oRs

	sSql = "SELECT pricediscountid, discountname FROM egov_price_discount "
	sSql = sSql & "WHERE orgid = " & iOrgId & " ORDER BY discountname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""pricediscountid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">No Discount Applied</option>"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("pricediscountid") & """ "
		If CLng(oRs("pricediscountid")) = CLng(iPriceDiscountId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("discountname") & "</option>"
		oRs.MoveNext
	Loop 
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowClassTimesold iClassId
'--------------------------------------------------------------------------------------------------
Sub ShowClassTimesold( ByVal iClassId )
	Dim sSql, oRs, iRow, sMin, sMax, sWaitlistmax

	iRow = 0
	sSql = "SELECT timeid, starttime, endtime, ISNULL(min,0) AS min, ISNULL(max, 0) AS max, ISNULL(waitlistmax,0) AS waitlistmax "
	sSql = sSql & "FROM egov_class_time WHERE classid = " & iClassId & " ORDER BY timeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRow = iRow + 1
		If clng(oRs("min")) = 0 Then
			sMin = ""
		Else
			sMin = oRs("min")
		End If 
		If clng(oRs("max")) = 0 Then
			sMax = ""
		Else
			sMax = oRs("max")
		End If 
		If clng(oRs("waitlistmax")) = 0 Then
			sWaitlistmax = ""
		Else
			sWaitlistmax = oRs("waitlistmax")
		End If 
		response.write vbcrlf & "<tr>"
		response.write "<input type=""hidden"" name=""timeid" & iRow & """ value=""" & oRs("timeid") & """ />"
		response.write "<td align=""center""><input type=""checkbox"" name=""removetime" & iRow & """ /></td>"
		response.write "<td><input type=""text"" name=""starttime" & iRow & """ value=""" & oRs("starttime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""endtime" & iRow & """ value=""" & oRs("endtime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""min" & iRow & """ value=""" & sMin & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""max" & iRow & """ value=""" &sMax & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""waitlistmax" & iRow & """ value=""" & sWaitlistmax & """ size=""4"" maxlength=""5"" /></td>"
		response.write vbcrlf & "</tr>"
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

	response.write vbcrlf & "<input type=""hidden"" name=""timecount"" value=""" & iRow & """ />"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowAvailableClasses
'--------------------------------------------------------------------------------------------------
Sub ShowAvailableClasses( )
	Dim sSql, oRs

	sSql = "SELECT classid, classname FROM egov_class "
	sSql = sSql & " WHERE classtypeid = 3 AND orgid = " & Session("OrgID") & " ORDER BY classname"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select border=""0"" name=""classid"">"
	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("classid") & """>" & oRs("classname") & "</option>"
		oRs.MoveNext
	Loop 
	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' integer ShowEarlyRegistrationClassSeasons( iClassSeasonId )
'--------------------------------------------------------------------------------------------------
Function ShowEarlyRegistrationClassSeasons( ByVal iClassSeasonId )
	Dim sSql, oRs, iSelectedSeasonId

	iSelectedSeasonId = CLng(iClassSeasonId)

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear desc, S.displayorder DESC, C.seasonname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1
	
	If Not oRs.EOF Then
		Do While NOT oRs.EOF
			If iSelectedSeasonId = CLng(0) Then
				iSelectedSeasonId = CLng(oRs("classseasonid"))
			End If 
			response.write vbcrlf & "<option value=""" & oRs("classseasonid") & """"  
			If CLng(iClassSeasonId) = CLng(oRs("classseasonid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("seasonname") & "</option>"
			oRs.MoveNext
		Loop
	End If

	oRs.Close
	Set oRs = Nothing

	ShowEarlyRegistrationClassSeasons = iSelectedSeasonId

End Function 


'--------------------------------------------------------------------------------------------------
' ShowEarlRegistrationClasses iClassSeasonId, iClassId
'--------------------------------------------------------------------------------------------------
Sub ShowEarlRegistrationClasses( ByVal iClassSeasonId, ByVal iClassId )
	Dim sSql, oRs

	sSql = "SELECT C.classid, C.classname "
	sSql = sSql & " FROM egov_class C, egov_class_status S, egov_registration_option RO "
	sSql = sSql & " WHERE C.statusid = S.statusid AND S.statusname = 'ACTIVE' AND C.classseasonid = " & iClassSeasonId
	sSql = sSql & " AND RO.optionid = C.optionid AND RO.canpurchase = 1 "
	sSql = sSql & " AND C.orgid = " & SESSION("orgid") & " ORDER BY C.classname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		response.write "<select id=""earlyregistrationclassid"" name=""earlyregistrationclassid"" multiple=""multiple"" size=""20"">"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value=""" & oRs("classid") & """"
			If ClassIsInEarlyRegistration( iClassId, CLng(oRs("classid")) ) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("classname") & "</option>"
			oRs.MoveNext
		Loop
		response.write "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean bIsEarlyRegistration = ClassIsInEarlyRegistration( iClassId, iPotentialClassId )
'--------------------------------------------------------------------------------------------------
Function ClassIsInEarlyRegistration( ByVal iClassId, ByVal iPotentialClassId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(earlyregistrationclassid) AS hits FROM egov_class_earlyregistrations "
	sSql = sSql & " WHERE classid = " & iClassId & " AND earlyregistrationclassid = " & iPotentialClassId

	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then 
		If CLng(oRs("hits")) > CLng(0) Then
			ClassIsInEarlyRegistration = True 
		Else
			ClassIsInEarlyRegistration = False 
		End If 
	Else
		ClassIsInEarlyRegistration = False 
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Function 


'--------------------------------------------------------------------------------------------------
' integer iTimeRows = ShowActivitiesAndRentals( iClassId, iLocationId )
'--------------------------------------------------------------------------------------------------
Function ShowActivitiesAndRentals( ByVal iClassId, ByVal iLocationId )
	Dim sSql, oRs, iRowCount

	iRowCount = 0

	sSql = "SELECT timeid, activityno, ISNULL(rentalid,0) AS rentalid, ISNULL(reservationid,0) AS reservationid "
	sSql = sSql & "FROM egov_class_time WHERE classid = " & iClassId & " ORDER BY activityno, timeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		iRowCount = iRowCount + 1
		response.write vbcrlf & "<tr>"
		response.write "<td>" & oRs("activityno") & "</td>"
		'response.write "<td colspan=""3"">"
		ShowDaysAndTime oRs("timeid")
		'response.write "</td>"
		response.write "<td align=""center"">"
		response.write "<input type=""hidden"" id=""rentaltimeid" & iRowCount & """ name=""rentaltimeid" & iRowCount & """ value=""" & oRs("timeid") & """ />"
		If CLng(oRs("reservationid")) > CLng(0) And CLng(oRs("rentalid")) > CLng(0) Then
			response.write "<input type=""hidden"" id=""rentalid" & iRowCount & """ name=""rentalid" & iRowCount & """ value=""" & oRs("rentalid") & """ />"
			ShowRentalName oRs("rentalid")
		Else 
			ShowLocationRentalsDropDown iLocationId, oRs("rentalid"), iRowCount
		End If 
		response.write "</td>"
		response.write "<td>"
		If CLng(oRs("reservationid")) > CLng(0) Then
			response.write "<input type=""button"" class=""button"" value=""View Reservation"" onclick=""location.href='../rentals/reservationedit.asp?reservationid=" & oRs("reservationid") & "';"" />"
		Else 
			response.write "<input type=""button"" class=""button"" value=""Make Reservation"" onclick=""showDateSelection( '" & iRowCount & "' );"" />"
		End If 
		response.write "</td>"
		response.write "</tr>"
		oRs.MoveNext
	Loop
	
	oRs.Close
	Set oRs = Nothing 

	ShowActivitiesAndRentals = iRowCount

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowLocationRentalsDropDown iLocationId, iRentalId
'--------------------------------------------------------------------------------------------------
Sub ShowLocationRentalsDropDown( ByVal iLocationId, ByVal iRentalId, ByVal iRowCount )
	Dim sSql, oRs

	sSql = "SELECT rentalid, rentalname "
	sSql = sSql & "FROM egov_rentals "
	sSql = sSql & "WHERE locationid = " & iLocationId & " AND orgid = " & session("orgid")
	sSql = sSql & "ORDER BY rentalname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write vbcrlf & "<select id=""rentalid" & iRowCount & """ name=""rentalid" & iRowCount & """>"

	If Not oRs.EOF Then 
		Do While Not oRs.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oRs("rentalid") & """ "
			If CLng(oRs("rentalid")) = CLng(iRentalId) Then 
				response.write " selected=""selected"" "
			End If 
			response.write ">" & oRs("rentalname") & "</option>"
			oRs.MoveNext
		Loop 
	End If 

	response.write vbcrlf & "</select>"

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' ShowDaysAndTime iTimeId
'--------------------------------------------------------------------------------------------------
Sub ShowDaysAndTime( ByVal iTimeId )
	Dim sSql, oRs, sDayCell, sStartCell, sEndCell, bHaveDay

	sDayCell = ""
	sStartCell = ""
	sEndCell = ""

	sSql = "SELECT sunday, monday, tuesday, wednesday, thursday, friday, saturday, sunday, starttime, endtime "
	sSql = sSql & "FROM egov_class_time_days WHERE timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		bHaveDay = False 
		If sDayCell <> "" Then 
			sDayCell = sDayCell & "<br />"
		End If 
		If oRs("sunday") Then 
			sDayCell = sDayCell & "Su"
			bHaveDay = True 
		End If 
		If oRs("monday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "Mo"
		End If 
		If oRs("tuesday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "Tu"
		End If 
		If oRs("wednesday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "We"
		End If 
		If oRs("thursday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "Th"
		End If 
		If oRs("friday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "Fr"
		End If 
		If oRs("saturday") Then 
			If bHaveDay Then
				sDayCell = sDayCell & ","
			Else
				bHaveDay = True 
			End If 
			sDayCell = sDayCell & "Sa"
		End If 
		If sStartCell <> "" Then 
			sStartCell = sStartCell & "<br />"
		End If 
		sStartCell = sStartCell & oRs("starttime")
		If sEndCell <> "" Then 
			sEndCell = sEndCell & "<br />"
		End If 
		sEndCell = sEndCell & oRs("endtime")
		oRs.MoveNext
	Loop

	response.write "<td align=""center"">" & sDayCell
	response.write "</td>"
	response.write "<td align=""center"">" & sStartCell
	response.write "</td>"
	response.write "<td align=""center"">" & sEndCell
	response.write "</td>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub


'------------------------------------------------------------------------------
' void ShowRentalName iRentalId
'------------------------------------------------------------------------------
Sub ShowRentalName( ByVal iRentalId )
	Dim sSql, oRs

	sSql = "SELECT rentalname FROM egov_rentals "
	sSql = sSql & "WHERE rentalid = " & iRentalId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		response.write oRs("rentalname")
	Else
		response.write ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 


'------------------------------------------------------------------------------
' void setupWeekdayCheckbox p_displayTD, p_isFirstDay, p_name, p_rowno, p_value
'------------------------------------------------------------------------------
Sub setupWeekdayCheckbox( ByVal p_displayTD, ByVal p_isFirstDay, ByVal p_name, ByVal p_rowno, ByVal p_value )
	Dim lcl_checkbox_checked

	'Determine if the checkbox is surrounded by TD tags.
	If p_displayTD Then 
		'Now check to see if this is the "first day" in the list.
		If p_isFirstDay Then 
			response.write "<td class=""firstday"" align=""center"">"
		Else 
			response.write "<td align=""center"">"
		End If 
	End If 

	'Determine if the value is checked.
	lcl_checkbox_checked = ""

	If p_value <> "" Then 
		If p_value Then 
			lcl_checkbox_checked = " checked=""checked"" "
		End If 
	End If 

	response.write "<input type=""checkbox"" name=""" & p_name & p_rowno & """" & lcl_checkbox_checked & " />"

	If p_displayTD = "Y" Then 
		response.write "</td>"
	End If 

End Sub 


'------------------------------------------------------------------------------
' string GetSeasonalStartDate( iClassSeasonId, iPriceTypeId )
'------------------------------------------------------------------------------
Function GetSeasonalStartDate( ByVal iClassSeasonId, ByVal iPriceTypeId )
	Dim sSql, oRs 

	sSql = "SELECT registrationstartdate "
	sSql = sSql & " FROM egov_class_seasons_to_pricetypes_dates "
	sSql = sSql & " WHERE classseasonid = " & iClassSeasonId
	sSql = sSql & " AND pricetypeid = " & iPriceTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		GetSeasonalStartDate = DateValue(oRs("registrationstartdate"))
	Else 
		GetSeasonalStartDate = ""
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
sub setupRosterOptions(iClassID, iFieldType, iFieldLabel, iFieldEnabled, iShowInputFields, iFieldInputType)

  dim lcl_classid, lcl_fieldtype, lcl_fieldlabel, lcl_fieldenabled, lcl_showInputFields, lcl_fieldinputtype
  dim lcl_selected_enabled_both, lcl_selected_enabled_internalonly, lcl_selected_enabled_disabled
  dim lcl_selected_inputtype_lov, lcl_selected_inputtype_text

  lcl_classid                       = 0
  lcl_fieldtype                     = ""
  lcl_fieldlabel                    = ""
  lcl_fieldenabled                  = ""
  lcl_showInputFields               = true
  lcl_fieldinputtype                = ""
  lcl_selected_enabled_both         = ""
  lcl_selected_enabled_internalonly = ""
  lcl_selected_enabled_disabled     = ""
  lcl_selected_inputtype_lov        = ""
  lcl_selected_inputtype_text       = ""

  if iClassID <> "" then
     lcl_classid = clng(iClassID)
  end if

  if iFieldType <> "" then
     if not containsApostrophe(iFieldType) then
        lcl_fieldtype = ucase(iFieldType)
     end if
  end if

  if iFieldLabel <> "" then
     if not containsApostrophe(iFieldLabel) then
        lcl_fieldlabel = ucase(iFieldLabel)
     end if
  end if

 'Determine which value to select in the "enabled" dropdown list
  if iFieldEnabled <> "" then
     if not containsApostrophe(iFieldEnabled) then
        lcl_fieldenabled = ucase(iFieldEnabled)
     end if
  end if

  if lcl_fieldenabled = "INTERNAL ONLY" then
     lcl_selected_enabled_internalonly = " selected=""selected"""
  elseif lcl_fieldenabled = "DISABLED" then
     lcl_selected_enabled_disabled = " selected=""selected"""
  else
     lcl_selected_enabled_both = " selected=""selected"""
  end if

 'Determine which value to select in the tshirt dropdown list
  if iFieldInputType <> "" then
     if not containsApostrophe(iFieldInputType) then
        lcl_fieldinputtype = ucase(iFieldInputType)
     end if
  end if

  if iShowInputFields <> "" then
     lcl_showInputFields = iShowInputFields

     if lcl_showInputFields then
        if lcl_fieldinputtype = "TEXT" then
           lcl_selected_inputtype_text = " selected=""selected"""
        else
           lcl_selected_inputtype_lov  = " selected=""selected"""
        end if
     end if
  end if

  if lcl_fieldtype <> "" then
     response.write "<input type=""hidden"" name=""teamreg_" & lcl_fieldtype & "_accessorytype"" id=""teamreg_" & lcl_fieldtype & "_accessorytype"" value=""" & lcl_fieldtype & """ />" & vbcrlf
     response.write "<select name=""teamreg_" & lcl_fieldtype & "_enabled"" id=""teamreg_" & lcl_fieldtype & "_enabled"">" & vbcrlf
     response.write "  <option value=""BOTH"""          & lcl_selected_enabled_both         & ">Both</option>" & vbcrlf
     response.write "  <option value=""INTERNAL ONLY""" & lcl_selected_enabled_internalonly & ">Internal Only</option>" & vbcrlf
     response.write "  <option value=""DISABLED"""      & lcl_selected_enabled_disabled     & ">Disabled</option>" & vbcrlf
     response.write "</select>" & vbcrlf

     if lcl_showInputFields then
        response.write "<select name=""teamreg_" & lcl_fieldtype & "_inputtype"" id=""teamreg_" & lcl_fieldtype & "_inputtype"" onchange=""enableDisableMaintainButton('" & lcl_fieldtype & "');"">" & vbcrlf
        response.write "  <option value=""LOV"""  & lcl_selected_inputtype_lov  & ">Drop Down List</option>" & vbcrlf
        response.write "  <option value=""TEXT""" & lcl_selected_inputtype_text & ">Input Text Field</option>" & vbcrlf
        response.write "</select>&nbsp;" & vbcrlf
        response.write "<input type=""button"" name=""teamreg_" & lcl_fieldtype & "_input_button"" id=""teamreg_" & lcl_fieldtype & "_input_button"" value=""Maintain " & lcl_fieldlabel & " Options"" class=""button"" onclick=""openWin('class_accessoryoptions_list.asp?classid=" & lcl_classid & "&atype=" & lcl_fieldtype & "');"" />" & vbcrlf
     end if
  end if

end sub
%>
