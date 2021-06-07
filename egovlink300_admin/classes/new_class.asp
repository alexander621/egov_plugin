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
' 1.0	04/19/06 Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 2.0	03/09/07	Steve Loar - Total make over for Menlo Park Project
' 2.1	11/29/07	Steve Loar - Added yui tabs
' 2.2	02/15/08	Steve Loar - Early Registration added
' 2.3  12/30/08 David Boyer - Added "DisplayRosterPublic" checkbox for Craig, CO custom registration fields.
' 2.4  06/17/09 David Boyer - Added "Show Terms" checkbox
' 2.7	12/2/2009	Steve Loar - Option to only allow purchases on admin but display on public
' 2.8	10/10/2011	Steve Loar - Added gender restrictions
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim iFirstSeasonId, iEarlySeasonId, lcl_orghasfeature_class_supervisors
Dim lcl_orghasfeature_gl_accounts, lcl_orghasfeature_discounts
Dim lcl_orghasfeature_custom_registration_craigco, bHasGenderRestrictions
Dim iInitialGenderRestriction

sLevel = "../"  'Override of value from common.asp

If Not UserHasPermission( Session("UserId"), "create single" ) Then 
	response.redirect sLevel & "permissiondenied.asp"
End If 

'Check for org features
lcl_orghasfeature_class_supervisors           = orghasfeature("class supervisors")
lcl_orghasfeature_gl_accounts                 = orghasfeature("gl accounts")
lcl_orghasfeature_discounts                   = orghasfeature("discounts")
lcl_orghasfeature_custom_registration_craigco = orghasfeature("custom_registration_CraigCO") 
bHasGenderRestrictions = orgHasFeature("gender restriction") 

iInitialGenderRestriction = GetGenderNotRequiredId( )


blnHasWP = hasWordPress()
sHomeWebsiteURL = getOrganization_WP_URL(session("orgid"), "OrgPublicWebsiteURL")
%>
<html lang="en">
<head>
	<meta charset="UTF-8">
	
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1" />
	
	<title>E-Gov Administration Console {Create Class/Event}</title>

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
	
	<link rel="stylesheet" type="text/css" href="../yui/build/tabview/assets/skins/sam/tabview.css" />
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

  	<script src="//code.jquery.com/jquery-1.12.4.js"></script>
   	<script src="//code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
	<!--#include file="../includes/wp-image-picker.asp"-->

	<script type="text/javascript" src="../yui/yahoo-dom-event.js"></script>  
	<script type="text/javascript" src="../yui/element-min.js"></script>  
	<script type="text/javascript" src="../yui/tabview-min.js"></script>

	<script language="javascript" src="../scripts/ajaxLib.js"></script>
	<script language="javascript" src="../scripts/formatnumber.js"></script>
	<script language="javascript" src="../scripts/removespaces.js"></script>
	<script language="javascript" src="../scripts/removecommas.js"></script>
	<script language="javascript" src="../scripts/setfocus.js"></script>
	<script language="javascript" src="../scripts/isvaliddate.js"></script>
	<script language="javascript" src="../scripts/textareamaxlength.js"></script>

	<script language="javascript">
	<!--
		var tabView;

		(function() {
			tabView = new YAHOO.widget.TabView('demo');
			tabView.set('activeIndex', 0); 

		})();

		function ClassChange() 
		{
			// Try to get a drop down of names
			doAjax('getseasonclasses.asp', 'classseasonid=' + $("#earlyregistrationclassseasonid").val(), 'UpdateClasses', 'get', '0');
		}

		function UpdateClasses( sResult )
		{
			$("#earlyclass").html(sResult);
		}


		function selectAll(selectBox,selectAll) 
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

		function ValidatePrice( oPrice )
		{
			var bValid = true;
			var total = 0.00;

			// Remove any extra spaces
			oPrice.value = removeSpaces(oPrice.value);
			//Remove commas that would cause problems in validation
			oPrice.value = removeCommas(oPrice.value);

			// Validate the format of the price
			if (oPrice.value != "")
			{
				var rege = /^\d*\.?\d{0,2}$/
				var Ok = rege.exec(oPrice.value);
				if ( Ok )
				{
					oPrice.value = format_number(Number(oPrice.value),2);
				}
				else 
				{
					oPrice.value = "";
					alert("Prices must be numbers in currency format or blank.\nPlease correct to continue.");
					setfocus(oPrice);
					return false;
				}
			}
		}

		function NewTimeRow()
		{
			document.ClassForm.maxtimeid.value = parseInt(document.ClassForm.maxtimeid.value) + 1;
			document.ClassForm.maxdayid.value = parseInt(document.ClassForm.maxdayid.value) + 1;
			var tbl = document.getElementById("seriestime");
			var lastRow = tbl.rows.length;
			var newRow = parseInt(document.ClassForm.maxtimeid.value);
			var row = tbl.insertRow(lastRow);

			var cellZero = row.insertCell(0);
			cellZero.className = 'ref';

			var e1 = document.createElement('input');
			e1.type = 'hidden';
			e1.name = 'timeid' + newRow;
			e1.value = newRow;
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
			var slength = document.getElementById("instructorid0").length;
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
			var e5 = document.createElement('input');
			e5.type = 'text';
			e5.name = 'waitlistmax' + newRow;
			e5.value = '';
			e5.size = 4;
			e5.maxLength = 5;
			cell4.appendChild(e5);

			var cell5 = row.insertCell(5);
			cell5.className = 'firstday';
			var e6 = document.createElement('input');
			e6.type = 'hidden';
			e6.name = 'dayid' + newRow;
			e6.value = newRow;
			cell5.appendChild(e6);

			var e7 = document.createElement('input');
			e7.type = 'checkbox';
			e7.name = 'su' + newRow;
			cell5.appendChild(e7);

			var cell6 = row.insertCell(6);
			var e8 = document.createElement('input');
			e8.type = 'checkbox';
			e8.name = 'mo' + newRow;
			cell6.appendChild(e8);

			var cell7 = row.insertCell(7);
			var e9 = document.createElement('input');
			e9.type = 'checkbox';
			e9.name = 'tu' + newRow;
			cell7.appendChild(e9);

			var cell8 = row.insertCell(8);
			var e10 = document.createElement('input');
			e10.type = 'checkbox';
			e10.name = 'we' + newRow;
			cell8.appendChild(e10);

			var cell9 = row.insertCell(9);
			var e11 = document.createElement('input');
			e11.type = 'checkbox';
			e11.name = 'th' + newRow;
			cell9.appendChild(e11);

			var cell10 = row.insertCell(10);
			var e12 = document.createElement('input');
			e12.type = 'checkbox';
			e12.name = 'fr' + newRow;
			cell10.appendChild(e12);

			var cell11 = row.insertCell(11);
			var e13 = document.createElement('input');
			e13.type = 'checkbox';
			e13.name = 'sa' + newRow;
			cell11.appendChild(e13);

			var cell12 = row.insertCell(12);
			cell12.align = 'center';
			var e14 = document.createElement('input');
			e14.type = 'text';
			e14.id = 'starttime' + newRow;
			e14.name = 'starttime' + newRow;
			e14.value = '';
			e14.size = 8;
			e14.maxLength = 7;
			cell12.appendChild(e14);

			var cell13 = row.insertCell(13);
			cell13.align = 'center';
			var e15 = document.createElement('input');
			e15.type = 'text';
			e15.id = 'endtime' + newRow;
			e15.name = 'endtime' + newRow;
			e15.value = '';
			e15.size = 8;
			e15.maxLength = 7;
			cell13.appendChild(e15);

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
			for (var s=0; s <= parseInt(document.ClassForm.maxtimeid.value) ; s++)
			{
				op = document.createElement('OPTION');
				newText = document.createTextNode( document.ClassForm.firstname.value + ' ' + document.ClassForm.lastname.value );
				op.appendChild( newText );
				op.setAttribute( 'value', sNewInstructor );
				document.getElementById("instructorid"+s).appendChild(op);
			}
			document.ClassForm.firstname.value = "";
			document.ClassForm.lastname.value = "";
			alert('Instructor Added');
		}

		function GetSeasonDefaults()
		{
			var csi = document.ClassForm.classseasonid.options[document.ClassForm.classseasonid.selectedIndex].value;
			// Fire off Ajax routine
			doAjax('getseasondefaults.asp', 'csi=' + csi, 'UpdateSeasonDefaults', 'get', '0');
		}

		function UpdateSeasonDefaults( sString )
		{
			var exists;
			//alert( sString );
			// Blank out the date fields
			document.ClassForm.registrationstartdate.value = '';
			document.ClassForm.registrationenddate.value = '';
			document.ClassForm.publishstartdate.value = '';
			document.ClassForm.publishenddate.value = '';
			// Split the returned string
			var results = sString.split("|");
			if (results[0] != 'NULL')
			{
				document.ClassForm.registrationstartdate.value = results[0];
			}
			if (results[1] != 'NULL')
			{
				document.ClassForm.publishstartdate.value = results[1];
			}
			if (results[2] != 'NULL')
			{
				document.ClassForm.publishenddate.value = results[2];
			}
			if (results[3] != 'NULL')
			{
				document.ClassForm.registrationenddate.value = results[3];
			}
			
			if (results.length > 4)
			{
				for (x=4; x<=(results.length - 1); x++ )
				{
					pricetype = results[x].split(";");
					//alert('pricetype:[' + pricetype[0] + '] date: [' + pricetype[1] + ']');
					if ($("#registrationstartdate" + pricetype[0]).length > 0)
					{
						//alert($("registrationstartdate" + i).getAttribute("type"));
						if (document.getElementById("registrationstartdate" + pricetype[0]).getAttribute("type") != 'hidden')
						{
							document.getElementById("registrationstartdate" + pricetype[0]).setAttribute("value",pricetype[1]);
						}
					}
				}
			}
			else 
			{
				// Fill in the price registration start dates
				for (var i = parseInt(document.ClassForm.minpricetypeid.value); i < parseInt(document.ClassForm.maxpricetypeid.value)+1; i++ )
				{
					if ($("#registrationstartdate" + i).length > 0)
					{
						//alert($("registrationstartdate" + i).getAttribute("type"));
						if (document.getElementById("registrationstartdate" + i).getAttribute("type") != 'hidden')
						{
							document.getElementById("registrationstartdate" + i).setAttribute("value","");
						}
					}
				}
			}
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
			var h = (screen.height - 250)/2;
			popupWin = eval('window.open(url, name,"resizable,width=800,height=650,left=' + 100 + ',top=' + h + '")');
		}

		function CopyClass(iClassid)
		{
			if (confirm('Copy this to a new Class/Event?'))
			{
				location.href='class_copyclass.asp?classid=' + iClassid;
			}
		}

		function CopyToChild(iClassid)
		{
			if (confirm('Create a new Individual?'))
			{
				location.href='class_copyaschild.asp?classid=' + iClassid;
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
			if (document.ClassForm.searchkeywords.value.length > document.ClassForm.searchkeywords.getAttribute('maxlength'))
			{
				tabView.set('activeIndex',0);
				alert("The maximum length of search keywords is " + document.ClassForm.searchkeywords.getAttribute('maxlength') + " characters. \nPlease make this smaller.");
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
				if (! isValidDate(document.ClassForm.publishenddate.value))
				{
					tabView.set('activeIndex',4);
					alert("Publication end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.publishenddate.focus();
					return;
				}
			}

			// check registration end date
			if (document.ClassForm.registrationenddate.value != "")
			{
				if (! isValidDate(document.ClassForm.registrationenddate.value))
				{
					tabView.set('activeIndex',4);
					alert("Registration end date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.registrationenddate.focus();
					return;
				}
			}

			// check evaluation date
			if (document.ClassForm.evaluationdate.value != "")
			{
				if (! isValidDate(document.ClassForm.evaluationdate.value))
				{
					tabView.set('activeIndex',4);
					alert("Evaluation date should be a valid date in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.evaluationdate.focus();
					return;
				}
			}

			// Check pricing on the ones they checked
			for (var p = parseInt(document.ClassForm.minpricetypeid.value); p <= parseInt(document.ClassForm.maxpricetypeid.value); p++)
			{
				// Does it exist
				if ($("#pricetypeid" + p).length > 0)
				{
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
			for (var t=0; t <= parseInt(document.ClassForm.maxtimeid.value); t++)
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
			//return false;

			// all is ok, so submit and save the changes
			document.ClassForm.submit();
		}

		function SetUpPage()
		{
			setMaxLength();
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
		<font size="+1"><strong>Create Classes and Events</strong></font><br />
	</p>
	<!--END: PAGE TITLE-->

<%
	Dim sSql, oClass, iClassId, bIsSeriesParent

	iClassTypeId = request("classtypeid")

	If iClassTypeId = 1 Then 
		bIsSeriesParent = True 
		sClassTypeName  = "Series"
	Else 
		bIsSeriesParent = False
		sClassTypeName  = "Single"
	End If 

	response.write vbcrlf & "<form name=""ClassForm"" id=""ClassForm""  accept-charset=""UTF-8"" method=""post"" action=""class_createclass.asp"">" & vbcrlf
	response.write "<input type=""hidden"" name=""classtypeid"" id=""classtypeid"" value=""" & iClassTypeId & """ />" & vbcrlf
	response.write "<p>" & vbcrlf
	response.write "Name: <input type=""text"" id=""classname"" name=""classname"" value="""" size=""50"" maxlength=""50"" />" & vbcrlf
	response.write "&nbsp; <strong>This is a " & sClassTypeName & " Class/Event </strong>" & vbcrlf
	response.write "</p>" & vbcrlf
	response.write "<p>" & vbcrlf
	response.write "Season: "

	iFirstSeasonId = ShowSeasonPicks(0) 'In class_global_functions.asp

	response.write "</p>" & vbcrlf
%>
	<div id="demo" class="yui-navset">
		<ul class="yui-nav">
			<li><a href="#tab1"><em>Information</em></a></li>
			<li><a href="#tab2"><em>Categories</em></a></li>
			<li><a href="#tab3"><em>Waivers</em></a></li>
			<li><a href="#tab4"><em>Instructors</em></a></li>
			<li><a href="#tab5"><em>Dates</em></a></li>
			<li><a href="#tab6"><em>Purchasing</em></a></li>
			<li><a href="#tab7"><em>Occurs</em></a></li>
			<li><a href="#tab8"><em>Early Registration</em></a></li>
		</ul>            
		<div class="yui-content">

			<div id="tab1"> <!-- General Information -->
				<p>
					Description:<br /><textarea name="classdescription" id="classdescription" maxlength="6000" wrap="soft"></textarea>
				</p>
				<p>
					Image:
					<img src="adf" id="imgurlpic" name="imgurlpic" border="0" align="middle" width="180" height="180"  onerror="this.src = '../images/placeholder.png';" />
					<input type="<%if blnHasWP then %>hidden<%else%>text<%end if%>" name="imgurl" class="imageurl" id="imgurl" size="50" maxlength="255" />
					<% if blnHasWP then %>
						<input type="button" class="button" value="Change" onclick="showModal('Pick Image',65,80,'imgurl');" />
					<% else%>
						&nbsp; <input type="button" class="button" value="Browse..." onclick="javascript:doPicker('ClassForm.imgurl');" />
					 	&nbsp; &nbsp; <input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
					<% end if %>
					<br /><span id="imgalttag">Image Alt Tag:</span> <input type="text" name="imgalttag" value="" size="50" maxlength="255" />
					<br /><span id="imgalttabdesc"> The Image Alt Tag is a description used for ADA compliance.</span>
				</p>
				<p>
					Search Keywords:<br />
					<textarea name="searchkeywords" id="searchkeywords" maxlength="1024" wrap="soft"></textarea>
				</p>
				<p>
					Minimum Age: <input type="text" name="minage" value="" size="5" maxlength="4" />  &nbsp; 
					<% ShowAgeCheckPrecision 0, "minageprecisionid" %> &nbsp;  &nbsp;  &nbsp; 
					Maximum Age: <input type="text" name="maxage" value="" size="5" maxlength="4" />  &nbsp; 
					<% ShowAgeCheckPrecision 0, "maxageprecisionid" %>
				</p>
				<p>
					Check their age against this date: <input type="text" maxlength="10" class="datefield" name="agecomparedate" value="" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('agecomparedate');" /></span>
				</p>

<%				If bHasGenderRestrictions Then		%>
				<p>
					Gender Restriction: <% ShowGenderRestrictions iInitialGenderRestriction %>
				</p>
<%				Else								
					response.write vbcrlf & "<input type=""hidden"" name=""genderrestrictionid"" value=""" & iInitialGenderRestriction & """ />"
				End If								%>

				<p>
					Location: <% ShowLocationPicks Session("OrgID") %>
				</p>
				<p>
					Point of Contact: <% ShowPOCPicks Session("OrgID") %>
				</p>
		<%
			   'Supervisor ----------------------------------------------------------------
				If lcl_orghasfeature_class_supervisors Then 
				   response.write "<p>" & vbcrlf
				   response.write "Supervisor: " & vbcrlf

				   ShowSupervisorPicks 0  'In class_global_functions.asp

				   response.write "</p>" & vbcrlf
				Else 
				   response.write "<input type=""hidden"" name=""supervisorid"" id=""supervisorid"" value=""0"" />" & vbcrlf
				End if

			   'External URL and Text -----------------------------------------------------
				response.write "<p>" & vbcrlf
				response.write "   External URL: <input type=""text"" name=""externalurl"" id=""externalurl"" value="""" size=""50"" maxlength=""255"" />" & vbcrlf
				response.write "   <br />" & vbcrlf
				response.write "   External URL Text: <input type=""text"" name=""externallinktext"" value="""" size=""50"" maxlength=""255"" />" & vbcrlf
				response.write "</p>" & vbcrlf

			   'Receipt Notes -------------------------------------------------------------
				response.write "<p>" & vbcrlf
				response.write "   Receipt Notes:<br />" & vbcrlf
				response.write "			<textarea name=""notes"" id=""receiptnotes"" maxlength=""1024"" wrap=""soft""></textarea>" & vbcrlf
				response.write "</p>" & vbcrlf

			   'Display Roster to Public (Craig, CO custom request) -------------------
				If lcl_orghasfeature_custom_registration_craigco Then 
				   response.write "<p>" & vbcrlf
				   response.write "   Display Roster to Public: " & vbcrlf
				   response.write "   <input type=""checkbox"" name=""displayrosterpublic"" id=""displayrosterpublic"" value=""on"" />" & vbcrlf
				   response.write "</p>" & vbcrlf
				End If 
  %>
			</div>

			<div id="tab2"> <!-- Categories -->
				<div style="display: block;">
					<div>
						<% ShowCategories session("orgid") %>
					</div>
				</div>
			</div>

			<div id="tab3"> <!-- Waivers -->
	  		<div>
			  		<div class="rightbuttons">
 				  		<input type="button" class="assignbuttons" name="waivermgr" value="Manage Waivers" onclick="openWin2('class_waivers.asp','_blank')" /><br />
						<input type="checkbox" name="showTerms" id="showTerms" value="on" checked="checked" /> Show Terms
  					</div>
		  			<% ShowWaiverPicks 0 %>
  					<br /><input type="button" class="assignbuttons" name="NoInstructors" value="Clear Selection" onclick="selectAll(document.getElementById('waiverid'),false)" />
  					<div id="waivernote">
    					Note: To add new waivers, click on Manage Waivers, and create the waiver.
   						The new waiver will not appear in this list until after you save your changes to this page.
  					</div>
 				</div>
			</div>

			<div id="tab4"> <!-- Instructors -->
				<div>
					<div class="rightbuttons">
					<fieldset class="edit"><legend><strong> New Instructor </strong></legend>
						<table id="newinstructor" cellspacing="0" cellpadding="0" border="0">
							<tr><th align="center">First Name</th><th align="center">Last Name</th><th>&nbsp;</th></tr>
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
					<% ShowInstructorPicks 0 %>
					<br />(select for display only)
					<br /><input type="button" class="assignbuttons" name="NoWaivers" value="Clear Selection" onclick="selectAll(document.getElementById('instructorid'),false)" />
				</div>
			</div>

			<div id="tab5"> <!-- Critical Dates -->
				<p><br />
				<table id="criticaldates" border="0" cellpadding="1" cellspacing="3">
					<tr>
						<td align="right">Class/Event Starts:</td><td><input type="text" maxlength="10" class="datefield" name="startdate" value="" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span></td>
						<td align="right">Class/Event Ends:</td><td><input type="text" maxlength="10" class="datefield" name="enddate" value="" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span></td>
					</tr>
					<tr>
						<td align="right">Publication Starts:</td><td><input type="text" maxlength="10" class="datefield" name="publishstartdate" value="<%=GetDefaultSeasonDate( iFirstSeasonId, "publicationstartdate" ) %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishstartdate');" /></span></td>
						<td align="right">Publication Ends:</td><td><input type="text" maxlength="10" class="datefield" name="publishenddate" value="<%= GetDefaultSeasonDate( iFirstSeasonId, "publicationenddate" ) %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishenddate');" /></span>
					</td></tr>
					<tr>
						<td align="right">Registration Starts:<br />(for display only)</td><td><input type="text" maxlength="10" class="datefield" name="registrationstartdate" value="<%= GetDefaultSeasonDate( iFirstSeasonId, "registrationstartdate" ) %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationstartdate');" /></span></td>
						<td align="right">Registration Ends:</td><td><input type="text" maxlength="10" class="datefield" name="registrationenddate" value="<%= GetDefaultSeasonDate( iFirstSeasonId, "registrationenddate" ) %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationenddate');" /></span>
						</td>
					</tr>
					<tr>
						<td align="right">Send Evaluation:</td><td><input type="text" maxlength="10" class="datefield" name="evaluationdate" value="" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('evaluationdate');" /></span></td>
						<td>&nbsp;</td>
						</td>
					</tr>
				</table>
				</p>
			</div>

			<div id="tab6"> <!-- Purchasing -->
				<p>
					<strong>Pricing:</strong><br /><br />
					<table id="pricingtable" border="0" cellpadding="0" cellspacing="0">
	  				<tr>
						<th>Type</th>
     		<%
						If lcl_orghasfeature_gl_accounts Then 
							response.write "<th>Account</th>" & vbcrlf
						End If 
       %>
				       <th>Price</th>
       <%
						If ClassCanNeedMemberships() Then   'In class_global_functions.asp
							response.write "<th>Membership</th>" & vbcrlf
						Else 
							response.write "<th>&nbsp;</th>" & vbcrlf
						End If 
       %>			
						<th>Instructor %</th><th>Registration Starts</th></tr>
						<% GetPricing Session("OrgID"), iFirstSeasonId %>
					</table>
				</p>

				<p>
					Requires: <% ShowRegistrationPicks %>
				</p>
				<p>
					<input type="checkbox" id="publiccanonlyview" name="publiccanonlyview" /> Allow public to view but not purchase
				</p>

		<%
				If lcl_orghasfeature_discounts Then 
					response.write "<p>" & vbcrlf
					response.write "Discount: " & vbcrlf

					If lcl_orghasfeature_discounts Then 
						ShowPriceDiscountPicks Session("OrgID") 
					End If 

					response.write "</p>" & vbcrlf
				Else 
					response.write "<input type=""hidden"" name=""pricediscountid"" value=""0"" />" & vbcrlf
				End If 
  %> 
			</div>

			<div id="tab7"> <!-- Occurs -->
				<p>
					<input type="button" class="button" value="Add Row" id="addref" onClick="NewTimeRow()" />
					<table id="seriestime" border="0" cellpadding="0" cellspacing="0">
					<tr><th align="center" colspan="5">Activities</th><th align="center" colspan="9" class="firstday">Days</th></tr>
					<tr><th>Activity #</th><th>Instructor</th><th>Min</th><th>Max</th><th>Waitlist<br />Max</th>
					<th class="firstday">Su</th><th>Mo</th><th>Tu</th><th>We</th><th>Th</th><th>Fr</th><th>Sa</th>
					<th>Start Time</th><th>End Time</th></tr>
					<tr>
						<td class="ref">
							<input type="hidden" name="timeid0" value="0" />
							<input type="text" id="activity0" name="activity0" value="" size="10" maxlength="10" />
						</td>
						<td align="center"><% ShowInitialInstructorPicks 0, 0 %></td>
						<td align="center"><input type="text" name="min0" value="" size="4" maxlength="5" /></td>
						<td align="center"><input type="text" name="max0" value="" size="4" maxlength="5" /></td>
						<td align="center"><input type="text" name="waitlistmax0" value="" size="4" maxlength="5" /></td>
						<td class="firstday">
							<input type="hidden" name="dayid0" value="0" />
							<input type="checkbox" name="su0" />
						</td>
						<td><input type="checkbox" name="mo0" /></td>
						<td><input type="checkbox" name="tu0" /></td>
						<td><input type="checkbox" name="we0" /></td>
						<td><input type="checkbox" name="th0" /></td>
						<td><input type="checkbox" name="fr0" /></td>
						<td><input type="checkbox" name="sa0" /></td>
						<td align="center"><input type="text" id="starttime0" name="starttime0" value="" size="8" maxlength="7" /></td>
						<td align="center"><input type="text" id="endtime0" name="endtime0" value="" size="8" maxlength="7" /></td>
					</tr>
					</table>
					<input type="hidden" name="maxtimeid" value="0" />
					<input type="hidden" name="maxdayid" value="0" />
				</p>
			</div>

			<div id="tab8"> <!-- Early Registration -->
				<p>
					<input type="checkbox" id="allowearlyregistration" name="allowearlyregistration" />
					&nbsp; Allow Early Registration For This Class/Event
				</p>
				<p>
					Early Registration Start Date: &nbsp; <input type="text" id="earlyregistrationdate" name="earlyregistrationdate" value="" size="10" maxlength="10" />
					&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('earlyregistrationdate');" /></span>
				</p>
				<p>
					<strong>Select The Class/Event That Gets Early Registration:</strong><br /><br />
					<table cellpadding="5" cellspacing="0" border="0">
						<tr>
							<td align="right" valign="top" class="earlyreglabel">Season: &nbsp;</td>
							<td valign="top"><select id="earlyregistrationclassseasonid" name="earlyregistrationclassseasonid" onchange="javascript:ClassChange();">
										<% iEarlySeasonId = ShowEarlyRegistrationClassSeasons( ) %>
  								   </select> &nbsp; &nbsp;
							</td>
						</tr>
						<tr>
							<td align="right" valign="top" class="earlyreglabel">Class/Event: &nbsp</td>
							<td><span id="earlyclass"><% ShowEarlRegistrationClasses iEarlySeasonId %></span>
								<br />* multiple class selection allowed
								<br /><input type="button" class="assignbuttons" name="NoEarlyClasses" value="Clear Selection" onclick="selectAll(document.getElementById('earlyregistrationclassid'),false)" />
							</td>
						</tr>
					</table>
				</p>
			</div>

		</div>
	</div>
		
		<p>
			<input type="button" class="button" name="create" value="Create" onclick="ValidateForm();" />
		<p>
		</form>
		
	</div>
</div>
<!--END: PAGE CONTENT-->

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>


<%
'------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Sub GetPricing( iOrgID, iClassSeasonId )
'------------------------------------------------------------------------------
Sub GetPricing( iOrgID, iClassSeasonId )
	Dim sSql, oPrice, iMax, iMin

	iMax = 0
	iMin = 10000
	sSql = "SELECT pricetypeid, pricetypename, ismember, ISNULL(instructorpercent,0) AS instructorpercent, "
	sSql = sSql & " needsregistrationstartdate, isfee, isdropin, basepricetypeid FROM egov_price_types "
	sSql = sSql & " WHERE isactiveforclasses = 1 AND orgid = " & iOrgId & " ORDER BY displayorder"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSql, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		If CLng(oPrice("pricetypeid")) > CLng(iMax) Then 
			iMax = oPrice("pricetypeid")
		End If 
		If CLng(oPrice("pricetypeid")) < CLng(iMin) Then 
			iMin = oPrice("pricetypeid")
		End If 
		response.write vbcrlf & "<tr><td class=""type""><input type=""checkbox"" name=""pricetypeid""  id=""pricetypeid" & oPrice("pricetypeid") & """ value=""" & oPrice("pricetypeid") & """ />&nbsp;" & oPrice("pricetypename") 
		If oPrice("isfee") Then
			response.write " (fee)"
		ElseIf oPrice("isdropin") Then
			response.write " (one time)"
		ElseIf Not IsNull(oPrice("basepricetypeid")) Then
			response.write " (+)" 
		End If 
		response.write "</td>"

		if lcl_orghasfeature_gl_accounts then
			response.write "<td align=""center"">"
			ShowAccountPicks 0, oPrice("pricetypeid")  ' In class_global_functions.asp
			response.write "</td>"
		Else
			response.write vbcrlf & "<input type=""hidden"" name=""accountid" & oPrice("pricetypeid") & """ value=""0"" />"
		End If 

		response.write "<td align=""center""><input type=""text"" name=""amount" & oPrice("pricetypeid") & """ id=""amount" & oPrice("pricetypeid") & """ value="""" size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" /></td>"
		
		response.write "<td align=""center"">"
		If ClassCanNeedMemberships() Then
			If oPrice("ismember") Then
				' Show the membership picks for the one that requires membership
				'response.write "Membership: &nbsp; "
				ShowClassMembershipPicks 0, oPrice("pricetypeid")  ' In class_global_functions.asp
			Else
				response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oPrice("pricetypeid") & """ value=""0"" />"
			End If 
		Else
			response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oPrice("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		response.write "<td align=""center""><input type=""text"" name=""instructorpercent" & oPrice("pricetypeid") & """ id=""instructorpercent" & oPrice("pricetypeid") & """ value=""" & oPrice("instructorpercent") & """ size=""3"" maxlength=""3"" /></td>"

		If oPrice("needsregistrationstartdate") Then 
			response.write "<td align=""center""><input type=""text"" maxlength=""10"" class=""datefield"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value=""" & GetPriceTypeStartDate( iClassSeasonId, oPrice("pricetypeid") ) & """ />&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('registrationstartdate" & oPrice("pricetypeid") & "');"" /></span></td>"
		Else
			response.write "<td> &nbsp;<input type=""hidden"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value="""" /> </td>"
		End If 
		response.write "</tr>"	

		oPrice.MoveNext
	Loop 
	response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMin & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMax & """ />"

	oPrice.Close
	Set oPrice = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub ShowSeriesChildren( iParentClassId )
'------------------------------------------------------------------------------
Sub ShowSeriesChildren( ByVal iParentClassId )
	Dim sSql, oChildren, iRowCount

	sSql = "SELECT C.classid, C.classname, C.startdate, C.enddate, S.statusname "
	sSql = sSql & " FROM egov_class C, egov_class_status S "
	sSql = sSql & " WHERE C.parentclassid = " & iParentClassId & " AND C.statusid = S.statusid ORDER BY C.startdate, C.classname"
	iRowCount = 1

	Set oChildren = Server.CreateObject("ADODB.Recordset")
	oChildren.Open sSql, Application("DSN"), 0, 1

	Do While Not oChildren.EOF
		iRowCount = iRowCount + 1
		If ( iRowCOunt Mod 2 ) = 0 Then 
			response.write vbcrlf & "<tr class=""alt_row"">"
		Else 
			response.write vbcrlf & "<tr>"
		End If 
		response.write vbcrlf & "<td>" & oChildren("classname") & "</td><td align=""center"">" & oChildren("statusname") & "</td><td align=""center"">" & oChildren("startdate") & "</td><td align=""center"">" & oChildren("enddate") & "</td>"
		response.write "<td align=""center""><input type=""button"" class=""button"" name=""editchild"" value=""Edit"" onclick=""javascript:location.href='edit_class.asp?classid=" & oChildren("classid") & "';"" /></td></tr>"
		oChildren.MoveNext
	Loop 

	oChildren.Close
	Set oChildren = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub ShowLocationPicks( iOrgId )
'------------------------------------------------------------------------------
Sub  ShowLocationPicks( ByVal iOrgId )
	Dim sSql, oLocation

	sSql = "SELECT locationid, name FROM egov_class_location WHERE orgid = " & iOrgId & " ORDER BY name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSql, Application("DSN"), 0, 1

	If Not oLocation.EOF Then
		response.write vbcrlf & "<select name=""locationid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a Location</option>"
		Do While Not oLocation.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oLocation("locationid") & """ "
			response.write ">" & oLocation("name") & "</option>"
			oLocation.MoveNext
		Loop 
		response.write "</select>"
	End If 

	oLocation.Close
	Set oLocation = Nothing
End Sub 


'------------------------------------------------------------------------------
' Sub  ShowPOCPicks( iOrgId )
'------------------------------------------------------------------------------
Sub  ShowPOCPicks( ByVal iOrgId )
	Dim sSql, oPOC

	sSql = "SELECT pocid, name FROM egov_class_pointofcontact WHERE orgid = " & iOrgId & " ORDER BY name"

	Set oPOC = Server.CreateObject("ADODB.Recordset")
	oPOC.Open sSql, Application("DSN"), 0, 1

	If Not oPOC.EOF Then
		response.write vbcrlf & "<select name=""pocid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a POC</option>"
		Do While Not oPOC.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oPOC("pocid") & """ "
			response.write ">" & oPOC("name") & "</option>"
			oPOC.MoveNExt
		Loop 
		response.write "</select>"
	End If 

	oPOC.Close
	Set oPOC = Nothing
End Sub 


'------------------------------------------------------------------------------
' Sub  ShowInstructors( iClassId )
'------------------------------------------------------------------------------
Sub ShowInstructors( ByVal iClassId )
	Dim sSql, oInstructor, iRow

	iRow = 0
	sSql = "SELECT L.lastname, L.firstname FROM egov_class_instructor L, egov_class_to_instructor C"
	sSql = sSql & " WHERE L.instructorid = C.instructorid"
	sSql = sSql & " AND C.classid = " & iClassId & " ORDER BY L.lastname, L.firstname"

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSql, Application("DSN"), 0, 1

	If Not oInstructor.EOF Then
		Do While Not oInstructor.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oInstructor("firstname") & " " & oInstructor("lastname")
			oInstructor.MoveNext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No Instructor Assigned"
	End If 

	oInstructor.Close
	Set oInstructor = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub  ShowWaivers( iClassId )
'------------------------------------------------------------------------------
Sub ShowWaivers( ByVal iClassId )
	Dim sSql, oWaiver, iRow

	iRow = 0
	sSql = "SELECT waivername FROM egov_class_waivers W, egov_class_to_waivers C"
	sSql = sSql & " WHERE W.waiverid = C.waiverid"
	sSql = sSql & " AND C.classid = " & iClassId & " ORDER BY waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSql, Application("DSN"), 0, 1

	If Not oWaiver.EOF Then
		Do While Not oWaiver.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oWaiver("waivername") 
			oWaiver.movenext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No 3rd Party Waivers Assigned"
	End If 

	oWaiver.close
	Set oWaiver = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub ShowRegistrationPicks( )
'------------------------------------------------------------------------------
Sub ShowRegistrationPicks( )
	Dim sSql, oOption

	sSql = "SELECT optionid, optionname, optiondescription, requirestime FROM egov_registration_option ORDER BY optionid"

	Set oOption = Server.CreateObject("ADODB.Recordset")
	oOption.Open sSql, Application("DSN"), 1, 1

	If Not oOption.EOF Then
		response.write vbcrlf & "<select name=""optionid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select an Option</option>"
		Do While Not oOption.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oOption("optionid") & """ "
			response.write ">" & oOption("optionname") & " &ndash; " & oOption("optiondescription") & "</option>"
			oOption.movenext
		Loop 
		response.write "</select>"

		' Put in the hidden fields to check for time requirements
		oOption.MoveFirst
		Do While Not oOption.EOF
			response.write vbcrlf & "<input type=""hidden"" name=""requirestime" & oOption("optionid") & """ value=""" & oOption("requirestime") & """ />"
			oOption.MoveNext
		Loop

	End If 

	oOption.Close
	Set oOption = Nothing

End Sub 


'------------------------------------------------------------------------------
' Function GetPriceAmount( iPriceTypeId, iClassId )
'------------------------------------------------------------------------------
Function GetPriceAmount( iPriceTypeId, iClassId )
	Dim sSql, oAmount

	sSql = "SELECT amount FROM egov_class_pricetype_price WHERE pricetypeid = " & iPriceTypeId & " AND classid = " & iClassId

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSql, Application("DSN"), 0, 1

	If Not oAmount.EOF Then
		GetPriceAmount = FormatNumber(oAmount("amount"),2)
	Else 
		GetPriceAmount = ""
	End If 

	oAmount.close
	Set oAmount = Nothing

End Function


'------------------------------------------------------------------------------
' Function CheckDayOfWeek( iClassId, iDayOfWeek )
'------------------------------------------------------------------------------
Function CheckDayOfWeek( ByVal iClassId, ByVal iDayOfWeek)
	Dim sSql, oDOW

	sSql = "SELECT rowid FROM egov_class_dayofweek WHERE dayofweek = " & iDayOfWeek & " AND classid = " & iClassId

	Set oDOW = Server.CreateObject("ADODB.Recordset")
	oDOW.Open sSql, Application("DSN"), 0, 1

	If Not oDOW.EOF Then
		CheckDayOfWeek = " checked=""checked"" "
	Else 
		CheckDayOfWeek = ""
	End If 

	oDOW.Close
	Set oDOW = Nothing

End Function 


'------------------------------------------------------------------------------
' Sub ShowClassCategories( iClassId )
'------------------------------------------------------------------------------
Sub ShowClassCategories( ByVal iClassId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "SELECT categorytitle FROM egov_class_categories C, egov_class_category_to_class G "
	sSql = sSql & " WHERE C.categoryid = G.categoryid"
	sSql = sSql & " AND classid = " & iClassId & " ORDER BY categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSql, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		If iRow > 0 Then
			response.write ", "
		End If 
		response.write oCategory("categorytitle") 
		oCategory.movenext
		iRow = iRow + 1
	Loop 

	oCategory.Close
	Set oCategory = Nothing

End Sub 


'------------------------------------------------------------------------------
' Sub ShowPriceDiscountPicks( iOrgId )
'------------------------------------------------------------------------------
Sub ShowPriceDiscountPicks( ByVal iOrgId )
	Dim sSql, oDiscounts, iDiscountId

	'iDiscountId = GetClassPriceDiscount( iClassId )

	sSql = "SELECT pricediscountid, discountname FROM egov_price_discount WHERE orgid = " & iOrgId & " ORDER BY discountname"

	Set oDiscounts = Server.CreateObject("ADODB.Recordset")
	oDiscounts.Open sSql, Application("DSN"), 0, 1

'	If Not oDiscounts.EOF Then
	response.write vbcrlf & "<select name=""pricediscountid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">No Discount Applied</option>"
	Do While Not oDiscounts.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oDiscounts("pricediscountid") & """ "
		response.write ">" & oDiscounts("discountname") & "</option>"
		oDiscounts.MoveNext
	Loop 
	response.write vbcrlf & "</select>"
'	Else
'		response.write " &nbsp; No Discounts Exist"
'	End If 

	oDiscounts.Close
	Set oDiscounts = Nothing
End Sub 


'------------------------------------------------------------------------------
' Sub ShowClassTimes( iClassId )
'------------------------------------------------------------------------------
Sub ShowClassTimes( iClassId )
	Dim sSql, oTimes, iRow, sMin, sMax, sWaitlistmax

	iRow = 0
	sSql = "SELECT timeid, starttime, endtime, ISNULL(min,0) AS min, ISNULL(max, 0) AS max, "
	sSql = sSql & " ISNULL(waitlistmax,0) AS waitlistmax FROM egov_class_time WHERE classid = " & iClassId & " ORDER BY timeid"

	Set oTimes = Server.CreateObject("ADODB.Recordset")
	oTimes.Open sSql, Application("DSN"), 0, 1

	Do While Not oTimes.EOF
		iRow = iRow + 1
		If clng(oTimes("min")) = 0 Then
			sMin = ""
		Else
			sMin = oTimes("min")
		End If 
		If clng(oTimes("max")) = 0 Then
			sMax = ""
		Else
			sMax = oTimes("max")
		End If 
		If clng(oTimes("waitlistmax")) = 0 Then
			sWaitlistmax = ""
		Else
			sWaitlistmax = oTimes("waitlistmax")
		End If 
		response.write vbcrlf & "<tr>"
		response.write "<input type=""hidden"" name=""timeid" & iRow & """ value=""" & oTimes("timeid") & """ />"
		response.write "<td align=""center""><input type=""checkbox"" name=""removetime" & iRow & """ /></td>"
		response.write "<td><input type=""text"" name=""starttime" & iRow & """ value=""" & oTimes("starttime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""endtime" & iRow & """ value=""" & oTimes("endtime") & """ size=""8"" maxlength=""7"" /></td>"
		response.write "<td><input type=""text"" name=""min" & iRow & """ value=""" & sMin & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""max" & iRow & """ value=""" &sMax & """ size=""4"" maxlength=""5"" /></td>"
		response.write "<td><input type=""text"" name=""waitlistmax" & iRow & """ value=""" & sWaitlistmax & """ size=""4"" maxlength=""5"" /></td>"
		response.write vbcrlf & "</tr>"
		oTimes.movenext
	Loop 

	oTimes.Close
	Set oTimes = Nothing
	response.write vbcrlf & "<input type=""hidden"" name=""timecount"" value=""" & iRow & """ />"

End Sub 


'------------------------------------------------------------------------------
' Sub ShowCategories( iOrgId )
'------------------------------------------------------------------------------
Sub ShowCategories( ByVal iOrgId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "SELECT categoryid, categorytitle FROM egov_class_categories WHERE isroot = 0 AND orgid = " & iOrgId
	sSql = sSql & " ORDER BY categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSql, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		response.write vbcrlf & "<input type=""checkbox"" name=""categorycheckid"" value=""" & oCategory("categoryid") & """ "
		response.write " /> " & oCategory("categorytitle") & " <br />"
		oCategory.MoveNext 
	Loop 

	oCategory.Close
	Set oCategory = Nothing
End Sub 


'------------------------------------------------------------------------------
' Function GetDefaultSeasonDate( iClassSeasonId, sDateField )
'------------------------------------------------------------------------------
Function GetDefaultSeasonDate( ByVal iClassSeasonId, ByVal sDateField )
	Dim sSql, oSeason

	sSql = "SELECT " & sDateField & " FROM egov_class_seasons WHERE classseasonid = " & iClassSeasonId

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSql, Application("DSN"), 0, 1

	if not oSeason.EOF then
		If IsNull(oSeason(sDateField)) Then 
  			GetDefaultSeasonDate = ""
		Else 
  			GetDefaultSeasonDate = oSeason(sDateField)
		End If 
	End If 

	oSeason.Close
	Set oSeason = Nothing
End Function 


'------------------------------------------------------------------------------
' Function ShowEarlyRegistrationClassSeasons( )
'------------------------------------------------------------------------------
Function ShowEarlyRegistrationClassSeasons( )
	Dim sSql, oSeasons, iSelectedSeasonId

	iSelectedSeasonId = CLng(0)

	sSql = "SELECT C.classseasonid, C.seasonname FROM egov_class_seasons C, egov_seasons S  "
	sSql = sSql & " WHERE C.seasonid = S.seasonid AND orgid = " & SESSION("orgid")
	sSql = sSql & " ORDER BY C.seasonyear desc, S.displayorder DESC, C.seasonname"

	Set oSeasons = Server.CreateObject("ADODB.Recordset")
	oSeasons.Open sSql, Application("DSN"), 0, 1
	
	If Not oSeasons.EOF Then
		Do While NOT oSeasons.EOF
			If iSelectedSeasonId = CLng(0) Then
				iSelectedSeasonId = CLng(oSeasons("classseasonid"))
			End If 
			response.write vbcrlf & "<option value=""" & oSeasons("classseasonid") & """"  
			response.write ">" & oSeasons("seasonname") & "</option>"
			oSeasons.MoveNext
		Loop
	End If

	oSeasons.Close
	Set oSeasons = Nothing

	ShowEarlyRegistrationClassSeasons = iSelectedSeasonId

End Function 


'------------------------------------------------------------------------------
' Sub ShowEarlRegistrationClasses( iClassSeasonId )
'------------------------------------------------------------------------------
Sub ShowEarlRegistrationClasses( ByVal iClassSeasonId )
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
				response.write ">" & oRs("classname") & "</option>"
			oRs.MoveNext
		Loop
		response.write "</select>"
	End If 
	
	oRs.Close
	Set oRs = Nothing 

End Sub 


%>


