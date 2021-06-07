<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="class_global_functions.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: edit_class.asp
' AUTHOR: Steve Loar
' CREATED: 04/19/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This page allows the editing of classes and events
'
' MODIFICATION HISTORY
' 1.0   4/19/2006   Steve Loar - INITIAL VERSION
' 1.1	10/11/06	Steve Loar - Security, Header and nav changed
' 2.0	3/9/2007	Steve Loar - Total make over for Menlo Park Project
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
Dim iMaxTimeRows

sLevel = "../" ' Override of value from common.asp
iMaxTimeRows = 0

If Not UserHasPermission( Session("UserId"), "manage classes" ) Then
	response.redirect sLevel & "permissiondenied.asp"
End If 

%>


<html>
<head>
	<title>E-Gov Administration Console</title>
	
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="classes.css" />

	<script language="JavaScript" src="../scripts/ajaxLib.js"></script>
	<script language="JavaScript" src="../scripts/formatnumber.js"></script>
	<script language="JavaScript" src="../scripts/removespaces.js"></script>
	<script language="JavaScript" src="../scripts/removecommas.js"></script>
	<script language="JavaScript" src="../scripts/setfocus.js"></script>

<script language="Javascript">
<!--
	function selectAll(selectBox,selectAll) {
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
			if (document.getElementById("instructorid"+s))
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
			alert('Please enter a description.');
			document.ClassForm.classdescription.focus();
			return;
		}

		var bHasAge = false; 

		// check the minimum age
		if (document.ClassForm.minage.value.length > 0)
		{
			bHasAge = true;
			rege = /^\d+\.?\d?$/;
			Ok = rege.test(document.ClassForm.minage.value);

			if (! Ok)
			{
				alert("The minimum age must be a number with at most one decimal place, or be blank.");
				document.ClassForm.minage.focus();
				return;
			}
		}
		// check the maximum age
		if (document.ClassForm.maxage.value.length > 0)
		{
			bHasAge = true;
			rege = /^\d+\.?\d?$/;
			Ok = rege.test(document.ClassForm.maxage.value);

			if (! Ok)
			{
				alert("The maximum age must be a number with at most one decimal place, or be blank.");
				document.ClassForm.maxage.focus();
				return;
			}
		}

		if (bHasAge)
		{
			// check the agecomparedate
			if (document.ClassForm.agecomparedate.value == "")
			{
				alert("Please enter an age comparison date");
				document.ClassForm.agecomparedate.focus();
				return;
			}
			else
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.ClassForm.agecomparedate.value);
				if (! Ok)
				{
					alert("The age comparison date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
					document.ClassForm.agecomparedate.focus();
					return;
				}
			}
		}

		// check the minimum grade
//		if (document.ClassForm.mingrade.value.length > 0)
//		{
//			if (document.ClassForm.mingrade.value != 'K' && document.ClassForm.mingrade.value != 'k')
//			{
//				rege = /^\d+$/;
//				Ok = rege.test(document.ClassForm.mingrade.value);
//
//				if (! Ok)
//				{
//					alert("The minimum grade must be 'K', a number(1-12), or blank.");
//					document.ClassForm.mingrade.focus();
//					return;
//				}
//				if (parseInt(document.ClassForm.mingrade.value) > 12 || parseInt(document.ClassForm.mingrade.value) < 1)
//				{
//					alert("The minimum grade must be 'K', a number(1-12), or blank.");
//					document.ClassForm.mingrade.focus();
//					return;
//				}
//			}
//		}
		// check the maximum grade
//		if (document.ClassForm.maxgrade.value.length > 0)
//		{
//			if (document.ClassForm.maxgrade.value != 'K' && document.ClassForm.maxgrade.value != 'k')
//			{
//				rege = /^\d+$/;
//				Ok = rege.test(document.ClassForm.maxgrade.value);
//
//				if (! Ok)
//				{
//					alert("The maximum grade must be 'K', a number(1-12), or blank.");
//					document.ClassForm.maxgrade.focus();
//					return;
//				}
//				if (parseInt(document.ClassForm.maxgrade.value) > 12 || parseInt(document.ClassForm.maxgrade.value) < 1)
//				{
//					alert("The maximum grade must be 'K', a number(1-12), or blank.");
//					document.ClassForm.maxgrade.focus();
//					return;
//				}
//			}
//		}

		// check the length of the search key words
		if (document.ClassForm.searchkeywords.value.length > 1024)
		{
			alert("The maximum length of search keywords is 1024 characters. \nPlease make this smaller.");
			document.ClassForm.searchkeywords.focus();
			return;
		}

		// check the startdate
		if (document.ClassForm.startdate.value == "")
		{
			alert("Please enter a start date");
			document.ClassForm.startdate.focus();
			return;
		}
		else
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.startdate.value);
			if (! Ok)
			{
				alert("Start date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.startdate.focus();
				return;
			}
		}

		// check the enddate
		if (document.ClassForm.enddate.value == "")
		{
			alert("Please enter an end date");
			document.ClassForm.enddate.focus();
			return;
		}
		else
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.enddate.value);
			if (! Ok)
			{
				alert("End date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.enddate.focus();
				return;
			}
		}

		// check publish start date
		if (document.ClassForm.publishstartdate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.publishstartdate.value);
			if (! Ok)
			{
				alert("Publication start date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.publishstartdate.focus();
				return;
			}
		}

		// check publish end date
		if (document.ClassForm.publishenddate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.publishenddate.value);
			if (! Ok)
			{
				alert("Publication end date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.publishenddate.focus();
				return;
			}
		}

		// check evaluation date
		if (document.ClassForm.evaluationdate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.evaluationdate.value);
			if (! Ok)
			{
				alert("Evaluation date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.evaluationdate.focus();
				return;
			}
		}
		
		// check registration end date
//		if (document.ClassForm.registrationenddate.value == "")
//		{
//			alert("Please enter a registration end date");
//			document.ClassForm.registrationenddate.focus();
//			return;
//		}
//		else
		if (document.ClassForm.registrationenddate.value != "")
		{
			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
			Ok = rege.test(document.ClassForm.registrationenddate.value);
			if (! Ok)
			{
				alert("Registration end date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
				document.ClassForm.registrationenddate.focus();
				return;
			}
		}

		// check alternate date
//		if (document.ClassForm.alternatedate.value != "")
//		{
//			rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
//			Ok = rege.test(document.ClassForm.alternatedate.value);
//			if (! Ok)
//			{
//				alert("Alternate date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
//				document.ClassForm.alternatedate.focus();
//				return;
//			}
//		}

		// Check pricing on the ones they checked
		for (var p = parseInt(document.ClassForm.minpricetypeid.value); p <= parseInt(document.ClassForm.maxpricetypeid.value); p++)
		{
			// Does it exist
			if (document.getElementById("pricetypeid" + p))
			{
				//Is is checked
				if(document.getElementById("pricetypeid" + p).checked)
				{
					// Remove any extra spaces
					document.getElementById("amount" + p).value = removeSpaces(document.getElementById("amount" + p).value);
					//Remove commas that would cause problems in validation
					document.getElementById("amount" + p).value = removeCommas(document.getElementById("amount" + p).value);

					// Is the price formated correctly
					rege = /^\d+\.\d{2}$/;
					Ok = rege.test(document.getElementById("amount" + p).value);
					if (! Ok)
					{
						alert("Selected prices cannot be blank and must be in currency format.");
						document.getElementById("amount" + p).focus();
						return;
					}
					//Is the instructor % formatted correctly
					rege = /^\d+$/;
					Ok = rege.test(document.getElementById("instructorpercent" + p).value);
					if (! Ok)
					{
						alert("Selected instructor percentages cannot be blank and must be in the range of 0-100.");
						document.getElementById("instructorpercent" + p).focus();
						return;
					}
					if (parseInt(document.getElementById("instructorpercent" + p).value) > 100)
					{
						alert("Selected instructor percentages cannot be blank and must be in the range of 0-100.");
						document.getElementById("instructorpercent" + p).focus();
						return;
					}
					// check that there is a registration start date
					if (document.getElementById("registrationstartdate" + p).getAttribute("type") != 'hidden')
					{
						rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
						Ok = rege.test(document.getElementById("registrationstartdate" + p).value);
						if (! Ok)
						{
							alert("Selected registration start dates cannot be blank and should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
							document.getElementById("registrationstartdate" + p).focus();
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
			if (parseInt(document.getElementById("timeid" + t).value) != 0)
			{
				// If not marked for delete
				if (document.getElementById("delete" + t).checked == false)
				{
					if (document.getElementById("activity" + t).value != '')
					{
						if (document.getElementById("starttime" + t).value != '')
						{
							timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
							timeOk = timerege.test(document.getElementById("starttime" + t).value);

							if (! timeOk)
							{
								alert("The start time must be formatted as HH:MM(AM|PM).");
								document.getElementById("starttime" + t).focus();
								return;
							}
						}
						if (document.getElementById("endtime" + t).value != '')
						{
							timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
							timeOk = timerege.test(document.getElementById("endtime" + t).value);

							if (! timeOk)
							{
								alert("The end time must be formatted as HH:MM(AM|PM).");
								document.getElementById("endtime" + t).focus();
								return;
							}
						}
					}
					else
					{
						// The activity number is blank.
						alert("The activity number cannot be blank for existing activities.");
						document.getElementById("activity" + t).focus();
						return;
					}
				}
			}
			else
			{
				// handle new activities
				if (document.getElementById("activity" + t).value != '')
				{
					if (document.getElementById("starttime" + t).value != '')
					{
						timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
						timeOk = timerege.test(document.getElementById("starttime" + t).value);

						if (! timeOk)
						{
							alert("The start time must be formatted as HH:MM(AM|PM).");
							document.getElementById("starttime" + t).focus();
							return;
						}
					}
					if (document.getElementById("endtime" + t).value != '')
					{
						timerege = /^\d{1,2}:{1}\d{2}[aApP]{1}[mM]{1}$/;
						timeOk = timerege.test(document.getElementById("endtime" + t).value);

						if (! timeOk)
						{
							alert("The end time must be formatted as HH:MM(AM|PM).");
							document.getElementById("endtime" + t).focus();
							return;
						}
					}
				}
			}
		}

		//alert('OK');
		//return;

		// all is ok, so submit and save the changes
		document.ClassForm.submit();
	}

//-->
</script>

</head>

<body>

 
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 


<!--BEGIN PAGE CONTENT-->
<div id="content">
	<div id="centercontent">
	<p>
		<a href="javascript:location.href='class_list.asp'"><img src="../images/arrow_2back.gif" align="absmiddle" border="0" />&nbsp;Return to Class/Event List</a>
	</p>
	<p><input type="button" class="button" name="update1" value="Save Changes" onclick="ValidateForm();" /></p>
<%
		Dim sSql, oClass, iClassId, bIsSeriesParent

		iClassId = request("classid")
		bIsSeriesParent = False

		sSql = "Select classname, isnull(classseasonid,0) as classseasonid, classdescription, isnull(startdate,'') as startdate, isnull(enddate,'') as enddate, "
		sSql = sSql & " isnull(publishstartdate,'') as publishstartdate, isnull(publishenddate,'') as publishenddate, "
		sSql = sSql & " isnull(registrationstartdate,'') as registrationstartdate, isnull(registrationenddate,'') as registrationenddate, "
		sSql = sSql & " isnull(promotiondate,'') as promotiondate, isnull(evaluationdate,'') as evaluationdate, isnull(locationid,0) as locationid, "
		sSql = sSql & " isnull(alternatedate,'') as alternatedate, sexrestriction, searchkeywords, externalurl, externallinktext, noenddate, "
		sSql = sSQL & " isparent, egov_class.classtypeid, isnull(parentclassid,0) as parentclassid, statusid, cancelreason, "
		sSql = sSQL & " isnull(imgurl,'') as imgurl, isnull(imgalttag,'') as imgalttag, optionid, isnull(pricediscountid,0) as pricediscountid, "
		sSql = sSQL & " isnull(minage,0) as minage, isnull(maxage,99) as maxage, mingrade, maxgrade, isnull(pocid,0) as pocid, T.classtypename, "
		sSql = sSql & " isnull(membershipid,0) as membershipid, isnull(supervisorid,0) as supervisorid, notes, "
		sSql = sSql & " isnull(minageprecisionid, 0) as minageprecisionid, isnull(maxageprecisionid, 0) as maxageprecisionid, isnull(agecomparedate,'') as agecomparedate "
		sSql = sSQL & " FROM egov_class, egov_class_type T "
		sSql = sSQL & " WHERE T.classtypeid = egov_class.classtypeid and classid = " & iClassId & " and orgid = " & Session("OrgID")

		Set oClass = Server.CreateObject("ADODB.Recordset")
		oClass.Open sSQL, Application("DSN"), 3, 1
	
		If Not oClass.eof Then 
%>
		<form name="ClassForm" method="post" action="update_class.asp">
		<input type="hidden" name="classid" value="<%=iClassId%>" />
		<p>
			Name: <input type="text" name="classname" value="<%=oClass("classname")%>" size="50" maxlength="50" />
			
			<% 
				response.write "&nbsp; <strong>This is a " & oClass("classtypename") 
				If clng(oClass("parentclassid")) > 0 Then
					response.write " Individual Class/Event </strong>"
					response.write "&nbsp; <input type=""button"" class=""button"" name=""gotoparent"" value=""Edit Series"" onclick=""javascript:location.href='edit_class.asp?classid=" & oClass("parentclassid") & "';"" />"
					cCopyText = "Copy to New Series Individual"
				Else
					response.write " Class/Event </strong>"
					If clng(oClass("classtypeid")) = 1 Then 
						cCopyText = "Copy to New Series"
						bIsSeriesParent = True
					Else 
						cCopyText = "Copy to New Class/Event"
					End If 
				End If 
			%>
		</p>
		<p>
			<div id="cancelreason"> <%=oClass("cancelreason") %></div>
			Status: <strong><%=GetStatusName( oClass("statusid") )%></strong>  &nbsp; &nbsp;
			<% If clng(oClass("statusid")) = 1 Then %>
				<input type="button" class="button" name="cancel" value="Cancel Class/Event"  onclick="javascript:location.href='class_cancel.asp?classid=<%=iClassId%>';" />	&nbsp; 
			<% Else %>
				<input type="button" class="button" name="activate" value="Activate Class/Event"  onclick="javascript:location.href='class_changestatus.asp?classid=<%=iClassId%>&statusid=1';" />	&nbsp; 
			<% End If %>

		</p>
		<p>
			Season: <%	ShowClassSeasonFilterPicks oClass("classseasonid") ' In class_global_functions.asp %>
		</p>

		<fieldset class="edit"><legend><strong><a name="C"> Categories </a></strong></legend>
		<p>
			<% If CLng(session("orgid")) = CLng(0) Then %>
			<div class="categories"><% ShowClassCategories iClassId %></div>
			<div class="rightbuttons">
				<input type="button" class="assignbuttons" name="categories" value="Assign Categories" onclick="openWin1('list_picker.asp?classid=<%=iClassId%>&listtype=C','_blank')" />
				&nbsp; <input type="button" class="assignbuttons" name="categorymgr" value="Manage Categories" onclick="openWin2('category_mgmt.asp','_blank')" /> 
			</div>
			<% Else %>
			<div class="categories"><% ShowCategories iClassId, session("orgid") %></div>
			<div class="rightbuttons">
			&nbsp; <input type="button" class="assignbuttons" name="categorymgr" value="Manage Categories" onclick="openWin2('category_mgmt.asp','_blank')" /> 
			</div>
			<% End If %>
		</p>
		</fieldset>
		<fieldset class="edit"><legend><strong> General Information </strong></legend>
		<p>
			Description:<br /><textarea name="classdescription" id="classdescription"><%=oClass("classdescription")%></textarea>
		</p>
		<p>
			Image: <%	If oClass("imgurl") <> "" Then 
							response.write "<img src=""" & oClass("imgurl") & """ border=""0"" alt=""" & oClass("imgalttag") & """ /><br />"
						End If 
					%> URL: <input type="text" name="imgurl" value="<%=oClass("imgurl")%>" size="50" maxlength="255" /> &nbsp; 
					
					<input type="button" class="button" value="Browse..." onclick="javascript:doPicker('ClassForm.imgurl');" />
					 &nbsp; &nbsp; <input type="button" class="button" name="upload" value="Upload" onclick="openWin2('../docs/default.asp','_blank')" />
					<br /><span id="imgalttag">Image Alt Tag:</span> <input type="text" name="imgalttag" value="<%=oClass("imgalttag")%>" size="50" maxlength="255" />
					<br /><span id="imgalttabdesc"> The Image Alt Tag is a description used for ADA compliance.</span>
		</p>
		<p>
			Search Keywords:<br />
			<textarea name="searchkeywords" id="searchkeywords" maxlength="1024"><%=oClass("searchkeywords")%></textarea>
		</p>
		<p>
			Minimum Age:  
			<input type="text" name="minage" value="<%If clng(oClass("minage")) <> 0 Then 
																		response.write oClass("minage")
																	End If %>" size="4" maxlength="4" /> &nbsp; 
			<% ShowAgeCheckPrecision oClass("minageprecisionid"), "minageprecisionid" %> &nbsp;  &nbsp;  &nbsp; 

			Maximum Age:  
			<input type="text" name="maxage" value="<%If clng(oClass("maxage")) <> 99 Then 
																		response.write oClass("maxage")
																	End If %>" size="4" maxlength="4" /> &nbsp; 
			<% ShowAgeCheckPrecision oClass("maxageprecisionid"), "maxageprecisionid" %>
		</p>
		<p>
			Check registrant age against this date: <input type="text" maxlength="10" class="datefield" name="agecomparedate" value="<%	If oClass("agecomparedate") <> "1/1/1900" Then 
																						response.write oClass("agecomparedate")
																					End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('agecomparedate');" /></span>
		</p>
<!--
		<p>
			Minimum Grade: <input type="text" name="mingrade" value="<%'=oClass("mingrade")%>" size="3" maxlength="2" /> &nbsp; 
			Maximum Grade: <input type="text" name="maxgrade" value="<%'=oClass("maxgrade")%>" size="3" maxlength="2" /> 
		</p>
-->
		<p>
			<div class="rightbuttons">
				<input type="button" class="assignbuttons" name="locationmgr" value="Manage Locations" onclick="openWin2('location_mgmt.asp','_blank')" /> 
			</div>
			Location: <% ShowLocationPicks oClass("locationid"), Session("OrgID") %>
		</p>
		<p>
			Point of Contact: <% ShowPOCPicks oClass("pocid"), Session("OrgID") %>
		</p>
<%		If OrgHasFeature( "class supervisors" ) Then %>
			<p>
				Supervisor: <% ShowSupervisorPicks oClass("supervisorid")  ' In class_global_functions.asp%>
			</p>
<%		Else	%>
			<input type="hidden" name="supervisorid" value="0" />
<%		End If %>

		<p>
			External URL: <input type="text" name="externalurl" value="<%=oClass("externalurl")%>" size="50" maxlength="255" />
			<br />
			External URL Text: <input type="text" name="externallinktext" value="<%=oClass("externallinktext")%>" size="50" maxlength="255" />
		</p>
		<p>
			Receipt Notes:<br />
			<textarea name="notes" id="receiptnotes"><%=oClass("notes")%></textarea>
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong><a name="W"> Waivers </a></strong></legend>
			<p>
				<div class="categories">
					<% ShowWaiverPicks iClassId %>
					<br /><input type="button" class="assignbuttons" name="NoWaivers" value="Clear Selection" onclick="selectAll(document.getElementById('waiverid'),false)" />
				</div>
				<div class="rightbuttons">
					<!--<input type="button" class="assignbuttons" name="waivers" value="Assign Waiver Links" onclick="openWin1('list_picker.asp?classid=<%=iClassId%>&listtype=W','_blank')" />&nbsp;-->
					<input type="button" class="assignbuttons" name="waivermgr" value="Manage Waivers" onclick="openWin2('class_waivers.asp','_blank')" /> 
				</div>
				<div id="waivernote">
					Note: To add new waivers, click on Manage Waivers, and create the waiver.
					The new waiver will not appear in this list until after you save your changes to this page.
				</div>
			</p>
		</fieldset>

		<fieldset class="edit"><legend><strong><a name="I"> Instructors </a></strong></legend>
		<p>
			<div class="categories">
				<% ShowInstructorPicks iClassId %>
				<br />(select for display only)
				<br /><input type="button" class="assignbuttons" name="NoInstructors" value="Clear Selection" onclick="selectAll(document.getElementById('instructorid'),false)" />
			</div>
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
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong><a name="criticaldates"> Critical Dates </a></strong></legend>
		<p>
		<table id="criticaldates" border="0" cellpadding="1" cellspacing="3">
			<tr>
			<td align="right">Class/Event Starts:</td><td><input type="text" maxlength="10" class="datefield" name="startdate" value="<%	If oClass("startdate") <> "1/1/1900" Then 
																						response.write oClass("startdate")
																					End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('startdate');" /></span>
			</td>																							
			<td align="right">Class/Event Ends:</td><td><input type="text" maxlength="10" class="datefield" name="enddate" value="<%	If oClass("enddate") <> "1/1/1900" Then 
																					response.write oClass("enddate")
																				End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('enddate');" /></span>
			</td>
		</tr>
		<tr>
			<td align="right">Publication Starts:</td><td><input type="text" maxlength="10" class="datefield" name="publishstartdate" value="<%	If oClass("publishstartdate") <> "1/1/1900" Then 
																											response.write oClass("publishstartdate")
																										End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishstartdate');" /></span>
			<td align="right">Publication Ends:</td><td><input type="text" maxlength="10" class="datefield" name="publishenddate" value="<%	If oClass("publishenddate") <> "1/1/1900" Then 
																										response.write oClass("publishenddate")
																									End If %>" />&nbsp;<a href="#criticaldates"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publishenddate');" /></a>
		</td></tr>
		<tr>
			<td align="right">Registration Starts:<br />(for display only)</td><td><input type="text" maxlength="10" class="datefield" name="registrationstartdate" value="<% If oClass("registrationstartdate") <> "1/1/1900" Then 
																												response.write oClass("registrationstartdate")
																											End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationstartdate');" /></span>
			<td align="right">Registration Ends:</td><td><input type="text" maxlength="10" class="datefield" name="registrationenddate" value="<%	If oClass("registrationenddate") <> "1/1/1900" Then 
																												response.write oClass("registrationenddate")
																											End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('registrationenddate');" /></span>
		</td></tr>
		<tr>
			<td align="right">Send Evaluation:</td><td><input type="text" maxlength="10" class="datefield" name="evaluationdate" value="<% If oClass("evaluationdate") <> "1/1/1900" Then 
																										response.write oClass("evaluationdate")
																									End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('evaluationdate');" /></span>
<!--			
			<td align="right">Alternate Date:</td><td><input type="text" class="datefield" name="alternatedate" value="<%'	If oClass("alternatedate") <> "1/1/1900" Then 
																									'response.write oClass("alternatedate")
		 																								'End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('alternatedate');" /></span>
			</td>
-->
			<td>&nbsp;</td>
		</tr>
<!--
		<tr>
			<td align="right">Promotion Date:</td><td><input type="text" class="datefield" name="promotiondate" value="<% 'If oClass("promotiondate") <> "1/1/1900" Then 
																										'response.write oClass("promotiondate")
																									'End If %>" />&nbsp;<span class="calendarimg" style="cursor:hand;"><img src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('promotiondate');" /></span>
			<td align="right">&nbsp;</td><td>&nbsp;
		</td></tr>
-->
		</table>
		</p>
		</fieldset>

		<fieldset class="edit"><legend><strong> Purchasing </strong></legend>
		<p>
			Requires: <% ShowRegistrationPicks oClass("optionid") %>
		</p>
		<p>
			<table id="pricingtable" border="0" cellpadding="0" cellspacing="0">
				<caption>Pricing:</caption>
			<tr><th>Type</th>
<%			If OrgHasFeature( "gl accounts" ) Then %>
				<th>Account</th>
<%			End If %>
			<th>Price</th>
<%			If ClassCanNeedMemberships() Then  ' In class_global_functions.asp %>			
				<th>Membership</th>
<%			Else %>
				<th>&nbsp;</th>
<%			End If %>
			<th>Instructor %</th><th>Registration Starts</th></tr>
				<% GetPricing iClassId, Session("OrgID"), oClass("membershipid"), oClass("classseasonid") %>
			</table>
		</p>
<%		If OrgHasFeature("discounts") Then %>
		<p>
			Discount: <% 
						If OrgHasFeature("discounts") Then
							ShowPriceDiscountPicks oClass("pricediscountid"), Session("OrgID")
						End If %> 
		</p>
<%		Else %>
			<input type="hidden" name="pricediscountid" value="0" />
<%		End If %>
		</fieldset>

		<fieldset class="edit"><legend><strong> Occurs </strong></legend>
			<p>
				<input type="button" class="button" value="Add Row" id="addref" onClick="NewTimeRow()" />
				<table id="seriestime" border="0" cellpadding="0" cellspacing="0">
				<tr><th align="center" colspan="7">Activities</th><th align="center" colspan="10" class="firstday">Days</th></tr>
				<tr><th>Activity #</th><th>Instructor</th><th>Min</th><th>Max</th><th>Enrld</th><th>Waitlist<br />Max</th><th>Can-<br />celed</th>
				<th class="firstday">Su</th><th>Mo</th><th>Tu</th><th>We</th><th>Th</th><th>Fr</th><th>Sa</th>
				<th>Start Time</th><th>End Time</th><th>Delete<br />Days*</th></tr>
<%				iMaxTimeRows = ShowClassTimes( iClassId ) %>
				</table>
				<strong>* Activities cannot be completely deleted when there are people on the roster.</strong>
				<input type="hidden" name="maxtimeid" value="<%=iMaxTimeRows%>" />
				<input type="hidden" name="maxtimedayid" value="<%=iMaxTimeRows%>" />
			</p>
		</fieldset>



		<p>
			<input type="button" class="button" name="update2" value="Save Changes" onclick="ValidateForm();" />
			<!--<input type="button" class="button" name="copy" value="<%=cCopyText%>" onclick="CopyClass('<%=iClassId%>');" />-->
		</p>
		</form>
		<fieldset class="edit"><legend><strong> Copy </strong></legend>
			<form name="copyForm" method="post" action="class_copyclass.asp">
				<input type="hidden" name="classid" value="<%=iClassId%>" />
			<p>
				To Season: <% ShowClassSeasonFilterPicks 0  ' In class_global_functions.asp %> &nbsp; 
				<input type="checkbox" name="copyattendees" /> Bring Attendees as Waitlist &nbsp; 
				<input type="button" class="button" name="copy" value="<%=cCopyText%>" onclick="CopyClass('<%=iClassId%>');" />
			</p>
			</form>
		</fieldset>
		<% If bIsSeriesParent Then %>
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
				<form name="SingleForm" method="post" action="class_addsingletoseries.asp">
				<input type="hidden" name="parentclassid" value="<%=iClassId%>" />
				<% ShowAvailableClasses %> &nbsp; 
				<input type="button" class="button" name="addexisting" value="Add Existing Single" onclick="AddSingle();" />
				</form>
			</p>
			</fieldset>
		<% End If %>

<%		Else
			response.write "<p>No information could be found for that Class/Event.</p>"
		End If 
		
		oClass.close
		Set oClass = Nothing 
%>
	</div>
</div>
<!--END: PAGE CONTENT-->


<!--#Include file="../admin_footer.asp"-->  

</body>


</html>


<!--#Include file="class_global_functions.asp"-->  

<%
'--------------------------------------------------------------------------------------------------
' USER DEFINED SUBROUTINES AND FUNCTIONS
'--------------------------------------------------------------------------------------------------

'--------------------------------------------------------------------------------------------------
' Function ShowClassTimes( ByVal iClassId )
'--------------------------------------------------------------------------------------------------
Function ShowClassTimes( ByVal iClassId )
	Dim sSql, oTimes, sOldActivityId, iRowNo, iActivityCount

	sSql = "select T.timeid, T.activityno, T.min, T.max, T.waitlistmax, isnull(T.instructorid,0) as instructorid, T.enrollmentsize, T.iscanceled, "
	sSql = sSql & " D.timedayid, sunday, monday, tuesday, wednesday, thursday, friday, saturday, D.starttime, D.endtime" 
	sSql = sSql & " from egov_class_time T, egov_class_time_days D where T.timeid = D.timeid and T.classid = " & iClassId
	sSql = sSql & " order by T.timeid, D.timedayid"

	Set oTimes = Server.CreateObject("ADODB.Recordset")
	oTimes.Open sSQL, Application("DSN"), 0, 1

	If Not oTimes.EOF Then 
		sOldActivityId = ":a:"
		iRowNo = -1
		iActivityCount = 0
		Do While Not oTimes.EOF
			iRowNo = iRowNo + 1
			' Build the table of times that were originally created
			If sOldActivityId <> oTimes("activityno") Then 
				iActivityCount = iActivityCount + 1
				If iActivityCount Mod 2 = 0 Then 
					sRowClass = " class=""altrow"" "
				Else
					sRowClass = ""
				End If 
			End If 

			response.write vbcrlf & vbtab & "<tr" & sRowClass & ">"
			response.write "<td class=""ref""><input type=""hidden"" id=""timeid" & iRowNo & """  name=""timeid" & iRowNo & """ value=""" & oTimes("timeid") & """ />"

			If sOldActivityId <> oTimes("activityno") Then 
				response.write "<input type=""text"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""" & oTimes("activityno") & """ size=""10"" maxlength=""10"" />"
			Else
				'response.write oTimes("activityno") & " <input type=""hidden"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""" & oTimes("activityno") & """ />"
				response.write "&nbsp; <input type=""hidden"" id=""activity" & iRowNo & """ name=""activity" & iRowNo & """ value=""skip"" />"
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" 
			If sOldActivityId <> oTimes("activityno") Then 
				ShowInitialInstructorPicks iRowNo, oTimes("instructorid") 
			Else
				response.write "&nbsp;"
			End If 
			response.write "</td>"
			If sOldActivityId <> oTimes("activityno") Then 
				response.write "<td align=""center""><input type=""text"" name=""min" & iRowNo & """ value=""" & oTimes("min") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center""><input type=""text"" name=""max" & iRowNo & """ value=""" & oTimes("max") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center"">" & oTimes("enrollmentsize") & "<input type=""hidden"" id=""enrollmentsize" & iRowNo & """ name=""enrollmentsize" & iRowNo & """ value=""" & oTimes("enrollmentsize") & """ /></td>"
				response.write "<td align=""center""><input type=""text"" name=""waitlistmax" & iRowNo & """ value=""" & oTimes("waitlistmax") & """ size=""4"" maxlength=""5"" /></td>"
				response.write "<td align=""center""><input type=""checkbox""  id=""iscanceled" & iRowNo & """ name=""iscanceled" & iRowNo & """ "
				If oTimes("iscanceled") Then
					response.write "checked=""checked"" "
				End If 
				response.write "/>"
			Else
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
				response.write "<td align=""center"">&nbsp;</td>"
			End If
			response.write "<td class=""firstday"">"
			response.write "<input type=""hidden"" name=""timedayid" & iRowNo & """ value=""" & oTimes("timedayid") & """ />"
			response.write "<input type=""checkbox"" name=""su" & iRowNo & """"
			If oTimes("sunday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""mo" & iRowNo & """"
			If oTimes("monday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""tu" & iRowNo & """"
			If oTimes("tuesday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""we" & iRowNo & """"
			If oTimes("wednesday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""th" & iRowNo & """"
			If oTimes("thursday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""fr" & iRowNo & """"
			If oTimes("friday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td><input type=""checkbox"" name=""sa" & iRowNo & """"
			If oTimes("saturday") Then 
				response.write " checked=""checked"" "
			End If 
			response.write " /></td>"
			response.write "<td align=""center""><input type=""text"" id=""starttime" & iRowNo & """ name=""starttime" & iRowNo & """ value=""" & oTimes("starttime") & """ size=""8"" maxlength=""7"" /></td>"
			response.write "<td align=""center""><input type=""text"" id=""endtime" & iRowNo & """ name=""endtime" & iRowNo & """ value=""" & oTimes("endtime") & """ size=""8"" maxlength=""7"" /></td>"
			response.write "<td align=""center""><input type=""checkbox"" id=""delete" & iRowNo & """ name=""delete" & iRowNo & """ /></td>"
			response.write "</tr>"
			sOldActivityId = oTimes("activityno")
			oTimes.MoveNext
		Loop 
	Else
		' They do not have any time rows
		response.write vbcrlf & vbtab & "<tr>"
		response.write "<td class=""ref"">"
		response.write "<input type=""hidden"" id=""timeid0"" name=""timeid0"" value=""0"" />"
		response.write "<input type=""text"" id=""activity0"" name=""activity0"" value="""" size=""10"" maxlength=""10"" />"
		response.write "</td>"
		response.write "<td align=""center"">"
		ShowInitialInstructorPicks 0, 0 
		response.write "</td>"
		response.write "<td align=""center""><input type=""text"" name=""min0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center""><input type=""text"" name=""max0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center"">&nbsp;</td>"
		response.write "<td align=""center""><input type=""text"" name=""waitlistmax0"" value="""" size=""4"" maxlength=""5"" /></td>"
		response.write "<td align=""center""><input type=""checkbox"" id=""iscanceled0"" name=""iscanceled0"" /></td>"
		response.write "<td class=""firstday"">"
		response.write "<input type=""hidden"" id=""timedayid0"" name=""timedayid0"" value=""0"" />"
		response.write "<input type=""checkbox"" name=""su0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""mo0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""tu0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""we0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""th0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""fr0"" /></td>"
		response.write "<td><input type=""checkbox"" name=""sa0"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""starttime0"" name=""starttime0"" value="""" size=""8"" maxlength=""7"" /></td>"
		response.write "<td align=""center""><input type=""text"" id=""endtime0"" name=""endtime0"" value="""" size=""8"" maxlength=""7"" /></td>"
		response.write "</tr>"
		iRowNo = 0
	End If 

	oTimes.close
	Set oTimes = Nothing 

	ShowClassTimes = iRowNo

End Function  


'--------------------------------------------------------------------------------------------------
' Sub GetPricing( iOrgID )
'--------------------------------------------------------------------------------------------------
Sub GetPricing( iClassId, iOrgID, iMembershipId, iClassSeasonId )
	Dim sSql, oPrice, iMax, iMin

	iMax = 0
	iMin = 10000
	sSql = "Select pricetypeid, pricetypename, ismember, isnull(instructorpercent,0) as instructorpercent, needsregistrationstartdate, isfee, isdropin, basepricetypeid from egov_price_types "
	sSql = sSql & " where isactiveforclasses = 1 and orgid = " & iOrgId & " Order By displayorder"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		If clng(oPrice("pricetypeid")) > clng(iMax) Then 
			iMax = oPrice("pricetypeid")
		End If 
		If clng(oPrice("pricetypeid")) < clng(iMin) Then 
			iMin = oPrice("pricetypeid")
		End If 
		response.write vbcrlf & "<tr><td class=""type""><input type=""checkbox"" name=""pricetypeid""  id=""pricetypeid" & oPrice("pricetypeid") & """ value=""" & oPrice("pricetypeid") & """"
		If ClassHasPriceType( iClassId, clng(oPrice("pricetypeid")) ) Then
			response.write " checked=""checked"" "
		End If 
		response.write " />&nbsp;" & oPrice("pricetypename") 
		If oPrice("isfee") Then
			response.write " (fee)"
		ElseIf oPrice("isdropin") Then
			response.write " (one time)"
		ElseIf Not IsNull(oPrice("basepricetypeid")) Then
			response.write " (+)" 
		End If 
		response.write "</td>"

		If OrgHasFeature( "gl accounts" ) Then
			iAccountId = GetClassPricetypeAccount( iClassId, clng(oPrice("pricetypeid")) )
			response.write "<td align=""center"">"
			ShowAccountPicks iAccountId, oPrice("pricetypeid")  ' In common.asp
			response.write "</td>"
		Else
			response.write vbcrlf & "<input type=""hidden"" name=""accountid" & oPrice("pricetypeid") & """ value=""0"" />"
		End If 

		response.write "<td align=""center""><input type=""text"" name=""amount" & oPrice("pricetypeid") & """ id=""amount" & oPrice("pricetypeid") & """ value=""" & GetPriceAmount( oPrice("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" onchange=""ValidatePrice(this);"" /></td>"
		
		response.write "<td align=""center"">"
		If ClassCanNeedMemberships() Then
			If oPrice("ismember") Then
				' Show the membership picks for the one that requires membership
				ShowClassMembershipPicks iMembershipId, oPrice("pricetypeid")  ' In class_global_functions.asp
			Else
				response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oPrice("pricetypeid") & """ value=""0"" />"
			End If 
		Else 
			response.write "&nbsp;<input type=""hidden"" name=""membershipid" & oPrice("pricetypeid") & """ value=""0"" />"
		End If 
		response.write "</td>"

		response.write "<td align=""center""><input type=""text"" name=""instructorpercent" & oPrice("pricetypeid") & """ id=""instructorpercent" & oPrice("pricetypeid") & """ value=""" & GetClassPricetypeInstructorPercent( iClassId, oPrice("pricetypeid"), oPrice("instructorpercent") ) & """ size=""3"" maxlength=""3"" /></td>"

		If oPrice("needsregistrationstartdate") Then 
			response.write "<td align=""center""><input type=""text"" maxlength=""10"" class=""datefield"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value=""" & GetClassPricetypeRegistrationStart( iClassId, oPrice("pricetypeid"), GetDefaultSeasonDate( iClassSeasonId, "registrationstartdate" ) ) & """ />&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('registrationstartdate" & oPrice("pricetypeid") & "');"" /></span></td>"
		Else
			response.write "<td> &nbsp;<input type=""hidden"" name=""registrationstartdate" & oPrice("pricetypeid") & """ id=""registrationstartdate" & oPrice("pricetypeid") & """ value="""" /> </td>"
		End If 
		response.write "</tr>"	

		oPrice.movenext
	Loop 
	response.write vbcrlf & "<input type=""hidden"" name=""minpricetypeid"" value=""" & iMin & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""maxpricetypeid"" value=""" & iMax & """ />"

	oPrice.close
	Set oPrice = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function ClassHasPriceType( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function ClassHasPriceType( iClassId, iPriceTypeId )
	Dim sSql, oPrice

	sSql = "Select count(pricetypeid) as hits from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If clng(oPrice("hits")) > clng(0) Then
		ClassHasPriceType = True 
	Else 
		ClassHasPriceType = False 
	End If 

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeAccount( iClassId, iPriceTypeId )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeAccount( iClassId, iPriceTypeId )
	Dim sSql, oPrice

	sSql = "Select isnull(accountid,0) as accountid from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then
		GetClassPricetypeAccount = oPrice("accountid")
	Else 
		GetClassPricetypeAccount = 0
	End If 

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeInstructorPercent( iClassId, iPriceTypeId, iDefaultPercent )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeInstructorPercent( iClassId, iPriceTypeId, iDefaultPercent )
	Dim sSql, oPrice

	sSql = "Select isnull(instructorpercent,0) as instructorpercent from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then
		GetClassPricetypeInstructorPercent = oPrice("instructorpercent")
	Else 
		GetClassPricetypeInstructorPercent = iDefaultPercent
	End If 

	oPrice.close
	Set oPrice = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Function GetClassPricetypeRegistrationStart( iClassId, iPriceTypeId, sDefaultDate )
'--------------------------------------------------------------------------------------------------
Function GetClassPricetypeRegistrationStart( iClassId, iPriceTypeId, sDefaultDate )
	Dim sSql, oPrice

	sSql = "Select registrationstartdate from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	If Not oPrice.EOF Then
		GetClassPricetypeRegistrationStart = oPrice("registrationstartdate")
	Else 
		GetClassPricetypeRegistrationStart = sDefaultDate
	End If 

	oPrice.close
	Set oPrice = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function GetPriceAmount( iPriceTypeId, iClassId )
'--------------------------------------------------------------------------------------------------
Function GetPriceAmount( iPriceTypeId, iClassId )
	Dim sSql, oAmount, sAmount

	sSql = "Select amount from egov_class_pricetype_price where pricetypeid = " & iPriceTypeId & " and classid = " & iClassId

	Set oAmount = Server.CreateObject("ADODB.Recordset")
	oAmount.Open sSQL, Application("DSN"), 0, 1

	If Not oAmount.EOF Then
		sAmount = FormatNumber(oAmount("amount"),2)
		GetPriceAmount = Replace(sAmount,",","") ' take out commas
	Else 
		GetPriceAmount = ""
	End If 

	oAmount.close
	Set oAmount = Nothing

End Function


'--------------------------------------------------------------------------------------------------
' Function GetDefaultSeasonDate( iClassSeasonId, sDateField )
'--------------------------------------------------------------------------------------------------
Function GetDefaultSeasonDate( iClassSeasonId, sDateField )
	Dim sSql, oSeason

	sSql = "Select " & sDateField & " from egov_class_seasons where classseasonid = " & iClassSeasonId

	Set oSeason = Server.CreateObject("ADODB.Recordset")
	oSeason.Open sSQL, Application("DSN"), 0, 1

	If Not oSeason.EOF Then 
		If IsNull(oSeason(sDateField)) Then 
			GetDefaultSeasonDate = ""
		Else 
			GetDefaultSeasonDate = oSeason(sDateField)
		End If 
	Else
		GetDefaultSeasonDate = ""
	End If 

	oSeason.close
	Set oSeason = Nothing
End Function 


'--------------------------------------------------------------------------------------------------
' Sub GetPricing1( iClassId, iOrgID, iMembershipId )
'--------------------------------------------------------------------------------------------------
Sub GetPricing1( iClassId, iOrgID, iMembershipId )
	Dim sSql, oPrice, iRow

	iRow = 0
	sSql = "Select pricetypeid, pricetypename, ismember from egov_price_types where orgid = " & iOrgId & " Order By displayorder"

	Set oPrice = Server.CreateObject("ADODB.Recordset")
	oPrice.Open sSQL, Application("DSN"), 0, 1

	Do While Not oPrice.EOF
		iRow = iRow + 1
		response.write vbcrlf & "<tr><td><input type=""hidden"" name=""pricetypeid" & iRow & """ value=""" & oPrice("pricetypeid") & """ />" & oPrice("pricetypename") & "</td>"
		response.write "<td><input type=""text"" name=""amount" & iRow & """ value=""" & GetPriceAmount( oPrice("pricetypeid"), iClassId ) & """ size=""10"" maxlength=""9"" /></td><td>"
		If oPrice("ismember") Then
			' Show the membership picks for the one that requires membership
			response.write "Membership: &nbsp; "
			ShowMembershipPicks iMembershipId
		Else
			response.write "&nbsp;"
		End If 
		response.write "</td></tr>"	
		oPrice.movenext
	Loop 
	response.write vbcrlf & "<input type=""hidden"" name=""pricetypeidcount"" value=""" & iRow & """ />"

	oPrice.close
	Set oPrice = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowSeriesChildren( iParentClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowSeriesChildren( iParentClassId )
	Dim sSql, oChildren, iRowCOunt

	sSql = "Select C.classid, C.classname, C.startdate, C.enddate, S.statusname from egov_class C, egov_class_status S "
	sSql = sSql & " where C.parentclassid = " & iParentClassId & " and C.statusid = S.statusid Order By C.startdate, C.classname"
	iRowCOunt = 1

	Set oChildren = Server.CreateObject("ADODB.Recordset")
	oChildren.Open sSQL, Application("DSN"), 0, 1

	Do While Not oChildren.EOF
		iRowCOunt = iRowCOunt + 1
		If ( iRowCOunt Mod 2 ) = 0 Then 
			response.write vbcrlf & "<tr class=""alt_row"">"
		Else 
			response.write vbcrlf & "<tr>"
		End If 
		response.write vbcrlf & "<td>" & oChildren("classname") & "</td><td align=""center"">" & oChildren("statusname") & "</td><td align=""center"">" & oChildren("startdate") & "</td><td align=""center"">" & oChildren("enddate") & "</td>"
		response.write "<td align=""center""><input type=""button"" class=""button"" name=""editchild"" value=""Edit"" onclick=""javascript:location.href='edit_class.asp?classid=" & oChildren("classid") & "';"" /></td></tr>"
		oChildren.movenext
	Loop 

	oChildren.close
	Set oChildren = Nothing

'<tr class="alt_row"><td>Summer Music and Mixer - May 14th</td><td align="center">5/14/2006</td><td align="center">5/14/2006</td><td align="center"><input type="button" name="editchild" value="Edit" /></td></tr>
'					<tr><td>Summer Music and Mixer - June 11th</td><td align="center">6/11/2006</td><td align="center">6/11/2006</td><td align="center"><input type="button" name="editchild" value="Edit" /></td></tr>
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowLocationPicks( iLocationId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub  ShowLocationPicks( iLocationId, iOrgId )
	Dim sSql, oLocation

	sSql = "Select locationid, name from egov_class_location where orgid = " & iOrgId & " Order By name"

	Set oLocation = Server.CreateObject("ADODB.Recordset")
	oLocation.Open sSQL, Application("DSN"), 0, 1

	If Not oLocation.EOF Then
		response.write vbcrlf & "<select name=""locationid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a Location</option>"
		Do While Not oLocation.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oLocation("locationid") & """ "
				If clng(oLocation("locationid")) = clng(iLocationId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oLocation("name") & "</option>"
			oLocation.movenext
		Loop 
		response.write "</select>"
	End If 

	oLocation.close
	Set oLocation = Nothing
End Sub 



'--------------------------------------------------------------------------------------------------
' Sub  ShowPOCPicks( iPocId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub  ShowPOCPicks( iPocId, iOrgId )
	Dim sSql, oPOC

	sSql = "Select pocid, name from egov_class_pointofcontact where orgid = " & iOrgId & " Order By name"

	Set oPOC = Server.CreateObject("ADODB.Recordset")
	oPOC.Open sSQL, Application("DSN"), 0, 1

	If Not oPOC.EOF Then
		response.write vbcrlf & "<select name=""pocid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select a POC</option>"
		Do While Not oPOC.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oPOC("pocid") & """ "
				If clng(oPOC("pocid")) = clng(iPocId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oPOC("name") & "</option>"
			oPOC.movenext
		Loop 
		response.write "</select>"
	End If 

	oPOC.close
	Set oPOC = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowInstructors( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowInstructors( iClassId )
	Dim sSql, oInstructor, iRow

	iRow = 0
	sSql = "Select L.lastname, L.firstname from egov_class_instructor L, egov_class_to_instructor C"
	sSql = sSql & " where L.instructorid = C.instructorid"
	sSql = sSql & " and C.classid = " & iClassId & " Order By L.lastname, L.firstname"

	Set oInstructor = Server.CreateObject("ADODB.Recordset")
	oInstructor.Open sSQL, Application("DSN"), 0, 1

	If Not oInstructor.EOF Then
		Do While Not oInstructor.EOF
			If iRow > 0 Then
				response.write ", "
			End If 
			response.write oInstructor("firstname") & " " & oInstructor("lastname")
			oInstructor.movenext
			iRow = iRow + 1
		Loop 
	Else 
		response.write " No Instructor Assigned"
	End If 

	oInstructor.close
	Set oInstructor = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowWaivers( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowWaivers( iClassId )
	Dim sSql, oWaiver, iRow

	iRow = 0
	sSql = "Select waivername from egov_class_waivers W, egov_class_to_waivers C"
	sSql = sSql & " where W.waiverid = C.waiverid"
	sSql = sSql & " and C.classid = " & iClassId & " Order By waivername"

	Set oWaiver = Server.CreateObject("ADODB.Recordset")
	oWaiver.Open sSQL, Application("DSN"), 0, 1

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


'--------------------------------------------------------------------------------------------------
' Sub ShowRegistrationPicks( iOptionId )
'--------------------------------------------------------------------------------------------------
Sub ShowRegistrationPicks( iOptionId )
	Dim sSql, oOption

	sSql = "Select optionid, optionname, optiondescription, requirestime from egov_registration_option Order By optionid"

	Set oOption = Server.CreateObject("ADODB.Recordset")
	oOption.Open sSQL, Application("DSN"), 1, 1

	If Not oOption.EOF Then
		response.write vbcrlf & "<select name=""optionid"" >"
		'response.write vbcrlf & vbtab & "<option value=""0"">Select an Option</option>"
		Do While Not oOption.EOF
			response.write vbcrlf & vbtab & "<option value=""" & oOption("optionid") & """ "
				If clng(oOption("optionid")) = clng(iOptionId) Then 
					response.write " selected=""selected"" "
				End If 
			response.write ">" & oOption("optionname") & " &ndash; " & oOption("optiondescription") & "</option>"
			oOption.movenext
		Loop 
		response.write vbcrlf & "</select>"

		' Put in the hidden fields to check for time requirements
		oOption.movefirst
		Do While Not oOption.EOF
			response.write vbcrlf & "<input type=""hidden"" name=""requirestime" & oOption("optionid") & """ value=""" & oOption("requirestime") & """ />"
			oOption.movenext
		Loop 

	End If 

	oOption.close
	Set oOption = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Function CheckDayOfWeek(iClassId, iDayOfWeek)
'--------------------------------------------------------------------------------------------------
Function CheckDayOfWeek(iClassId, iDayOfWeek)
	Dim sSql, oDOW

	sSql = "Select rowid from egov_class_dayofweek where dayofweek = " & iDayOfWeek & " and classid = " & iClassId

	Set oDOW = Server.CreateObject("ADODB.Recordset")
	oDOW.Open sSQL, Application("DSN"), 0, 1

	If Not oDOW.EOF Then
		CheckDayOfWeek = " checked=""checked"" "
	Else 
		CheckDayOfWeek = ""
	End If 

	oDOW.close
	Set oDOW = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Function CheckClassInCategory( iClassId, iCategoryId )
'--------------------------------------------------------------------------------------------------
Function CheckClassInCategory( iClassId, iCategoryId )
	Dim sSql, oCat

	sSql = "Select rowid from egov_class_category_to_class where categoryid = " & iCategoryId & " and classid = " & iClassId

	Set oCat = Server.CreateObject("ADODB.Recordset")
	oCat.Open sSQL, Application("DSN"), 0, 1

	If Not oCat.EOF Then
		CheckClassInCategory = " checked=""checked"" "
	Else 
		CheckClassInCategory = ""
	End If 

	oCat.close
	Set oCat = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' Sub ShowCategories( iClassId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowCategories( iClassId, iOrgId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "Select categoryid, categorytitle from egov_class_categories where isroot = 0 and orgid = " & iOrgId & " Order By categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		'<input type="checkbox" name="dayofweek" value="1" CheckDayOfWeek(iClassId, 1) /> Sunday
		response.write vbcrlf & "<input type=""checkbox"" name=""categoryid"" value=""" & oCategory("categoryid") & """ "
		response.write CheckClassInCategory( iClassId, oCategory("categoryid") )
		response.write " /> " & oCategory("categorytitle") & " <br />"
		oCategory.movenext 
	Loop 

	oCategory.close
	Set oCategory = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassCategories( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassCategories( iClassId )
	Dim sSql, oCategory, iRow

	iRow = 0
	sSql = "Select categorytitle from egov_class_categories C, egov_class_category_to_class G where C.categoryid = G.categoryid"
	sSql = sSql & " and classid = " & iClassId & " Order By categorytitle"

	Set oCategory = Server.CreateObject("ADODB.Recordset")
	oCategory.Open sSQL, Application("DSN"), 0, 1

	Do While Not oCategory.EOF
		If iRow > 0 Then
			response.write ", "
		End If 
		response.write oCategory("categorytitle") 
		oCategory.movenext
		iRow = iRow + 1
	Loop 

	oCategory.close
	Set oCategory = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowPriceDiscountPicks( iClassId, iOrgId )
'--------------------------------------------------------------------------------------------------
Sub ShowPriceDiscountPicks( iPriceDiscountId, iOrgId )
	Dim sSql, oDiscounts

'	iDiscountId = GetClassPriceDiscount( iClassId )

	sSql = "Select pricediscountid, discountname from egov_price_discount Where orgid = " & iOrgId & " Order By discountname"

	Set oDiscounts = Server.CreateObject("ADODB.Recordset")
	oDiscounts.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select name=""pricediscountid"" >"
	response.write vbcrlf & vbtab & "<option value=""0"">No Discount Applied</option>"
	Do While Not oDiscounts.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oDiscounts("pricediscountid") & """ "
			If clng(oDiscounts("pricediscountid")) = clng(iPriceDiscountId) Then 
				response.write " selected=""selected"" "
			End If 
		response.write ">" & oDiscounts("discountname") & "</option>"
		oDiscounts.movenext
	Loop 
	response.write vbcrlf & "</select>"
'	Else
'		response.write " &nbsp; No Discounts Exist"
'	End If 

	oDiscounts.close
	Set oDiscounts = Nothing
End Sub 


'--------------------------------------------------------------------------------------------------
' Sub ShowClassTimes( iClassId )
'--------------------------------------------------------------------------------------------------
Sub ShowClassTimesold( iClassId )
	Dim sSql, oTimes, iRow, sMin, sMax, sWaitlistmax

	iRow = 0
	sSql = "Select timeid, starttime, endtime, isnull(min,0) as min, isnull(max, 0) as max, isnull(waitlistmax,0) as waitlistmax from egov_class_time Where classid = " & iClassId & " Order By timeid"

	Set oTimes = Server.CreateObject("ADODB.Recordset")
	oTimes.Open sSQL, Application("DSN"), 0, 1

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

	oTimes.close
	Set oTimes = Nothing
	response.write vbcrlf & "<input type=""hidden"" name=""timecount"" value=""" & iRow & """ />"

End Sub 


'--------------------------------------------------------------------------------------------------
' Sub  ShowAvailableClasses( )
'--------------------------------------------------------------------------------------------------
Sub ShowAvailableClasses( )
	Dim sSql, oList

	sSql = "Select classid, classname from egov_class "
	sSql = sSql & " where classtypeid = 3 and orgid = " & Session("OrgID") & " Order By classname"
	'response.write sSql & "<br />"

	Set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 0, 1

	response.write vbcrlf & "<select border=""0"" name=""classid"">"
	Do While Not oList.EOF
		response.write "<option value=""" & oList("classid") & """>" & oList("classname") & "</option>"
		oList.movenext
	Loop 
	response.write vbcrlf & "</select>"

	oList.close
	Set oList = Nothing

End Sub 
%>


