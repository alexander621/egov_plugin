<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: class_signup.asp
' AUTHOR: Steve Loar
' CREATED: 03/16/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This handles the signup process for classes and events.
'
' MODIFICATION HISTORY
' 1.0	03/16/2006  Steve Loar  - Initial version
' 1.1	05/10/2006	Steve Loar - Added citizen search
' 1.2	10/17/2006	Steve Loar - Security, Header and nav changed
' 2.0	01/22/2007	Steve Loar - New Family structure applied
' 2.1	02/19/2008	Steve Loar - Changes for Early Registration
' 2.2	05/28/2008  David Boyer - Added Override Discount
' 2.3	01/07/2008  David Boyer - Added "DisplayRosterPublic" check for Craig, CO custom registration fields.
' 2.4	11/30/2009  David Boyer - Added "pants size" to team registration section
' 2.5	11/30/2009  David Boyer - Now pull team/pants sizes from database
' 2.6	04/26/2010	Steve Loar - Added prompt for registrations on classes/events that are in the past
' 3.0	03/30/2011	Steve Loar - Name search re-worked to work off AJAX to fill selects
' 3.1	07/22/2011	Steve Loar - Added AJAX call to getuserinfo.asp on user name pick change
' 3.2	10/10/2011	Steve Loar - Added Gender Restriction
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
Dim bRegistrationBlocked, iAvailability, sUserType, iUserid, sResidentDesc, iMemberCount
Dim iTimeId, bMultiWeeks, sSearchName, sResults, bIsInPast, bPreLoadSelects, bHasGenderRestrictions
Dim iGenderNotRequiredId, sRequiredGender
		intMaxEnroll = 0
		intEnrolled = 0
		intWaitlist = 0

'If they are coming directly to this page without a selected class, take them to the roster list so they can pick one.
If request("classid") = "" Then 
	response.redirect "roster_list.asp"
End If 

bRegistrationBlocked = False 
iAvailability        = 1
sLevel               = "../"  'Override of value from common.asp
lcl_setupCoachFields = "N"
bIsInPast            = False 
bPreLoadSelects      = False 

'Check the page availability and user access rights in one call
PageDisplayCheck "registration", sLevel	 'In common.asp

iItemTypeId = GetItemTypeId( "recreation activity" )  'This is what kind of thing they are buying - in class_global_functions.asp

session("RedirectPage") = "../classes/class_signup.asp?classid=" & request("classid") & "&timeid=" & request("timeid")
session("RedirectLang") = "Return to Class/Events Signup"

sUserType    = "P"
bMultiWeeks  = False 
iMemberCount = 0

If request("egovuserid") <> "" Then 
	iUserId = CLng("0" & request("egovuserid"))
Else 
	
	If session("eGovUserId") <> "" Then 
		iUserid = Session("eGovUserId")
	Else 
		iUserId = CLng(0)
	End If 
End If 

If IsNull(iUserid) Then
	iUserid = CLng(0)
End If 

'First find out what resident type they are
 sUserType = GetUserResidentType( iUserid )

'If they are not one of these (R, N), we have to figure which they are
 If sUserType <> "R" And sUserType <> "N" Then 
  	'This leaves E and B - See if they are a resident, also
   	sUserType = GetResidentTypeByAddress( iUserid, Session("OrgID") )
 End If 

 sResidentDesc = GetResidentTypeDesc( sUserType )	' in class_global_functions.asp

'See if a timeid was passed
 If request("timeid") <> "" Then 
   	iTimeId = request("timeid")
 Else 
   	iTimeId = 0
 End If 

'Get the availability of the selected time
 iAvailability = GetActivityAvailability( iTimeId )		' In class_global_functions.asp

'See if a search term was passed
 If request("searchname") <> "" Then 
	sSearchName = request("searchname")
 Else 
	If session("searchname") <> "" Then 
		sSearchName = session("searchname")
		bPreLoadSelects = True 
	Else 
		sSearchName = ""
	End If 
 End If 
'See if a search term was passed
 If request("searchname2") <> "" Then 
	sSearchName2 = request("searchname2")
 Else 
	If session("searchname2") <> "" Then 
		sSearchName2 = session("searchname2")
		bPreLoadSelects = True 
	Else 
		sSearchName2 = ""
	End If 
 End If 
 'response.write "sSearchName = " & session("searchname")

 If request("results") <> "" Then 
   	sResults = request("results")
 Else 
   	sResults = ""
 End If 

 If request("searchstart") <> "" Then 
   	sSearchStart = request("searchstart")
 Else 
   	sSearchStart = -1
 End If 

'Check for org features
 lcl_orghasfeature_residency_verification      = orghasfeature("residency verification")
 lcl_orghasfeature_custom_registration_craigco = orghasfeature("custom_registration_craigco")
 bHasGenderRestrictions = orgHasFeature("gender restriction")

%>
<html lang="en">
<head>
	<meta charset="UTF-8">
 	<title>E-Gov Administration Console {Class/Event Signup}</title>

	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="classes.css" />

	<script src="https://code.jquery.com/jquery-1.5.min.js"></script>

	<script src="../scripts/ajaxLib.js"></script>
	<script src="../scripts/formatnumber.js"></script>
	<script src="../scripts/removespaces.js"></script>
	<script src="../scripts/removecommas.js"></script>
	<script src="../scripts/setfocus.js"></script>
	<script src="../scripts/formvalidation_msgdisplay.js"></script>

	 <script>
	 <!--

		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=PurchaseForm", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

		function SearchCitizens( iSearchStart )
		{
			var optiontext;
			var optionchanged;
			//alert(document.BuyerForm.searchname.value);
			var searchtext = document.BuyerForm.searchname.value;
			var searchchanged = searchtext.toLowerCase();

			iSearchStart = parseInt(iSearchStart) + 1;
			
			for (x=iSearchStart; x < document.BuyerForm.egovuserid.length ; x++)
			{
				optiontext = document.BuyerForm.egovuserid.options[x].text;
				optionchanged = optiontext.toLowerCase();
				if (optionchanged.indexOf(searchchanged) != -1)
				{
					document.BuyerForm.egovuserid.selectedIndex = x;
					document.BuyerForm.results.value = 'Possible Match Found.';
					document.getElementById('searchresults').innerHTML = 'Possible Match Found.';
					document.BuyerForm.searchstart.value = x;
					document.BuyerForm.submit();
					return;
				}
			}
			document.BuyerForm.results.value = 'No Match Found.';
			document.getElementById('searchresults').innerHTML = 'No Match Found.';
			document.BuyerForm.searchstart.value = -1;
		}

		function ClearSearch()
		{
			document.BuyerForm.searchstart.value = -1;
		}

		function UserPick()
		{
			document.BuyerForm.searchname.value = '';
			document.BuyerForm.results.value = '';
			document.getElementById('searchresults').innerHTML = '';
			document.BuyerForm.searchstart.value = -1;
			document.BuyerForm.submit();
		}

		function UpdateFamily( )
		{
			//location.href='../dirs/family_members.asp?userid=' + iUserId;
			location.href='../dirs/family_list.asp?userid=' + $("#egovuserid").val();
		}

		function EditUser( )
		{
			location.href='../dirs/update_citizen.asp?userid=' + $("#egovuserid").val();
		}

		function NewUser()
		{
			location.href='../dirs/register_citizen.asp';
		}


		function doInitialNameSearchChange()
		{
			var sType = $("#optiontype").val();
			if ($("#searchname").val() != "")
			{
				var sReturnTo = "UpdateInitialApplicants";
				var onChange = "doFamilyPick";
				if (sType == "tickets")
				{
					sReturnTo = "UpdateInitialRegistrants";
					onChange = "UpdateUserId";
				}
				// Try to get a drop down of citizen names
				doAjax('getcitizenpicks.asp', 'searchname=' + $("#searchname").val() + '&onchange=' + onChange, sReturnTo, 'get', '0');
			}
			else
			{
				alert('Please enter a name before searching.');
				$("#searchname").focus();
			}
		}

		function UpdateInitialApplicants( sResult )
		{
			//alert("Back");
			$("#applicant").html( sResult );
			if (sResult.substr(0,6) == "Select")
			{
				$("#edituserbtn").css("visibility", "visible");
				$("#egovuserid").val(<%=iUserId%>);
				doFamilyPick();
			}
			else
			{
				//$("#egovuserid").val(0);
				$("#edituserbtn").css("visibility", "hidden");
				$("#registrant").html( "<input type='hidden' name='familymemberid' id='familymemberid' value='0' /><span id='nomatch'>No Registrants Found.</span>" );
				$("#updatefamilymembersbtn").css("visibility", "hidden");
			}
		}

		function UpdateInitialRegistrants( sResult )
		{
			//alert("Back");
			$("#applicant").html( sResult );
			if (sResult.substr(0,6) == "Select")
			{
				$("#egovuserid").val(<%=iUserId%>);
				$("#edituserbtn").css("visibility", "visible");
				$('#userid').val( $("#egovuserid").val() );
			}
			else
			{
				$("#edituserbtn").css("visibility", "hidden");
			}
		}

		function doNameSearchChange()
		{
			var sType = $("#optiontype").val();
			if ($("#searchname").val() != "")
			{
				var sReturnTo = "UpdateApplicants";
				var onChange = "doFamilyPick";
				if (sType == "tickets")
				{
					sReturnTo = "UpdateRegistrants";
					onChange = "UpdateUserId";
				}
				// Try to get a drop down of citizen names
				doAjax('getcitizenpicks.asp', 'searchname=' + $("#searchname").val() + '&searchname2=' + $("#searchname2").val() + '&onchange=' + onChange, sReturnTo, 'get', '0');
			}
			else
			{
				alert('Please enter a name before searching.');
				$("#searchname").focus();
			}
		}

		function UpdateUserId( )
		{
			$('#userid').val( $("#egovuserid").val() );
			// fire off getpricecheck
			doAjax('getpricecheck.asp', 'userid=' + $("#egovuserid").val() + "&classid=" + $("#classid").val() + "&membershipid=" + $("#membershipid").val(), 'PickTicketPrice', 'get', '0');
		}

		function UpdateRegistrants( sResult )
		{
			//alert("Back");
			$("#applicant").html( sResult );
			if (sResult.substr(0,6) == "Select")
			{
				$("#edituserbtn").css("visibility", "visible");
				$('#userid').val( $("#egovuserid").val() );
				// fire off getpricecheck
				doAjax('getpricecheck.asp', 'userid=' + $("#egovuserid").val() + "&classid=" + $("#classid").val() + "&membershipid=" + $("#membershipid").val(), 'PickTicketPrice', 'get', '0');
			}
			else
			{
				$("#edituserbtn").css("visibility", "hidden");
			}
		}

		function PickTicketPrice( sResult )
		{
			var total = 0.00;
			var sTotalString;

			$("#completebtn").attr('disabled', '');

			if ( sResult != 'none')
			{
				// Uncheck all the pricetypes
				for (x=parseInt($("#minpricetypeid").val()); x <= parseInt($("#maxpricetypeid").val()); x++)
				{
					if ($("#pricetypeid" + x).length != 0) 
						$("#pricetypeid" + x).removeAttr("checked");
				}

				//alert( sResult );
				var  pricetypes = new Array();
				pricetypes = sResult.split(',');

				$.each( pricetypes, 
				function(index, value) { 
					$("#pricetypeid" + value).attr('checked','checked'); 
					total += Number(eval('document.PurchaseForm.amount' + value + '.value'));
					$("#totalprice").val( total );
				});
			}
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);

			// Try to update the user info here??
			doAjax('getuserinfo.asp', 'citizenuserid=' + $("#egovuserid").val(), 'updateUserInfoDisplay', 'get', '0');
		}

		function UpdateApplicants( sResult )
		{
			//alert("Back");
			$("#applicant").html( sResult );
			if (sResult.substr(0,6) == "Select")
			{
				$("#edituserbtn").css("visibility", "visible");
				//$('#userid').val( $("#egovuserid").val() );
				doFamilyPick();
			}
			else
			{
				$("#edituserbtn").css("visibility", "hidden");
				$("#registrant").html( "<input type='hidden' name='familymemberid' id='familymemberid' value='0' /><span id='nomatch'>No Registrants Found.</span>" );
				$("#updatefamilymembersbtn").css("visibility", "hidden");
			}
		}

		function doFamilyPick()
		{
			//alert($("#membershipid").val());
			$('#userid').val( $("#egovuserid").val() );

			$("#completebtn").attr('disabled', '');

			// call the familymember pick creation script
			//doAjax('getfamilypicks.asp', 'egovuserid=' + $("#egovuserid").val() + "&membershipid=" + $("#membershipid").val(), 'UpdateFamilyPicks', 'get', '0');
			if ($("#egovuserid").val() == "0")
				GetFamilyPicks( 'none' );
			else
				doAjax('getpricecheck.asp', 'userid=' + $("#egovuserid").val() + "&classid=" + $("#classid").val() + "&membershipid=" + $("#membershipid").val(), 'GetFamilyPicks', 'get', '0');
			//GetFamilyPicks( 'none' );
		}

		function GetFamilyPicks( sResult )
		{
			var total = 0.00;
			var sTotalString;

			if ( sResult != 'none')
			{
				// Uncheck all the pricetypes
				for (x=parseInt($("#minpricetypeid").val()); x <= parseInt($("#maxpricetypeid").val()); x++)
				{
					if ($("#pricetypeid" + x).length != 0) 
						$("#pricetypeid" + x).removeAttr("checked");
				}

				//alert( sResult );
				var  pricetypes = new Array();
				pricetypes = sResult.split(',');

				$.each( pricetypes, 
				function(index, value) { 
					$("#pricetypeid" + value).attr('checked','checked'); 
					total += Number(eval('document.PurchaseForm.amount' + value + '.value'));
					$("#totalprice").val( total );
				});
			}
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
			doAjax('getfamilypicks.asp', 'egovuserid=' + $("#egovuserid").val() + "&membershipid=" + $("#membershipid").val(), 'UpdateFamilyPicks', 'get', '0');
		}

		function UpdateFamilyPicks( sResult )
		{
			$("#registrant").html( sResult );
			if (sResult.substr(1,6) == "select")
			{
				$("#updatefamilymembersbtn").css("visibility", "visible");
			}
			else
			{
				$("#updatefamilymembersbtn").css("visibility", "hidden");
			}
			// Try to update the user info here??
			doAjax('getuserinfo.asp', 'citizenuserid=' + $("#egovuserid").val(), 'updateUserInfoDisplay', 'get', '0');
		}

		function updateUserInfoDisplay( sResult )
		{
			$("#userinfo").html( sResult );

			// if registration is bloced then disable the Add To Cart Button
			if ($("#registrationblocked").val() == 'yes')
			{
				$("#completebtn").attr('disabled', 'disabled');
			}
		}

		function AutoSelect(iTimeId)
		{
			// IF they start typing in a special price then check the Other Price radio
			var radioLength = document.PurchaseForm.pricetypeid.length;
			if(radioLength == undefined) {
				return;
			}
			var i = radioLength - 1;
			document.PurchaseForm.pricetypeid[i].checked = true;
		}

		function askThenValidate()
		{
			if ( confirm("This Class/Event has already ended. Are you certain you wish to add someone to this?") )
			{
				ValidateForm();
			}
		}

		function ValidateForm()	
		{
			var iPriceTypeCount  = 0;
			var lcl_return_false = "N";
			var lcl_focus        = "";

			// check that we have a purchaser
			if ($("#egovuserid").val() == "0")
			{
				lcl_focus = "searchname";
				inlineMsg(document.getElementById("searchname").id,'<strong>Required Field Missing: </strong>A valid purchaser must be selected.',10,'searchname');
				lcl_return_false = "Y";
			}
			else
			{
				// check that we have a registrant
				if ($("#familymemberid").val() == "0")
				{
					lcl_focus = "egovuserid";
					inlineMsg(document.getElementById("egovuserid").id,'<strong>Required Field Missing: </strong>A valid registrant must be selected. This user does not have any family members.',10,'egovuserid');
					lcl_return_false = "Y";
				}
			}

			// Check that a price is picked and that the amounts are formatted correctly if they are buying.
			if (document.PurchaseForm.buyorwait[0].checked) 
			{ 
				// Buy is 0 Wait is 1
				//alert(document.PurchaseForm.minpricetypeid.value);
				//alert(document.PurchaseForm.maxpricetypeid.value);
				for (var p = parseInt(document.PurchaseForm.minpricetypeid.value); p <= parseInt(document.PurchaseForm.maxpricetypeid.value); p++) 
				{
					// Does it exist
					if (document.getElementById("pricetypeid" + p)) 
					{
						// Is is checked
						if(document.getElementById("pricetypeid" + p).checked) 
						{
							iPriceTypeCount += 1;

							// Remove any extra spaces
							document.getElementById("amount" + p).value = removeSpaces(document.getElementById("amount" + p).value);

							// Remove commas that would cause problems in validation
							document.getElementById("amount" + p).value = removeCommas(document.getElementById("amount" + p).value);

							// Is the price formated correctly and not blank
							rege = /^\d+\.\d{2}$/;
							Ok = rege.test(document.getElementById("amount" + p).value);
							if (! Ok) 
							{
								lcl_focus = "amount" + p;
								inlineMsg(document.getElementById("amount" + p).id,'<strong>Required Field Missing: </strong>Selected prices cannot be blank and must be in currency format.',10,'amount' + p);
								lcl_return_false = "Y";
							}

							// Check if there is a dropindate entered and is in correct format
							if (document.getElementById("dropindate" + p)) 
							{
								rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
								Ok = rege.test(document.getElementById("dropindate" + p).value);
								if (! Ok) 
								{
									lcl_focus = "pricetypeid" + p;
									inlineMsg(document.getElementById("dropindate" + p).id,'<strong>Invalid Value: </strong>The Drop In date should be in the format of MM/DD/YYYY.',10,'dropindate' + p);
									lcl_return_false = "Y";
								}
							}
						}
					}
				}

				// Make sure that at least one thing was checked.
				if (iPriceTypeCount == 0) 
				{
					lcl_focus = "pricetypeid" + parseInt(document.PurchaseForm.minpricetypeid.value);
					inlineMsg(document.getElementById("pricetypeid" + parseInt(document.PurchaseForm.minpricetypeid.value)).id,'<strong>Required Field Missing: </strong>Please select at least one price.',10,'pricetypeid' + parseInt(document.PurchaseForm.minpricetypeid.value));
					lcl_return_false = "Y";
				}
			}

		<% if lcl_orghasfeature_custom_registration_craigco then %>

			//Validate Team Registration fields
			if(document.getElementById("displayrosterpublic")) {
   			if(document.getElementById("displayrosterpublic").value=="True") {

         if(document.getElementById("teamreg_coach_enabled").value == 'BOTH' || document.getElementById("teamreg_coach_enabled").value == 'INTERNAL ONLY') {
   					    //Check to see if a "coach type" has been selected.
    			   		//If so then Full Name and at least one of the phone numbers and/or email are required.
   	    				if(document.getElementById("rostercoachtype").value != "") {

         						//Build the daytime phone
				    							lcl_dayphone = document.getElementById("skip_volcoachday_areacode").value;
								    			lcl_dayphone = lcl_dayphone + document.getElementById("skip_volcoachday_exchange").value;
    											lcl_dayphone = lcl_dayphone + document.getElementById("skip_volcoachday_line").value;

    											//Build the cell phone
				    							lcl_cellphone = document.getElementById("skip_volcoachcell_areacode").value;
								    			lcl_cellphone = lcl_cellphone + document.getElementById("skip_volcoachcell_exchange").value;
    											lcl_cellphone = lcl_cellphone + document.getElementById("skip_volcoachcell_line").value;
	
     		  	  			//Atleast one method of contact is required
 			      	  		if(lcl_dayphone=="" && lcl_cellphone=="" && document.getElementById("rostervolunteercoachemail").value=="") {
     		      					lcl_focus = "skip_volcoachday_areacode";
					           		inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Required Field Missing: </strong>One method of contact must be entered.',10,'skip_volcoachday_line');
    	  		 		     	lcl_return_false = "Y";
	  	      		  	} else {
 		  	      	  			//Validate the Email
	   				  	      	if(document.getElementById("rostervolunteercoachemail").value != "" ) {
     					  	  	    	//var rege = /^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*\.(\w{2}|(com|net|org|edu|mil|gov|biz|us))$/;
									var rege = /.+@.+\..+/i;
    	  		   			  	  	var Ok = rege.test(document.getElementById("rostervolunteercoachemail").value);

    	  	  		   			  	if (! Ok) {
   	      	       		 			lcl_focus = "rostervolunteercoachemail";
	     	        		    			inlineMsg(document.getElementById("rostervolunteercoachemail").id,'<strong>Invalid Value: </strong>The volunteer coach email must be in a valid format.',10,'rostervolunteercoachemail');
					    	         	    lcl_return_false = "Y";
		  	  		  		  	     }
    	  					      }

    	  	  		  			 //Validate the Cell Phone
                  if(lcl_cellphone!="") {
             								lcl_cell_areacode = document.getElementById("skip_volcoachcell_areacode").value;
         			    					lcl_cell_exchange = document.getElementById("skip_volcoachcell_exchange").value;
         							    	lcl_cell_line     = document.getElementById("skip_volcoachcell_line").value;

             								if(lcl_cellphone.length < 10) {
         					  	  	  		lcl_focus = "skip_volcoachcell_areacode";
	  	   	      	  		  			inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Cell Phone.',10,'skip_volcoachcell_line');
	  	   	  	      	  				lcl_return_false = "Y";
	  	  	  					       } else {
         					  	  	  		var cellPhone = new Number(lcl_cell_areacode+lcl_cell_exchange+lcl_cell_line);
	  	         		  	  				if(cellPhone.toString() == "NaN") {
    	  	      	  		  	  			lcl_focus = "skip_volcoachcell_areacode";
	  	            	   							inlineMsg(document.getElementById("skip_volcoachcell_line").id,'<strong>Invalid Value: </strong>Cell Phone must be numeric.',10,'skip_volcoachcell_line');
	  	  	  	   	      	  				lcl_return_false = "Y";
	  	  	  	   	  				    }
    		  	  	  				   }
    		  	  	  			 }

    	  		  	  			 //Validate the Day Phone
	  	      		  			 if ( lcl_dayphone!="" ) {
            					  		lcl_day_areacode = document.getElementById("skip_volcoachday_areacode").value;
	  	  		   	    		  	lcl_day_exchange = document.getElementById("skip_volcoachday_exchange").value;
				  	  	     	    	lcl_day_line     = document.getElementById("skip_volcoachday_line").value;
	
    		  	  	  	  	 		if (lcl_dayphone.length < 10) {
        	  		  	  	  			lcl_focus = "skip_volcoachday_areacode";
	  	         	  		  				inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>One or more numbers are missing from the Day Phone.',10,'skip_volcoachday_line');
	  				   	      	  	  	lcl_return_false = "Y";
	  	  	  		   		    	} else {
     		  	  				  		    var dayPhone = new Number(lcl_day_areacode+lcl_day_exchange+lcl_day_line);
	
    		  	     	  	  				if (dayPhone.toString() == "NaN") {
               					  	  		lcl_focus = "skip_volcoachday_areacode";
	  	        	  		  	  	  		inlineMsg(document.getElementById("skip_volcoachday_line").id,'<strong>Invalid Value: </strong>Day Phone must be numeric.',10,'skip_volcoachday_line');
	  	  	  		       			  	  	lcl_return_false = "Y";
		  	 	     		        		}
    					  		 	      }
				    			       }
     			    			}

      	  		  		//Validate the Full Name
  	      	  			if (document.getElementById("rostervolunteercoachname").value=="") {
     	      	  			lcl_focus = "rostervolunteercoachname";
		   	  	  	    		inlineMsg(document.getElementById("rostervolunteercoachname").id,'<strong>Required Field Missing: </strong>Volunteer Coach - Full Name.',10,'rostervolunteercoachname');
    	  	 	  	  			lcl_return_false = "Y";
	  	     		  		}
					       }
         }

    					//Validate the Grade
         if(document.getElementById("teamreg_grade_enabled").value == 'BOTH' || document.getElementById("teamreg_grade_enabled").value == 'INTERNAL ONLY') {
   					    if (document.getElementById("rostergrade").value == "") {
      						    lcl_focus = "rostergrade";
    				  						inlineMsg(document.getElementById("rostergrade").id,'<strong>Required Field Missing: </strong>Grade.',10,'rostergrade');
		      				   	lcl_return_false = "Y";
				 			    } else {
				   					    var rosterGrade = new Number(document.getElementById("rostergrade").value);
    			 				    if (rosterGrade.toString() == "NaN") {
            							lcl_focus = "rostergrade";
	      	    			 			inlineMsg(document.getElementById("rostergrade").id,'<strong>Invalid Value: </strong>Grade must be numeric.',10,'rostergrade');
				    				    			lcl_return_false = "Y";
    			 				    }
				    				}
         }
    		}
   }
		<% end if %>

			// Check that the enrollment max available is not being exceeded and if so ask them about doing so

			// If the ticket field exists check that a quantity is entered
			var bexists = eval(document.PurchaseForm["quantity"]);
			if (bexists) 
			{
				if (document.getElementById("quantity").value == "") 
				{
					lcl_focus = "quantity";
					inlineMsg(document.getElementById("quantity").id,'<strong>Required Field Missing: </strong>Ticket quantity',10,'quantity');
					lcl_return_false = "Y";
				}
				else
				{
					var rege = /^\d+$/;
					var Ok   = rege.test(document.getElementById("quantity").value);

					if (! Ok) 
					{
						lcl_focus = "quantity";
						inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity must be a number.',10,'quantity');
						lcl_return_false = "Y";
					}
					else
					{
						/* Old redundant code
						// Check that the quantity is not more than what is available if they are buying
						if (document.PurchaseForm.buyorwait.value == 'B') 
						{
							var iTimeId;
							var iAvail, iQty;
	
							iTimeId = getSelectedRadioValue(document.PurchaseForm.timeid)
							//iTimeId = document.PurchaseForm.timeid.value;
							//alert(iTimeId);
							//get the availability for the select time 
							iAvail = Number(eval('document.PurchaseForm.avail' + iTimeId + '.value'));
							iQty   = Number(eval('document.PurchaseForm.quantity.value'));
							//check that the ticket qty is not greater than what is available;
							if (iQty > iAvail) 
							{
								lcl_focus = "quantity";
								inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity cannot be greater than the availability.',10,'quantity');
								lcl_return_false = "Y";
							}
						}
						*/
					}
				}

				if (parseInt(document.PurchaseForm.availability.value) <= 0) 
				{
					var response = confirm("This activity is full. Do you wist to continue this registration anyway?");
					if ( response == false ) 
					{
						lcl_return_false = "Y";
					}
				}
				else
				{
					if ((parseInt(document.PurchaseForm.availability.value) - parseInt(document.PurchaseForm.quantity.value)) < 0) 
					{
						var response = confirm("The quantity input exceeds the availablity of this activity. Do you wist to continue this registration anyway?");
						if ( response == false ) 
						{
							lcl_return_false = "Y";
						}
					}
				}

				if (lcl_return_false == "Y") 
				{
					if(lcl_focus != "")
					{
						document.getElementById(lcl_focus).focus();
					}
					return false;

				}
				else
				{	
					//alert('Successful');  // For Ticketed events
					document.PurchaseForm.submit();
				}

			}
			else
			{
				if (parseInt(document.PurchaseForm.availability.value) <= 0) 
				{
					var response = confirm("This activity is full. Do you wist to continue this registration anyway?");
					if ( response == false ) 
					{
						lcl_return_false = "Y";
					}
				}

				if (lcl_return_false=="Y") 
				{
					if(lcl_focus != "") 
					{
						document.getElementById(lcl_focus).focus();
					}
					return false;
				}
				else
				{
					//there is an error with ShowFamilyMembers that if none exist the validation errors on this ajax call
					//because the family member dropdown list does not display with no family members.
	
					// Fire off AJAX check of age restrictions for registrations
					doAjax('check_age_restrictions.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&classid=' + document.PurchaseForm.classid.value, 'AgeCheckReturn', 'get', '0');
				}
			}
		}

<%	If bHasGenderRestrictions Then	%>
		
		function AgeCheckReturn( sResult )
		{
			var passed = false;

			//alert( sResult );
			if (sResult == "PASSED")
			{
				passed = true;
			}
			else 
			{
				if (confirm("The selected family member does not meet the age requirements of this activity. \nDo you wish to register them anyway?"))
				{
					passed = true;
				}
			}

			if (passed == true)
			{
				// if the class requires a gender match
				if ($("#requiredgender").val() != 'N')
				{
					// Get the Gender of the Family Member 
					doAjax('getfamilymembergender.asp', 'familymemberid=' + $("#familymemberid").val(), 'GenderCheckReturn', 'get', '0');
				}
				else
				{
					// Put call here to check for duplicate enrollment since this class does not require a gender
					doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + $("#familymemberid").val() + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
				}
			}
			
		}

		function GenderCheckReturn( sGender )
		{
			// If the gender matches the Required gender then go on, else ask about it
			if (sGender == $("#requiredgender").val())
			{
				// Put call here to check for duplicate enrollment since this class does not require a gender
				doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + $("#familymemberid").val() + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
			}
			else
			{
				if (confirm("The selected family member does not meet the gender restrictions of this activity. \nDo you wish to register them anyway?"))
				{
					// Put call here to check for duplicate enrollment.
					doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + $("#familymemberid").val() + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
				}
			}
		}

<%	Else							%>
		function AgeCheckReturn( sResult )
		{
			//alert( sResult );
			if (sResult == "PASSED")
			{
				// Put call here to check for duplicate enrollment.
				doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
			}
			else 
			{
				if (confirm("The selected family member does not meet the age requirements of this activity. \nDo you wish to register them anyway?"))
				{
					// Put call here to check for duplicate enrollment.
					doAjax('check_duplicate_enrollment.asp', 'familymemberid=' + document.PurchaseForm.familymemberid.options[document.PurchaseForm.familymemberid.selectedIndex].value + '&timeid=' + document.PurchaseForm.timeid.value, 'DupCheckReturn', 'get', '0');
				}
			}
		}
<%	End If							%>

		function DupCheckReturn( sResult )
		{
			if (sResult == "NOTFOUND")
			{
				document.PurchaseForm.submit();
			}
			else 
			{
				if (confirm("The selected family member has already been registered for this activity. \nDo you wish to register them anyway?"))
				{
					document.PurchaseForm.submit();
				}
			}
		}

		function ValidateWait() 
		{
			// This does not care for pricing, as it is a wait list addition
			// If the ticket field exists check that something is entered
			var bexists = eval(document.PurchaseForm["quantity"]);
			var lcl_return_false = "N";
			var lcl_focus        = "";  
			if(bexists)
			{
				if (document.PurchaseForm.quantity.value == "")
				{
					lcl_focus = "quantity";
					inlineMsg(document.getElementById("quantity").id,'<strong>Required Field Missing: </strong>Ticket Quantity.',10,'quantity');
					lcl_return_false = "Y";
				}
				var rege = /^\d$/;
				var Ok = rege.test(document.PurchaseForm.quantity.value);

				if (! Ok) 
				{
					lcl_focus = "quantity";
					inlineMsg(document.getElementById("quantity").id,'<strong>Invalid Value: </strong>The ticket quantity must be a number.',10,'quantity');
					lcl_return_false = "Y";
				}
			}

			if(lcl_return_false=="Y") 
			{
				if(lcl_focus != "") 
				{
					document.getElementById(lcl_focus).focus();
				}
				return false;
			}
			else
			{	
				//alert('Successful');
				document.PurchaseForm.buyorwait.value = 'W'
				document.PurchaseForm.submit();
			}
		}

		function ViewCart()
		{
			location.href='class_cart.asp';
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
					oPrice.value = format_number(0,2);
					bValid = false;
				}
			}

			// Calculate a new total price
			if (document.PurchaseForm.pricetypeid.length)   // If there is more than one price checkbox
			{
				var checklength = document.PurchaseForm.pricetypeid.length;
				var i = checklength - 1;

				for (l = 0; l <= i; l++)
				{
					if (document.PurchaseForm.pricetypeid[l].checked)
					{ 
						//total += Number(document.frmStatus.pricetypeid[l].value);
						total += Number(eval('document.PurchaseForm.amount' + document.PurchaseForm.pricetypeid[l].value + '.value'));
					}
				}
			}
			else   // There is only one price checkbox
			{
				if (document.PurchaseForm.pricetypeid.checked)
				{
					total += Number(eval('document.PurchaseForm.amount' + document.PurchaseForm.pricetypeid.value + '.value'));
				}
			}

			document.PurchaseForm.totalprice.value = total;
			document.getElementById("displaytotalprice").innerHTML = format_number(total,2);

			if ( bValid == false ) 
			{
				document.getElementById(oPrice.id).focus();
				inlineMsg(oPrice.id,'<strong>Invalid Value: </strong>Prices should numbers in currency format.',10,oPrice.id);
				return false;
			}
			return true;
		}

		function UpdatePriceTotal( iPrice, bChecked )
		{
			var total = 0.00;

			if (iPrice != "")
			{
				total = Number(document.PurchaseForm.totalprice.value);
				if (bChecked)
				{
					total += Number(iPrice);
				}
				else
				{
					total -= Number(iPrice);
				}
				document.PurchaseForm.totalprice.value = total;
				document.getElementById("displaytotalprice").innerHTML = format_number(total,2);
			}
		}

		function setupCoachFields() {
 			//Check to see if a value has been selected in the "I would like to" volunteer coach field.
	 		//If one has been selected then enable the other volunteer coach fields.
		 	//If one has not then disable them.
			 lcl_type = document.getElementById("rostercoachtype").value;

 			if(lcl_type!="") {
   				//document.getElementById("volunteerCoachInfo").style.visibility="visible";
       $('#volunteerCoachInfo').show('slow');
		 	} else {
  	 			//document.getElementById("volunteerCoachInfo").style.visibility="hidden";
       $('#volunteerCoachInfo').hide('slow');
 			}
		}

		var isNN = (navigator.appName.indexOf("Netscape")!=-1);

		function autoTab(input,len, e) 
		{
			var keyCode = (isNN) ? e.which : e.keyCode; 
			var filter  = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];

			if(input.value.length >= len && !containsElement(filter,keyCode)) 
			{
				input.value = input.value.slice(0, len);
				var addNdx = 1;

				while(input.form[(getIndex(input)+addNdx) % input.form.length].type == "hidden") 
				{
					addNdx++;
					//alert(input.form[(getIndex(input)+addNdx) % input.form.length].type);
				}

				input.form[(getIndex(input)+addNdx) % input.form.length].focus();
			}

			function containsElement(arr, ele) 
			{
				var found = false, index = 0;

				while(!found && index < arr.length)
					if(arr[index] == ele)
						found = true;
					else
						index++;
				return found;
			}

			function getIndex(input)	
			{
				var index = -1, i = 0, found = false;
	
				while (i < input.form.length && index == -1)
					if (input.form[i] == input)index = i;
					else i++;
					return index;
			}

			return true;
		}


<%
	If sSearchName <> "" Then
%>
		$(document).ready(function() {
			doInitialNameSearchChange( );
		});
<%	End If			%>

	//-->
	</script>
</head>
<body>

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<!--BEGIN PAGE CONTENT-->

<div id="content">
	<div id="centercontent">

<%
'Check the cart.  If items exist then display the "View Cart" button
 If CartHasItems() Then 
    response.write vbcrlf & "<div id=""topbuttons"">" 
    response.write vbcrlf & "<input type=""button"" name=""viewcart"" id=""viewcart"" class=""button"" value=""View Cart"" onclick=""ViewCart();"" />"
    response.write vbcrlf & "</div>"
 End If 

'Display "Back" and "Return to Class/Event List" buttons
 response.write vbcrlf & "<p>" 
 response.write "<input type=""button"" class=""button"" name=""backbtn"" id=""backbtn"" value=""<< Back"" onclick=""location.href='class_offerings.asp?classid=" & request("classid") & "'"" /> &nbsp;&nbsp; "
 response.write "<input type=""button"" class=""button"" name=""returnToList"" id=""returnToList"" value=""Return to Class/Event List"" onclick=""location.href='roster_list.asp'"" />"
 response.write "</p><br />" 
%>
	<!--<a href="javascript:history.go(-1)"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%=langBackToStart%></a>-->
<!--	<a href="class_offerings.asp?classid=<%'request("classid")%>"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;<%'langBackToStart%></a> &nbsp; 
	<a href="roster_list.asp"><img src="../images/arrow_2back.gif" align="absmiddle" border="0">&nbsp;Return To Class/Event List</a><br /><br /> -->
	
	<% 
	iGenderNotRequiredId = GetGenderNotRequiredId( )
	lcl_enabled_tshirt   = "BOTH"
	lcl_enabled_pants    = "BOTH"
	lcl_enabled_grade    = "BOTH"
	lcl_enabled_coach    = "BOTH"
	lcl_inputtype_tshirt = "LOV"
	lcl_inputtype_pants  = "LOV"
	lcl_inputtype_grade  = "TEXT"

	'Get the class information to display
	sSql = "SELECT "
	sSql = sSql & " C.classname, "
	sSql = sSql & " isnull(C.startdate,'') AS startdate, "
	sSql = sSql & " C.classdescription, "
	sSql = sSql & " ISNULL(C.enddate,'') AS enddate, "
	sSql = sSql & " O.optionid, "
	sSql = sSql & " O.optionname, "
	sSql = sSql & " O.optiondescription, "
	sSql = sSql & " O.canpurchase, "
	sSql = sSql & " O.optiontype, "
	sSql = sSql & " C.isparent, "
	sSql = sSql & " C.classtypeid, "
	sSql = sSql & " ISNULL(C.membershipid,0) AS membershipid, "
	sSql = sSql & " L.name as locationname, "
	sSql = sSql & " L.address1, "
	sSql = sSql & " ISNULL(C.minage,0) AS minage, "
	sSql = sSql & " ISNULL(C.maxage,99) as maxage, "
	sSql = sSql & " ISNULL(C.pricediscountid,0) as pricediscountid, "
	sSql = sSql & " displayrosterpublic, "
	'sSql = sSql & " ISNULL(teamreg_tshirt_enabled,0) AS teamreg_tshirt_enabled, "
	'sSql = sSql & " ISNULL(teamreg_pants_enabled,0) AS teamreg_pants_enabled, "
	sSql = sSql & " teamreg_tshirt_enabled, "
	sSql = sSql & " teamreg_pants_enabled, "
	sSql = sSql & " teamreg_grade_enabled, "
	sSql = sSql & " teamreg_coach_enabled, "
	sSql = sSql & " teamreg_tshirt_inputtype, "
	sSql = sSql & " teamreg_pants_inputtype, "
	sSql = sSql & " teamreg_grade_inputtype, "
	sSql = sSql & " ISNULL(genderrestrictionid," & iGenderNotRequiredId & ") AS genderrestrictionid "
	sSql = sSql & " FROM egov_class C, "
	sSql = sSql &      " egov_registration_option O, "
	sSql = sSql &      " egov_class_location L "
	sSql = sSql & " WHERE classid = " & request("classid")
	sSql = sSql & " AND C.optionid = O.optionid "
	sSql = sSql & " AND C.locationid = L.locationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
   'Setup values
    if oRs("teamreg_tshirt_enabled") <> "" then
       lcl_enabled_tshirt = oRs("teamreg_tshirt_enabled")
    end if

    if oRs("teamreg_pants_enabled") <> "" then
       lcl_enabled_pants = oRs("teamreg_pants_enabled")
    end if

    if oRs("teamreg_grade_enabled") <> "" then
       lcl_enabled_grade = oRs("teamreg_grade_enabled")
    end if

    if oRs("teamreg_coach_enabled") <> "" then
       lcl_enabled_coach = oRs("teamreg_coach_enabled")
    end if

    if oRs("teamreg_tshirt_inputtype") <> "" then
       	lcl_inputtype_tshirt = oRs("teamreg_tshirt_inputtype")
    end if

    if oRs("teamreg_pants_inputtype") <> "" then
       lcl_inputtype_pants = oRs("teamreg_pants_inputtype")
    end if

    if oRs("teamreg_grade_inputtype") <> "" then
       lcl_inputtype_grade = oRs("teamreg_grade_inputtype")
    end if

  		response.write "<h3>" & oRs("classname") & " &nbsp; ( " & GetActivityNo( iTimeId ) & " )</h3>" 
  		response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend><strong> Details </strong></legend>" & vbcrlf

 		'Date
  		response.write "<div><p><strong>Class/Event Dates: " 

  		If oRs("startdate") <> "" Then 
    		response.write MonthName(Month(oRs("startdate"))) & " " & Day(oRs("startdate")) & ", " & Year(oRs("startdate")) 
		End If 

 		'handle enddate
  		If oRs("enddate") <> "" Then 
			'response.write CDate(DateValue(oRs("enddate"))) & "<br />"
			If CDate(DateValue(oRs("enddate"))) <> CDate(DateValue(oRs("startdate"))) Then 
				response.write " &ndash; " & MonthName(Month(oRs("enddate"))) & " " & Day(oRs("enddate")) & ", " & Year(oRs("enddate")) 

				If DateDiff("d", oRs("startdate"), oRs("enddate")) > 7 Then 
					bMultiWeeks = True 
				End If 
			End If 

			' figure out if this ends in the past or future - we want the futures normally for signups
			'response.write "<br /> " & DateDiff("d", CDate(DateValue(Now())), CDate(DateValue(oRs("enddate"))) ) & " " &  CDate(DateValue(Now())) & " " 
			'response.write "<br /> " & DateDiff("d", CDate(DateValue(oRs("enddate"))), CDate(DateValue(Now())) ) & " " &  CDate(DateValue(Now())) & " " 
			'If DateDiff("d", CDate(DateValue(Now())), CDate(DateValue(oRs("enddate")))) > 0 Then 
			If CDate(DateValue(oRs("enddate"))) < CDate(DateValue(Now())) Then 
				bIsInPast = True
				'response.write " *** In The Past ***"
			Else
				'response.write " *** NOT In The Past ***"
			End If 
		End If 

		 'Days of the week
  		response.write "</strong></p>" 

 		'Tell if registration, ticket, or free
  		response.write "<p><strong>" & oRs("optionname") & " &ndash; " & oRs("optiondescription") & "</strong></p>" 

 		'Tell about age restrictions
  		response.write vbcrlf & "<p><strong>Age Restrictions:</strong>" 
  		If CDbl(oRs("minage")) = CDbl(0.0) And CDbl(oRs("maxage")) = CDbl(99.0) Then 
    		response.write vbcrlf & "<br />&nbsp;&nbsp;&nbsp;None"
  		Else 
			If CDbl(oRs("minage")) <> CDbl(0.0) Then 
				response.write vbcrlf & "<br />&nbsp;&nbsp;&nbsp;Minimum: " & oRs("minage") & " years of age"
			End If 

			If CDbl(oRs("maxage")) <> CDbl(99.0) Then 
				response.write vbcrlf & "<br />&nbsp;&nbsp;&nbsp;Maximum: " & oRs("maxage") & " years of age"
			End If 
  		End If 

  		response.write vbcrlf & "</p>" 

		' Gender Restriction 
		If bHasGenderRestrictions Then 
			response.write vbcrlf & "<p><strong>Gender Restriction:</strong><br />&nbsp;&nbsp;&nbsp;"
			response.write GetGenderRestrictionText( oRs("genderrestrictionid") )
			sRequiredGender = GetGenderRestriction( oRs("genderrestrictionid") )
			response.write vbcrlf & "<input type=""hidden"" id=""requiredgender"" name=""requiredgender"" value="""
			If sRequiredGender = "M" Or sRequiredGender = "F" Then
				response.write sRequiredGender
			Else
				response.write "N"
			End If 
			response.write """ />"
			response.write vbcrlf & "</p>" 
		End If 

		' Location 
		response.write vbcrlf & "<p><strong>Location:</strong><br />"
		response.write "&nbsp;&nbsp;&nbsp;" & oRs("locationname") & "<br />&nbsp;&nbsp;&nbsp;" & oRs("address1") & "</p>"

		' Display Waiver Links
		response.write "<p><strong>Waivers:</strong>&nbsp; " 
		ShowClassWaiverLinks request("classid") 
		response.write "</p>"

		'response.write vbcrlf & "<p>You are considered a " & sUserType & " - " & sResidentDesc & "</p>"
		'response.write "</div><div id=""rightdetail"">"
		'response.write "<p><strong>Description:</strong><br />" & oRs("classdescription") & "</p>" 
		response.write vbcrlf & "</div></fieldset>"

 		Select Case oRs("optiontype")
			Case "register"		' Handle registration required
				' Show pick of registered users and their detail info.
				'ShowRegisteredUsers iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId

				response.write "<form id=""PurchaseForm"" name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" 
				response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend><strong> Select the Purchaser </strong></legend>"
				response.write "  <div id=""namesearch"">" & vbcrlf
    response.write "    <strong>Name Is Like:</strong>&nbsp;" & vbcrlf
    response.write "    <input type=""text"" id=""searchname"" name=""searchname"" value=""" & sSearchName & """ placeholder=""Enter part of the name"" size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange( );return false;}"" onchange=""clearMsg('searchname')"" />&nbsp;" & vbcrlf
    response.write "    <input type=""text"" id=""searchname2"" name=""searchname2"" value=""" & sSearchName2 & """ placeholder=""Enter another part of the name"" size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange( );return false;}"" onchange=""clearMsg('searchname2')"" />&nbsp;" & vbcrlf
    response.write "    <input type=""button"" class=""button"" value=""Search for a Name"" onclick=""doNameSearchChange( );"" />&nbsp;<input type=""button"" class=""button"" value=""New Public User"" onclick=""NewUser();"" />" & vbcrlf
    response.write "  </div>" & vbcrlf
				response.write "  <span id=""applicant"">" & vbcrlf
				'If iUserId > CLng(0) Then
					' Show registered user picks
				'	ShowCitizenPicks iUserId, sSearchName
				'End If 
				response.write "  <input type=""hidden"" value=""0"" name=""egovuserid"" id=""egovuserid"" /><span class=""nomatch"">Search for a name then select one from the resulting list.</span>" & vbcrlf
				response.write "  </span>&nbsp;" & vbcrlf
    response.write "  <input type=""button"" class=""button"" id=""edituserbtn"" value=""Edit User"" onclick=""EditUser();"" />" & vbcrlf

				' User Info table here
				response.write "<div id=""userinfo""></div>" & vbcrlf

				response.write "</fieldset>" & vbcrlf

				'response.write vbcrlf & "<form id=""PurchaseForm"" name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" 
				response.write "  <input type=""hidden"" name=""classid"" id=""classid"" value="""                               & request("classid")         & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""timeid"" id=""timeid"" value="""                                 & iTimeId                    & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""userid"" id=""userid"" value="""                                 & iUserId                    & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optionid"" id=""optionid"" value="""                             & oRs("optionid")            & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""optiontype"" id=""optiontype"" value="""                         & oRs("optiontype")          & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""isparent"" id=""isparent"" value="""                             & oRs("isparent")            & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classtypeid"" id=""classtypeid"" value="""                       & oRs("classtypeid")         & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""classname"" id=""classname"" value="""                           & oRs("classname")           & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""itemtypeid"" id=""itemtypeid"" value="""                         & iItemTypeId                & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""availability"" id=""availability"" value="""                     & iAvailability              & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""membershipid"" id=""membershipid"" value="""                     & oRs("membershipid")        & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""pricediscountid"" id=""pricediscountid"" value="""               & oRs("pricediscountid")     & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""displayrosterpublic"" id=""displayrosterpublic"" value="""       & oRs("displayrosterpublic") & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""teamreg_tshirt_enabled"" id=""teamreg_tshirt_enabled"" value=""" & lcl_enabled_tshirt         & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""teamreg_pants_enabled"" id=""teamreg_pants_enabled"" value="""   & lcl_enabled_pants          & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""teamreg_grade_enabled"" id=""teamreg_grade_enabled"" value="""   & lcl_enabled_grade          & """ />" & vbcrlf
				response.write "  <input type=""hidden"" name=""teamreg_coach_enabled"" id=""teamreg_coach_enabled"" value="""   & lcl_enabled_coach          & """ />" & vbcrlf

				'Family Member Drop Down --------------------------------------------------
				response.write "<fieldset id=""registrantselection"" class=""fieldset"">" & vbcrlf
    response.write "  <legend><strong>Select the Registrant </strong></legend>" & vbcrlf

				'bAllMembers = ShowFamilyMembers( iUserid, iMemberCount, oRs("membershipid") )
				response.write "<span id=""registrant"">" & vbcrlf
				response.write "<input type='hidden' name='familymemberid' id='familymemberid' value='0' /><span class=""nomatch"">Search for a purchaser to get possible registrants.</span>" & vbcrlf
				response.write "</span>" & vbcrlf

				response.write " &nbsp;<input id=""updatefamilymembersbtn"" name=""updatefamilymembersbtn"" type=""button"" class=""button"" onclick=""UpdateFamily()"" value=""Update Family Members"" />" & vbcrlf
				response.write "</fieldset>" & vbcrlf

			'BEGIN: Team Roster Registration Fields (Craig, CO - Custom Fields) --------
				lcl_label_tshirt = "T-Shirt"

				if lcl_orghasfeature_custom_registration_craigco AND oRs("displayrosterpublic") then
       if lcl_enabled_tshirt <> "DISABLED" _
       OR lcl_enabled_pants  <> "DISABLED" _
       OR lcl_enabled_grade  <> "DISABLED" _
       OR lcl_enabled_coach  <> "DISABLED" then

     					response.write "<fieldset class=""fieldset"">" & vbcrlf
          response.write "  <legend><strong>Team Registration - Additional Info&nbsp;</strong></legend>" & vbcrlf
     
          if lcl_enabled_tshirt <> "DISABLED" OR lcl_enabled_pants <> "DISABLED" OR lcl_enabled_grade <> "DISABLED" then
      				  	response.write "<p>" & vbcrlf
     	   				response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf
		     		   	response.write "  <tr>" & vbcrlf

            'Grade
             if lcl_enabled_grade <> "DISABLED" then
   			   	     	response.write "      <td>" & vbcrlf
        	   				response.write "          <span style=""color:#ff0000"">*</span>Grade:&nbsp;" & vbcrlf
		        		   	'response.write "          <input type=""text"" name=""rostergrade"" id=""rostergrade"" size=""3"" maxlength=""2"" onchange=""clearMsg('rostergrade')"" />" & vbcrlf

         						'Determine which input is to be displayed
			          			if lcl_inputtype_grade <> "" then
						      	      if ucase(lcl_inputtype_grade) = "TEXT" then
        	      							response.write "<input type=""text"" name=""rostergrade"" id=""rostergrade"" size=""5"" maxlength=""50"" />" & vbcrlf
						            	else
          					      	response.write "<select name=""rostergrade"" id=""rostergrade"" onchange=""clearMsg('rostergrade')"">" & vbcrlf
                                        displayTeamRosterAccessories session("orgid"), request("classid"), "GRADE"
            		   					response.write "</select>" & vbcrlf
           						 	end if
   			      		  else
      			      				response.write "  <input type=""hidden"" name=""rostergrade"" id=""rostergrade"" value="""" />" & vbcrlf
       					   	end if

   				        	response.write "      </td>" & vbcrlf
             end if

            'T-shirt
             if lcl_enabled_tshirt <> "DISABLED" then

         		 			'Check for an "edit display" for the T-shirt label
       		    			if orgHasDisplay(session("orgid"),"class_teamregistration_tshirt_label") then
      	    	   				lcl_label_tshirt = getOrgDisplay(session("orgid"),"class_teamregistration_tshirt_label")
           					end if

    	      					response.write "      <td>" & vbcrlf
       		   				response.write "          " & lcl_label_tshirt & " Size:&nbsp;" 

         						'Determine which input is to be displayed
			          			if lcl_inputtype_tshirt <> "" then
						      	      if ucase(lcl_inputtype_tshirt) = "TEXT" then
        	      							response.write "<input type=""text"" name=""rostershirtsize"" id=""rostershirtsize"" size=""20"" maxlength=""50"" />" & vbcrlf
						            	else
          					      	response.write "<select name=""rostershirtsize"" id=""rostershirtsize"" onchange=""clearMsg('rostershirtsize')"">" & vbcrlf
                                        displayTeamRosterAccessories session("orgid"), request("classid"), "TSHIRT"
            		   					response.write "</select>" & vbcrlf
           						 	end if
   			      		  else
      			      				response.write "  <input type=""hidden"" name=""rostershirtsize"" id=""rostershirtsize"" value="""" />" & vbcrlf
       					   	end if

         	 					response.write "      </td>" & vbcrlf
  	      				end if

            'Pants
        					if lcl_enabled_pants <> "DISABLED" then
          						response.write "      <td>" & vbcrlf
    		      				response.write "          Pants Size:&nbsp;" 

          					'Determine which input is to be displayed
          						if lcl_inputtype_pants <> "" then
				            			if ucase(lcl_inputtype_pants) = "TEXT" then
        	      							response.write "<input type=""text"" name=""rosterpantssize"" id=""rosterpantssize"" size=""20"" maxlength=""50"" />" & vbcrlf
      						      	else
		  		      			      	response.write "<select name=""rosterpantssize"" id=""rosterpantssize"" onchange=""clearMsg('rosterpantssize');"">" & vbcrlf
                                        displayTeamRosterAccessories session("orgid"), request("classid"), "PANTS"
              								response.write "</select>" & vbcrlf
            							end if
    		      				else
      			      				response.write "  <input type=""hidden"" name=""rosterpantssize"" id=""rosterpantssize"" value="""" />" & vbcrlf
   					       	end if

          						response.write "      </td>" & vbcrlf
  	      				else
    		      				response.write "  <input type=""hidden"" name=""rosterpantssize"" id=""rosterpantssize"" value="""" />" & vbcrlf
   				     	end if

        					response.write "  </tr>" & vbcrlf
		        			response.write "</table>" & vbcrlf
				        	response.write "</p>" & vbcrlf
          end if

         'Coach
          if lcl_enabled_coach <> "DISABLED" then
     	   				lcl_volunteercoach_text = getOrgDisplay(session("orgid"),"class_teamregistration_volunteercoachdesc")

        					if lcl_volunteercoach_text <> "" then
      		    				response.write "<div>" & lcl_volunteercoach_text & "</div><br />" & vbcrlf
  	      				end if

        					response.write "<div>" & vbcrlf
  	      				response.write "  I would like to:&nbsp;" & vbcrlf
		  		      	response.write "  <select name=""rostercoachtype"" id=""rostercoachtype"" onchange=""setupCoachFields();"">" & vbcrlf
     				   	response.write "    <option value=""""></option>" & vbcrlf
   		     			response.write "    <option value=""Head Coach"">Head Coach</option>" & vbcrlf
			   	     	response.write "    <option value=""Assistant Coach"">Assistant Coach</option>" & vbcrlf
     	   				response.write "  </select>" & vbcrlf
		     		   	response.write "</div>" & vbcrlf
     				   	response.write "<br />" & vbcrlf
      		  			response.write "<div id=""volunteerCoachInfo"">" & vbcrlf
			      	  	response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbcrlf
   			   	  	response.write "    <tr>" & vbcrlf
  	   			   	response.write "        <td>" & vbcrlf
   		     			response.write "            <span style=""color:#ff0000"">*</span>Full Name:" & vbcrlf
     	   				response.write "        </td>" & vbcrlf
		     		   	response.write "        <td width=""85%"">" & vbcrlf
   			   	  	response.write "            <input type=""text"" name=""rostervolunteercoachname"" id=""rostervolunteercoachname"" size=""40"" maxlength=""100"" onchange=""clearMsg('rostervolunteercoachname');"" />" & vbcrlf
     	   				response.write "        </td>" & vbcrlf
		     		   	response.write "    </tr>" & vbcrlf
   			   	  	response.write "    <tr>" & vbcrlf
     	   				response.write "        <td>" & vbcrlf
		     		   	response.write "            Daytime Phone:" & vbcrlf
   			   	  	response.write "        </td>" & vbcrlf
     	   				response.write "        <td>" & vbcrlf
		     		   	'response.write "            <input type=""hidden"" name=""rostervolunteercoachdayphone"" id=""rostervolunteercoachdayphone"" size=""10"" maxlength=""10"" />" 
   			   	  	response.write "            (<input type=""text"" name=""skip_volcoachday_areacode"" id=""skip_volcoachday_areacode"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />)" & vbcrlf
     				   	response.write "            <input type=""text"" name=""skip_volcoachday_exchange"" id=""skip_volcoachday_exchange"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />" & vbcrlf
   		  	   		response.write "            &ndash;" & vbcrlf
      				  	response.write "            <input type=""text"" name=""skip_volcoachday_line"" id=""skip_volcoachday_line"" size=""4"" maxlength=""4"" onKeyUp=""return autoTab(this, 4, event);"" onchange=""clearMsg('skip_volcoachday_line');"" />" & vbcrlf
  	      				response.write "        </td>" & vbcrlf
		  		      	response.write "    </tr>" & vbcrlf
      				  	response.write "    <tr>" & vbcrlf
  	      				response.write "        <td>" & vbcrlf
		  		      	response.write "            Cell Phone:" & vbcrlf
   				     	response.write "        </td>" & vbcrlf
     	   				response.write "        <td>" & vbcrlf
		     		   	'response.write "            <input type=""hidden"" name=""rostervolunteercoachcellphone"" id=""rostervolunteercoachcellphone"" size=""10"" maxlength=""10"" />" 
   			   	  	response.write "           (<input type=""text"" name=""skip_volcoachcell_areacode"" id=""skip_volcoachcell_areacode"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />)" & vbcrlf
  	   			   	response.write "            <input type=""text"" name=""skip_volcoachcell_exchange"" id=""skip_volcoachcell_exchange"" size=""3"" maxlength=""3"" onKeyUp=""return autoTab(this, 3, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />" & vbcrlf
   		  		   	response.write "            &ndash;" & vbcrlf
      				  	response.write "            <input type=""text"" name=""skip_volcoachcell_line"" id=""skip_volcoachcell_line"" size=""4"" maxlength=""4"" onKeyUp=""return autoTab(this, 4, event);"" onchange=""clearMsg('skip_volcoachcell_line');"" />" & vbcrlf
  	      				response.write "        </td>" & vbcrlf
		  		      	response.write "    </tr>" & vbcrlf
      				  	response.write "  </table>" & vbcrlf
  	      				response.write "  <div>" & vbcrlf
		  		      	response.write "    Please list an email address, so you can be contacted for more information:&nbsp;" & vbcrlf
   				     	response.write "    <input type=""text"" name=""rostervolunteercoachemail"" id=""rostervolunteercoachemail"" size=""50"" maxlength=""100"" onchange=""clearMsg('rostervolunteercoachemail');"" />" & vbcrlf
     	   				response.write "  </div>" & vbcrlf
		     		   	response.write "</div>" & vbcrlf

        					lcl_setupCoachFields = "Y"
          end if

   				  	response.write "</fieldset>" & vbcrlf
       end if

       if lcl_enabled_tshirt = "DISABLED" then
   		  			response.write "<input type=""hidden"" name=""rostershirtsize"" id=""rostershirtsize"" value="""" />" & vbcrlf
       end if

       if lcl_enabled_pants = "DISABLED" then
   		  			response.write "<input type=""hidden"" name=""rosterpantssize"" id=""rosterpantssize"" value="""" />" & vbcrlf
       end if

       if lcl_enabled_grade = "DISABLED" then
    			  	response.write "<input type=""hidden"" name=""rostergrade"" id=""rostergrade"" value="""" />" & vbcrlf
       end if

       if lcl_enabled_coach = "DISABLED" then
    			  	response.write "<input type=""hidden"" name=""rostercoachtype"" id=""rostercoachtype"" value="""" />" & vbcrlf
       end if

				else
  					response.write "  <input type=""hidden"" name=""rostergrade"" id=""rostergrade"" value="""" />" & vbcrlf
		  			response.write "  <input type=""hidden"" name=""rostershirtsize"" id=""rostershirtsize"" value="""" />" & vbcrlf
				  	response.write "  <input type=""hidden"" name=""rostercoachtype"" id=""rostercoachtype"" value="""" />" & vbcrlf
  					response.write "  <input type=""hidden"" name=""rostervolunteercoachname"" id=""rostervolunteercoachname"" value="""" />" & vbcrlf
		  			response.write "  <input type=""hidden"" name=""rostervolunteercoachdayphone"" id=""rostervolunteercoachdayphone"" value="""" />" & vbcrlf
				  	response.write "  <input type=""hidden"" name=""rostervolunteercoachcellphone"" id=""rostervolunteercoachcellphone"" value="""" />" & vbcrlf
  					response.write "  <input type=""hidden"" name=""rostervolunteercoachemail"" id=""rostervolunteercoachemail"" value="""" />" & vbcrlf
				end if
			'END: Team Roster Registration Fields (Craig, CO - Custom Fields)-----------

				'Availability and Pricing --------------------------------------------------
				response.write "<fieldset class=""fieldset"">" & vbcrlf
    response.write "  <legend><strong> Availability and Pricing&nbsp;</strong></legend>" & vbcrlf
				'Form for selecting either ticket qty, or selecting a family member
				response.write "<div>" & vbcrlf

				'Availability
				response.write "<p><strong>Availability:</strong><br />" & vbcrlf
				DisplayClassActivities request("classid"), iTimeId, False  'In class_global_functions.asp
				response.write "</p>" & vbcrlf


				' Time Options
'				response.write vbcrlf & "<p><strong>Time:</strong><br />"
'				ShowTimeOptions request("classid"), Session("OrgID"), oRs("isparent"), oRs("classtypeid"), iTimeId
'				response.write "</p>"

				' Price options
				response.write "<p><strong>Price:</strong><br />"
				ShowPriceOptions request("classid"), Session("OrgID"), sUserType, iMemberCount, oRs("membershipid"), oRs("pricediscountid"), iUserId
				'ShowCostOptions request("classid"), sUserType, Session("OrgID"), bAllMembers, iMemberCount, oRs("isparent"), oRs("classtypeid")
				response.write vbcrlf & "</p>"

				' Purchase or Waitlist 
				response.write "<p><strong>Select:</strong><br />" 
				response.write "<input type=""radio"" name=""buyorwait"" id=""buyorwait"" value=""B"" checked=""checked"" /> Purchase <br />" 
				response.write "<input type=""radio"" name=""buyorwait"" id=""buyorwait"" value=""W"" /> Add to Wait List" 
				response.write "</p>" 

				'Add to cart button
				response.write "<p>" 
				response.write "<input type=""button"" id=""completebtn"" name=""complete"" class=""button"" style=""width:140px;text-align:center;"" value=""Add To Cart"" "
				If bIsInPast Then
					response.write " onclick=""askThenValidate();"" "
				Else
					response.write " onclick=""ValidateForm();"" "
				End If 

				If bRegistrationBlocked Then 
  					response.write " disabled=""disabled"" "
				End If

				response.write "/>" 
				'response.write vbcrlf & "&nbsp;&nbsp;<strong>OR</strong>"
				'response.write vbcrlf & "&nbsp;&nbsp;<input type=""button"" name=""waitlist"" value=""Add to Wait List"" onclick=""ValidateWait();"" />"
				response.write vbcrlf & "</p>"

				response.write vbcrlf & "</div>"
		if intMaxEnroll > intEnrolled and intWaitlist > 0 and Session("OrgID") = "60" then
			response.write "<script>window.onload = function () {alert('There is now one or more openings in this class which has a waitlist.  Please contact participants on the waitlist to enroll in the class.');}</script>"
		end if

				' Show the availability
				'response.write vbcrlf & "<div id=""rightprice"">"
				'ShowAvailability request("classid"), oRs("isparent")
				'response.write vbcrlf & "</div>"
				response.write vbcrlf & "</fieldset>"
				response.write vbcrlf & "</form>"
				
			Case "tickets"		' Ticketed Event
				' Show pick of registered users and their detail info.
				'ShowRegisteredUsers iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId
				
				response.write "<form name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" 
				response.write "<fieldset id=""registrantselection""><legend><strong> Select the Purchaser </strong></legend>" 
				response.write vbcrlf & "<div id=""namesearch""><strong>Name Is Like:</strong>&nbsp;<input type=""text"" id=""searchname"" name=""searchname"" value=""" & sSearchName & """ placeholder=""Enter part of the name"" size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange( );return false;}"" />"
    response.write "    <input type=""text"" id=""searchname2"" name=""searchname2"" value=""" & sSearchName2 & """ placeholder=""Enter another part of the name"" size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange( );return false;}"" onchange=""clearMsg('searchname2')"" />&nbsp;" & vbcrlf
				response.write "&nbsp;<input type=""button"" class=""button"" value=""Search for a Name"" onclick=""doNameSearchChange( );"" />&nbsp;<input type=""button"" class=""button"" value=""New Public User"" onclick=""NewUser();"" /></div>"
				response.write "<span id=""applicant"">"
				response.write "<input type=""hidden"" value=""0"" name=""egovuserid"" id=""egovuserid"" /><span class=""nomatch"">Search for a name then select one from the resulting list.</span>"
				response.write "</span>&nbsp;<input type=""button"" class=""button"" id=""edituserbtn"" value=""Edit User"" onclick=""EditUser();"" />"
				
				' User Info table here
				response.write vbcrlf & "<div id=""userinfo""></div>"

				response.write vbcrlf & "</fieldset>"

				response.write vbcrlf & "<fieldset><legend><strong> Ticket Availability and Pricing </strong></legend>"
				' Form for selecting either ticket qty, or selecting a family member
				response.write  vbcrlf & "<div id=""leftprice"">" 
				'response.write "<form name=""PurchaseForm"" method=""post"" action=""class_addtocart.asp"">" 
				response.write vbcrlf & "<input type=""hidden"" id=""classid"" name=""classid"" value=""" & request("classid") & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />" 
				response.write vbcrlf & "<input type=""hidden"" id=""userid"" name=""userid"" value=""" & iUserId & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""optionid"" value=""" & oRs("optionid") & """ />" 
				response.write vbcrlf & "<input type=""hidden"" id=""optiontype"" name=""optiontype"" value=""" & oRs("optiontype") & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""isparent"" value=""" & oRs("isparent") & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""classtypeid"" value=""" & oRs("classtypeid") & """ />" 
				'response.write "  <input type=""hidden"" name=""buyorwait"" value=""B"" />" 
				response.write vbcrlf & "<input type=""hidden"" name=""classname"" value=""" & oRs("classname") & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""itemtypeid"" value=""" & iItemTypeId & """ />" 
				response.write vbcrlf & "<input type=""hidden"" name=""availability"" value=""" & iAvailability & """ />" 
				response.write vbcrlf & "<input type=""hidden"" id=""membershipid"" name=""membershipid"" value=""" & oRs("membershipid") & """ />"

				' Availability
				response.write vbcrlf & "<p><strong>Availability:</strong><br />"
				DisplayClassActivities request("classid"), iTimeId, False   ' In class_global_functions.asp
				'ShowAvailability request("classid"), oRs("isparent"), oRs("optionid"), iTimeId
				response.write vbcrlf & "</p>"

				' Ticket Quantity
				response.write "<p>No. of Tickets: &nbsp; <input type=""text"" name=""quantity"" id=""quantity"" value=""1"" size=""6"" maxlength=""6"" /></p>" 
				
				' Availability
'				response.write vbcrlf & "<p><strong>Availability:</strong><br />"
'				ShowAvailability request("classid"), oRs("isparent"), oRs("optionid"), iTimeId
'				response.write vbcrlf & "</p>"

				' Time Options
'				response.write vbcrlf & "<p><strong>Time:</strong><br />"
'				ShowTimeOptions request("classid"), Session("OrgID"), oRs("isparent"), oRs("classtypeid"), iTimeId
'				response.write "</p>"

				' Price Options
				response.write vbcrlf & "<p><strong>Price:</strong><br />"
				ShowPriceOptions request("classid"), Session("OrgID"), sUserType, iMemberCount, oRs("membershipid"), oRs("pricediscountid"), iUserId
				'ShowCostOptions request("classid"), sUserType, Session("OrgID"), bAllMembers, iMemberCount, oRs("isparent"), oRs("classtypeid")
				response.write vbcrlf & "</p>"

				' Purchase or Waitlist 
				response.write vbcrlf & "<p><strong>Select:</strong><br />"
				response.write vbcrlf & "<input type=""radio"" name=""buyorwait"" value=""B"" checked=""checked"" /> Purchase <br />"
				response.write vbcrlf & "<input type=""radio"" name=""buyorwait"" value=""W"" /> Add to Wait List"
				response.write vbcrlf & "</p>"

				' Add to cart button
				response.write vbcrlf & "<p>"
				response.write vbcrlf & "<input type=""button"" class=""button"" id=""completebtn"" name=""complete"" style=""width:140px;text-align:center;"" value=""Add To Cart"" "
				If bIsInPast Then
					response.write " onclick=""askThenValidate();"" "
				Else
					response.write " onclick=""ValidateForm();"" "
				End If 
				
				If bRegistrationBlocked Then 
					response.write " disabled=""disabled"" "
				End If 
				response.write "/>"
				'response.write vbcrlf & "&nbsp;&nbsp;<strong>OR</strong>"
				'response.write vbcrlf & "&nbsp;&nbsp;<input type=""button"" name=""waitlist"" value=""Add to Wait List"" onclick=""ValidateWait();"" />"
				response.write vbcrlf & "</p>"

				response.write vbcrlf & "</div></form>"

				' Show the availability
				'response.write vbcrlf & "<div id=""rightprice"">"
				'ShowAvailability request("classid"), oRs("isparent")
				'response.write vbcrlf & "</div>"
				response.write vbcrlf & "</fieldset>"


			Case "open"		' Open Attendance
				response.write "<p><strong>Ticketing or registration is not required.</strong></p>"
			Case Else 		' Information Only
				response.write "<p><strong>This listing is for information only.</strong>  See related classes/events to register or purchase tickets.</p>"
		End Select 
	Else
		response.write "<p>No information could be found for this class.</p>"
	End If 

	oRs.Close
	set oRs = Nothing 

 response.write "  </div>" 
 response.write "</div>" 

'Check for javascripts
 if lcl_setupCoachFields = "Y" then
    response.write "<script language=""javascript"">" 
    response.write "  setupCoachFields();" 
    response.write "</script>" 
 end if
%>

<!--#Include file="../admin_footer.asp"-->  

</body>
</html>

<!--#Include file="class_global_functions.asp"-->  


<%
'--------------------------------------------------------------------------------------------------
' void ShowRegisteredUsers iUserId, sUserType, sResidentDesc, sSearchName, sResults, sSearchStart, iTimeId 
'--------------------------------------------------------------------------------------------------
Sub ShowRegisteredUsers( ByVal iUserId, ByVal sUserType, ByVal sResidentDesc, ByVal sSearchName, ByVal sResults, ByVal sSearchStart, ByVal iTimeId )

	response.write vbcrlf & "<fieldset><legend><strong> Purchaser Information </strong></legend>"

	response.write vbcrlf & "<form name=""BuyerForm"" method=""post"" action=""class_signup.asp"">"

	response.write vbcrlf & "<p>Name Search: <input type=""text"" name=""searchname"" value=""" & sSearchName & """ size=""25"" maxlength=""50"" onchange=""javascript:ClearSearch();"" />"
	response.write vbcrlf & "<input type=""button"" class=""button"" value=""Search"" onclick=""javascript:SearchCitizens(document.BuyerForm.searchstart.value);"" /> &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:NewUser();"" value=""New User"" />"
	response.write vbcrlf & "<input type=""hidden"" name=""results"" value="""" />"
	response.write vbcrlf & "<input type=""hidden"" name=""timeid"" value=""" & iTimeId & """ />"
	response.write vbcrlf & "<input type=""hidden"" name=""searchstart"" value=""" & sSearchStart & """ />"
	response.write vbcrlf & "<span id=""searchresults"">" & sResults & "</span>"
	response.write vbcrlf & "<br /><div id=""searchtip"">(last name, first name)</div>"
	response.write vbcrlf & "</p>"
	response.write vbcrlf & "<p><input type=""hidden"" name=""classid"" value=""" & request("classid") & """ />"

	response.write vbcrlf & "Select Name: <select name=""egovuserid"" onchange=""javascript:UserPick();"">"
	ShowUserDropDown iUserId 
	response.write vbcrlf & "</select>"

	response.write vbcrlf & " &nbsp;&nbsp; <input type=""button"" class=""button"" onclick=""javascript:EditUser();"" value=""Edit User Profile"" />"
	response.write vbcrlf & "</p></form>"

	ShowUserInfo iUserId, sUserType, sResidentDesc 

	response.write vbcrlf & "</fieldset>"

End Sub 


'------------------------------------------------------------------------------
' boolean ShowFamilyMembers( iUserid )
'------------------------------------------------------------------------------
Function ShowFamilyMembers( ByVal iUserid, ByRef iMemberCount, ByVal iMembershipId )
	Dim sSql, oRs, sMember, iCount, iMonths, iAge

	iCount = 0
	iMemberCount = 0 

	sSql = "SELECT firstname, lastname, familymemberid, relationship, birthdate, userid"
	sSql = sSql & " FROM egov_familymembers "
	sSql = sSql & " WHERE isdeleted = 0 "
	sSql = sSql & " AND belongstouserid = " & iUserid
	sSql = sSql & " ORDER BY birthdate ASC"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If Not oRs.EOF Then 
		response.write vbcrlf & "<select name=""familymemberid"" id=""familymemberid"">" 

		Do While Not oRs.EOF
			If CLng(iMembershipId) > CLng(0) Then 
				sMember = DetermineMembership(oRs("familymemberid"), iUserid, iMembershipId)   ' In class_global_functions.asp
			End If 

			If iCount = 0 Then 
				lcl_member_selected = " selected=""selected"""
			Else 
				lcl_member_selected = ""
			End If 

			response.write vbcrlf & "  <option value=""" & oRs("familymemberid") & """" & lcl_member_selected & ">" & oRs("firstname") & " " & oRs("lastname") & " &ndash; " 

			If CLng(oRs("userid")) = CLng(iUserid) Then
				response.write "Head of Household"
			Else 
				response.write oRs("relationship") 
			End If 

			If UCase(oRs("relationship")) = "CHILD" Then 
				iAge = GetChildAge(oRs("birthdate"))
				response.write " &ndash; Age: " & iAge & " yrs"
			Else
				If UCase(oRs("relationship")) <> "SITTER" Then
					response.write " &ndash; Adult"
				End If 
			End If 

			If CLng(iMembershipId) > CLng(0) Then 
				If sMember = "M" Then
					response.write " &ndash; Member" 
					iMemberCount = iMemberCount + 1
				Else 
					response.write " &ndash; NonMember"
				End If 
			End If 

			response.write "</option>" 
			iCount = iCount + 1

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>" 

	End If 
	
	oRs.Close
	Set oRs = Nothing

	If iMemberCount = iCount Then 
		ShowFamilyMembers = True 
	Else
		ShowFamilyMembers = False
	End If 

End Function  


'------------------------------------------------------------------------------
' void ShowTimeOptions iClassid, iorgid, bIsParent, iClassTypeId, iTimeId
'------------------------------------------------------------------------------
Sub ShowTimeOptions( ByVal iClassid, ByVal iorgid, ByVal bIsParent, ByVal iClassTypeId, ByVal iTimeId )
	Dim sSql, oRs, iCount

	iCount = 0

	sSql = "SELECT starttime, endtime, timeid FROM egov_class_time "
	sSql = sSql & " WHERE classid = " & iClassid & " ORDER BY starttime"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	'response.write "<table border=""0"" cellpadding=""0"" cellspacing=""0"">"
	Do While Not oRs.EOF
		' Display new time pick
		If iCount > 0 Then 
			response.write "<br />"
		End If 
		'response.write vbcrlf & "<tr><td><input type=""radio"" "
		response.write vbcrlf & "<input type=""radio"" "
		If (iTimeId = 0 And iCount = 0) Then 
			response.write " checked=""checked"" "
		Else 
			If CLng(iTimeId) = CLng(oRs("timeid")) Then 
				response.write " checked=""checked"" "
			End If 
		End If 
		response.write "name=""timeid"" value=""" & oRs("timeid") & """> " 
		' Handle Series
		If bIsParent And iClassTypeId = 1 Then
			response.write "Entire Series "
			If CheckIfFullSeries(iClassid) Then 
				response.write "<span class=""filledstatus""> &ndash; FILLED</span>"
			End If 
		Else 
			response.write oRs("starttime") 
			sOldTimes = oRs("starttime") 
			If oRs("endtime") <> oRs("starttime") Then
				response.write " &ndash; " & oRs("endtime")
			End If 
			If CheckIfFullSingle(oRs("timeid")) Then 
				response.write "<span class=""filledstatus""> &ndash; FILLED</span>"
			End If 
		End If 

		iCount = iCount + 1

		oRs.MoveNext 
	Loop 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' void ShowPriceOptions iClassid, iorgid, sResidentType, iMemberCount, iMembershipId, iPriceDiscountId, iUserId
'------------------------------------------------------------------------------
Sub ShowPriceOptions( ByVal iClassid, ByVal iorgid, ByVal sResidentType, ByVal iMemberCount, ByVal iMembershipId, ByVal iPriceDiscountId, ByVal iUserId )
	Dim sSql, oRs, iCount, sDiscount, bMemberTypematch, sMemberType, iMinPricetype, iMaxPriceType
	Dim iFamilyMemberId, sMemberCode, cTotalPrice

	iCount = 0
	cTotalPrice = CDbl(0.00)

	sDiscount = GetDiscountPhrase( iPriceDiscountId )

	bResTypeMatch = CheckResTypeExists(iClassid, iorgid, sResidentType)

	' IF at least one person in the family is a member, then set up for member pricing match
	If iMemberCount > 0 Then 
		sMemberType = "M"
	Else
		sMemberType = "O"
	End If 

	sSql = "SELECT P.pricetypeid, T.pricetypename, T.ismember, P.amount, T.pricetype, P.accountid, "
	sSql = sSql & "T.isfee, T.isbaseprice, T.checkmembership, P.membershipid, T.isdropin "
	sSql = sSql & " FROM egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " WHERE T.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND orgid = " & iorgid & " AND P.classid = " & iClassid & " ORDER BY P.pricetypeid"
	'response.write sSql & "<br />"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	'response.write "<!--egov_class_pricetype_price.pricetypeid -->"

	If Not oRs.EOF Then 
		iMinPricetype = CLng(oRs("pricetypeid"))
		iMaxPriceType = CLng(oRs("pricetypeid"))

		response.write vbcrlf & "<table id=""pricetable"" border=""0"" cellpadding=""0"" cellspacing=""0"">"

		Do While Not oRs.EOF
  			If CLng(oRs("pricetypeid")) < iMinPricetype Then
		    	iMinPricetype = CLng(oRs("pricetypeid"))
  			End If 

  			If CLng(oRs("pricetypeid")) > iMaxPriceType Then
		    	iMaxPriceType = CLng(oRs("pricetypeid"))
  			End If 

			'Display new time pick
			response.write vbcrlf & "<tr>"
			response.write "<td class=""pricetd"" nowrap=""nowrap"" valign=""top"">" 
			response.write "<input type=""checkbox"" "

			If oRs("isfee") Then 
				'Always check a fee
				response.write " checked=""checked"" "
				cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
			Else 
				If oRs("isbaseprice") Then 
					'always check a base price
					response.write " checked=""checked"" "
       				cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
    			Else
					If oRs("pricetype") = sResidentType Then 
						'if the resident type requirement matches
						response.write " checked=""checked"" "
						cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
					Else
						If oRs("checkmembership") Then
							If oRs("pricetype") = sMemberType Then 
								response.write " checked=""checked"" "
								cTotalPrice = cTotalPrice + CDbl(oRs("amount"))
							End If 
						End If 
					End If 
				End If 
			End If 

			response.write "id=""pricetypeid" & oRs("pricetypeid") & """ name=""pricetypeid"" value=""" & oRs("pricetypeid") & """ onClick=""clearMsg('pricetypeid" & oRs("pricetypeid") & "');UpdatePriceTotal(document.PurchaseForm.amount" & oRs("pricetypeid") & ".value, this.checked);"" /> "
			response.write "&nbsp; " & oRs("pricetypename")
			response.write "</td>" 
			response.write "<td class=""priceentrytd"" valign=""top"">" 
			response.write "<input type=""text"" id=""amount" & oRs("pricetypeid") & """ name=""amount" & oRs("pricetypeid") & """ value=""" & Replace(FormatNumber(CDbl(oRs("amount")),2),",","") & """ size=""10"" maxlength=""9"" onchange=""clearMsg('amount" & oRs("pricetypeid") & "');ValidatePrice(this);"" />"
			response.write "</td>"
			response.write "<td class=""priceentrytd"">&nbsp;" & FormatCurrency(oRs("amount")) & "</td>"
			response.write "<td>"

			If sDiscount <> "" Then 
				response.write "(<input type=""checkbox"" name=""useOverrideDiscount" & oRs("pricetypeid") & """ value=""1"">Override Discount)"
			Else 
				response.write "<input type=""hidden"" name=""useOverrideDiscount" & oRs("pricetypeid") & """ value=""0"">"
				response.write "&nbsp;"
			End If 

			response.write "</td>" 
			response.write "<td class=""pricemember"">" 

			If oRs("ismember") Then 
				'Show the membership for the one that requires membership
				ShowMembership iMembershipId
			Else 
  				response.write " &nbsp; "
			End if  

			If oRs("isdropin") Then 
				'Input for drop in date
				response.write "Date: <input type=""text"" class=""datefield"" id=""dropindate" & oRs("pricetypeid") & """ name=""dropindate" & oRs("pricetypeid") & """ value=""" & FormatDateTime(date(),2) & """ />&nbsp;<span class=""calendarimg"" style=""cursor:hand;""><img src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""clearMsg('dropindate" & oRs("pricetypeid") & "');doCalendar('dropindate" & oRs("pricetypeid") & "');"" /></span>" 
			End If 

			response.write "</td>"
			response.write "<td class=""pricemember"">"

			If sDiscount <> "" Then 
  				response.write " (" & sDiscount & ")"
			Else 
		  		response.write " &nbsp; "
			End If 

			response.write "</td>"
			response.write "</tr>"

			iCount = iCount + 1
			oRs.MoveNext 
		Loop 

		' final display of an other price pick
	'	response.write vbcrlf & "<tr><td class=""pricetd"" nowrap=""nowrap""><input type=""radio"" name=""pricetypeid"" value=""0"" /> &nbsp; Other Price</td><td>"
	'	response.write "<input type=""text"" name=""amount"" onKeyUp=""AutoSelect();"" value="""" size=""6"" maxlength=""6"" /></td></tr>"
'		response.write vbcrlf & "<tr></tr>"
		response.write vbcrlf & "<tr>"
		response.write "<td>Total Price</td>"
		response.write "<td><span id=""displaytotalprice"">" & FormatNumber(cTotalPrice,2,,,0) & "</span></td>"
		response.write "<td colspan=""4"">&nbsp;</td>"
		response.write "</tr>"
		response.write "</table>" 
		response.write "<input type=""hidden"" id=""totalprice"" name=""totalprice"" value=""" & cTotalPrice & """ />"
		response.write "<input type=""hidden"" id=""minpricetypeid"" name=""minpricetypeid"" value=""" & iMinPricetype & """ />"
		response.write "<input type=""hidden"" id=""maxpricetypeid"" name=""maxpricetypeid"" value=""" & iMaxPriceType & """ />"
	End If 

	oRs.Close
	Set oRs = Nothing

End Sub 


'------------------------------------------------------------------------------
' boolean CheckResTypeExists( iClassid, iorgid, sResidentType )
'------------------------------------------------------------------------------
Function CheckResTypeExists( ByVal iClassid, ByVal iorgid, ByVal sResidentType )
	Dim sSql, oRs
	
	sSql = "SELECT COUNT(T.pricetype) AS hits "
	sSql = sSql & " FROM egov_price_types T, egov_class_pricetype_price P "
	sSql = sSql & " WHERE T.pricetypeid = P.pricetypeid "
	sSql = sSql & " AND orgid = " & iorgid & " AND P.classid = " & iClassid & " AND T.pricetype = '" & sResidentType & "'"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If CLng(oRs("hits")) > CLng(0) Then 
		CheckResTypeExists = True 
	Else
		CheckResTypeExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'------------------------------------------------------------------------------
' void ShowUserDropDown iUserId
'------------------------------------------------------------------------------
Sub ShowUserDropDown( ByVal iUserId )
	Dim oCmd, oRs

	Set oCmd = Server.CreateObject("ADODB.Command")
	With oCmd
		.ActiveConnection = Application("DSN")
	    .CommandText = "GetEgovUserWithAddressList"
	    .CommandType = 4
		.Parameters.Append oCmd.CreateParameter("@iOrgid", 3, 1, 4, Session("OrgID"))
	    Set oRs = .Execute
	End With

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("userid") & """"
		If CLng(iUserId) = CLng(oRs("userid")) Then
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("userlname") & ", " & oRs("userfname") & " &ndash; " & oRs("useraddress") & "</option>"
		oRs.MoveNext
	Loop 
		
	oRs.Close
	Set oRs = Nothing
	Set oCmd = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowUserInfo iUserId, sUserType, sResidentDesc
'--------------------------------------------------------------------------------------------------
Sub ShowUserInfo( ByVal iUserId, ByVal sUserType, ByVal sResidentDesc )
	Dim oRs, sSql

	sSql = "SELECT userfname, userlname, useraddress, useraddress2, usercity, userstate, "
	sSql = sSql & " userzip, usercountry, useremail, userhomephone, "
	sSql = sSql & " userworkphone, userfax, userbusinessname, userpassword, "
	sSql = sSql & " userregistered, residenttype, residencyverified, registrationblocked, "
	sSql = sSql & " blockeddate, blockedadminid, blockedexternalnote, blockedinternalnote "
	sSql = sSql & " FROM egov_users WHERE userid = " & iUserId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	response.write vbcrlf & "<table border=""0"" cellpadding=""5"" cellspacing=""0"" id=""signupuserinfo"">"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Name:</td><td >" & oRs("userfname") & " " & oRs("userlname") & "&nbsp;&nbsp;&nbsp;<strong>" & sResidentDesc & "</strong>"

	If Not oRs("residencyverified") And oRs("residenttype") = "R" Then 
		If lcl_orghasfeature_residency_verification Then 
			response.write " (not verified)"
		End If 
	End If 

	response.write "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Email:</td><td>" & oRs("useremail") & "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Phone:</td><td>" & FormatPhone(oRs("userhomephone")) & "</td></tr>"
	response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Address:</td><td>" & oRs("useraddress") & "<br />" 

	If oRs("useraddress2") <> "" Then 
  		response.write oRs("useraddress2") & "<br />" 
	End If 

	If oRs("usercity") <> "" Or oRs("userstate") <> "" Or oRs("userzip") <> "" Then 
		  response.write oRs("usercity") & ", " & oRs("userstate") & " " & oRs("userzip") 
	End If 

	response.write "</td></tr>"

	' Handle blocked
	If oRs("registrationblocked") Then
		bRegistrationBlocked = True 
		response.write vbcrlf & "<tr><td colspan=""2""><span id=""warningmsg""> *** Registration Blocked *** </span></td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">Date:</td><td>" & oRs("blockeddate") & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"">By:</td><td>" & GetAdminName( oRs("blockedadminid") ) & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">Internal Note:</td><td>" & oRs.Fields("blockedinternalnote") & "</td></tr>"
		response.write vbcrlf & "<tr><td align=""right"" valign=""top"" nowwrap=""nowrap"">External Note:</td><td>" & oRs.Fields("blockedexternalnote") & "</td></tr>"
	End If 

	response.write vbcrlf & "</table>"

	oRs.Close
	Set oRs = Nothing
	
End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentChoices
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentChoices()

	response.write vbcrlf & "<p><strong>Payment</strong><br />"
	
	response.write vbcrlf & "<select name=""paymenttypeid"" size=""1"">"
	ShowPaymentTypes
	response.write vbcrlf & "</select>&nbsp;&nbsp;&nbsp;&nbsp;"

	response.write vbcrlf & "<select name=""paymentlocationid"" size=""1"">"
	ShowPaymentLocations
	response.write vbcrlf & "</select>&nbsp;&nbsp;&nbsp;&nbsp;"

	response.write vbcrlf & "<input type=""button"" name=""complete"" class=""button"" style=""width:140px;text-align:center;"" value=""Complete Purchase"" onclick=""ValidateForm();"" />"
	response.write vbcrlf & "</p>"
	
	response.write vbcrlf & "<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<strong>OR</strong>"
	response.write vbcrlf & "&nbsp;&nbsp;&nbsp;<input type=""button"" name=""waitlist"" class=""button"" value=""Add to Wait List"" onclick=""ValidateWait();"" /></p>"

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentTypes
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentTypes()
	Dim sSql, oRs

	sSql = "SELECT paymenttypeid, paymenttypename FROM egov_paymenttypes ORDER BY paymenttypeid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymenttypeid") & """>" & oRs("paymenttypename") & "</option>"
		oRs.MoveNext 
	Loop

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' void ShowPaymentLocations
'--------------------------------------------------------------------------------------------------
Sub ShowPaymentLocations()
	Dim sSql, oRs

	sSql = "SELECT paymentlocationid, paymentlocationname FROM egov_paymentlocations ORDER BY paymentlocationid"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	Do While Not oRs.EOF
		response.write vbcrlf & "<option value=""" & oRs("paymentlocationid") & """>" & oRs("paymentlocationname") & "</option>"
		oRs.movenext 
	Loop

	oRs.Close
	Set oRs = Nothing

End Sub 


'--------------------------------------------------------------------------------------------------
' boolean CheckIfFullSingle( iTimeId )
'--------------------------------------------------------------------------------------------------
Function CheckIfFullSingle( ByVal iTimeId )
	Dim sSql, oRs

	sSql = "SELECT ISNULL(max,999999) AS max, enrollmentsize FROM egov_class_time WHERE timeid = " & iTimeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1

	If CLng(oRs("enrollmentsize")) >= CLng(oRs("max")) Then
		CheckIfFullSingle = True 
	Else
		CheckIfFullSingle = False 
	End If 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' boolean CheckIfFullSeries( iClassid )
'--------------------------------------------------------------------------------------------------
Function CheckIfFullSeries( ByVal iClassid )
	Dim sSql, oRs

	sSql = "SELECT T.timeid FROM egov_class C, egov_class_time T "
	sSql = sSql & "WHERE C.classid = T.classid AND C.parentclassid = " & iClassid

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	Do While Not oRs.EOF
		If CheckIfFullSingle( oRs("timeid") ) Then
			CheckIfFullSeries = True 
			Exit Do 
		Else
			CheckIfFullSeries = False 
		End If 
		oRs.MoveNext
	Loop 

	oRs.Close
	Set oRs = Nothing

End Function 


'--------------------------------------------------------------------------------------------------
' void ShowAvailability iClassid, bIsParent, iOptionid, iTimeId 
'--------------------------------------------------------------------------------------------------
Sub ShowAvailability( ByVal iClassid, ByVal bIsParent, ByVal iOptionid, ByVal iTimeId )
	Dim sSql, oRs

	If bIsParent Then 
		' Get the availability of the children events
		sSql = "SELECT T.timeid, T.starttime, T.endtime, ISNULL(T.min,0) AS min, ISNULL(T.max,0) AS max, "
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " FROM egov_class_time T, egov_class C"
		sSql = sSql & " WHERE C.parentclassid = " & iClassid
		sSql = sSql & " AND T.timeid = " & iTimeId 
		sSql = sSql & " AND C.classid = T.classid"
		sSql = sSql & " ORDER BY C.startdate, T.starttime"
	Else 
		' Get the single event
		sSql = "SELECT T.timeid, T.starttime, T.endtime, ISNULL(T.min,0) AS min, ISNULL(T.max,0) AS max,"
		sSql = sSql & " T.enrollmentsize, T.waitlistsize, C.startdate"
		sSql = sSql & " FROM egov_class_time T, egov_class C WHERE C.classid = " & iClassid 
		sSql = sSql & " AND T.timeid = " & iTimeId 
		sSql = sSql & " AND C.classid = T.classid ORDER BY T.starttime"
	End If 

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 3, 1
	
	If Not oRs.EOF Then 
		response.write vbcrlf & "<table id=""tableavail"" border=""0"" cellpadding=""2"" cellspacing""0"">"
		'response.write vbcrlf & "<caption>Availability</caption>"

		If bIsParent Then
			response.write vbcrlf & "<tr><th>Date</th><th>Time</th><th>Min</th><th>Max</th>"
		Else
			response.write vbcrlf & "<tr><th>Time</th><th>Min</th><th>Max</th>"
		End If
		If iOptionid = 1 then
			response.write "<th>Enrolled</th>"
		Else 
			response.write "<th>Attending</th>"
		End If 
		response.write "<th>Available</th><th>Waiting</th></tr>"
		 
		Do While Not oRs.EOF
			response.write vbcrlf & "<tr><td>" 
			If bIsParent Then
				response.write DatePart("m",oRs("startdate")) & "/" & DatePart("d",oRs("startdate")) & "</td><td>"
			End If 
			response.write oRs("starttime") 
			If oRs("endtime") <> oRs("starttime") Then
				response.write "&ndash;" & oRs("endtime")
			End If 
			response.write "</td>"
			response.write "<td align=""center"">" 
			If CLng(oRs("min")) = 0 Then
				response.write "none"
			Else
				response.write oRs("min")
			End If 
			response.write "</td><td align=""center"">" 
			If CLng(oRs("max")) = 0 Then
				response.write "none"
			Else
				response.write oRs("max")
			End If
			' enrollment
			response.write "</td><td align=""center"">" & oRs("enrollmentsize") & "</td>"
			' availability
			response.write "<td align=""center"">" 
			If CLng(oRs("max")) = 0 Then
				response.write "n/a"
			Else
				iAvail = CLng(oRs("max")) - CLng(oRs("enrollmentsize"))
				If iAvail < 0 Then 
					iAvail = 0
				End If 
				response.write CLng(oRs("max")) - CLng(oRs("enrollmentsize"))
			End If 
			response.write "</td>"
			' get waitlist here
			response.write "<td align=""center"">" & oRs("waitlistsize") & "</td>"
			response.write "</tr>"
			oRs.MoveNext 
		Loop 
		response.write vbcrlf & "</table>"
	End If 

	oRs.Close
	Set oRs = Nothing

End sub  

'------------------------------------------------------------------------------
sub displayTeamRosterAccessories(p_orgid, p_classid, p_type )

  dim sOrgID, sClassID, sType

  sOrgID   = 0
  sClassID = 0
  sType    = "TSHIRT"

  if p_orgid <> "" then
     sOrgID = clng(p_orgid)
  end if

  if p_classid <> "" then
     sClassID = clng(p_classid)
  end if

  if p_type <> "" then
     if not containsApostrophe(p_type) then
        lcl_type = ucase(p_type)
     end if
  else
     lcl_type = "TSHIRT"
  end if

  lcl_type = dbsafe(lcl_type)
  lcl_type = "'" & lcl_type & "'"

  sSql = "SELECT "
  sSql = sSql & " atc.accessoryid, "
  sSql = sSql & " a.accessoryname, "
  sSql = sSql & " a.accessoryvalue "
  sSql = sSql & " FROM egov_class_teamroster_accessories_to_class atc "
  sSql = sSql &      " INNER JOIN egov_class_teamroster_accessories a ON atc.accessoryid = a.accessoryid "
  sSql = sSql & " WHERE a.orgid = " & sOrgID
  sSql = sSql & " AND atc.classid = " & sClassID
  sSql = sSql & " AND UPPER(a.accessorytype) = " & lcl_type
  sSql = sSql & " ORDER BY isnull(atc.displayorder,a.displayorder) "

  set oClassAccessories = Server.CreateObject("ADODB.Recordset")
  oClassAccessories.Open sSql, Application("DSN"), 3, 1

  if not oClassAccessories.eof then
     do while not oClassAccessories.eof
        response.write "  <option value=""" & oClassAccessories("accessoryvalue") & """>" & oClassAccessories("accessoryname") & "</option>" 

        oClassAccessories.movenext
     loop
  end if

  oClassAccessories.close
  set oClassAccessories = nothing 

  'response.write "            <option value=""Youth - Small (6-8)"">Youth - Small (6-8)</option>" 
  'response.write "            <option value=""Youth - Medium (10-12)"">Youth - Medium (10-12)</option>" 
  'response.write "            <option value=""Youth - Large (14-16)"">Youth - Large (14-16)</option>" 
  'response.write "            <option value=""Adult - Small (34-36)"">Adult - Small (34-36)</option>" 
  'response.write "            <option value=""Adult - Medium (38-40)"">Adult - Medium (38-40)</option>" 
  'response.write "            <option value=""Adult - Large (40-42)"">Adult - Large (40-42)</option>" 
  'response.write "            <option value=""Adult - X-Large (44-46)"">Adult - X-Large (44-46)</option>" 

end sub


'--------------------------------------------------------------------------------------------------
' void ShowCitizenPicks iCitizenUserid, sSearchName
'--------------------------------------------------------------------------------------------------
Sub ShowCitizenPicks( ByVal iCitizenUserid, ByVal sName )
	Dim sSearchName, sSql, oRs

	sSearchName = dbsafe(sName)

	' This query finds last names that start with the search ahead of those that match anywhere
	sSql = "SELECT 1 AS foo, userid AS userid, userfname AS firstname, userlname AS lastname, "
	sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
	sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
	sSql = sSql & " AND userlname LIKE '" & sSearchName & "%' "
	sSql = sSql & " UNION "
	sSql = sSql & " SELECT 2 AS foo, userid, userfname AS firstname, userlname AS lastname, "
	sSql = sSql & " ISNULL(userlname,'') + ISNULL(userfname,'') AS sortname, useraddress AS address "
	sSql = sSql & " FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 AND headofhousehold = 1 AND userregistered = 1 "
	sSql = sSql & " AND ( userfname LIKE '%" & sSearchName & "%' OR userlname LIKE '%" & sSearchName & "%' ) "
	sSql = sSql & " AND userid NOT IN ( SELECT userid FROM egov_users WHERE orgid = " & session("orgid") & " AND isdeleted = 0 "
	sSql = sSql & " AND headofhousehold = 1 AND userregistered = 1 AND userlname LIKE '" & sSearchName & "%' ) "
	sSql = sSql & " ORDER BY foo, sortname"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRS.EOF Then
		response.write "Select a Name: <select name='egovuserid' id='egovuserid'>"
		Do While Not oRs.EOF
			response.write vbcrlf & "<option value='" & oRs("userid") & "'"
			If iCitizenUserid = CLng(oRs("userid")) Then
				response.write " selected=""selected"" "
			End If 
			response.write ">"
			response.write oRs("lastname") & ", " & oRs("firstname")
			If oRs("address") <> "" Then
				response.write " - " & oRs("address")
			End If 

			response.write "</option>"

			oRs.MoveNext
		Loop 
		response.write vbcrlf & "</select>"
	Else
		response.write "<input type='hidden' name='egovuserid' id='egovuserid' value='0' />No Matching Names Found"
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

%>
