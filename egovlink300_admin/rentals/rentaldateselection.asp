<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="rentalsguifunctions.asp" //-->
<!-- #include file="rentalscommonfunctions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: rentaldateselection.asp
' AUTHOR: Steve Loar
' CREATED: 10/12/2009
' COPYRIGHT: Copyright 2009 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Rental date and time selection page in the reservation process.
'
' MODIFICATION HISTORY
' 1.0 10/12/2009	Steve Loar - INITIAL VERSION
' 2.0 03/01/2012 David Boyer - Redesigned screen to handle multiple rentals.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
'Force the page to be re-loaded on back button
 response.Expires = 60
 response.Expiresabsolute = Now() - 1
 response.AddHeader "pragma","no-store"
 response.AddHeader "cache-control","private"
 response.CacheControl = "no-store" 'This prevents problems that occur when they hit the back button after a purchase

 dim iRentalId, iReservationTempId, sReservationType, sCitizenName, sStartDate, sStartTime
 dim sEndDate, sEndTime, sOccurs, sPeriodType, sPeriodTypeSelector, bIsReservation, iTotalRows
 dim sReservationTypeId, sSearchName, iRecreationCategoryId, iLocationId, sRentalName
 dim iPeriodTypeId, iStartHour, iStartMinute, sStartAmPm, iEndHour, iEndMinute, sEndAmPm, iEndDay
 dim sOccursChecked, iMonthlyPeriod, iMonthlyDOW, iOrderBy, sWantedDOWs, sLoadMsg
 dim iReservationTypeId, iRentalUserid, bIsPublicUsers, sNoCostPhrase, bHasUsers, iReservationId
 dim lcl_selected_rentalids, lcl_rental_linecount, lcl_scripts
 dim sSearchName2
 sSearchName2 = ""

 sLevel                 = "../"  'Override of value from common.asp
 iTotalRows             = 0
 sLoadMsg               = ""
 lcl_selected_rentalids = ""
 lcl_rental_linecount   = 0
 lcl_scripts            = ""
 lcl_createpath         = ""

'check if page is online and user has permissions in one call not two
 PageDisplayCheck "make reservations", sLevel	' In common.asp

If request("rti") = "" Then
	response.redirect "rentalsearch.asp"
Else 
	iReservationTempId = CLng(request("rti"))
End If 

If Not TempReservationExists( iReservationTempId ) Then
	response.redirect "rentalsearch.asp"
End If 

if request("selected_rentalids") <> "" then
   if not containsApostrophe(request("selected_rentalids")) then
      lcl_selected_rentalids = request("selected_rentalids")
   end if
end if
'iRentalId      = CLng(request("rentalid"))
iReservationId = CLng(request("rid"))

If request("reservationtypeid") <> "" Then 
	iReservationTypeId = CLng(request("reservationtypeid"))
Else
	iReservationTypeId = GetFirstReservationTypeInList( )
End If

'If RentalHasNoCosts( iRentalId ) Then
'	sNoCostPhrase = "<strong>There is no cost to rent this.</strong>"
'Else
'	sNoCostPhrase = ""
'End If 

If CLng(iReservationTypeId) > CLng(0) Then 
	GetUserFlagsFromReservationTypeId iReservationTypeId, bHasUsers, bIsPublicUsers
Else
	bHasUsers = False 
	bIsPublicUsers = False 
End If 

If request("searchname") <> "" Then
	sSearchName = request("searchname")
Else
	sSearchName = ""
End If 
If request("searchname2") <> "" Then
	sSearchName2 = request("searchname2")
Else
	sSearchName2 = ""
End If 

If bHasUsers Then 
	If request("rentaluserid") <> "" Then 
		iRentalUserid = CLng(request("rentaluserid"))
	Else
		iRentalUserid = CLng(0)
	End If 
Else
	iRentalUserid = CLng(0)
End If 

'Get Temp Rental Info from the table
 GetTempRentalValues iReservationTempId

'Determine which "create path" the user took to get here.
'  SIMPLE = Make Simple Reservations
'  ""     = Make Reservations
 if request("createpath") <> "" then
    if not containsApostrophe(request("createpath")) then
       lcl_createpath = ucase(request("createpath"))
    end if
 end if

'Set up the BACK button
 lcl_goBackOnClick = "goBack();"

 if lcl_createpath = "SIMPLE" then
    lcl_goBackOnClick = "goBackSimple();"
 end if
%>
<html>
<head>
	<title>E-Gov Administration Console</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="rentalsstyles.css" />

 <style>
   .reservationDatesDiv {
      display: none;
   }

   .reservationErrorsLegend {
      margin-bottom: 2px;
      color:         #ff0000;
      font-size:     12pt;
      font-weight:   bold;
   }

   .dateSelectionLegendText {
      padding-right:         5px;
      border:                1pt solid #c0c0c0 !important;
      -webkit-border-radius: 5px;
      -moz-border-radius:    5px;
   }

   .reservationErrorsDiv {
      margin-bottom: 10px;
   }

   .noCostPhrase {
      color:         #ff0000;
      font-weight:   bold;
      text-align:    right;
      margin-bottom: 5pt;
   }

   #continueMsg {
      color:     #ff0000;
      font-size: 13pt;
   }

   .fieldset {
      margin-bottom: 10px;
   }

   .reservationDate {
      color: #ff0000;
   }
 </style>

	<script language="javaScript" src="../prototype/prototype-1.6.0.2.js"></script>
 <script type="text/javascript" src="../scripts/jquery-1.7.2.min.js"></script>

	<script language="javaScript" src="../scripts/ajaxLib.js"></script>
 <script language="javascript" src="../scripts/formvalidation_msgdisplay.js"></script>

	<script language="javascript">
	<!--

  //This is set so that we can use jQuery with prototype.
  jQuery.noConflict();

		function EditApplicant()
		{
			var strPickedUserId = document.frmDateSelection.rentaluserid.options[document.frmDateSelection.rentaluserid.selectedIndex].value;
			var myRand = parseInt(Math.random() * 99999999 );
			eval('window.open("rentaluseredit.asp?userid=' + strPickedUserId + '&rand=' + myRand + '", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
		}


		function newUser()
		{
			//location.href='../dirs/register_citizen.asp';
			var myRand = parseInt(Math.random() * 99999999 );
			eval('window.open("../dirs/register_citizen.asp?rand=' + myRand + '", "_picker", "width=800,height=800,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=10,top=10")');
		}

		function searchName()
		{
			if ($("searchname").value != "")
			{
				//alert($("searchname").value);
				// Try to get a drop down of names
				doAjax('getcitizenpicks.asp', 'searchname=' + $("searchname").value, 'UpdateApplicants', 'get', '0');
			}
			else
			{
				alert('Please enter a name before searching.');
				$("searchname").focus();
			}
		}

		function UpdateApplicants( sResult )
		{
			//alert("Back");
			$("applicant").innerHTML = sResult;
			if (sResult.substr(0,6) == "Select")
			{
				$("edituserbtn").style.visibility = 'visible';
				document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
			}
			else
				$("edituserbtn").style.visibility = 'hidden';
		}

		function WaitAndValidate(iAction) {
     var lcl_action = 'CHECK_RESERVE';

     if((iAction != '') && (iAction != undefined)) {
        lcl_action = iAction.toUpperCase();
     }

     jQuery('#checkbutton').prop('disabled',true);
     jQuery('#continuebutton').prop('disabled',true);

  			// This causes the reservation to have a 0 to 3 sec wait before checking availability to try and break ties
		  	var ReturnTime = Math.floor(Math.random()*3000);
  			setTimeout('validate(\"' + lcl_action + '\")', ReturnTime);
		}

		function validate(iAction) {
     var lcl_total_rentals                = jQuery('#total_rentals').val();
     var lcl_total_dates                  = Number(0);
     var lcl_firstRentalID                = Number(0);
     var lcl_false_count                  = Number(0);
     var lcl_false_count_total            = Number(0);
     var lcl_totalErrors_alertsWarnings   = Number(0);
     var lcl_totalCount_checkDatesReserve = Number(0);
     var lcl_maxrows                      = Number(0);
     var lcl_false_count                  = Number(0);
     var lcl_rentalid                     = '';
     var lcl_action                       = 'CHECK_RESERVE';
     var lcl_stopProcessing               = false;

     jQuery('#falseCountTotal').val('');

     //Determine which button was pressed.
     if((iAction != '') && (iAction != undefined)) {
        lcl_action = iAction.toUpperCase();
     }

     if(jQuery('#totalCount_checkDatesReserve').val() != '') {
        lcl_totalCount_checkDatesReserve = Number(jQuery('#totalCount_checkDatesReserve').val());
     }

     //Setting this field to 'Y' enabled the .AjaxStop jQuery function to execute
     //which will submit the form when no errors exist.
     if(lcl_action == 'CHECK_RESERVE') {
        jQuery('#checkReserveValidate').val('Y');
     } else { //lcl_action == 'CHECK_DATES'
        jQuery('#checkReserveValidate').val('N');
     }

     if(lcl_false_count > 0) {
        lcl_stopProcessing = true;
     }

     //Stop all further validation IF user is attempting to "check and reserve"
     //and this is NOT the initial attempt to perform this action already AND
     //errors exist from a previous validation attempt AND there are action in
     //the errors that require the user to select an action.
     if(lcl_stopProcessing) {
        jQuery('#checkbutton').prop('disabled',false);
        jQuery('#continuebutton').prop('disabled',false);
     } else {
        jQuery('input[id^="okToContinue"]').each(function(index) {
           jQuery(this).css('display','none');
        });

        jQuery('fieldset[id^="fieldset_reservationerrors_"]').each(function(index) {
           jQuery(this).css('display','none');
        });

        jQuery('.reservationErrorsDiv').each(function(index) {
           jQuery(this).css('display','none');
        });

        if(jQuery('#firstRentalID').val() != '') {
           lcl_firstRentalID = Number(jQuery('#firstRentalID').val());
        }

        //We want to track the number of times this validation is executed.
        //This will help us later to determine if this is the intial time
        //the validation is executed and errors exist, or if errors already exist
        //and the user is attempting, again, to perform a "check and reserve".
        lcl_totalCount_checkDatesReserve = lcl_totalCount_checkDatesReserve + 1;
        jQuery('#totalCount_checkDatesReserve').val(lcl_totalCount_checkDatesReserve)

        //BEGIN: Expand all Rental reservation sections
        //------------------------------------------------------------------------
        //Loop through all of the rentals and expand any/all sections that are hidden
        //so that we can validate each field properly.
        for (var i = parseInt(lcl_total_rentals); i >= 1 ; i--) {
           lcl_rentalid = jQuery('#rentalid' + i).val();

           displayScreenMsg('Processing...');
           jQuery('#checkDateTimes'         + lcl_rentalid).prop('disabled',true);
           jQuery('#hideDateTimes'          + lcl_rentalid).prop('disabled',false);
           jQuery('#div_reservationdates_'  + lcl_rentalid).css('display','block');
           jQuery('#div_reservationerrors_' + lcl_rentalid).html('')
           //jQuery('#reservationErrorsDiv_' + lcl_rentalid).css('display','none');

           //If this is the first rental section (last in loop) then move on a validate the fields
           // or stop processing to allow the user to fix any/all errors.
           if(Number(lcl_rentalid) == lcl_firstRentalID) {

              //BEGIN: Loop through each rental section and validate the fields and
              //determine if any field validation errors exist.  This field validation
              //consists of ONLY:
              //  1. checking the start/end times to ensure they are NOT EQUAL
              //  2. the end time is GREATER THAN the start time
              //  3. the date is NOT blank
              //If errors exist, do NOT show them yet.  This is to just make sure
              //The sections without errors are collapsed so it doesn't mess up
              //the location of the error message(s).  We are simply counting the
              //errors until the end of the loop and then we move on with the processing.
              //------------------------------------------------------------------
              for (var v = parseInt(lcl_total_rentals); v >= 1 ; v--) {
                 lcl_false_count = 0;

                 //Get the rentalid for this row
                 lcl_rentalid    = jQuery('#rentalid' + v).val();
                 lcl_total_dates = jQuery('#maxrows' + lcl_rentalid).val();

                 //BEGIN: Validate fields in each section
                 //---------------------------------------------------------------
                 displayScreenMsg('Processing: Validating Fields...');
                 validateFields('HideErrors', lcl_rentalid, lcl_total_dates);

                 //Update the total errors for all sections
                 lcl_false_count     = jQuery('#totalErrors_fieldValidation' + lcl_rentalid).val();
                 lcl_falseCountTotal = jQuery('#falseCountTotal').val();

                 if(lcl_falseCountTotal != '') {
                    lcl_falseCountTotal = Number(lcl_falseCountTotal);
                 } else {
                    lcl_falseCountTotal = Number(0);
                 }

                 lcl_falseCountTotal = lcl_falseCountTotal + Number(lcl_false_count);

                 jQuery('#falseCountTotal').val(lcl_falseCountTotal);

                 if(lcl_falseCountTotal > 0) {
                    jQuery('#checkbutton').prop('disabled',false);
                    jQuery('#continuebutton').prop('disabled',false);
                 }

                 //If this is the first rental section (last in loop) then move on a validate
                 //the fields or stop processing to allow the user to fix any/all errors.
                 if(Number(lcl_rentalid) == lcl_firstRentalID) {

                    //BEGIN: Hide sections without errors
                    //------------------------------------------------------------
                    //Now loop through each rental section to hide it.
                    //Check for the total errors in each setion.
                    //If any exist then do NOT collapse the section
                    var lcl_totalErrors_fieldValidation = Number(0);

                    for (var h = parseInt(lcl_total_rentals); h >= 1 ; h--) {
                       lcl_rentalid                    = jQuery('#rentalid' + h).val();
                       lcl_totalErrors_fieldValidation = jQuery('#totalErrors_fieldValidation' + lcl_rentalid).val();

                       if(lcl_totalErrors_fieldValidation != '') {
                          if(lcl_totalErrors_fieldValidation == 0) {
                             hideDatesTimes('CSS',lcl_rentalid);
                          }
                       }

                       //BEGIN: Show validation error messages or continue processing
                       //---------------------------------------------------------
                       //If this is the last rental section, we should now have any sections that have at least 1 error
                       //  already expanded and the others collapsed.  We now need to loop through each section again
                       //  and re-validate the fields so that we can show any errors messages properly.
                       if(Number(lcl_rentalid) == lcl_firstRentalID) {
                          lcl_falseCountTotal = jQuery('#falseCountTotal').val();

                          //if there errors in ANY rental section then stop the processing.
                          //Otherwise, continue and check the date/times
                          if(Number(lcl_falseCountTotal) > 0) {

                             //BEGIN: Show error messages ---------------------------
                             for (var r = parseInt(lcl_total_rentals); r >= 1 ; r--) {
                                lcl_false_count = 0;

                                //Get the rentalid for this row
                                lcl_rentalid    = jQuery('#rentalid' + r).val();
                                lcl_total_dates = jQuery('#maxrows' + lcl_rentalid).val();

                                validateFields('ShowErrors', lcl_rentalid, lcl_total_dates);

                                //If this is the last rental section then determine if we are to process
                                //  the rental reservations or stop processing to allow the user to fix any/all errors.
                                if(Number(lcl_rentalid) == lcl_firstRentalID) {
                                   jQuery('#checkbutton').prop('disabled',false);
                                   jQuery('#continuebutton').prop('disabled',false);

                                   return false;
                                }
                             }
                          } else {
                             displayScreenMsg('Processing...Checking date/times...');

                             //First check to see if the user has already attempted to validate the
                             //reservation(s) and if errors exist as well as errors that require user
                             //interaction to continue processing.
                             jQuery('#totalErrors_rentals').val('0');

                             for (var r = parseInt(lcl_total_rentals); r >= 1 ; r--) {
                                lcl_rti               = jQuery('#rti').val();
                                lcl_rentalid          = jQuery('#rentalid' + r).val();
                                lcl_reservationtypeid = jQuery('#reservationtypeid').val();
                                lcl_rentaluserid      = jQuery('#rentaluserid').val();
                                lcl_maxrows           = jQuery('#maxrows'      + lcl_rentalid).val();

                                jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val('0');
                                lcl_totalErrors_alertsWarnings = Number(0);

                                for (var d = 1; d <= parseInt(lcl_maxrows); d++) {
                                   //If the "Check Dates" button has been pressed then it does not matter if any errors
                                   //exist as we simply revalidate the dates, even if the user has set the "Ok to continue" to "Yes".
                                   if(lcl_action == 'CHECK_DATES') {
                                      jQuery('#errorCheck_reservationDateTime_' + lcl_rentalid + '_' + d).val('');
                                      jQuery('#errorCheck_errorCode_'           + lcl_rentalid + '_' + d).val('');
                                      jQuery('#errorCheck_okToContinue_'        + lcl_rentalid + '_' + d).val('');
                                   }

                                   lcl_startdate                      = jQuery('#startdate_'   + lcl_rentalid + '_' + d).val();
                                   lcl_starthour                      = jQuery('#starthour_'   + lcl_rentalid + '_' + d).val();
                                   lcl_startminute                    = jQuery('#startminute_' + lcl_rentalid + '_' + d).val();
                                   lcl_startampm                      = jQuery('#startampm_'   + lcl_rentalid + '_' + d).val();
                                   lcl_endhour                        = jQuery('#endhour_'     + lcl_rentalid + '_' + d).val();
                                   lcl_endminute                      = jQuery('#endminute_'   + lcl_rentalid + '_' + d).val();
                                   lcl_endampm                        = jQuery('#endampm_'     + lcl_rentalid + '_' + d).val();
                                   lcl_endday                         = jQuery('#endday_'      + lcl_rentalid + '_' + d).val();
                                   lcl_errorCheck_reservationDateTime = jQuery('#errorCheck_reservationDateTime_' + lcl_rentalid + '_' + d).val();
                                   lcl_errorCheck_errorCode           = jQuery('#errorCheck_errorCode_'           + lcl_rentalid + '_' + d).val();
                                   lcl_errorCheck_okToContinue        = jQuery('#errorCheck_okToContinue_'        + lcl_rentalid + '_' + d).val();
                                   lcl_includeReservationTime         = false;

                                   if(document.getElementById('includereservationtime_' + lcl_rentalid + '_' + d).checked) {
                                      lcl_includeReservationTime = true;
                                   }

                                   if(lcl_includeReservationTime) {
                                      jQuery('#lineCount').val(d);

                                      jQuery.post('checkselecteddates.asp', {
                                         rentalid:                       lcl_rentalid,
                                         linenum:                        d,
                                         maxrows:                        lcl_maxrows,
                                         rti:                            lcl_rti,
                                         reservationtypeid:              lcl_reservationtypeid,
                                         rentaluserid:                   lcl_rentaluserid,
                                         startdate:                      lcl_startdate,
                                         starthour:                      lcl_starthour,
                                         startminute:                    lcl_startminute,
                                         startampm:                      lcl_startampm,
                                         endhour:                        lcl_endhour,
                                         endminute:                      lcl_endminute,
                                         endampm:                        lcl_endampm,
                                         endday:                         lcl_endday,
                                         includereservationtime:         lcl_includeReservationTime,
                                         includeRentalID:                true,
                                         errorCheck_reservationDateTime: lcl_errorCheck_reservationDateTime,
                                         errorCheck_errorCode:           lcl_errorCheck_errorCode,
                                         errorCheck_okToContinue:        lcl_errorCheck_okToContinue
                                      }, function(results) {
                                         var lcl_underscore_position1 = Number(0);
                                         var lcl_underscore_position2 = Number(0);
                                         var lcl_linenum              = Number(0);
                                         var lcl_totalErrors_rentals  = Number(0);
                                         var lcl_increaseErrorCount   = Number(1);
                                         var lcl_rentalid             = '';
                                         var lcl_error                = '';
                                         var lcl_results              = '';

                                         if(results != '') {
                                            lcl_results = results;
                                         }
                                         //Return value (result) format: [RentalID]_[lineNum]_[errorCode]
                                         //Find the position of the first underscore so that we can get the rentalid
                                         if(lcl_results.indexOf('_') > -1) {
                                            lcl_underscore_position1 = lcl_results.indexOf('_');
                                         }

                                         if(lcl_underscore_position1 > 0) {
                                            lcl_rentalid = lcl_results.substr(0,lcl_underscore_position1);
                                            lcl_results  = lcl_results.replace(lcl_rentalid + '_','');
                                         }

                                         //Now that the RentalID and the first underscore have been trimmed off
                                         //we now want to get the linenum and error code so we know which row we are working with.
                                         if(lcl_results.indexOf('_') > -1) {
                                            lcl_underscore_position2 = lcl_results.indexOf('_');

                                            if(lcl_underscore_position2 > 0) {
                                               lcl_linenum = lcl_results.substr(0,lcl_underscore_position2);
                                               lcl_results = lcl_results.replace(lcl_linenum + '_','');
                                               lcl_error   = lcl_results;
                                            } else {
                                               lcl_error = lcl_results;
                                            }
                                         } else {
                                            lcl_error = lcl_results;
                                         }

                                         if(lcl_error.substr(0,2) != 'OK') {
                                            showDatesTimes(lcl_rentalid);

                                            var lcl_dateErrors = jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val();

                                            if(lcl_dateErrors != '') {
                                               lcl_dateErrors = Number(lcl_dateErrors) + 1;
                                            } else {
                                               lcl_dateErrors = Number(1);
                                            }

                                            jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val(lcl_dateErrors);

                                            //BEGIN: Show the error message(s)
                                            var lcl_errorcode       = lcl_error;
                                            var lcl_reservationdate = '';

                                            if (lcl_error.substr(0,14) == 'shortnoconfirm') {
                                                lcl_errorcode = 'shortnoconfirm';
                                            } else if (lcl_error.substr(0,20) == 'buffershortnoconfirm') {
                                                lcl_errorcode = 'buffershortnoconfirm';
                                            } else if (lcl_error.substr(0,15) == 'buffernoconfirm') {
                                                lcl_errorcode = 'buffernoconfirm';
                                            } else if (lcl_error.substr(0,6) == 'closed') {
                                                lcl_errorcode = 'closed';
                                            } else if (lcl_error.substr(0,6) == 'nouser') {
                                                lcl_errorcode = 'nouser';
                                            } else if (lcl_error.substr(0,8) == 'conflict') {
                                                lcl_errorcode = 'conflict';
                                            //User Interaction CONFIRMED errorCodes (i.e. User HAS okayed date/times)
                                            } else if(lcl_error.substr(0,7) == 'shortOK') {
                                                lcl_errorcode = 'shortok';
                                            } else if (lcl_error.substr(0,13) == 'buffershortOK') {
                                                lcl_errorcode = 'buffershortok';
                                            } else if (lcl_error.substr(0,8) == 'bufferOK') {
                                                lcl_errorcode = 'bufferok';
                                            //User Interaction Needed errorCodes (i.e. Okay to continue)
                                            } else if(lcl_error.substr(0,5) == 'short') {
                                                lcl_errorcode = 'short';
                                            } else if (lcl_error.substr(0,11) == 'buffershort') {
                                                lcl_errorcode = 'buffershort';
                                            } else if (lcl_error.substr(0,6) == 'buffer') {
                                                lcl_errorcode = 'buffer';
                                            //} else if (lcl_error.substr(0,2) == 'OK') {
                                            //    lcl_errorcode = 'OK';
                                            }

                                            if(lcl_errorcode != '') {
                                               //lcl_reservationdate = lcl_error.substr(lcl_errorcode.length);
                                               lcl_reservationdate = jQuery('#startdate_' + lcl_rentalid + '_' + lcl_linenum).val();
                                               lcl_errortext       = getErrorText(lcl_errorcode);

                                               jQuery('#errorCheck_errorCode_'       + lcl_rentalid + '_' + lcl_linenum).val(lcl_errorcode);
                                               jQuery('#fieldset_reservationerrors_' + lcl_rentalid).css('display','block');
                                               jQuery('#div_reservationerrors_'      + lcl_rentalid).css('display','block');

                                               lcl_reservationerrors = jQuery('#div_reservationerrors_' + lcl_rentalid).html();

                                               if(lcl_reservationerrors == '') {
                                                  lcl_reservationerrors = lcl_reservationerrors + '<div id="reservationError' + lcl_rentalid + '_' + lcl_linenum + '"><span class=\"reservationDate\">' + lcl_reservationdate + ':</span>&nbsp;' + lcl_errortext;
                                               } else {
                                                  lcl_reservationerrors = lcl_reservationerrors + '</div><div id="reservationError' + lcl_rentalid + '_' + lcl_linenum + '"><span class=\"reservationDate\">' + lcl_reservationdate + ':</span>&nbsp;' + lcl_errortext;
                                               }

                                               //Determine if error code requires input from the user.
                                               //ErrorCodes that require user input: (short, buffer, buffershort)
                                               //ErrorCodes that require user input but have been 'okayed' and validated: (shortok, bufferok, buffershortok)
                                               if((lcl_errorcode == 'short') || (lcl_errorcode == 'buffer') || (lcl_errorcode == 'buffershort') || (lcl_errorcode == 'shortok') || (lcl_errorcode == 'bufferok') || (lcl_errorcode == 'buffershortok')) {
                                                  lcl_reservationerrors = lcl_reservationerrors + '&nbsp;Okay to continue?&nbsp;';

                                                  if((lcl_errorcode == 'shortok') || (lcl_errorcode == 'bufferok') || (lcl_errorcode == 'buffershortok')) {
                                                     lcl_reservationerrors = lcl_reservationerrors + '<span style=\"color:#ff0000;\">Yes</span>';
                                                     lcl_reservationerrors = lcl_reservationerrors + '<input type=\"hidden\" name=\"okToContinue_' + lcl_rentalid + '_' + lcl_linenum + '\" id=\"okToContinue_' + lcl_rentalid + '_' + lcl_linenum + '\" value=\"Y\" size=\"3\" maxlength=\"10\" />';

                                                     lcl_increaseErrorCount = 0;

                                                  } else {
                                                     lcl_reservationerrors = lcl_reservationerrors + '<select name=\"okToContinue_' + lcl_rentalid + '_' + lcl_linenum + '\" id=\"okToContinue_' + lcl_rentalid + '_' + lcl_linenum + '\" onchange=\"setupOkToContinue(\'' + lcl_rentalid + '\',\'' + lcl_linenum + '\');\">';
                                                     lcl_reservationerrors = lcl_reservationerrors +   '<option value=\"N\">No</option>';
                                                     lcl_reservationerrors = lcl_reservationerrors +   '<option value=\"Y\">Yes</option>';
                                                     lcl_reservationerrors = lcl_reservationerrors + '</select>';

                                                     jQuery('#errorCheck_okToContinue_' + lcl_rentalid + '_' + lcl_linenum).val('');
                                                  }

                                                  //Update the errorCheck_reservationDateTime with the current field values so that
                                                  //this field can be checked when the reservation is re-validated that the user has selected
                                                  //if it is "okay to continue" AND to ensure that none of the field values have changed
                                                  //since the last time the reservation was validated.
                                                  lcl_startdate   = jQuery('#startdate_'   + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_starthour   = jQuery('#starthour_'   + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_startminute = jQuery('#startminute_' + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_startampm   = jQuery('#startampm_'   + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_endhour     = jQuery('#endhour_'     + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_endminute   = jQuery('#endminute_'   + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_endampm     = jQuery('#endampm_'     + lcl_rentalid + '_' + lcl_linenum).val();
                                                  lcl_endday      = jQuery('#endday_'      + lcl_rentalid + '_' + lcl_linenum).val();

                                                  lcl_errorCheck_reservationDateTime  = lcl_startdate;
                                                  lcl_errorCheck_reservationDateTime += '_';
                                                  lcl_errorCheck_reservationDateTime += lcl_starthour + ':' + lcl_startminute + '_' + lcl_startampm;
                                                  lcl_errorCheck_reservationDateTime += '_';
                                                  lcl_errorCheck_reservationDateTime += lcl_endhour + ':' + lcl_endminute + '_' + lcl_endampm;
                                                  lcl_errorCheck_reservationDateTime += '_';
                                                  lcl_errorCheck_reservationDateTime += lcl_endday;

                                                  jQuery('#errorCheck_reservationDateTime_' + lcl_rentalid + '_' + lcl_linenum).val(lcl_errorCheck_reservationDateTime);
                                               }

                                               lcl_reservationerrors = lcl_reservationerrors + '</div>';

                                               jQuery('#div_reservationerrors_' + lcl_rentalid).html(lcl_reservationerrors);
                                               //END: Show the error message(s)

                                               //Update the "total errors" count for the rental.  This count includes:
                                               //  1. Any non-user interaction needed alerts
                                               //  2. User Interaction Needed alerts that have NOT been "okayed to continue"
                                               //
                                               //  NOTE: We check this count in the final validation (jQuery.ajaxStop) routine
                                               //  to determine if we can submit the form or not.
                                               if(jQuery('#totalErrors_rentals').val() != '') {
                                                  lcl_totalErrors_rentals = Number(jQuery('#totalErrors_rentals').val());
                                               }

                                               lcl_totalErrors_rentals = lcl_totalErrors_rentals + lcl_increaseErrorCount;
                                               jQuery('#totalErrors_rentals').val(lcl_totalErrors_rentals);
                                            }
                                         }

                                      });
                                   }
                                }
                             }
                          }
                       }
                       //END: Show validation error messages or continue processing

                    }
                    //END: Hide sections without errors --------------------------
                 }

           			}
              //END: Loop through each rental section and validate the fields ----
           }
        }
        //END: Expand all Rental reservation sections ----------------------------

     }
		}

  function getErrorText(iErrorCode) {
     var lcl_return = '';

     if(iErrorCode != '') {

        if(iErrorCode == 'shortnoconfirm') {
           lcl_return = 'Warning: The duration is for less than the allowed minimum time.';
        } else if(iErrorCode == 'buffernoconfirm') {
           lcl_return = 'Warning: There is a conflict with the buffering between reservations.';
        } else if(iErrorCode == 'buffershortnoconfirm') {
           lcl_return = 'Warning: The duration is less than allowed and there is a conflict with the buffering.';
        } else if(iErrorCode == 'conflict') {
           lcl_return = 'There is a conflict with an existing reservation.';
        } else if(iErrorCode == 'closed') {
           lcl_return = 'The rental is not open, or the time requested is beyond operating hours.';
        } else if(iErrorCode == 'nouser') {
           lcl_return = 'This type of reservation requires the selection of a person to complete.';
        } else if(iErrorCode == 'OK') {
           lcl_return = 'The selected time checks out fine for this reservation.';
        //User Internaction Needed and Confirmed errorCodes (i.e. Okay to Continue)
        } else if((iErrorCode == 'short') || (iErrorCode == 'shortok')) {
           lcl_return = 'You have selected a time interval that is less than the allowed minimum.';
        } else if((iErrorCode == 'buffer') || (iErrorCode == 'bufferok')) {
           lcl_return = 'There is a conflict with the buffering between reservations.';
        } else if((iErrorCode == 'buffershort') || (iErrorCode == 'buffershortok')) {
           lcl_return = 'The duration is less than allowed and there is a conflict with the buffering.';
        }
     }

     return lcl_return;
  }

  function validateFields(iMode, iRentalID, iTotalDates) {
     var lcl_false_count = Number(0);

     // Loop through all of the date/times for each rental
     for (var j = parseInt(iTotalDates); j >= 1 ; j--) {

        var lcl_startdate_exists = false;

        //No validation needed if the date/time row has NOT been included (is unchecked)
        var lcl_includereservationtime = document.getElementById('includereservationtime_' + iRentalID + '_' + j);

        if(lcl_includereservationtime.checked) {
           // See if a row exists for this one
          	if (jQuery('#startdate_' + iRentalID + '_' + j)) {

              //Start Date
              if (jQuery('#startdate_' + iRentalID + '_' + j).val() == '') {

                 if(iMode != 'HideErrors') {
                    jQuery('#startdate_' + iRentalID + '_' + j).focus();
                    inlineMsg(document.getElementById('startdate_' + iRentalID + '_' + j).id,'<strong>Required Field Missing: </strong> Date.',10,'startdate_' + iRentalID + '_' + j);
                 }

                 lcl_false_count = lcl_false_count + 1;
              } else {
                 lcl_startdate_exists = true
              }

              if(lcl_startdate_exists) {
                 var lcl_startdate = '';
                 var lcl_enddate   = '';
                 var lcl_endday    = jQuery('#endday_' + iRentalID + '_' + j).val();

                 lcl_startdate  = jQuery('#startdate_'   + iRentalID + '_' + j).val();
                 lcl_startdate += ' ';
                 lcl_startdate += jQuery('#starthour_'   + iRentalID + '_' + j).val();
                 lcl_startdate += ':';
                 lcl_startdate += jQuery('#startminute_' + iRentalID + '_' + j).val();
                 lcl_startdate += ' ';
                 lcl_startdate += jQuery('#startampm_'   + iRentalID + '_' + j).val();

               		lcl_enddate  = jQuery('#startdate_' + iRentalID + '_' + j).val();
                 lcl_enddate += ' ';
                 lcl_enddate += jQuery('#endhour_'   + iRentalID + '_' + j).val();
                 lcl_enddate += ':';
                 lcl_enddate += jQuery('#endminute_' + iRentalID + '_' + j).val();
                 lcl_enddate += ' ';
                 lcl_enddate += jQuery('#endampm_'   + iRentalID + '_' + j).val();

                 var dtStart = new Date(lcl_startdate);
                	var dtEnd   = new Date(lcl_enddate);

                	if (lcl_endday == '1') {
                  		dtEnd.setDate(dtEnd.getDate()+1);
                 }

                	var difference_in_milliseconds = dtEnd - dtStart;

                	if (difference_in_milliseconds <= 0) {
                    if(iMode != 'HideErrors') {
                      	//alert("One of the end times is not after the start time. Please correct this and try again.");
                       jQuery('#endday_' + iRentalID + '_' + j).focus();
                       inlineMsg(document.getElementById('endday_' + iRentalID + '_' + j).id,'<strong>Invalid Value: </strong> The end time is not after the start time.',10,'endday_' + iRentalID + '_' + j);
                    }

                    lcl_false_count = lcl_false_count + 1;
             				}
              }
           }
        }
     }

     //Check for any errors (false_counts) and update the field for the section with the total errors.
     jQuery('#totalErrors_fieldValidation' + iRentalID).val(lcl_false_count);

  }

		function doNameSearchChange()
		{
			if ($("searchname").value != "")
			{
				document.frmSearchReturn.searchname.value = $("searchname").value;
				document.frmSearchReturn.searchname2.value = $("searchname2").value;
				if ($("rentaluserid").value != '')
				{
					document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
				}
				document.frmSearchReturn.reservationtypeid.value = $("reservationtypeid").value;
				doUserPickChange();
			}
			else
			{
				alert('Please enter a name before searching.');
				$("searchname").focus();
			}
		}

		function doUserPickChange()
		{
			var iReservationTypeId = $("reservationtypeid").value;
			document.frmSearchReturn.reservationtypeid.value = $("reservationtypeid").value;
			if ($("searchname") != '' || $("rentaluserid") != '0')
			{
				// Fire off job to get the reservationtype type 
				doAjax('getreservationtype.asp', 'reservationtypeid=' + iReservationTypeId , 'changePickers', 'get', '0');
			}
		}

		function changePickers( sReturn )
		{
			//alert( sReturn);
			
			if (sReturn == 'public')
			{
				if ($("searchname").value != "")
				{
					// Try to get a drop down of citizen names
					doAjax('getcitizenpicks.asp', 'searchname=' + $("searchname").value + '&searchname2=' + $("searchname2").value, 'UpdateApplicants', 'get', '0');
				}
				else
				{
					$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />Search for a name then select one from the resulting list.";
					$("edituserbtn").style.visibility = 'hidden';
				}
			}
			else
			{
				if (sReturn == 'admin')
				{
					if ($("searchname").value != "")
					{
						// Try to get a drop down of citizen names
						doAjax('getadminpicks.asp', 'searchname=' + $("searchname").value, 'UpdateAdminApplicants', 'get', '0');
					}
					else
					{
						$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />Search for a name then select one from the resulting list.";
						$("edituserbtn").style.visibility = 'hidden';
					}
				}
				else
				{
					// for anything else, blank out the picks and put things back to nothing 
					$("applicant").innerHTML = "<input type='hidden' name='rentaluserid' id='rentaluserid' value='0' />This reservation type does not need a renter.";
					$("edituserbtn").style.visibility = 'hidden';
				}
			}
		}

		function UpdateAdminApplicants( sResult )
		{
			//alert(sResult);
			$("applicant").innerHTML = sResult;
			$("edituserbtn").style.visibility = 'hidden';
			document.frmSearchReturn.rentaluserid.value = $("rentaluserid").value;
		}

  function showDatesTimes(iRentalID) {
    var lcl_datesTimesLoaded = 'N';
    var lcl_rentalID         = iRentalID;

    //Determine if the date/times section has already been loaded or if it needs to be built
    if(jQuery('#dateTimesLoaded' + lcl_rentalID).val() != '') {
       lcl_datesTimesLoaded = jQuery('#dateTimesLoaded' + lcl_rentalID).val();
    }

    if(lcl_datesTimesLoaded == 'Y') {
       jQuery('#checkDateTimes' + lcl_rentalID).prop('disabled',true);
       jQuery('#hideDateTimes'  + lcl_rentalID).prop('disabled',false);
       jQuery('#div_reservationdates_' + lcl_rentalID).slideDown('slow',function() {
          enableButtons();
       });
    } else {
       jQuery('#checkDateTimes' + lcl_rentalID).val('View Date/Times');
       jQuery('#checkDateTimes' + lcl_rentalID).prop('disabled',true);
       jQuery('#hideDateTimes'  + lcl_rentalID).prop('disabled',false);
       jQuery('#div_reservationdates_' + lcl_rentalID).slideDown('slow');
       jQuery('#div_reservationdates_' + lcl_rentalID).html('<span style="color:#800000">Processing...</span>');
       jQuery.post('showRentalAvailibilityDetails.asp', {
          orgid:              '<%=session("orgid")%>',
          reservationtempid:  '<%=iReservationTempID%>',
          periodtypeselector: '<%=sPeriodTypeSelector%>',
          rentalid:           lcl_rentalID
       }, function(result) {
          jQuery('#div_reservationdates_' + lcl_rentalID).html(result);
          jQuery('#dateTimesLoaded' + lcl_rentalID).val('Y');
             enableButtons();
       });
    }
  }

  function hideDatesTimes(iMode, iRentalID) {
    var lcl_mode     = iMode;
    var lcl_rentalID = iRentalID;
    var lcl_maxrows  = jQuery('#maxrows' + lcl_rentalID).val();

    jQuery('#hideDateTimes' + lcl_rentalID).prop('disabled',true);

    //Clear any/all error messages for the section.
    for (var m = 1; m <= parseInt(lcl_maxrows); m++) {
       clearMsg('endday_' + lcl_rentalID + '_' + m);
    }

    if(iMode == 'CSS') {
       jQuery('#div_reservationdates_' + lcl_rentalID).css('display','none');
       jQuery('#checkDateTimes'        + lcl_rentalID).prop('disabled',false);
    } else {
       jQuery('#div_reservationdates_' + lcl_rentalID).slideUp('slow', function() {
         jQuery('#checkDateTimes' + lcl_rentalID).prop('disabled',false);
       });
    }
  }

  //Using jQuery "ready" function to set up the page instead of the BODY.onload
  jQuery(document).ready(function(){

     //Disable all "Hide Date/Times" button(s).
     jQuery('input[name^="hideDateTimes"]').each(function(index) {
        jQuery(this).prop('disabled',true);
     });

     jQuery('fieldset[id^="fieldset_reservationerrors_"]').each(function(index) {
        jQuery(this).css('display','none');
     });

     jQuery('.reservationErrorsDiv').each(function(index) {
        jQuery(this).css('display','none');
     });

     //Check to see if there is at least one rental date/time section.  If "yes" then expand it.
     var lcl_firstRentalID = '';

     if(jQuery('#firstRentalID').val != '') {
        lcl_firstRentalID = jQuery('#firstRentalID').val();
        showDatesTimes(lcl_firstRentalID);
     }

     jQuery('#checkbutton').prop('disabled',true);
     jQuery('#continuebutton').prop('disabled',true);

     //This function executes after any/all jQuery functions are finished running.
     //We are just wanting it to run for us when the user is attempting to reserve the date/time(s).
     jQuery('#totalErrors_rentals').ajaxStop(function(){
        var lcl_isFinalValidate              = jQuery('#checkReserveValidate').val();
        var lcl_totalCount_checkDatesReserve = Number(0);
        var lcl_false_count                  = Number(0);
        var lcl_maxrows                      = Number(0);

        if(lcl_isFinalValidate == 'Y') {
           if(jQuery('#totalCount_checkDatesReserve').val() != '') {
              lcl_totalCount_checkDatesReserve = Number(jQuery('#totalCount_checkDatesReserve').val());
           }

           if(lcl_totalCount_checkDatesReserve > 0) {
              var lcl_totalErrors_rentals = Number(jQuery('#totalErrors_rentals').val());

              if(lcl_totalErrors_rentals > 0) {
                 lcl_false_count = lcl_false_count + 1;
              }
           }

           if(lcl_false_count > 0) {
              return false;
           } else {
              jQuery('#frmDateSelection').submit();
           }
        }
     });
  });

  function enableButtons() {
    //Determine if all sections have been viewed.
    //If all sections have been "viewed" then enable the "check dates" and "check and reserve" buttons
    var lcl_false_count    = Number(0);
    var lcl_dateTimeLoaded = '';

    jQuery('input[name^="dateTimesLoaded"]').each(function(index) {
       lcl_dateTimeLoaded = jQuery(this).val();

       if(lcl_dateTimeLoaded != 'Y') {
          lcl_false_count = lcl_false_count + 1;
       }
    });

    if(lcl_false_count == 0) {
       jQuery('#continueMsg').hide('slow',function() {
          jQuery('#checkbutton').prop('disabled',false);
          jQuery('#continuebutton').prop('disabled',false);
       });
    }
  }

  function enableDisableFields(iRentalRowID) {
     //The format for iRentalRowID: rentalid + '_' + linenum
     var lcl_includereservationtime = document.getElementById('includereservationtime_' + iRentalRowID);
     var lcl_isDisabled             = false;
     var lcl_totalErrors_rentals    = Number(jQuery('#totalErrors_rentals').val());

     if(! lcl_includereservationtime.checked) {
        lcl_isDisabled = true;
     }

     //jQuery('#startdate_' + iRentalRowID).prop('disabled',true);
     //jQuery('#startdate_popup_' + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#starthour_'   + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#startminute_' + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#startampm_'   + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#endhour_'     + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#endminute_'   + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#endampm_'     + iRentalRowID).prop('disabled',lcl_isDisabled);
     jQuery('#endday_'      + iRentalRowID).prop('disabled',lcl_isDisabled);

     //If "disabled" then reduce the total errors.
     if(lcl_isDisabled) {
        lcl_totalErrors_rentals = lcl_totalErrors_rentals - 1;

        if(lcl_totalErrors_rentals < 0) {
           lcl_totalErrors_rentals = 0;
        }

        jQuery('#totalErrors_rentals').val(lcl_totalErrors_rentals);
        jQuery('#errorCheck_reservationDateTime_' + iRentalRowID).val('');
        jQuery('#errorCheck_errorCode_'           + iRentalRowID).val('');
        jQuery('#errorCheck_okToContinue_'        + iRentalRowID).val('');
     }

     //Check to see if an error exists for this row.  If "yes" and this field has been unchecked to be unincluded then
     //hide the "error" message.
     if(jQuery('#okToContinue_' + iRentalRowID)) {
        if(lcl_isDisabled) {
           clearMsg('okToContinue_' + iRentalRowID);
           jQuery('#reservationError' + iRentalRowID).slideUp('slow',function() {
              jQuery('#reservationError' + iRentalRowID).html('');

              var lcl_rentalid            = Number(0);
              var lcl_underscore_position = Number(0);

              if(iRentalRowID.indexOf('_') > -1) {
                 lcl_underscore_position = iRentalRowID.indexOf('_');
              }

              if(lcl_underscore_position > 0) {
                 lcl_rentalid = iRentalRowID.substr(0,lcl_underscore_position);
              }

              //If this error message is hidden we may have some maintenance to do to the
              //Alerts/Warnings area for the rental.
              //1st: Make sure that the field is not blank.  If "no" then pull the total errors for the rental
              //2nd: If the errors are GREATER THAN zero then we want to subtract one from the count.
              //3rd: If the error total is set to zero then we want to hide the entire Alerts/Warnings section for the rental
              var lcl_totalErrors_alertsWarnings = Number(0);

              if(jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val() != '' && jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val() != undefined) {
                 lcl_totalErrors_alertsWarnings = Number(jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val());
              }
              if(lcl_totalErrors_alertsWarnings > 0) {
                 lcl_totalErrors_alertsWarnings = lcl_totalErrors_alertsWarnings - 1;

                 jQuery('#totalErrors_alertsWarnings' + lcl_rentalid).val(lcl_totalErrors_alertsWarnings);
              }

              if(lcl_totalErrors_alertsWarnings == 0) {
                 jQuery('#fieldset_reservationerrors_' + lcl_rentalid).slideUp('slow');
              }

           });

        }
     }
  }

  function setupOkToContinue(iRentalID, iLineNumber) {
     clearMsg('okToContinue_' + iRentalID + '_' + iLineNumber);
     jQuery('#errorCheck_okToContinue_' + iRentalID + '_' + iLineNumber).val(jQuery('#okToContinue_' + iRentalID + '_' + iLineNumber).val());
  }

		function displayScreenMsg(iMsg) 
		{
			if(iMsg!="") 
			{
				$("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
				window.setTimeout("clearScreenMsg()", (10 * 1000));
			}
		}

		function clearScreenMsg() {
 			$("screenMsg").innerHTML = "&nbsp;";
		}

		function doCalendar( sField ) {
 			var w = (screen.width - 350)/2;
	 		var h = (screen.height - 350)/2;
		 	var sSelectedDate = $(sField).value;

			 eval('window.open("calendarpicker.asp?date=' + sSelectedDate + '&p=1&updatefield=' + sField + '&updateform=frmDateSelection", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}
		
		function goBack() {
		  document.frmSearchReturn.submit();
		}

  function goBackSimple() {
     var lcl_url = 'rentalavailability.asp?rti=<%=iReservationTempId%>';

     location.href = lcl_url;
  }

		function loader() {
  		<%=sLoadMsg%>
		}
	//-->
	</script>

</head>

<body onload="loader();">

 <% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<div id=""content"">" & vbcrlf
  response.write "  <div id=""centercontent"">" & vbcrlf
  response.write "<p><font size=""+1""><strong>Rental Date Selection</strong></font></p>" & vbcrlf
  response.write "<p><span id=""screenMsg"">&nbsp;</span></p>" & vbcrlf
  response.write "<form name=""frmDateSelection"" id=""frmDateSelection"" method=""post"" action=""rentalreservationmake.asp"">" & vbcrlf
  'response.write "  <input type=""hidden"" name=""rentalid"" id=""rentalid"" value=""" & iRentalId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""selected_rentalids"" id=""selected_rentalids"" value=""" & lcl_selected_rentalids & """ size=""5"" maxlength=""500"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rti"" id=""rti"" value=""" & iReservationTempId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rid"" id=""rid"" value=""" & iReservationId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""createpath"" id=""createpath"" value=""" & lcl_createpath & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""firstRentalID"" id=""firstRentalID"" value="""" size=""3"" maxlength=""50"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""falseCountTotal"" id=""falseCountTotal"" value="""" size=""3"" maxlength=""50"" />" & vbcrlf
  response.write "  <!--<br />Line Count: --><input type=""hidden"" name=""lineCount"" id=""lineCount"" value=""0"" size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <!--<br />Total Errors: --><input type=""hidden"" name=""totalErrors_rentals"" id=""totalErrors_rentals"" value=""0"" size=""3"" maxlength=""10"" />" & vbcrlf
  response.write "  <!--<br />Final Validation: --><input type=""hidden"" name=""checkReserveValidate"" id=""checkReserveValidate"" value=""N"" size=""1"" maxlength=""1"" />" & vbcrlf
  response.write "  <!--<br />Checked # of times: --><input type=""hidden"" name=""totalCount_checkDatesReserve"" id=""totalCount_checkDatesReserve"" value=""0"" size=""1"" maxlength=""10"" />" & vbcrlf
  response.write "<table id=""reservationtempinfo"" cellpadding=""0"" cellspacing=""1"" border=""0"">" & vbcrlf

  if iReservationId = CLng(0) then
     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""labelcolumn""><strong>Reservation Type:</strong></td>" & vbcrlf
     response.write "      <td class=""pickcolumn"" align=""left"">" & vbcrlf
                               ShowRentalReservationTypes iReservationTypeId
     response.write "      </td>" & vbcrlf
     response.write "      <td class=""labelcolumn2"">&nbsp;</td>" & vbcrlf
     response.write "      <td>&nbsp;</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "  				<td class=""labelcolumn"">" & vbcrlf
     response.write "          <strong>Name Is Like:</strong>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td colspan=""3"">" & vbcrlf
     response.write "          <input type=""text"" name=""searchname"" id=""searchname"" value=""" & sSearchName & """ size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange();return false;}"" />" & vbcrlf
     response.write "          <input type=""text"" name=""searchname2"" id=""searchname2"" value=""" & sSearchName2 & """ size=""25"" maxlength=""25"" onkeypress=""if(event.keyCode=='13'){doNameSearchChange();return false;}"" />" & vbcrlf
     response.write "     					<input type=""button"" class=""button"" value=""Search for a Name"" onclick=""doNameSearchChange();"" />" & vbcrlf
     response.write "          <input type=""button"" class=""button"" value=""New Public User"" onclick=""newUser();"" />" & vbcrlf
     response.write "  				</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td colspan=""4"">" & vbcrlf
     response.write "  				    <span id=""applicant"">" & vbcrlf

     if bHasUsers then
        if iRentalUserid > CLng(0) then
           if bIsPublicUsers then
 												'Show registered user picks
												  ShowCitizenPicks iRentalUserid, sSearchName
											else
												 'Show admin picks
												  ShowAdminPicks iRentalUserid, sSearchName
											end if
        else
           response.write "          <input type=""hidden"" value=""0"" name=""rentaluserid"" id=""rentaluserid"" />Search for a name then select one from the resulting list." & vbcrlf
        end if
     else
        response.write "          <input type=""hidden"" value=""0"" name=""rentaluserid"" id=""rentaluserid"" />This reservation type does not need a renter." & vbcrlf
     end if

     response.write "</span> <input type=""button"" class=""button"" id=""edituserbtn"" value=""Edit User"" onclick=""EditApplicant();"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf

  else  'Adding to an Existing Reservation		
					sReservationTypeId = GetReservationTypeId( iReservationId )    'in rentalscommonfunctions.asp
     sReservationType   = GetReservationType( sReservationTypeId )  'in rentalscommonfunctions.asp
     sRenterName        = GetRenterName( iReservationId )

     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""labelcolumn"">" & vbcrlf
     response.write "          <strong>Reservation Type:</strong>" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td class=""pickcolumn"" align=""left"" colspan=""3"">" & vbcrlf
					response.write            sReservationType & vbcrlf
     response.write "          <input type=""hidden"" name=""reservationtypeid"" id=""reservationtypeid"" value=""" & sReservationTypeId & """ />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""labelcolumn"">" & vbcrlf
     response.write "          <strong>Renter Is:</strong>" & vbcrlf
     response.write "          <input type=""hidden"" name=""searchname"" id=""searchname"" value="""" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "      <td class=""pickcolumn"" align=""left"" colspan=""3"">" & vbcrlf
     response.write            sRenterName & vbcrlf
     response.write "          <input type=""hidden"" value=""" & GetReservationRentalUserId( iReservationId ) & """ name=""rentaluserid"" id=""rentaluserid"" />" & vbcrlf
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
		end if

  if lcl_createpath <> "SIMPLE" then
     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""tempinfolabel""><strong>Start Date:</strong></td>" & vbcrlf
     response.write "      <td class=""datacolumn"">" & sStartDate & "</td>" & vbcrlf
     response.write "      <td class=""tempinfolabel2"" align=""center""><strong>End Date:</strong></td>" & vbcrlf
     response.write "      <td>" & sEndDate & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf

     if sPeriodTypeSelector = "selectedperiod" then
        response.write "  <tr>" & vbcrlf
        response.write "      <td class=""tempinfolabel""><strong>Start Time:</strong></td>" & vbcrlf
        response.write "      <td class=""datacolumn"">" & sStartTime & "</td>" & vbcrlf
        response.write "      <td class=""tempinfolabel2"" align=""center""><strong>End Time:</strong></td>" & vbcrlf
        response.write "      <td>" & sEndTime & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     else
        response.write "  <tr>" & vbcrlf
        response.write "      <td class=""tempinfolabel""><strong>Time Period:</strong></td>" & vbcrlf
        response.write "      <td colspan=""3"">" & sPeriodType & "</td>" & vbcrlf
        response.write "  </tr>" & vbcrlf
     end if

     response.write "  <tr>" & vbcrlf
     response.write "      <td class=""tempinfolabel""><strong>Occurs:</strong></td>" & vbcrlf
     response.write "      <td colspan=""3"">" & sOccurs & "</td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if

  response.write "</table>" & vbcrlf
  response.write "<p>" & vbcrlf
  response.write "  <input type=""button"" class=""button"" id=""back"" name=""back"" value=""<< Back"" onclick=""" & lcl_goBackOnClick & """ /> &nbsp;" & vbcrlf
  'response.write "  <input type=""button"" class=""button"" id=""back"" name=""back"" value=""<< Back"" onclick=""goBack();"" /> &nbsp;" & vbcrlf
  'response.write "  <input type=""button"" class=""button"" id=""adddate"" name=""adddate"" value=""Add A Date"" onclick=""AddDateRow();"" /> &nbsp;" & vbcrlf
  'response.write "  <input type=""button"" class=""button"" id=""removedates"" name=""removedates"" value=""Remove Selected Dates"" onclick=""RemoveDateRow();"" />" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "<p>" & vbcrlf
  'response.write "<table id=""reservationtempdates"" cellpadding=""0"" cellspacing=""0"" border=""0"">" & vbcrlf
  'response.write "  <tr><th class=""firstcell"">Date</th><th>Start Time</th><th>End Time</th><th class=""lastcell"">Available</th></tr>" & vbcrlf

 'BEGIN: Cycle through all of the rentals that have been selected -------------
 	sSQLr = "SELECT DISTINCT R.rentalid "
  sSQLr = sSQLr & " FROM egov_rentals R "
 	sSQLr = sSQLr & " WHERE R.orgid = " & session("orgid")
  sSQLr = sSQLr & " AND R.rentalid IN (" & lcl_selected_rentalids & ") "

 	set oGetRentalIDs = Server.CreateObject("ADODB.Recordset")
 	oGetRentalIDs.Open sSQLr, Application("DSN"), 3, 1

 	if not oGetRentalIDs.eof then
     do while not oGetRentalIDs.eof
        lcl_rental_linecount = lcl_rental_linecount + 1

        if lcl_rental_linecount = 1 then
           lcl_scripts = lcl_scripts & "document.getElementById('firstRentalID').value='" & oGetRentalIDs("rentalid") & "';" & vbcrlf
        end if

 						'Pull the wanted dates list here
    				ShowRentalAvailabilityDetails iReservationTempId, _
                                      oGetRentalIDs("rentalid"), _
                                      sPeriodTypeSelector, _
                                      lcl_rental_linecount

        oGetRentalIDs.movenext
     loop
	 end if

 	oGetRentalIDs.Close
	 set oGetRentalIDs = nothing 
 'END: Cycle through all of the rentals that have been selected ---------------

  response.write "</p>" & vbcrlf
  response.write "<p>" & vbcrlf
  'response.write "  <input type=""button"" class=""button"" id=""checkbutton"" name=""checkbutton"" value=""Check Dates"" onclick=""CheckDates()"" />&nbsp;" & vbcrlf
  response.write "		<input type=""button"" class=""button"" id=""checkbutton"" name=""checkbutton"" value=""Check Dates"" onclick=""validate('CHECK_DATES')"" />&nbsp;" & vbcrlf
  response.write "		<input type=""button"" class=""button"" id=""continuebutton"" name=""continuebutton"" value=""Check and Reserve"" onclick=""WaitAndValidate('')"" />" & vbcrlf
  response.write "  <input type=""hidden"" name=""total_rentals"" id=""total_rentals"" value=""" & lcl_rental_linecount & """ size=""3"" maxlength=""5"" />" & vbcrlf
  response.write "  <div id=""continueMsg"">* Please click the ""View and Confirm Dates/Times"" button for each section to continue.</div>" & vbcrlf
  response.write "</p>" & vbcrlf
  response.write "</form>" & vbcrlf
  response.write "<form name=""frmSearchReturn"" id=""frmSearchReturn"" method=""post"" action=""rentalsearch.asp"">" & vbcrlf
  'response.write "<input type=""hidden"" name=""rentalid"" value=""" & iRentalId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rti"" id=""search_rti"" value="""                                   & iReservationTempId    & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""reservationtypeid"" id=""search_reservationtypeid"" value="""       & sReservationTypeId    & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""searchname"" id=""search_searchname"" value="""                     & sSearchName           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""searchname2"" id=""search_searchname2"" value="""                     & sSearchName           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""recreationcategoryid"" id=""search_recreationcategoryid"" value=""" & iRecreationCategoryId & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""locationid"" id=""search_locationid"" value="""                     & iLocationId           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rentalname"" id=""search_rentalname"" value="""                     & sRentalName           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""startdate"" id=""search_startdate"" value="""                       & sStartDate            & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""enddate"" id=""search_enddate"" value="""                           & sEndDate              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rentaluserid"" id=""search_rentaluserid"" value="""                 & iRentalUserid         & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""periodtypeid"" id=""search_periodtypeid"" value="""                 & iPeriodTypeId         & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""starthour"" id=""search_starthour"" value="""                       & iStartHour            & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""startminute"" id=""search_startminute"" value="""                   & iStartMinute          & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""startampm"" id=""search_startampm"" value="""                       & sStartAmPm            & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""endhour"" id=""search_endhour"" value="""                           & iEndHour              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""endminute"" id=""search_endminute"" value="""                       & iEndMinute            & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""endampm"" id=""search_endampm"" value="""                           & sEndAmPm              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""endday"" id=""search_endday"" value="""                             & iEndDay               & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""occurs"" id=""search_occurs"" value="""                             & sOccursChecked        & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""monthlyperiodid"" id=""search_monthlyperiodid"" value="""           & iMonthlyPeriod        & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""monthlydow"" id=""search_monthlydow"" value="""                     & iMonthlyDOW           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""orderby"" id=""search_orderby"" value="""                           & iOrderBy              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""weeklydays"" id=""search_weeklydays"" value="""                     & sWantedDOWs           & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""msg"" id=""search_msg"" value="""                                   & sMessage              & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""rid"" id=""search_rid"" value="""                                   & iReservationId        & """ />" & vbcrlf
  response.write "  <input type=""hidden"" name=""createpath"" id=""createpath"" value="""                            & lcl_createpath        & """ />" & vbcrlf
  response.write "</form>" & vbcrlf
  response.write "  </div>" & vbcrlf
  response.write "</div>" & vbcrlf

  if lcl_scripts <> "" then
     response.write "<script language=""javascript"">" & vbcrlf
     response.write lcl_scripts
     response.write "</script>" & vbcrlf
  end if
%>
	<!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

'------------------------------------------------------------------------------
Sub GetTempRentalValues( ByVal iReservationTempId )
	Dim sSQL, oRs

	sSQL = "SELECT requestedstartdate, "
 sSQL = sSQL & " requestedstarthour, "
 sSQL = sSQL & " dbo.AddLeadingZeros(isnull(requestedstartminute,0),2) AS requestedstartminute, "
	sSQL = sSQL & " requestedstartampm, "
 sSQL = sSQL & " requestedenddate, "
 sSQL = sSQL & " requestedendhour, "
	sSQL = sSQL & " dbo.AddLeadingZeros(isnull(requestedendminute,0),2) AS requestedendminute, "
 sSQL = sSQL & " requestedendampm, "
 sSQL = sSQL & " requestedendday, "
	sSQL = sSQL & " occurs, "
 sSQL = sSQL & " weeklydays, "
 sSQL = sSQL & " rentalmonthlyperiodid, "
 sSQL = sSQL & " monthlydow, "
 sSQL = sSQL & " P.periodtype, "
 sSQL = sSQL & " P.periodtypeselector, "
	sSQL = sSQL & " R.userlike, "
 sSQL = sSQL & " R.recreationcategoryid, "
 sSQL = sSQL & " R.locationid, "
 sSQL = sSQL & " R.rentallike, "
 sSQL = sSQL & " R.periodtypeid, "
 sSQL = sSQL & " R.orderby, "
	sSQL = sSQL & " ISNULL(R.reservationid,0) AS reservationid "
	sSQL = sSQL & " FROM egov_rentalreservationstemp R, "
 sSQL = sSQL &      " egov_rentalperiodtypes P "
	sSQL = sSQL & " WHERE R.periodtypeid = P.periodtypeid "
 sSQL = sSQL & " AND reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSQL, Application("DSN"), 0, 1

	If Not oRs.EOF Then

		sPeriodType         = oRs("periodtype")
		sPeriodTypeSelector = oRs("periodtypeselector")
		sStartDate          = oRs("requestedstartdate")
		sEndDate            = oRs("requestedenddate")

		If sPeriodTypeSelector = "selectedperiod" Then 
			sStartTime = oRs("requestedstarthour") & ":" & oRs("requestedstartminute") & " " & oRs("requestedstartampm")
			sEndTime   = oRs("requestedendhour")   & ":" & oRs("requestedendminute")   & " " & oRs("requestedendampm")
		Else
			sStartTime = ""
			sEndTime   = ""
		End If 

		Select Case oRs("occurs")
			Case "o"
				sOccurs = "Just This Once"
			Case "d"
				sOccurs = "Daily"
			Case "w"
				x = 1
				Do While x < 8
					If InStr(oRs("weeklydays"),x) > 0 Then
						If sOccurs <> "" Then
							sOccurs = sOccurs & ", "
						End If 
						sOccurs = sOccurs & WeekDayName(x)
					End If
					x = x + 1
				Loop 
				sOccurs = "Weekly On These Days: " & sOccurs
			Case "m"
				sOccurs = "Monthly On The "
				Select Case oRs("rentalmonthlyperiodid")
					Case "1"
						sOccurs = sOccurs & "First "
					Case "2"
						sOccurs = sOccurs & "Second "
					Case "3"
						sOccurs = sOccurs & "Third "
					Case "4"
						sOccurs = sOccurs & "Fourth "
					Case "5"
						sOccurs = sOccurs & "Last "
				End Select 
				sOccurs = sOccurs & " " & WeekDayName(oRs("monthlydow"))
		End Select 

		iRecreationCategoryId = oRs("recreationcategoryid")
		iLocationId           = oRs("locationid")
		sRentalName           = oRs("rentallike")
		iPeriodTypeId         = oRs("periodtypeid")
		iStartHour            = oRs("requestedstarthour")
		iStartMinute          = CStr(clng(oRs("requestedstartminute")))
		sStartAmPm            = oRs("requestedstartampm")
		iEndHour              = oRs("requestedendhour")
		iEndMinute            = CStr(clng(oRs("requestedendminute")))
		sEndAmPm              = oRs("requestedendampm")
		iEndDay               = oRs("requestedendday")
		sOccursChecked        = oRs("occurs")
		iMonthlyPeriod        = oRs("rentalmonthlyperiodid")
		iMonthlyDOW           = oRs("monthlydow")
		iOrderBy              = oRs("orderby")
		sWantedDOWs           = oRs("weeklydays")
		iReservationId        = CLng(oRs("reservationid"))
	Else
		sStartDate     = ""
		sStartTime     = ""
		sEndDate       = ""
		sEndTime       = ""
		sOccurs        = ""
		sPeriodType    = ""
		iReservationid = CLng(0)
	End If
	
	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
sub ShowRentalAvailabilityDetails( ByVal iReservationTempId, ByVal iRentalId, ByVal sPeriodTypeSelector, ByVal iRentalLineCount )
	Dim sSQL, oRs, iRowCount, sAmPm, sReservationStartTime, sReservationEndTime, iEndDay, bOffSeasonFlag
	Dim bIsAllDayOnly, sDisabledOption, lcl_rental_namelocation, sRentalLineCount
	Dim aWantedDates(1,0)

	iRowCount               = 0
 sRentalLineCount        = iRentalLineCount
 lcl_rental_namelocation = getRentalNameAndLocation(iRentalId)

 response.write "<fieldset class=""fieldset"">" & vbcrlf
 response.write "  <legend class=""dateSelectionLegendText"">" & vbcrlf
 response.write      lcl_rental_namelocation & vbcrlf
 response.write "  </legend>" & vbcrlf
 response.write "  <fieldset id=""fieldset_reservationerrors_" & iRentalID & """ class=""fieldset"">" & vbcrlf
 response.write "     <legend class=""reservationErrorsLegend"">Alerts/Warnings</legend>" & vbcrlf
 response.write "     <div id=""div_reservationerrors_" & iRentalID & """ class=""reservationErrorsDiv""></div>" & vbcrlf
 response.write "  </fieldset>" & vbcrlf
 response.write "  <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""dateTimesTable"">" & vbcrlf
 response.write "    <tr valign=""top"">" & vbcrlf
 response.write "        <td>" & vbcrlf
 response.write "            <input type=""button"" name=""checkDateTimes"  & iRentalID & """ id=""checkDateTimes"  & iRentalID & """ class=""button"" value=""View and Confirm Dates/Times"" onclick=""showDatesTimes('" & iRentalID & "')"" />" & vbcrlf
 response.write "            <input type=""button"" name=""hideDateTimes"   & iRentalID & """ id=""hideDateTimes"   & iRentalID & """ class=""button"" value=""Hide Dates/Times"" onclick=""hideDatesTimes('JQUERY','" & iRentalID & "')"" />" & vbcrlf
 response.write "            <input type=""hidden"" name=""dateTimesLoaded" & iRentalID & """ id=""dateTimesLoaded" & iRentalID & """ value=""N"" size=""1"" maxlength=""5"" />" & vbcrlf
 response.write "            <input type=""hidden"" name=""rentalid" & sRentalLineCount & """ id=""rentalid" & sRentalLineCount & """ value=""" & iRentalID & """ size=""3"" maxlength=""5"" />" & vbcrlf
 response.write "        </td>" & vbcrlf
 response.write "        <td>" & vbcrlf

 if RentalHasNoCosts( iRentalId ) then
   	'sNoCostPhrase = "<strong>There is no cost to rent this.</strong>"
    response.write "<div class=""noCostPhrase"">* There is no cost to rent this. *</div>" & vbcrlf
 end if

 response.write "        </td>" & vbcrlf
 response.write "    </tr>" & vbcrlf
 response.write "  </table>" & vbcrlf
 response.write "  <div id=""div_reservationdates_" & iRentalID & """ class=""reservationDatesDiv""></div>" & vbcrlf
 response.write "</fieldset>" & vbcrlf
end sub

'------------------------------------------------------------------------------
Sub ShowRentalReservationTypes( ByVal iReservationTypeId )
	Dim oRs, sSql

	sSql = "SELECT reservationtypeid, reservationtype FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE displayindropdown = 1 AND orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	response.write "<select id=""reservationtypeid"" name=""reservationtypeid"" onchange=""doUserPickChange();"">"
	Do While Not oRs.EOF
		response.write vbcrlf & vbtab & "<option value=""" & oRs("reservationtypeid") & """ "
		If CLng(oRs("reservationtypeid")) = CLng(iReservationTypeId) Then 
			response.write " selected=""selected"" "
		End If 
		response.write ">" & oRs("reservationtype") & "</option>"
		oRs.MoveNext 
	Loop
	response.write vbcrlf & "</select>"
	
	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
Function GetFirstReservationTypeInList( )
	Dim oRs, sSql

	sSql = "SELECT reservationtypeid FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE displayindropdown = 1 AND orgid = " & session("orgid") & " ORDER BY displayorder"

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		GetFirstReservationTypeInList = oRs("reservationtypeid")
	Else
		GetFirstReservationTypeInList = 0
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function 

'------------------------------------------------------------------------------
Function IsAReservation( ByVal iReservationTypeId )
	Dim oRs, sSql

	sSql = "SELECT isreservation FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationtypeid = " & iReservationTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("isreservation") Then
			IsAReservation = True 
		Else
			IsAReservation = False 
		End If 
	Else
		IsAReservation = False 
	End If 

	oRs.Close
	Set oRs = Nothing
	
End Function 

'------------------------------------------------------------------------------
Sub GetUserFlagsFromReservationTypeId( ByVal iReservationTypeId, ByRef bHasUsers, ByRef bIsPublicUsers )
	Dim oRs, sSql

	sSql = "SELECT hasusers, haspublicusers FROM egov_rentalreservationtypes "
	sSql = sSql & "WHERE orgid = " & session("orgid") & " AND reservationtypeid = " & iReservationTypeId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If oRs("hasusers") Then
			bHasUsers = True 
		Else
			bHasUsers = False 
		End If 
		If oRs("haspublicusers") Then
			bIsPublicUsers = True 
		Else
			bIsPublicUsers = False 
		End If 
	Else
		bHasUsers = False 
		bIsPublicUsers = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Sub 

'------------------------------------------------------------------------------
Function TempReservationExists( ByVal iReservationTempId )
	Dim sSql, oRs

	sSql = "SELECT COUNT(reservationtempid) AS hits FROM egov_rentalreservationstemp "
	sSql = sSql & "WHERE reservationtempid = " & iReservationTempId

	Set oRs = Server.CreateObject("ADODB.Recordset")
	oRs.Open sSql, Application("DSN"), 0, 1

	If Not oRs.EOF Then
		If clng(oRs("hits")) > clng(0) Then
			TempReservationExists = True 
		Else
			TempReservationExists = False 
		End If 
	Else
		TempReservationExists = False 
	End If 

	oRs.Close
	Set oRs = Nothing 

End Function
%>
