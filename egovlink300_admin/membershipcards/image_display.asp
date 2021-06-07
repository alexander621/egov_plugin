<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="membership_card_functions.asp" //-->
<!-- #include file="../poolpass/poolpass_global_functions.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

Dim lcl_member_id, lcl_action

sLevel = "../" ' Override of value from common.asp
lcl_member_id  = request("memberid")
lcl_action     = request("action")
lcl_poolpassid = request("poolpassid")
lcl_rateid     = getRateID(lcl_poolpassid)

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo            = request("demo")
 lcl_demo_page_title = ""
 lcl_demo_url        = ""

'set up demo variables
 if lcl_demo = "Y" then
    lcl_demo_page_title = " (DEMO)"
    lcl_demo_url        = "&demo=" & lcl_demo
 end if

'Set up Session variable for DISPLAY include file.
 session("CARD_PRINT") = "N"
 session("MEMBERID")   = lcl_member_id
 session("poolpassid") = lcl_poolpassid

 sPrintCardURL = "image_display.asp"
 sPrintCardURL = sPrintCardURL & "?memberid=" & lcl_member_id
 sPrintCardURL = sPrintCardURL & "&poolpassid=" & lcl_poolpassid
 sPrintCardURL = sPrintCardURL & "&action=CARD_PRINTED"
 sPrintCardURL = sPrintCardURL & lcl_demo_url
 sPrintCardURL = sPrintCardURL & "&card_layout=p"

 if lcl_action = "CANCEL" then
    if lcl_demo <> "Y" then
       remove_image(lcl_member_id)
    end if

   	response.redirect(Session("RedirectPage"))

 elseif lcl_action = "REPRINT" then
    lcl_status = lcl_action
   	lcl_action = "CARD_PRINTED"

 elseif lcl_action = "SAVE" then
    if lcl_demo <> "Y" then
       save_card(lcl_member_id)
    end if

   	response.redirect(sPrintCardURL)

 elseif lcl_action = "PRINT_CARD" then
    if lcl_demo <> "Y" then
      	save_card(lcl_member_id)
      	print_card(lcl_member_id)
    end if

    response.redirect(sPrintCardURL)

 end if

'Check for org features.
lcl_orghasfeature_card_layout_multiplelayouts = orghasfeature("card_layout_multiplelayouts")
lcl_orghasfeature_memberships_usekeycards     = orghasfeature("memberships_usekeycards")
lcl_orghasfearure_w7_takepic                  = orghasfeature("create membership cards new")

Dim sTakePicUrl

'If lcl_orghasfearure_w7_takepic Then
  'sTakePicUrl = "image_takepic_new.asp"
'Else
  sTakePicUrl = "image_takepic.asp"
'End If 

'Retrieve the card layout attributes
 lcl_layout_maint = "N"
 lcl_title            = ""
 lcl_subtitle         = ""
 lcl_year_text        = ""
 lcl_display_date     = 0
 lcl_custom_image_url = ""
 lcl_quote            = ""
 lcl_color1           = "ffffff"
 lcl_color2           = "ffffff"
 lcl_text_color1      = "000000"
 lcl_text_color2      = "000000"
 lcl_back_text        = ""
 lcl_back_text_color  = "000000"

 sSQL = "SELECT m.cardid, "
 sSQL = sSQL & " m.title, "
 sSQL = sSQL & " m.subtitle, "
 sSQL = sSQL & " m.year_text, "
 sSQL = sSQL & " m.display_date, "
 sSQL = sSQL & " m.custom_image_url, "
 sSQL = sSQL & " m.quote_text, "
 sSQL = sSQL & " m.main_color, "
 sSQL = sSQL & " m.secondary_color, "
 sSQL = sSQL & " m.main_text_color, "
 sSQL = sSQL & " m.secondary_text_color, "
 sSQL = sSQL & " m.back_text, "
 sSQL = sSQL & " m.back_text_color "
 sSQL = sSQL & " FROM egov_membershipcard_layout m "

 if lcl_orghasfeature_card_layout_multiplelayouts then
    sSQL = sSQL & " LEFT OUTER JOIN egov_poolpassrates r ON m.cardid = r.cardid "
    sSQL = sSQL & " WHERE r.rateid = " & lcl_rateid
    sSQL = sSQL & " AND "
 else
    sSQL = sSQL & " WHERE "
 end if

 sSQL = sSQL & " m.orgid = " & session("orgid")

 set oCardLayout = Server.CreateObject("ADODB.Recordset")
 oCardLayout.Open sSQL, Application("DSN"), 3, 1

 if not oCardLayout.eof then
    lcl_title            = oCardLayout("title")
    lcl_subtitle         = oCardLayout("subtitle")
    lcl_year_text        = oCardLayout("year_text")
    lcl_display_date     = oCardLayout("display_date")
    lcl_custom_image_url = oCardLayout("custom_image_url")
    lcl_quote            = oCardLayout("quote_text")
    lcl_color1           = oCardLayout("main_color")
    lcl_color2           = oCardLayout("secondary_color")
    lcl_text_color1      = oCardLayout("main_text_color")
    lcl_text_color2      = oCardLayout("secondary_text_color")
    lcl_back_text        = oCardLayout("back_text")
    lcl_back_text_color  = oCardLayout("back_text_color")
end if

oCardLayout.close
set oCardLayout = nothing

 'Setup Buttons
  if lcl_action <> "CARD_PRINTED" then
     lcl_print_label    = "Print Card"
  	  lcl_print_msg      = "Prints, and saves, the membership card.  Please have your color printer ready."
     lcl_cancel_label   = "Cancel"
  	  lcl_cancel_msg     = "Return to the &quot;Create a Membership List&quot; without saving any of the data."
  	  lcl_cancel_url     = "remove_image()"
     sDisplaySaveMsg    = "<li><strong>Save Card: </strong>ONLY saves the membership card data so that it can be printed at another time.</li>" & vbcrlf
     sDisplaySaveButton = "<input type=""button"" value=""Save Card"" class=""button noprint"" id=""card_save"" name=""card_save"" onclick=""card_save();"" />" & vbrlf
  else
     lcl_print_label    = "Reprint Card"
     lcl_print_msg      = "Prints, and saves, the membership card.  Please have your color printer ready."
     lcl_cancel_label   = "Card Completed"
	    lcl_cancel_msg     = "The new membership card has been printed and the data has saved.  Click to<br />return to &quot;Create a Membership List&quot; results screen."
     lcl_cancel_url     = "goBack()"
     sDisplaySaveMsg    = ""
     sDisplaySaveButton = ""
  end if

 'Setup the card printed cound
  sDisplayCardPrintedCount = ""

  if lcl_demo = "Y" then
     if lcl_action = "CARD_PRINTED" then
          sDisplayCardPrintedCount = "<div align=""center""># times Membership Card has been printed: 2</div>" & vbcrlf
       end if
    else
       dim oCardCount

      'determine if this is the first time that a card has been printed
       sSqlc = "SELECT printed_count "
       sSqlc = sSqlc & " FROM egov_poolpassmembers "
       sSqlc = sSqlc & " WHERE memberid = " & clng(lcl_member_id)

       Set oCardCount = Server.CreateObject("ADODB.Recordset")
       oCardCount.Open sSqlc, Application("DSN"), 0, 1

       if not oCardCount.eof and clng(oCardCount("printed_count")) > 0 then
          sDisplayCardPrintedCount = "<div align=""center""># times Membership Card has been printed: " & oCardCount("printed_count") & "</div>" & vbcrlf
       end if

       oCardCount.close
       set oCardCount = nothing
    end if
%>
<html>
<head>
<title>E-Gov Administration Console {Membership Photo and ID Creation}</title>

  <meta name="viewport" content="width=device-width, minimum-scale=1, maximum-scale=1" />

  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="membership_card.css" />	

<style type="text/css">
.fieldset {
    background-color: #ffffff;
    border: 1pt solid #c0c0c0;
      border-radius: 5px;
    margin: 0px 10px;
}

.fieldset legend {
    background-color: #ffffff;
    border: 1pt solid #c0c0c0;
      border-radius: 5px;
    padding: 4px 10px;
    font-size: 1.25em;
    color: #800000;
}

#displayCard {
    position: relative;
    width: 300px;
    height: 176px;
}

#barcodeMsg {
    text-align: right;
    font-weight: bold;
    color: #ff0000;
}

#barcodeTD {
    border: 1pt solid #c0c0c0;
      border-radius: 5px;
    background-color: #eeeeee;
}

.barcodeID {
    width: 200px;
}

#fieldsetKeyCards {
    margin: 10px 10px;
}

#keyCardsTable {
    border: 1pt solid #c0c0c0;
    border-radius: 5px;
}

#keyCardsTable th {
    padding: 4px;
    vertical-align: bottom;
    white-space: nowrap;
}

.rowHighlight {
    background-color: #eeeeee;
}

#keyCardsTable td {
    border-top: 1pt solid #c0c0c0;
    padding: 4px;
    text-align: center;
    white-space: nowrap;
}

#keyCardsTable .columnAlignLeft {
    text-align: left;
}

#keyCardsTable .columnAlignCenter {
    text-align: center;
}

#keyCardsTable .columnTextArea {
    width: 100%;
}

#keyCardsTable textarea {
    width: 96% !important;
}

input.button {
    border: 1px solid #777777;
    border-radius: 4px;
    color: #000000;
    cursor: pointer;
    padding: 4px;
}

input.button:hover {
    background-color: #666666;
    color: #ffffff;
}

input {
    font-family: Verdana,Tahoma,Arial;
    font-size: 11px;
}

#barcodesButtonRow {
    margin: 0px 0px 10px 0px;
}

.barcodeData {
    border: 0pt solid #000000 !important;
}

.barcodeData td {
    border: 0pt solid #000000 !important;
    text-align: left !important;
}

.barcodeDisplayInfo {
    text-align: left;
}

.displayAddUpdateMsg {
    color: #ff0000;
}

/* ------------------------------------------------------------------------- */
@media screen and (max-width: 800px)
{
    .barcodeID {
        width: 100px;
    }
}
/* ------------------------------------------------------------------------- */
</style>

  <script type="text/javascript" src="validator.js"></script>
  <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

<script type="text/javascript">
  $(document).ready(function() {
    $('input[id^="barcode"]').css('display','none');
    $('select[id^="status"]').css('display','none');
    $('textarea[id^="comments"]').css('display','none');
    $('input[id^="removeBarcode"]').css('display','none');
    $('input[id^="buttonSave"]').css('display','none');
    $('div[id^="displayAddUpdateMsg"]').html('');
    $('div[id^="displayAddUpdateMsg"]').css('display','none');

    $('#backButton').click(function() {
        goBack();
    });

    $('#maintainBarcodeButton').click(function() {
       var lcl_url  = 'barcodeStatusesMaint.asp';
           lcl_url += '?memberid=<%=lcl_member_id%>';
           lcl_url += '&action=<%=lcl_action%>';
           lcl_url += '&poolpassid=<%=lcl_poolpassid%>';

       location.href = lcl_url;
    });

    $('#addBarcodeButton').click(function() {
        var lcl_total_barcodes = $('#totalBarcodes').val();
        var lcl_bgcolor        = $('#barcodeRow' + lcl_total_barcodes).prop('bgColor');

        if(lcl_total_barcodes == '') {
           lcl_total_barcodes = 0;
        }

        if(lcl_bgcolor == "#eeeeee") {
           lcl_bgcolor = "#ffffff";
        } else {
           lcl_bgcolor = "#eeeeee";
        }

        //Determine what the rowid and total row count is
        var num              = new Number(lcl_total_barcodes);
        var lcl_new_total    = (num + 1);
        var lcl_new_rowid    = lcl_new_total.toString();
        var lcl_row_html     = '';
        var lcl_displayBarcodeStatuses = <%=buildBarcodeStatusOptions(session("orgid"), 0, "Y") %>;

        lcl_row_html  = '<tr id="barcodeRow' + lcl_new_rowid + '" class="barcodeRow" valign="top" bgcolor="' + lcl_bgcolor + '">';
        lcl_row_html += '    <td class="column' + lcl_new_rowid + '">';
        lcl_row_html += '        <input type="button" name="buttonEdit' + lcl_new_rowid + '" id="buttonEdit' + lcl_new_rowid + '" value="&nbsp;Edit&nbsp;" class="button" onclick="editKeycardRow(\'EDIT\', \'' + lcl_new_rowid + '\');" />';
        lcl_row_html += '        <input type="button" name="buttonSave' + lcl_new_rowid + '" id="buttonSave' + lcl_new_rowid + '" value="&nbsp;Add&nbsp;" class="button" onclick="editKeycardRow(\'SAVE\', \'' + lcl_new_rowid + '\');" />';
        lcl_row_html += '    </td>';
        lcl_row_html += '    <td class="column' + lcl_new_rowid + '">';
        lcl_row_html += '        <input type="hidden" name="memberBarcodeID'        + lcl_new_rowid + '" id="memberBarcodeID'        + lcl_new_rowid + '" value="0" />';
        lcl_row_html += '        <input type="hidden" name="isBarcodeStatusEnabled' + lcl_new_rowid + '" id="isBarcodeStatusEnabled' + lcl_new_rowid + '" value="Y" />';
        lcl_row_html += '        <input type="text" name="barcode' + lcl_new_rowid + '" id="barcode' + lcl_new_rowid + '" value="" />';
        lcl_row_html += '        <div id="displayBarcode' + lcl_new_rowid + '" class="barcodeDisplayInfo"></div>';
        lcl_row_html += '    </td>';
        lcl_row_html += '    <td class="column' + lcl_new_rowid + '">';
        lcl_row_html += '        <select name="status' + lcl_new_rowid + '" id="status' + lcl_new_rowid + '">';
        lcl_row_html +=            lcl_displayBarcodeStatuses;
        lcl_row_html += '        </select>';
        lcl_row_html += '        <div id="displayStatus' + lcl_new_rowid + '" class="barcodeDisplayInfo"></div>';
        lcl_row_html += '    </td>';
        lcl_row_html += '    <td class="column' + lcl_new_rowid + ' columnTextArea">';
        lcl_row_html += '        <textarea rows="2" name="comments' + lcl_new_rowid + '" id="comments' + lcl_new_rowid + '"></textarea>';
        lcl_row_html += '        <div id="displayComments' + lcl_new_rowid + '" class="barcodeDisplayInfo"></div>';
        lcl_row_html += '    </td>';
        lcl_row_html += '    <td class="column' + lcl_new_rowid + ' columnAlignCenter">';
        lcl_row_html += '        <input type="checkbox" name="removeBarcode' + lcl_new_rowid + '" id="removeBarcode' + lcl_new_rowid + '" value="Y" onclick="modifySaveButtonText(\'' + lcl_new_rowid + '\', \'Delete\');" />';
        lcl_row_html += '        <div id="displayAddUpdateMsg' + lcl_new_rowid + '" class="displayAddUpdateMsg"></div>';
        lcl_row_html += '    </td>';
        lcl_row_html += '</tr>';
        
        //Append the new row to the table and increment the barcodes total.
        $('#keyCardsTable').append(lcl_row_html);
        $('#totalBarcodes').val(lcl_new_rowid);

        $('#buttonEdit'          + lcl_new_rowid).css('display','none');
        $('#displayAddUpdateMsg' + lcl_new_rowid).html('');
        $('#displayAddUpdateMsg' + lcl_new_rowid).css('display','none');
        $('#barcode'             + lcl_new_rowid).focus();
    });
  });

function editKeycardRow(iStatus, iRowID) {
  var lcl_isStatusEnabled = $('#isBarcodeStatusEnabled' + iRowID).val();

  if(iStatus == 'EDIT') {
     $('#buttonEdit' + iRowID).fadeOut('slow', function() {
        $('#buttonSave' + iRowID).fadeIn('slow');
        modifySaveButtonText(iRowID, 'Save');
     });

     $('#displayBarcode' + iRowID).fadeOut('slow', function() {
        $('#displayBarcode' + iRowID).html('');
        $('#barcode' + iRowID).prop('disabled',false);
        $('#barcode' + iRowID).fadeIn('slow');
     });

  
//     if(lcl_isStatusEnabled == 'Y') {
        $('#displayStatus' + iRowID).fadeOut('slow', function() {
           $('#displayStatus' + iRowID).html('');
           $('#status' + iRowID).prop('disabled',false);
           $('#status' + iRowID).fadeIn('slow');
        });
//     }
//     else
//     {
//        var lcl_status_msg  = '[' + $('#displayStatus' + iRowID).html() + ']';
//            lcl_status_msg += '<br />';
//            lcl_status_msg += 'This status has';
//            lcl_status_msg += '<br />';
//            lcl_status_msg += 'been disabled.';

//        $('#displayStatus' + iRowID).html(lcl_status_msg)
//        $('#displayStatus' + iRowID).css('color', '#ff0000');
//        $('#status' + iRowID).prop('disabled',false);
//        $('#status' + iRowID).fadeIn('slow');
//     }

     $('#displayComments' + iRowID).fadeOut('slow', function() {
        $('#displayComments' + iRowID).html('');
        $('#comments' + iRowID).prop('disabled',false);
        $('#comments' + iRowID).fadeIn('slow');
     });

     $('#removeBarcode' + iRowID).prop('disabled',false);
     $('#removeBarcode' + iRowID).fadeIn('slow');
  }
  else
  {
     var lcl_memberBarcodeID = $('#memberBarcodeID' + iRowID).val();
     var lcl_barcode         = $('#barcode'         + iRowID).val();
     var lcl_status          = $('#status'          + iRowID).val();
     var lcl_comments        = $('#comments'        + iRowID).val();
     var lcl_removeBarcode   = $('#removeBarcode'   + iRowID);
     var lcl_removeIsChecked = 'N';
     var lcl_canProcess      = true;

     if(lcl_removeBarcode.prop('checked')) {
        lcl_removeIsChecked = 'Y';
     }

     $('#barcodeMsg').css('display', 'none');

//lcl_url = 'saveBarcodeToMemberID.asp';
//lcl_url += '?orgid=<%=session("orgid")%>';
//lcl_url += '&memberid=<%=lcl_member_id%>';
//lcl_url += '&userid=<%=session("userid")%>';
//lcl_url += '&memberBarcodeID=' + lcl_memberBarcodeID;
//lcl_url += '&statusid=' + lcl_status;
//lcl_url += '&barcode=' + lcl_barcode;
//lcl_url += '&comments=' + lcl_comments;

//alert(lcl_url);

     if(lcl_barcode == '' && !lcl_removeBarcode.prop('checked')) {
        $('#barcodeMsg').html('*** Error: A barcode is required. ***');
        $('#barcodeMsg').fadeIn('slow');
        window.setTimeout("clearScreenMsg('barcodeMsg')", (10 * 1000));

        $('#barcode' + iRowID).focus();
     } else {
        if(lcl_removeBarcode.prop('checked')) {
           if (!confirm("Are you sure you want to delete \"" + lcl_barcode + "\"?")) { 
              lcl_canProcess = false;
           }
        }

        if(lcl_canProcess)
        {
           $.post('saveBarcodeToMemberID.asp', {
              orgid:           '<%=session("orgid")%>',
              memberid:        '<%=lcl_member_id%>',
              userid:          '<%=session("userid")%>',
              memberBarcodeID: lcl_memberBarcodeID,
              statusid:        lcl_status,
              barcode:         lcl_barcode,
              comments:        lcl_comments,
              removeIsChecked: lcl_removeIsChecked
           }, function(result){
              var lcl_success = false;

              if(result == 'ActiveBarcodeForMemberExists')
              {
                 $('#barcodeMsg').html('*** Error: An ACTIVE barcode already exists for this member. ***');
                 $('#barcodeMsg').fadeIn('slow');
                 window.setTimeout("clearScreenMsg('barcodeMsg')", (10 * 1000));
              }
              else if (result == 'DuplicateBarcodeExists')
              {
                 $('#barcodeMsg').html('*** Error: The barcode entered is a duplicate of another barcode. ***');
                 $('#barcodeMsg').fadeIn('slow');
                 window.setTimeout("clearScreenMsg('barcodeMsg')", (10 * 1000));
              }
              else if (result == 'SD')
              {
                 $('#barcodeMsg').html('*** Success: The barcode has been successfully deleted. ***');
                 $('#barcodeMsg').fadeIn('slow');
                 window.setTimeout("clearScreenMsg('barcodeMsg')", (10 * 1000));

                 $('td[class^="column' + iRowID + '"]').fadeOut('slow');

                 //Recalculate the TotalBarcodes
                 var lcl_total_barcodes = $('#totalBarcodes').val();
                 var num           = new Number(lcl_total_barcodes);
                 var lcl_new_total = (num - 1);

                 if(lcl_new_total < 0) {
                    lcl_new_total = 0;
                 }

                 $('#barcodeRow' + iRowID).remove();
                 $('#totalBarcodes').val(lcl_new_total);
              }
              else
              {
                  $('#displayAddUpdateMsg' + iRowID).html(result);
                  $('#displayAddUpdateMsg' + iRowID).fadeIn('slow', function() {
                      window.setTimeout("clearScreenMsg('displayAddUpdateMsg" + iRowID + "')", (10 * 1000));
                  });

                  lcl_success = true;
              }

              if(lcl_success)
              {
                 $('#displayBarcode'  + iRowID).html($('#barcode'  + iRowID).val());
                 $('#displayStatus'   + iRowID).html($('#status'   + iRowID + ' option:selected').text());
                 $('#displayComments' + iRowID).html($('#comments' + iRowID).val());

                 $('#barcode' + iRowID).fadeOut('slow', function() {
                    $('#displayBarcode' + iRowID).fadeIn('slow');
                 });

                 $('#status' + iRowID).fadeOut('slow', function() {
                    $('#displayStatus' + iRowID).fadeIn('slow');
                 });

                 $('#comments' + iRowID).fadeOut('slow', function() {
                    $('#displayComments' + iRowID).fadeIn('slow');
                 });

                 $('#removeBarcode' + iRowID).fadeOut('slow');

                 $('#buttonSave' + iRowID).fadeOut('slow', function() {
                     $('#buttonEdit' + iRowID).fadeIn('slow');
                 });
              }
           });
        }
     }
  }
}

function clearScreenMsg(iFieldID) {
  $('#' + iFieldID).fadeOut('slow');
}

function goBack() {
  var lcl_return_page = '<%=session("redirectpage")%>';

  if (lcl_return_page != "") {
      location.href = lcl_return_page;
  } else {
      history.go(-1);
  }
}

function card_print() {
  var lcl_url_print  = 'card_print.asp';
      lcl_url_print += '?memberid=<%=lcl_member_id%>';
      lcl_url_print += '&poolpassid=<%=lcl_poolpassid%>';
      lcl_url_print += '&status=PRINT<%=lcl_demo_url%>';
      lcl_url_print += '&card_layout=<%=request("card_layout")%>';
      lcl_url_print += '&initPrint=Y';
      lcl_url_print += '&OS=XP';

  var lcl_url_display  = 'image_display.asp';
      lcl_url_display += '?memberid=<%=lcl_member_id%>';
      lcl_url_display += '&poolpassid=<%=lcl_poolpassid%>';
      lcl_url_display += '&action=PRINT_CARD<%=lcl_demo_url%>';
      lcl_url_display += '&card_layout=p';

  window.open(lcl_url_print);
  location.href = lcl_url_display;
}

function retake_picture() {
  var lcl_url_retake  = '<%=sTakePicUrl%>';
      lcl_url_retake += '?memberid=<%=lcl_member_id%>';
      lcl_url_retake += '&poolpassid=<%=lcl_poolpassid%>';
      lcl_url_retake += '&reload_pic=Y<%=lcl_demo_url%>';

  location.href = lcl_url_retake;
}

function reload_picture() {
  if ("Y"=="<%=Session("RELOAD_PIC")%>") {
      window.location.reload();
<% session("RELOAD_PIC") = "N" %>
  }else{
     return true;
  }
}

function remove_image() {
  var lcl_url_remove  = 'image_display.asp';
      lcl_url_remove += '?memberid=<%=lcl_member_id%>';
      lcl_url_remove += '&poolpassid=<%=lcl_poolpasid%>';
      lcl_url_remove += '&action=CANCEL<%=lcl_demo_url%>';

  location.href = lcl_url_remove;
}

function card_save() {
  var lcl_url_save  = 'image_display.asp';
      lcl_url_save += '?memberid=<%=lcl_member_id%>';
      lcl_url_save += '&poolpassid=<%=lcl_poolpassid%>';
      lcl_url_save += '&action=SAVE<%=lcl_demo_url%>';
      lcl_url_save += '&card_layout=p';

  location.href = lcl_url_save;
}

function modifySaveButtonText(iRowID, iValue) {

   var lcl_buttonText = iValue;

   if(lcl_buttonText == 'Delete') {
      var lcl_fieldsDisabled = false;
      var lcl_fieldsBGColor  = '';

      if($('#removeBarcode' + iRowID).prop('checked'))
      {
         lcl_fieldsDisabled = true;
         lcl_fieldsBGColor  = '#eeeeee';
      }
      else
      {
         lcl_buttonText = 'Save';
      }

      $('#barcode'  + iRowID).css('background-color', lcl_fieldsBGColor);
      $('#status'   + iRowID).css('background-color', lcl_fieldsBGColor);
      $('#comments' + iRowID).css('background-color', lcl_fieldsBGColor);

      $('#barcode'  + iRowID).prop('disabled', lcl_fieldsDisabled);
      $('#status'   + iRowID).prop('disabled', lcl_fieldsDisabled);
      $('#comments' + iRowID).prop('disabled', lcl_fieldsDisabled);
   }

   $('#buttonSave' + iRowID).val(lcl_buttonText);

}

</script>
</head>

<body onload="reload_picture()">
	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 
<%
  response.write "<table border=""0"" cellpadding=""6"" cellspacing=""0"" class=""start"" width=""100%"">" & vbcrlf
  response.write "  <tr>" & vbcrlf
  response.write "      <td valign=""top"">" & vbcrlf
  response.write "          <h3>Print Membership Card" & lcl_demo_page_title & "</h3>" & vbcrlf
  response.write "          <input type=""button"" name=""backButton"" id=""backButton"" value=""Back"" class=""button"" />" & vbrlf

 'BEGIN: Display Card ---------------------------------------------------------
  response.write "          <p>" & vbrlf
  response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""2"">" & vbrlf
  response.write "            <tr valign=""top"">" & vbrlf
  response.write "              		<td width=""320"">" & vbrlf
  response.write "              		    <div id=""displayCard"">" & vbrlf
%>
  <!--#include file="membership_card.asp"-->
<%
  response.write "              		    </div>" & vbrlf
  response.write "                </td>" & vbrlf
 'END: Display Card -----------------------------------------------------------

 'BEGIN: Buttons --------------------------------------------------------------
  response.write "                <td height=""200"">" & vbrlf
  response.write "                    <fieldset class=""fieldset"">" & vbcrlf
  response.write "                      <legend>Button Instructions</legend>" & vbrlf
  response.write "                		    <ul>" & vbrlf
  response.write "                        <li><strong>" & lcl_print_label & ": </strong>" & lcl_print_msg & "</li>" & vbrlf
  response.write                          sDisplaySaveMsg
  response.write "                  				  <li><strong>Retake Picture: </strong>Click to retake the picture.</li>" & vbrlf
  response.write "                  				  <li><strong>" & lcl_cancel_label & ": </strong>" & lcl_cancel_msg & "</li>" & vbrlf
  response.write "                      </ul>" & vbrlf
  response.write "                      <div align=""center"">" & vbrlf
  response.write "                    		  <input type=""button"" value=""" & lcl_print_label & """ id=""card_print"" name=""card_print"" onclick=""card_print();"" class=""button noprint"" />" & vbrlf
  response.write                          sDisplaySaveButton
  response.write "                    				<input type=""button"" value=""Retake Picture"" id=""retake_picture"" name=""retake_picture"" onclick=""retake_picture();"" class=""button noprint"" />" & vbrlf
  response.write "                    				<input type=""button"" value=""" & lcl_cancel_label & """ class=""button noprint"" id=""remove_image"" name=""remove_image"" onclick=""" & lcl_cancel_url & """ />" & vbrlf
  response.write "                    		</div>" & vbrlf
  response.write "                    </fieldset>" & vbcrlf
  response.write "                </td>" & vbrlf
 'END: Buttons ----------------------------------------------------------------

  response.write "            </tr>" & vbrlf

 'BEGIN: Key Cards ------------------------------------------------------------
  if lcl_orghasfeature_memberships_usekeycards then
     response.write "            <tr>" & vbrlf
     response.write "                <td colspan=""2"" id=""barcodeMsg""></td>" & vbcrlf
     response.write "            </tr>" & vbcrlf
     response.write "            <tr valign=""top"">" & vbrlf
     response.write "                <td colspan=""2"" id=""barcodeTD"">" & vbrlf
     response.write "                    <fieldset id=""fieldsetKeyCards"" class=""fieldset"">" & vbrlf
     response.write "                      <legend>Pre-Printed Keycards</legend>" & vbcrlf
     response.write "                      <div id=""barcodesButtonRow"">" & vbcrlf
     response.write "                        <input type=""button"" name=""maintainBarcodeButton"" id=""maintainBarcodeButton"" value=""Maintain Barcode Statuses"" class=""button"" />" & vbcrlf
     response.write "                        <input type=""button"" name=""addBarcodeButton"" id=""addBarcodeButton"" value=""Add Barcode"" class=""button"" />" & vbcrlf
     response.write "                      </div>" & vbcrlf
     response.write "                      <table id=""keyCardsTable"" border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
     response.write "                        <tr class=""columnAlignLeft"">" & vbcrlf
     response.write "                            <th>&nbsp;</th>" & vbcrlf
     response.write "                            <th>Barcode</th>" & vbcrlf
     'response.write "                            <th>Rate</th>" & vbcrlf
     response.write "                            <th>Status</th>" & vbcrlf
     response.write "                            <th>Comments</th>" & vbcrlf
     response.write "                            <th class=""columnAlignCenter"">Delete</th>" & vbcrlf
     'response.write "                            <th>&nbsp;</th>" & vbcrlf
     response.write "                        </tr>" & vbcrlf

     sSQLb = "SELECT memberbarcodeid, "
     sSQLb = sSQLb & " orgid, "
     sSQLb = sSQLb & " memberid, "
     sSQLb = sSQLb & " barcode, "
     sSQLb = sSQLb & " barcode_statusid, "
     sSQLb = sSQLb & " barcode_comments, "
     sSQLb = sSQLb & " isnull(createdbyid, 0) as createdbyid, "
     sSQLb = sSQLb & " isnull(createddate, '') as createddate, "
     sSQLb = sSQLb & " isnull(lastmodifiedbyid, 0) as lastmodifiedbyid, "
     sSQLb = sSQLb & " isnull(lastmodifieddate, '') as lastmodifieddate "
     sSQLb = sSQLb & " FROM egov_poolpassmembers_to_barcodes "
     sSQLb = sSQLb & " WHERE orgid = " & session("orgid")
     sSQLb = sSQLb & " AND memberid = " & lcl_member_id

     set oMemberBarcodes = Server.CreateObject("ADODB.Recordset")
     oMemberBarcodes.Open sSQLb, Application("DSN"), 3, 1

     if not oMemberBarcodes.eof then
        sLineCount            = 0
        sBGColor              = "#eeeeee"
        sMemberBarcodeID      = 0
        sBarcode              = ""
	       sBarcodeStatusID      = 0
      	 sBarcodeStatusName    = ""
        sBarcodeComments      = ""
        sBarcodeStatusEnabled = ""
        sIsJavascript         = "N"
        'sDisplayBarcodeRates    = buildBarcodeRateOptions(session("orgid"), _
        '                                                  lcl_member_id)

        do while not oMemberBarcodes.eof
           sLineCount              = sLineCount + 1
           sMemberBarcodeID        = oMemberBarcodes("memberbarcodeid")
           sBarcode                = oMemberBarcodes("barcode")
           sBarcodeStatusID        = oMemberBarcodes("barcode_statusid")
           sBarcodeComments        = oMemberBarcodes("barcode_comments")
           sBarcodeStatusEnabled   = ""
           sDisplayBarcodeStatuses = buildBarcodeStatusOptions(session("orgid"), _
                                                               sBarcodeStatusID, _
                                                               sIsJavascript)

           sBarcodeStatusName = getBarcodeStatusName(session("orgid"), _
                                                     sBarcodeStatusID)

           sIsStatusEnabled = isStatusEnabled(session("orgid"), _
                                              sBarcodeStatusID)

           if sIsStatusEnabled then
              sBarcodeStatusEnabled = "Y"
           end if

           response.write "                     <tr id=""barcodeRow" & sLineCount & """ class=""barcodeRow"" valign=""top"" bgcolor=""" & sBGColor & """>" & vbcrlf
           response.write "                         <td class=""column" & sLineCount & """>" & vbcrlf
           response.write "                             <input type=""button"" name=""buttonEdit" & sLineCount & """ id=""buttonEdit" & sLineCount & """ value=""&nbsp;Edit&nbsp;"" class=""button"" onclick=""editKeycardRow('EDIT', '" & sLineCount & "');"" />" & vbcrlf
           response.write "                             <input type=""button"" name=""buttonSave" & sLineCount & """ id=""buttonSave" & sLineCount & """ value=""Save"" class=""button"" onclick=""editKeycardRow('SAVE', '" & sLineCount & "');"" />" & vbcrlf
           response.write "                         </td>" & vbcrlf
           response.write "                         <td class=""column" & sLineCount & """>" & vbcrlf
           response.write "                             <input type=""hidden"" name=""memberBarcodeID"        & sLineCount & """ id=""memberBarcodeID"        & sLineCount & """ value=""" & sMemberBarcodeID      & """ />" & vbcrlf
           response.write "                             <input type=""hidden"" name=""isBarcodeStatusEnabled" & sLineCount & """ id=""isBarcodeStatusEnabled" & sLineCount & """ value=""" & sBarcodeStatusEnabled & """ />" & vbcrlf
           response.write "                             <input type=""text"" name=""barcode" & sLineCount & """ id=""barcode" & sLineCount & """ value=""" & sBarcode & """ class=""barcodeID"" />" & vbcrlf
           response.write "                             <div id=""displayBarcode" & sLineCount & """ class=""barcodeDisplayInfo"">" & sBarcode & "</div>" & vbcrlf
           response.write "                         </td>" & vbcrlf
           'response.write "                         <td>" & vbcrlf
           'response.write "                             <select name=""poolpassid" & sLineCount & """ id=""poolpassid" & sLineCount & """>" & vbcrlf
           'response.write                                 sDisplayBarcodeRates & vbcrlf
           'response.write "                             </select>" & vbcrlf
           'response.write "                         </td>" & vbcrlf
           response.write "                         <td class=""column" & sLineCount & """>" & vbcrlf
           response.write "                             <select name=""status" & sLineCount & """ id=""status" & sLineCount & """>" & vbcrlf
           response.write                                 sDisplayBarCodeStatuses & vbcrlf
           response.write "                             </select>" & vbcrlf
           response.write "                             <div id=""displayStatus" & sLineCount & """ class=""barcodeDisplayInfo"">" & sBarcodeStatusName & "</div>" & vbcrlf
           response.write "                         </td>" & vbcrlf
           response.write "                         <td class=""column" & sLineCount & " columnTextArea"">" & vbcrlf
           response.write "                             <textarea rows=""2"" name=""comments" & sLineCount & """ id=""comments" & sLineCount & """>" & sBarcodeComments &"</textarea>" & vbcrlf
           response.write "                             <div id=""displayComments" & sLineCount & """ class=""barcodeDisplayInfo"">" & sBarcodeComments & "</div>" & vbcrlf
           response.write "                         </td>" & vbcrlf
           response.write "                         <td class=""column" & sLineCount & " columnAlignCenter"">" & vbcrlf
           response.write "                             <input type=""checkbox"" name=""removeBarcode" & sLineCount & """ id=""removeBarcode" & sLineCount & """ value=""Y"" onclick=""modifySaveButtonText('" &sLineCount & "', 'Delete');"" />" & vbcrlf
           response.write "                             <div id=""displayAddUpdateMsg" & sLineCount & """ class=""displayAddUpdateMsg""></div>" & vbcrlf
           response.write "                         </td>" & vbcrlf
           response.write "                     </tr>" & vbcrlf

           sBGColor = changeBGColor(sBGColor, "#eeeeee", "#ffffff")

           oMemberBarcodes.movenext
        loop
     end if

     oMemberBarcodes.close
     set oMemberBarcodes = nothing

     response.write "                      </table>" & vbcrlf
     response.write "                      <input type=""hidden"" name=""totalBarcodes"" id=""totalBarcodes"" value=""" & sLineCount & """ />" & vbcrlf
     response.write "                    </fieldset>" & vbcrlf
     response.write "                </td>" & vbrlf
     response.write "            </tr>" & vbrlf
  end if
 'END: Key Cards --------------------------------------------------------------

  response.write "          </table>" & vbrlf
  response.write            sDisplayCardPrintedCount
  response.write "      </td>" & vbcrlf
  response.write "    </tr>" & vbcrlf
  response.write "</table>" & vbcrlf
%>
  <!--#Include file="../admin_footer.asp"-->  
<%
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf

 '-----------------------------------------------------------------------------
function buildBarcodeStatusOptions(iOrgID, _
                                   iCurrentStatusID, _
                                   iIsJavascript)

  dim lcl_return, sSQL, sCurrentStatusID, lcl_selected_status, sLineCount

  lcl_return       = ""
  sCurrentStatusID = 0
  sLineCount       = 0

  lcl_selected_status = ""

  if iCurrentStatusID <> "" then
     if isnumeric(iCurrentStatusID) then
        sCurrentStatusID = clng(iCurrentStatusID)
     end if
  end if

  sSQL = "SELECT statusid, "
  sSQL = sSQL & " statusname "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND isEnabled = 1 "
  sSQL = sSQL & " ORDER BY isEnabled DESC, statusname "

  set oBarcodeStatus = Server.CreateObject("ADODB.Recordset")
  oBarcodeStatus.Open sSQL, Application("DSN"), 3, 1

  if not oBarcodeStatus.eof then
     do while not oBarcodeStatus.eof
        sLineCount = sLineCount + 1
        lcl_selected_status = ""

        if clng(oBarcodeStatus("statusid")) = sCurrentStatusID then
           lcl_selected_status = " selected=""selected"""
        end if

        if iIsJavascript = "Y" then
           if sLineCount > 1 then
              lcl_return = lcl_return & "+"
           end if

           lcl_return = lcl_return & "'<option value=""" & oBarcodeStatus("statusid") & """" & lcl_selected_status & ">" & oBarcodeStatus("statusname") & "</option>'"
        else
           lcl_return = lcl_return & "<option value=""" & oBarcodeStatus("statusid") & """" & lcl_selected_status & ">" & oBarcodeStatus("statusname") & "</option>" & vbcrlf
        end if

        oBarcodeStatus.movenext
     loop
  end if

  oBarcodeStatus.close
  set oBarcodeStatus = nothing

  if lcl_return = "" then
     if iIsJavascript = "Y" then
        lcl_return = "'"
     end if

     lcl_return = lcl_return & "<option value="""">No Statuses Available</option>"

     if iIsJavascript = "Y" then
        lcl_return = lcl_return & "'"
     end if
  end if

  buildBarcodeStatusOptions = lcl_return

end function

'------------------------------------------------------------------------------
function buildBarcodeRateOptions(iOrgID, _
                                 iMemberID)

  dim lcl_return, sSQL, sBarcodesExist

  lcl_return = ""

  sBarcodesExist        = checkBarcodesExist(iMemberID)
  sFamilyMemberID       = getFamilyMemberID(iMemberID)
  sBelongsToUserID      = getFamilyMembersInfo_byFamilyMemberID("belongstouserid", sFamilyMemberID)
  sUserID               = getFamilyMembersInfo_byFamilyMemberID("userid", sFamilyMemberID)
  sUserName             = GetCitizenName(sUserID)
  sDisplayRateInfo      = ""
  lcl_rate_description  = ""
  lcl_rate_residenttype = ""

  sSQL = "SELECT ppp.poolpassid, "
  sSQL = sSQL & " ppp.rateid, "
  sSQL = sSQL & " ppp.expirationdate, "
  sSQL = sSQL & " ppr.description, "
  sSQL = sSQL & " m.membershipdesc, "
  sSQL = sSQL & " mp.period_desc "
  sSQL = sSQL & " FROM egov_poolpasspurchases ppp "
  sSQL = sSQL &      " INNER JOIN egov_membership_periods mp ON mp.periodid = ppp.periodid "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_poolpassrates ppr ON ppr.rateid = ppp.rateid "
  sSQL = sSQL &      " LEFT OUTER JOIN egov_memberships m ON m.membershipid = ppp.membershipid "
  sSQL = sSQL & " WHERE mp.orgid = ppp.orgid "
  sSQL = sSQL & " AND ppp.userid = " & sBelongsToUserID
  sSQL = sSQL & " ORDER BY ppp.expirationdate DESC "

  set oBarcodeRates = Server.CreateObject("ADODB.Recordset")
  oBarcodeRates.Open sSQL, Application("DSN"), 3, 1

  if not oBarcodeRates.eof then
     do while not oBarcodeRates.eof

        getRateInfo oBarcodeRates("rateid"), _
                    lcl_rate_description, _
                    lcl_rate_residenttype

        sDisplayRateInfo = ""

        if oBarcodeRates("membershipdesc") <> "" then
           sDisplayRateInfo = oBarcodeRates("membershipdesc")
        end if

        if oBarcodeRates("period_desc") <> "" then
           if sDisplayRateInfo <> "" then
              sDisplayRateInfo = sDisplayRateInfo & " - "
           end if

           sDisplayRateInfo = sDisplayRateInfo & oBarcodeRates("period_desc")
        end if

        if lcl_rate_description = "" then
           lcl_rate_description = "Rate Not Available"
        end if

        sDisplayRateInfo = sDisplayRateInfo & " [" & lcl_rate_description & "] "

        sDisplayRateInfo = sDisplayRateInfo & "<strong>&nbsp;expires: " & datevalue(oBarcodeRates("expirationdate")) & "</strong>"

        lcl_return = lcl_return & "  <option value=""" & oBarcodeRates("poolpassid") & "_" & iMemberID & """>" & sDisplayRateInfo & "</option>" & vbcrlf

        oBarcodeRates.movenext
     loop
  end if

  oBarcodeRates.close
  set oBarcodeRates = nothing

  buildBarcodeRateOptions = lcl_return

end function

'------------------------------------------------------------------------------
function checkBarcodesExist(iMemberID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT count(poolpassid) as totalCount "
  sSQL = sSQL & " FROM egov_poolpassmembers "
  sSQL = sSQL & " WHERE memberid = " & iMemberID

  set oBarcodesExist = Server.CreateObject("ADODB.Recordset")
  oBarcodesExist.Open sSQL, Application("DSN"), 3, 1

  if not oBarcodesExist.eof then
     if oBarcodesExist("totalCount") > 0 then
        lcl_return = true
     end if
  end if

  oBarcodesExist.close
  set oBarcodesExist = nothing

  checkBarcodesExist = lcl_return

end function

'------------------------------------------------------------------------------
function getFamilyMemberID(iMemberID)
  dim lcl_return, sSQL

  lcl_return = 0

  sSQL = "SELECT max(familymemberid) as familymemberid "
  sSQL = sSQL & " FROM egov_poolpassmembers "
  sSQL = sSQL & " WHERE memberid = " & iMemberID

  set oGetFamilyMemberID = Server.CreateObject("ADODB.Recordset")
  oGetFamilyMemberID.Open sSQL, Application("DSN"), 3, 1

  if not oGetFamilyMemberID.eof then
     lcl_return = oGetFamilyMemberID("familymemberid")
  end if

  oGetFamilyMemberID.close
  set oGetFamilyMemberID = nothing

  getFamilyMemberID = lcl_return

end function

'------------------------------------------------------------------------------
function getFamilyMembersInfo_byFamilyMemberID(iColumnName, _
                                               iFamilyMemberID)

  dim lcl_return, sSQL, sColumnName

  lcl_return = ""

  if iColumnName <> "" then
     sColumnName = dbsafe(iColumnName)

     sSQL = "SELECT " & sColumnName & " as dbColumn "
     sSQL = sSQL & " FROM egov_familymembers "
     sSQL = sSQL & " WHERE familymemberid = " & iFamilyMemberID

     set oGetFamilyMemberInfo = Server.CreateObject("ADODB.Recordset")
     oGetFamilyMemberInfo.Open sSQL, Application("DSN"), 3, 1

     if not oGetFamilyMemberInfo.eof then
        lcl_return = oGetFamilyMemberInfo("dbColumn")
     end if

     oGetFamilyMemberInfo.close
     set oGetFamilyMemberInfo = nothing
  end if

  getFamilyMembersInfo_byFamilyMemberID = lcl_return

end Function

'------------------------------------------------------------------------------
function getBarcodeStatusName(iOrgID, _
                              iStatusID)

  dim lcl_return, sSQL

  lcl_return = ""

  sSQL = "SELECT statusname "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND statusid = " & iStatusID

  set oGetBarcodeStatusName = Server.CreateObject("ADODB.Recordset")
  oGetBarcodeStatusName.Open sSQL, Application("DSN"), 3, 1

  if not oGetBarcodeStatusName.eof then
     lcl_return = oGetBarcodeStatusName("statusname")
  end if

  oGetBarcodeStatusName.close
  set oGetBarcodeStatusName = nothing

  getBarcodeStatusName = lcl_return

end function

'------------------------------------------------------------------------------
function isStatusEnabled(iOrgID, _
                         iStatusID)

  dim lcl_return, sSQL

  lcl_return = false

  sSQL = "SELECT isEnabled "
  sSQL = sSQL & " FROM egov_poolpassmembers_barcode_statuses "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND statusid = " & iStatusID

  set oIsStatusEnabled = Server.CreateObject("ADODB.Recordset")
  oIsStatusEnabled.Open sSQL, Application("DSN"), 3, 1

  if not oIsStatusEnabled.eof then
     if oIsStatusEnabled("isEnabled") then
        lcl_return = true
     end if
  end if

  oIsStatusEnabled.close
  set oIsStatusEnabled = nothing

  isStatusEnabled = lcl_return

end function
%>
