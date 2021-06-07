<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="news_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: new_item.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module adds News Scroller Items.
'
' MODIFICATION HISTORY
' 1.0 09/11/06	Steve Loar - Initial Version.
' 1.2	02/25/08	Steve Loar - Added textarea limit JavaScript and increased size to 400 from 200
' 1.3 09/23/08 David Boyer - Added new error handling
' 1.4 05/07/09 David Boyer - Combined "edit_item.asp" and "new_item.asp"
' 1.5 05/08/09 David Boyer - Converted "Find a Link" from the "site linker" format to the "link picker" format.
' 1.6 07/09/09 David Boyer - Added "newstype" to split News and News Scroller items.
' 1.7 06/11/10 David Boyer - Now hide "Date" field (top) when adding/editing Community Link news items.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 sLevel = "../"  'Override of value from common.asp

'Determine the newstype.
'NEWS     = News Items (CommunityLink)
'SCROLLER = News Scroller Items
 if request("newstype") <> "" then
    lcl_newstype = UCASE(request("newstype"))
 else
    lcl_newstype = "SCROLLER"
 end if

'Setup page variables
 if lcl_newstype = "NEWS" then
    lcl_feature                 = "edit_news_items"
    lcl_userpermission_required = "edit_news_items"
    lcl_pagetitle               = "News Item"
    lcl_messagelength           = "8000"
    lcl_returnAsHTMLLink        = "Y"
 else
    lcl_feature                 = "edit scroller"
    lcl_userpermission_required = "edit scroller"
    lcl_pagetitle               = "News Scroller Item"
    lcl_messagelength           = "400"
    if session("orgid") = "113" then lcl_messagelength           = "8000"
    lcl_returnAsHTMLLink        = "N"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission_required) then
 	  response.redirect sLevel & "permissiondenied.asp"
 end if

 Dim sSQL, oItem

 if request("newsitemid") <> "" AND request("newsitemid") > 0 then
    lcl_newsitemid  = request("newsitemid")
    lcl_page_label  = "Edit"
    lcl_sendToLabel = "Update"
    lcl_screen_mode = "EDIT"

 else
    lcl_newsitemid  = 0
    lcl_page_label  = "Create"
    lcl_sendToLabel = "Create"
    lcl_screen_mode = "ADD"
 end if

'Set up the page variables
 lcl_itemtitle        = ""
 lcl_itemdate         = ""
 lcl_itemtext         = ""
 lcl_itemlinkurl      = ""
 lcl_publicationstart = ""
 lcl_publicationend   = ""

'Get the new item data
 if lcl_newsitemid > 0 then
   	sSQL = "SELECT newsitemid, "
    sSQL = sSQL & " itemtitle, "
    sSQL = sSQL & " itemdate, "
    sSQL = sSQL & " itemtext, "
    sSQL = sSQL & " isnull(itemlinkurl,'') as itemlinkurl, "
    sSQL = sSQL & " publicationstart, "
    sSQL = sSQL & " publicationend, "
    sSQL = sSQL & " itemdisplay "
   	sSQL = sSQL & " FROM egov_news_items "
    sSQL = sSQL & " WHERE newsitemid = " & request("newsitemid")
    sSQL = sSQL & " AND orgid = " & session("orgid")

   	set oItem = Server.CreateObject("ADODB.Recordset")
   	oItem.Open sSQL, Application("DSN"), 3, 1

    if not oItem.eof then
       lcl_newsitemid       = oItem("newsitemid")
       lcl_itemtitle        = oItem("itemtitle")
       lcl_itemdate         = oItem("itemdate")
       lcl_itemtext         = oItem("itemtext")
       lcl_itemlinkurl      = oItem("itemlinkurl")
       lcl_publicationstart = oItem("publicationstart")
       lcl_publicationend   = oItem("publicationend")

    			if oItem("itemdisplay") then
 			      lcl_checked_itemdisplay = " checked=""checked"""
 			   else
     			  lcl_checked_itemdisplay = ""
 						end if

    end if

   	oItem.close
   	set oItem = nothing

 end if

'Check for org features
 lcl_orghasfeature_rssfeeds_news = orghasfeature("rssfeeds_news")

'Check for user permissions
 lcl_userhaspermission_rssfeeds_news = userhaspermission(session("userid"),"rssfeeds_news")

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Determine if there is any additional processing needed from the past update
 if lcl_orghasfeature_rssfeeds_news AND lcl_userhaspermission_rssfeeds_news AND (lcl_success = "SU" OR lcl_success = "SA") then
    if request("sendTo_RSS") <> "" then
       lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
    end if
 end if
%>
<html>
<head>
<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />

	<script type="text/javascript" src="../scripts/ajaxLib.js"></script>
	<script type="text/javascript" src="../scripts/isvaliddate.js"></script>
	<script type="text/javascript" src="../scripts/textareamaxlength.js"></script>
 <script type="text/javascript" src="../scripts/formvalidation_msgdisplay.js"></script>
 <script type="text/javascript" src="../scripts/jquery-1.9.1.min.js"></script>

	<script type="text/javascript">
	<!--

		function doCalendar(sField) 
		{
		  var w = (screen.width - 350)/2;
		  var h = (screen.height - 350)/2;
		  eval('window.open("calendarpicker.asp?p=1&updatefield=' + sField + '&updateform=frmEditItem", "_calendar", "width=350,height=250,toolbar=0,statusbar=0,scrollbars=1,menubar=0,location=0,left=' + w + ',top=' + h + '")');
		}

		function doSitePicker(sFormField) 
		{
			w = (screen.width - 350)/2;
			h = (screen.height - 350)/2;
			eval('window.open("../sitelinker/default.asp?name=' + sFormField + '", "_picker", "width=600,height=470,toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + w + ',top=' + h + '")');
		}

function doPicker(sFormField, p_displayDocuments, p_displayActionLine, p_displayPayments, p_displayURL) {
  w = 600;
  h = 400;
  l = (screen.AvailWidth/2)-(w/2);
  t = (screen.AvailHeight/2)-(h/2);
  lcl_showFolderStart = "";
  lcl_folderStart     = 0;

  //Determine which options will be displayed
  if((p_displayDocuments=="")||(p_displayDocuments==undefined)) {
      lcl_displayDocuments = "";
  }else{
      lcl_displayDocuments = "&displayDocuments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayActionLine=="")||(p_displayActionLine==undefined)) {
      lcl_displayActionLine = "";
  }else{
      lcl_displayActionLine = "&displayActionLine=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayPayments=="")||(p_displayPayments==undefined)) {
      lcl_displayPayments = "";
  }else{
      lcl_displayPayments = "&displayPayments=Y";
      lcl_folderStart = lcl_folderStart + 1;
  }

  if((p_displayURL=="")||(p_displayURL==undefined)) {
      lcl_displayURL = "";
  }else{
      lcl_displayURL = "&displayURL=Y";
  }

  if(lcl_folderStart > 0) {
<% 'lcl_showFolderStart = "&folderStart=published_documents"; %>
     lcl_showFolderStart = "&folderStart=CITY_ROOT";
  }

  pickerURL  = "../picker_new/default.asp";
  pickerURL += "?name=" + sFormField;
  pickerURL += "&returnAsHTMLLink=<%=lcl_returnAsHTMLLink%>";
  pickerURL += lcl_showFolderStart;
  pickerURL += lcl_displayDocuments;
  pickerURL += lcl_displayActionLine;
  pickerURL += lcl_displayPayments;
  pickerURL += lcl_displayURL;

  eval('window.open("' + pickerURL + '", "_picker", "width=' + w + ',height=' + h + ',toolbar=0,statusbar=0,scrollbars=1,menubar=0,left=' + l + ',top=' + t + '")');
}

		 function storeCaret (textEl) 
		 {
		   if (textEl.createTextRange) 
			 textEl.caretPos = document.selection.createRange().duplicate();
		 }

		 function insertAtCaret (textEl, text) 
		 {
		   if (textEl.createTextRange && textEl.caretPos) {
			 var caretPos = textEl.caretPos;
			 caretPos.text =
			   caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
				 text + ' ' : text;
		   }
		   else
   <%
     if lcl_newstype = "NEWS" then
        response.write "textEl.value = textEl.value + text;" & vbcrlf
     else
        response.write "textEl.value = text;" & vbcrlf
     end if
   %>
		 }

		function fnCheckSubject()
		{
			if (document.NewEvent.Subject.value != '') {
				return true;
			}
			else
			{
				return false;
			}
		}

		function ValidateForm() 
		{
			var rege;
			var OK;
   var lcl_false_count = 0;
   var lcl_publicationStartDate_isValidDate = false;
			
			// validate the item title
			if (document.frmEditItem.itemtitle.value == "")
			{
//				alert("Please enter a title for this News Item.");
    inlineMsg(document.getElementById("itemtitle").id,'<strong>Required field missing: </strong> Title.',10,'itemtitle');
				document.frmEditItem.itemtitle.focus();
				return;
			}
<%
  'Only perform this check IF this is a News Scroller item.
   if lcl_newstype <> "NEWS" then
%>
			// validate the item date
			if (document.frmEditItem.itemdate.value == "")
			{
//				alert("Please enter a date for this News Item.");
    inlineMsg(document.getElementById("itemdate").id,'<strong>Required field missing: </strong> Date.',10,'itemdate');
				document.frmEditItem.itemdate.focus();
				return;
			}
			else
			{
				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				Ok = rege.test(document.frmEditItem.itemdate.value);
				if (! Ok)
				{
//					alert("The date should be in the format of MM/DD/YYYY.  \nPlease enter it again.");
     inlineMsg(document.getElementById("itemdate").id,'<strong>Invalid Value: </strong> The date should be in the form of MM/DD/YYYY.',10,'itemdate');
					document.frmEditItem.itemdate.focus();
					return;
				}
			}
<% end if %>
			// Validate the message text
			if (document.frmEditItem.itemtext.value == "")
			{
//				alert("Please enter a message for this News Item.");
    inlineMsg(document.getElementById("itemtext").id,'<strong>Required field missing: </strong> Message.',10,'itemmessage');
				document.frmEditItem.itemtext.focus();
				return;
			}
			else 
			{
				if (document.frmEditItem.itemtext.value.length > document.frmEditItem.itemtext.getAttribute('maxlength'))
				{
      lcl_msg_length = document.getElementById("itemtext").getAttribute('maxlength');
//					alert("The maxium length for the Message is " + document.frmEditItem.itemtext.getAttribute('maxlength') + " characters.\nPlease correct this and try saving again.");
     inlineMsg(document.getElementById("itemtext").id,'<strong>Invalid Value: </strong> Message.  The maxium length for the Message is ' + lcl_msg_length + ' characters.',10,'itemtext');
					document.frmEditItem.itemtext.focus();
					return;
				}
			}

			// check the publication start date
			if (document.getElementById("publicationstart").value != "")
			{
    if (! isValidDate(document.getElementById("publicationstart").value)) {
          document.getElementById("publicationstart").focus();
          inlineMsg(document.getElementById("datePicker_publicationStart").id,'<strong>Invalid Value: </strong> Publication Start Date',10,'datePicker_publicationStart');
          lcl_false_count = lcl_false_count + 1;
  		}else{
          lcl_publicationStartDate_isValidDate = true;
          clearMsg("datePicker_publicationStart");
    }
			}

			if (document.getElementById("publicationend").value != "")
			{
    if (! isValidDate(document.getElementById("publicationend").value)) {
       document.getElementById("publicationend").focus();
       inlineMsg(document.getElementById("datePicker_publicationEnd").id,'<strong>Invalid Value: </strong> Publication End Date',10,'datePicker_publicationEnd');
       lcl_false_count = lcl_false_count + 1;
  		}else{
    			if (document.getElementById("publicationstart").value != "")
    			{
          if(lcl_publicationStartDate_isValidDate)
          {
             var startDate = $('#publicationstart').val();
             var endDate   = $('#publicationend').val();

             var arrStartDate = startDate.split('/');
             var arrEndDate   = endDate.split('/');

             var lcl_error = 0;

             if(parseInt(arrStartDate[2]) > parseInt(arrEndDate[2]))  //year
             {
                lcl_error = 1;
             }
             else
             {
                if(parseInt(arrStartDate[2]) == parseInt(arrEndDate[2]))  //year
                {
                   if(parseInt(arrStartDate[0]) > parseInt(arrEndDate[0]))  //month
                   {
                      lcl_error = 2;
                   }
                   else
                   {
                      if(parseInt(arrStartDate[0]) == parseInt(arrEndDate[0]))  //month
                      {
                         if(parseInt(arrStartDate[1]) > parseInt(arrEndDate[1]))  //day
                         {
                            lcl_error = 3;
                         }
                      }
                   }
                }
             }

             if(lcl_error > 0)
             {
                inlineMsg(document.getElementById("datePicker_publicationStart").id,'<strong>Invalid Value: </strong> The Publication Start Date cannot be after the Publication End Date',10,'datePicker_publicationStart');
                document.getElementById("publicationstart").focus();
                lcl_false_count = lcl_false_count + 1;
             }
             else
             {
                clearMsg("datePicker_publicationEnd");
             }
          }
       }

       clearMsg("datePicker_publicationEnd");
    }
			}


			//if (document.frmEditItem.publicationstart.value != "")
				//rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
				//Ok = rege.test(document.frmEditItem.publicationstart.value);
				//if (! Ok)
				//{
//					alert("The Publication Start date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
    // inlineMsg(document.getElementById("publicationstart").id,'<strong>Invalid Value: </strong> The Publication Start Date should be in the format of MM/DD/YYYY or \"blank\".',10,'publicationstart');
				//	document.frmEditItem.publicationstart.focus();
				//	return;
				//}
   //}

 		// check the publication end date
//			if (document.frmEditItem.publicationend.value != "")
//			{
//				rege = /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/;
//				Ok = rege.test(document.frmEditItem.publicationend.value);
//				if (! Ok)
//				{
					//alert("The Publication End date should be in the format of MM/DD/YYYY, or be blank.  \nPlease enter it again.");
//     inlineMsg(document.getElementById("publicationend").id,'<strong>Invalid Value: </strong>  The Publication End Date should be in the format of MM/DD/YYYY or \"blank\".',10,'publicationend');
// 				document.frmEditItem.publicationend.focus();
//					return;
//				}
//			}

   if(lcl_false_count > 0) {
      return false;
   }else{
      //document.frmEditItem.submit();
   			document.getElementById("frmEditItem").submit();
      return true;
 		}

		}

<% if lcl_orghasfeature_rssfeeds_news AND lcl_userhaspermission_rssfeeds_news then %>
function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&isAjax=Y';

  doAjax('news_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}
<% end if %>

function displayScreenMsg(iMsg) {
  if(iMsg!="") {
     document.getElementById("screenMsg").innerHTML = "*** " + iMsg + " ***&nbsp;&nbsp;&nbsp;";
     window.setTimeout("clearScreenMsg()", (10 * 1000));
  }
}

function clearScreenMsg() {
  document.getElementById("screenMsg").innerHTML = "";
}

	//-->
	</script>

</head>

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

<body onload="setMaxLength();<%=lcl_onload%>">

<div id="content">
	<div id="centercontent">

 <table border="0" cellspacing="0" cellpadding="0" width="100%">
   <tr>
       <td><font size="+1"><strong><%=session("sOrgName")%>&nbsp;<%=lcl_page_label%>&nbsp<%=lcl_pagetitle%></strong></font></td>
       <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
   </tr>
 </table>

 <input type="button" name="backButton" id="backButton" value="Return to List" class="button" onclick="location.href='list_items.asp?newstype=<%=lcl_newstype%>';" />
	<p>
   <% displayButtons lcl_newsitemid, lcl_newstype %>
	</p>
		<!--BEGIN: New Item -->
<form name="frmEditItem" id="frmEditItem" action="save_item.asp" method="post">
  <input type="hidden" name="newsitemid" value="<%=request("newsitemid")%>" />
  <input type="hidden" name="newstype" id="newstype" value="<%=lcl_newstype%>" />
<%
 'Build a hidden field for the "itemdate" since it's hidden for CommunityLink News Items.
  if lcl_newstype = "NEWS" then
     response.write "<input type=""hidden"" name=""itemdate"" id=""itemdate"" size=""15"" maxlength=""10"" value=""" & lcl_itemdate & """ />" & vbcrlf
  end if
%>
<p>
<div class="shadow">
<table id="newsitemedit" border="0" cellpadding="6" cellspacing="0">
		<tr>
      <td>
					     <strong>Title:</strong>
          <input type="text" name="itemtitle" id="itemtitle" value="<%=lcl_itemtitle%>" maxlength="100" size="100" onchange="clearMsg('itemtitle');" />
  				</td>
  </tr>
<%
  'Only display this date if this is a News Scroller item.
   if lcl_newstype <> "NEWS" then
      response.write "		<tr>" & vbcrlf
      response.write "      <td>" & vbcrlf
      response.write "					     <strong>Date:</strong> (MM/DD/YYYY)" & vbcrlf
      response.write "          <input type=""text"" name=""itemdate"" id=""itemdate"" value=""" & lcl_itemdate & """ maxlength=""10"" size=""15"" onchange=""clearMsg('itemdate');"" />&nbsp;" & vbcrlf
      response.write "     					<img class=""calendarimg"" src=""../images/calendar.gif"" height=""16"" width=""16"" border=""0"" onclick=""javascript:void doCalendar('itemdate');"" />" & vbcrlf
      response.write "  				</td>" & vbcrlf
      response.write "  </tr>" & vbcrlf
   end if
%>
		<tr>
      <td>
          <strong>Message:</strong> &nbsp;* You May Use Simple HTML for formatting
      </td>
  </tr>
		<tr>
      <td>
        <%
          if lcl_newstype = "NEWS" then
             response.write "<div align=""right"">" & vbcrlf
             response.write "<input type=""button"" value=""Add a Link"" class=""button"" onClick=""doPicker('frmEditItem.itemtext','Y','Y','Y','Y');"" />" & vbcrlf
             response.write "</div>" & vbcrlf
          end if
        %>
					     <textarea name="itemtext" id="itemtext" rows="20" cols="120" maxlength="<%=lcl_messagelength%>" wrap="soft" onchange="clearMsg('itemtext');"><%=LEFT(lcl_itemtext,lcl_messagelength)%></textarea>
  				</td>
  </tr>
<% if lcl_newstype <> "NEWS" then %>
		<tr>
      <td>
					     <strong>Link URL:</strong>
          <input type="text" name="itemlinkurl" id="itemlinkurl" value="<%=lcl_itemlinkurl%>" maxlength="500" size="50" />
					     <input type="button" value="Add a Link" class="button" onClick="doPicker('frmEditItem.itemlinkurl','Y','Y','Y','Y');" />
					     <!--<input type="button" value="Find a Link" class="button" onClick="doSitePicker('frmEditItem.itemlinkurl');" />-->
  				</td>
  </tr>
<% end if %>
		<tr>
      <td>
          <strong>Publication Start: </strong>
          <input type="text" name="publicationstart" id="publicationstart" value="<%=lcl_publicationstart%>" onchange="clearMsg('datePicker_publicationStart');" />
      				<span class="calendarimg" style="cursor:pointer;"><img id="datePicker_publicationStart" src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationstart');" /></span>
  				</td>
  </tr>
		<tr>
      <td>
          <strong>Publication End: </strong>
          <input type="text" name="publicationend" id="publicationend" value="<%=lcl_publicationend%>" onchange="clearMsg('datePicker_publicationStart');clearMsg('datePicker_publicationEnd');" />
      				<span class="calendarimg" style="cursor:pointer;"><img id="datePicker_publicationEnd" src="../images/calendar.gif" height="16" width="16" border="0" onclick="javascript:void doCalendar('publicationend');" /></span>
  				</td>
  </tr>
		<tr>
      <td>
          <strong>Display: </strong>
          <input type="checkbox" name="itemdisplay" value="on"<%=lcl_checked_itemdisplay%> />
  				</td>
  </tr>
<%
  if lcl_orghasfeature_rssfeeds_news AND lcl_userhaspermission_rssfeeds_news then
     response.write "  <tr valign=""top"">" & vbcrlf
     response.write "      <td nowrap=""nowrap"">" & vbcrlf
     response.write "<strong>On " & lcl_sendToLabel & " Send To:</strong>" & vbcrlf

                     displaySendToOption "RSS", _
                                         lcl_screen_mode, _
                                         "N", _
                                         lcl_orghasfeature_rssfeeds_news, _
                                         lcl_userhaspermission_rssfeeds_news
     response.write "      </td>" & vbcrlf
     response.write "  </tr>" & vbcrlf
  end if
%>
</table>
</div>
</p>
<p>
   <% displayButtons lcl_newsitemid, lcl_newstype %>
</p>
</form>
<!--END: news item -->
	</div>
</div>
	
	<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub displayButtons(p_newsitemid, p_newstype)

  if p_newsitemid > 0 then
     lcl_button_label = "Update This News Item"
  else
     lcl_button_label = "Create News Item"
  end if

  response.write "<input type=""button"" class=""button"" value=""" & lcl_button_label & """ onclick=""ValidateForm();"" />" & vbcrlf

end sub
%>
