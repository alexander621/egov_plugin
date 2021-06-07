<!DOCTYPE html>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="news_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: list_items.asp
' AUTHOR: Steve Loar
' CREATED: 10/31/2006
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists the News Scroller Items.
'
' MODIFICATION HISTORY
' 1.0 10/31/06	Steve Loar - Initial Version Created
' 1.1 05/07/09 David Boyer - Added screen "success" message
' 1.2 05/08/09 David Boyer - Modified screen: links to buttons.
' 1.3 07/09/09 David Boyer - Added "newstype" to split News and News Scroller items.
' 1.4 07/23/09 David Boyer - Changed the "Suggest a News Item" dropdown list to rely on Action Line and not CommunityLink feature.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim blnCanEditForms 
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
    lcl_feature                 = "news_items"
    lcl_userpermission_required = "edit_news_items"
    lcl_pagetitle               = "News Item"
    session("RSSType")          = "NEWS"  'Used in custom reporting
    lcl_postCommentLabel        = "Suggest a News Item"
 else
    lcl_feature                 = "news scroller"
    lcl_userpermission_required = "edit scroller"
    lcl_pagetitle               = "News Scroller Item"
    session("RSSType")          = ""
    lcl_postCommentLabel        = ""
 end if

 if not userhaspermission(session("userid"),lcl_userpermission_required) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 blnCanEditForms = True 
 iMaxOrder       = GetMaxItemOrder(session("orgid"))

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_rssfeeds_news                          = orghasfeature("rssfeeds_news")
 lcl_orghasfeature_action_line                            = orghasfeature("action line")
 lcl_orghasfeature_maintainactionlineform_suggestnewsitem = orghasfeature("maintainactionlineform_suggestnewsitem")
 'lcl_orghasfeature_communitylink = orghasfeature("communitylink")

'Check for user permissions
 lcl_userhaspermission_rssfeeds_news                          = userhaspermission(session("userid"),"rssfeeds_news")
 lcl_userhaspermission_action_line                            = userhaspermission(session("userid"),"action line")
 lcl_userhaspermission_maintainactionlineform_suggestnewsitem = userhaspermission(session("userid"),"maintainactionlineform_suggestnewsitem")
 'lcl_userhaspermission_communitylink = userhaspermission(session("userid"),"communitylink")
%>
<html>
<head>

	<title>E-Gov Administration Console {<%=lcl_pagetitle%>s}</title>

	<link rel="stylesheet" href="../global.css" />
	<link rel="stylesheet" href="../menu/menu_scripts/menu.css" />
 <link rel="stylesheet" href="../custom/css/tooltip.css" />

	<script src="../scripts/ajaxLib.js"></script>
 <script src="../scripts/tooltip_new.js"></script>

<script>
<!--
function mouseOverRow( oRow ) {
  oRow.style.backgroundColor='#93bee1';
		oRow.style.cursor='pointer';
}

function mouseOutRow( oRow ) {	
  oRow.style.backgroundColor='';
		oRow.style.cursor='';
}

function confirm_delete(iItemId) {
  lcl_name = document.getElementById("itemtitle_" + iItemId).innerHTML; 
		if (confirm("Are you sure you want to delete " + lcl_name + "?")) { 
  				location.href='delete_item.asp?newsitemid=' + iItemId + '&newstype=<%=lcl_newstype%>';
		}
}

function ChangeOrder(newsitemid,itemorder,iDirection) {
  location.href='order_items.asp?newstype=<%=lcl_newstype%>&itemorder='+ itemorder + '&newsitemid=' + newsitemid + '&iDirection=' + iDirection;
}

function ChangeDisplay( newsitemid, itemdisplay ) {
  location.href='display_items.aspx?newstype=<%=lcl_newstype%>&newsitemid=' + newsitemid + '&itemdisplay=' + itemdisplay;
}
<% if lcl_newstype = "NEWS" then %>
function viewRSSLog(pID) {
  lcl_width  = 900;
  lcl_height = 500;
  lcl_left   = (screen.availWidth/2) - (lcl_width/2);
  lcl_top    = (screen.availHeight/2) - (lcl_height/2);
		popupWin = window.open("../customreports/customreports.asp?CR=RSSLOG&id=" + pID, "_blank","scrollbars=1,resizable=1,width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top);
}

function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&isAjax=Y';

  doAjax('news_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function updatePostComment() {
  lcl_formid = document.getElementById("CL_postcomment_formid").value;

  //Build the parameter string
  var sParameter = 'orgid='     + encodeURIComponent("<%=session("orgid")%>");
  sParameter    += '&newstype=' + encodeURIComponent("<%=lcl_newstype%>");
  sParameter    += '&feature='  + encodeURIComponent("<%=lcl_feature%>");
  sParameter    += '&formid='   + encodeURIComponent(lcl_formid);
  sParameter    += '&isAjaxRoutine=Y';

  doAjax('../communitylink/saveCommunityLinkOptions.asp', sParameter, 'displayScreenMsg', 'post', '0');
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
<body onload="<%=lcl_onload%>">

	<% ShowHeader sLevel %>
	<!--#Include file="../menu/menu.asp"--> 

	<div id="content">
		<div id="centercontent">

<table border="0" cellspacing="0" cellpadding="0" width="100%">
  <tr>
      <td><font size="+1"><strong><%=session("sOrgName")%>&nbsp;<%=lcl_pagetitle%>s</strong></font></td>
      <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" style="margin-top:5px; margin-bottom:5px;">
  <tr>
      <td>
          <input type="button" name="createButton" id="createButton" value="Create a <%=lcl_pagetitle%>" class="button" onclick="location.href='edit_item.asp?newstype=<%=lcl_newstype%>';" />
      </td>
      <td align="right">
      <%
       'Post a Comment - Action Line form
       'This is ONLY a feature for items where newstype is NEWS
        if lcl_newstype = "NEWS" AND lcl_orghasfeature_action_line AND lcl_orghasfeature_maintainactionlineform_suggestnewsitem then
        'if lcl_newstype = "NEWS" AND lcl_orghasfeature_communitylink AND lcl_userhaspermission_communitylink then
           lcl_feature_label   = GetFeatureName(lcl_feature)
           lcl_comments_formid = getCommentsFormID(session("orgid"), "", lcl_feature)

           response.write "<span style=""color:#800000"">" & vbcrlf
           'response.write "CommunityLink: Link for """ & lcl_postCommentLabel & """ (" & lcl_feature_label & "):" & vbcrlf
           response.write "Action Line Request to be used for """ & lcl_postCommentLabel & """" & vbcrlf
           response.write "</span><br />" & vbcrlf

          'If the user has the proper permission then allow then to maintain the list.
          'If not then display the form selected.
           if lcl_userhaspermission_maintainactionlineform_suggestnewsitem then
              response.write "<select name=""CL_postcomment_formid"" id=""CL_postcomment_formid"" onchange=""updatePostComment();"">" & vbcrlf
                                displayActionLineForms session("orgid"), lcl_comments_formid, "Y"
              response.write "</select>" & vbcrlf
           else
              lcl_actionline_formname = getActionLineFormName(session("orgid"), lcl_comments_formid)

              response.write "[" & lcl_actionline_formname & "]" & vbcrlf
           end if
        else
           response.write "&nbsp;" & vbcrlf
        end if
      %>
      </td>
  </tr>
</table>
<p><% ListItems session("orgid"), lcl_newstype, lcl_orghasfeature_rssfeeds_news, lcl_userhaspermission_rssfeeds_news %></p>

		</div>
	</div>
	
	<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub ListItems(iOrgID, iNewsType, iOrgHasFeature_RSSFeeds_News, iUserHasPermission_RSSFeeds_News)
 	Dim sSql, oItems, iRowCount

 	sSQL = "SELECT newsitemid, orgid, itemtitle, itemdate, itemtext, itemlinkurl, "
  sSQL = sSQL & " itemdisplay, itemorder, publicationstart, publicationend "
 	sSQL = sSQL & " FROM egov_news_items "
  sSQL = sSQL & " WHERE orgid = " & iOrgID
  sSQL = sSQL & " AND UPPER(newstype) = '" & UCASE(iNewsType) & "' "
  sSQL = sSQL & " ORDER BY itemorder "

	 set oItems = Server.CreateObject("ADODB.Recordset")
	 oItems.Open sSQL, Application("DSN"), 3, 1
	
	 if not oItems.eof then
   		iRowCount = 0

     if lcl_newstype = "SCROLLER" then
        lcl_colspan = "2"
     else
        lcl_colspan = "1"
     end if

   		response.write "<div class=""shadow"">" & vbcrlf
   		response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"">" & vbcrlf
   		response.write "  <tr>" & vbcrlf
   		response.write "      <th>Title</th>"       & vbcrlf
   		response.write "      <th>Publication</th>" & vbcrlf
   		response.write "      <th>Display</th>"     & vbcrlf
   		response.write "      <th colspan=""" & lcl_colspan & """>&nbsp;</th>" & vbcrlf

     if UCASE(iNewsType) = "NEWS" AND iOrgHasFeature_RSSFeeds_News AND iUserHasPermission_RSSFeeds_News then
        response.write "      <th>Send to<br />RSS</th>"  & vbcrlf
        response.write "      <th>RSS Send<br />Log</th>" & vbcrlf
     end if

   		response.write "  </tr>" & vbcrlf

     lcl_bgcolor = "#ffffff"

   		do while not oItems.eof
  	   		iRowCount   = iRowCount + 1
        lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

       'Check if we need to show the Move Up button
		     	if iRowCount <> 1 then
           lcl_move_up = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & oItems("newsitemid") & "," & oItems("itemorder") & ",-1);"" /><br />" & vbcrlf
        else
           lcl_move_up = ""
        end if

       'Check if we need to show the Move Down button
		     	if oItems("itemorder") <> iMaxOrder then
           lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""ChangeOrder(" & oItems("newsitemid") & "," & oItems("itemorder") & ",1);"" />" & vbcrlf
        else
           lcl_move_down = ""
        end if

       'If neither button is shown we need to put in a "space" so that the border works properly if this record just happens to be
       'the last record for the category.
        if lcl_move_down = "" AND lcl_move_up = "" then
           lcl_move_down = "&nbsp;"
        end if

       'Setup ItemDisplay
  						if oItems("itemdisplay") then
       			 lcl_checked_itemdisplay = " checked=""checked"""
  			   else
       			 lcl_checked_itemdisplay = ""
  						end if

  						if oItems("itemdisplay") then
  			  				iItemDisplay = 1
  						else
		  	  				iItemDisplay = 0
  						end if

       'Build the publication start and/or end date(s)
        lcl_display_publicationdate = ""

  						if not isnull(oItems("publicationstart")) then
		  	  				lcl_display_publicationdate = oItems("publicationstart")
  			   end if

  						if not isnull(oItems("publicationend")) then
  			  				lcl_display_publicationdate = lcl_display_publicationdate & " until " & oItems("publicationend")
  			   end if

       'Setup row onclick to edit a news item
        lcl_edit_onclick = " onClick=""location.href='edit_item.asp?newstype=" & lcl_newstype & "&newsitemid=" & oItems("newsitemid") & "';"""

     			response.write "  <tr id=""" & iRowCount & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"" bgcolor=""" & lcl_bgcolor & """>" & vbcrlf
     			response.write "      <td class=""formlist"" title=""click to edit""" & lcl_edit_onclick & ">" & vbcrlf
  			   response.write "          &nbsp;<span id=""itemtitle_" & oItems("newsitemid") & """>" & oItems("itemtitle") & "</span>" & vbcrlf
  			   response.write "      </td>" & vbcrlf
  						response.write "      <td align=""center"" title=""click to edit""" & lcl_edit_onclick & ">" & lcl_display_publicationdate & "</td>" & vbcrlf
								response.write "      <td align=""center""><input type=""checkbox"" name=""itemdisplay"" value=""" & oItems("newsitemid") & """ onclick=""ChangeDisplay(" & oItems("newsitemid") & "," & iItemDisplay & ");""" & lcl_checked_itemdisplay & " /></td>" & vbcrlf

       'Only News SCROLLER items can be reorderd.
        if lcl_newstype = "SCROLLER" then
   					   response.write "      <td>" & lcl_move_up & lcl_move_down & "</td>" & vbcrlf
        end if

					   response.write "      <td><input type=""button"" name=""deleteButton" & oItems("newsitemid") & """ id=""deleteButton" & oItems("newsitemid") & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oItems("newsitemid") & "');"" /></td>" & vbcrlf

					   if UCASE(iNewsType) = "NEWS" AND iOrgHasFeature_RSSFeeds_News AND iUserHasPermission_RSSFeeds_News then
     					 response.write "      <td align=""center""><input type=""button"" name=""sendToRSS" & iRowCount & """ id=""sendToRSS"   & iRowCount & """ value=""Send"" class=""button"" onclick=""sendToRSS('" & oItems("newsitemid") & "');"" /></td>" & vbcrlf

					     'Check to see if a log exists for this row
					      if checkRSSLogExists(iOrgID,oItems("newsitemid"),"NEWS") then
     					    response.write "      <td align=""center""><input type=""button"" name=""viewRSSLog" & iRowCount & """ id=""viewRSSLog" & iRowCount & """ value=""View"" class=""button"" onclick=""viewRSSLog('" & oItems("newsitemid") & "');"" /></td>" & vbcrlf
					      else
     					    response.write "      <td align=""center"">&nbsp;</td>" & vbcrlf
					      end if
					   end if

								response.write "  </tr>" & vbcrlf

								oItems.movenext
		 		loop

   		response.write "</table>" & vbcrlf
   		response.write "</div>" & vbcrlf
  else
   		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No News Items have been created.</p>" & vbcrlf
  end if

 	oItems.close
	 set oItems = nothing 

end sub

'------------------------------------------------------------------------------
function GetMaxItemOrder(iOrgID)
 	dim sSql, oMax

 	sSQL = "SELECT max(itemorder) as maxOrder "
  sSQL = sSQL & " FROM egov_news_items "
  sSQL = sSQL & " WHERE OrgID = " & iOrgID

 	set oMax = Server.CreateObject("ADODB.Recordset")
 	oMax.Open sSQL, Application("DSN"), adOpenStatic, adLockReadOnly

 	if IsNull(oMax("MaxOrder")) then
 		  GetMaxItemOrder = 0
 	else
 		  GetMaxItemOrder = oMax("MaxOrder")
 	end if

 	oMax.close
 	set oMax = nothing
end function
%>