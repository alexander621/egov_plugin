<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<!-- #include file="faq_global_functions.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: list_faq.asp
' AUTHOR: ??
' CREATED: ??
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  This module lists the Frequently Asked Questions (FAQ).
'
' MODIFICATION HISTORY
' 1.1 09/11/06 Steve Loar - Changes for categories.
' 1.2 10/09/06 Steve Loar - Security, Header and Nav changed
' 1.3 03/20/09 David Boyer - Added "faqtype" for the new "Rumor Mill" data
' 1.4 06/10/09	David Boyer - Added checkbox for "send to" function.  (Send to features like RSS and eventually Twitter, etc.)
' 1.5 07/23/09 David Boyer - Changed the "Ask a Question" dropdown list to rely on Action Line and not CommunityLink feature.
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
 Dim blnCanEditForms 
 sLevel = "../"  'Override of value from common.asp

'Check for the faqtype
 if request("faqtype") <> "" then
    lcl_faqtype = UCASE(request("faqtype"))
 else
    lcl_faqtype = "FAQ"
 end if

'Used in custom reporting
 session("RSSType") = lcl_faqtype

'Based on the faqtype check for the proper permission
 if lcl_faqtype = "RUMORMILL" then
    lcl_feature          = "rumormill"
    lcl_userpermission   = "rumormill_manage"
    lcl_feature_rssfeeds = "rssfeeds_rumormill"
    lcl_pagetitle        = "Rumor Mill"
    lcl_postCommentLabel = "Submit a Rumor"
    lcl_postCommentRule  = "maintainactionlineform_submitarumor"
 else
    lcl_feature          = "faq"
    lcl_userpermission   = "manage faq"
    lcl_feature_rssfeeds = "rssfeeds_faqs"
    lcl_pagetitle        = "FAQ"
    lcl_postCommentLabel = "Ask a Question"
    lcl_postCommentRule  = "maintainactionlineform_askaquestion"
 end if

 if not userhaspermission(session("userid"),lcl_userpermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if

 blnCanEditForms = True 

'Check for a screen message
 lcl_onload  = ""
 lcl_success = request("success")

 if lcl_success <> "" then
    lcl_msg    = setupScreenMsg(lcl_success)
    lcl_onload = "displayScreenMsg('" & lcl_msg & "');"
 end if

'Check for org features
 lcl_orghasfeature_action_line     = orghasfeature("action line")
 lcl_orghasfeature_rssfeeds_faqs   = orghasfeature(lcl_feature_rssfeeds)
 lcl_orghasfeature_postCommentRule = orghasfeature(lcl_postCommentRule)
 'lcl_orghasfeature_communitylink = orghasfeature("communitylink")

'Check for user permissions
 lcl_userhaspermission_action_line     = userhaspermission(session("userid"),"action line")
 lcl_userhaspermission_rssfeeds_faqs   = userhaspermission(session("userid"),lcl_feature_rssfeeds)
 lcl_userhaspermission_postCommentRule = userhaspermission(session("userid"),lcl_postCommentRule)
 'lcl_userhaspermission_communitylink = userhaspermission(session("userid"),"communitylink")

'Determine if there is any additional processing needed from the past update
 if lcl_orghasfeature_rssfeeds_faqs AND lcl_userhaspermission_rssfeeds_faqs AND (lcl_success = "SU" OR lcl_success = "SA") then
    if request("sendTo_RSS") <> "" then
       lcl_onload = lcl_onload & "sendToRSS('" & request("sendTo_RSS") & "');"
    end if
 end if

'Build BODY onload
 lcl_onload = lcl_onload & "enableDisableLabel();"
%>
<html>
<head>
 	<title>E-Gov Administration Console {<%=lcl_pagetitle%>}</title>

	 <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	 <link rel="stylesheet" type="text/css" href="../global.css" />
  <link rel="stylesheet" type="text/css" href="../custom/css/tooltip.css" />

  <script language="javascript" src="../scripts/modules.js"></script>
 	<script language="javascript" src="../scripts/ajaxLib.js"></script>
  <script language="javascript" src="../scripts/tooltip_new.js"></script>

<script language="javascript">
<!--
function viewRSSLog(pID) {
  lcl_width  = 900;
  lcl_height = 400;
  lcl_left   = (screen.availWidth/2) - (lcl_width/2);
  lcl_top    = (screen.availHeight/2) - (lcl_height/2);
		popupWin = window.open("../customreports/customreports.asp?CR=RSSLOG&id=" + pID, "_blank","resizable,width=" + lcl_width + ",height=" + lcl_height + ",left=" + lcl_left + ",top=" + lcl_top);
}

function confirm_delete(ifaqid) {
  lcl_faq = document.getElementById("faq"+ifaqid).innerHTML;

 	if (confirm("Are you sure you want to delete " + lcl_faq + " ?")) { 
  				//DELETE HAS BEEN VERIFIED
		  		location.href='delete_faq.asp?faqtype=<%=lcl_faqtype%>&ifaqid='+ ifaqid;
		}
}

function sendToRSS(pID) {
  var sParameter = 'id=' + encodeURIComponent(pID);
  sParameter    += '&faqtype=<%=lcl_faqtype%>';
  sParameter    += '&isAjax=Y';

  doAjax('faq_sendToRSS.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function updatePostComment(p_field) {
  var lcl_value = '';
  var lcl_param = 'formid';

  //Determine which value we are working with
  if(p_field == "LABEL") {
     lcl_fieldid = "CL_postcomment_label";
  }else{
     lcl_fieldid = "CL_postcomment_formid";
  }

  lcl_value = document.getElementById(lcl_fieldid).value;
  lcl_param = p_field;

  //Build the parameter string
  var sParameter = 'orgid='     + encodeURIComponent("<%=session("orgid")%>");
  sParameter    += '&feature='  + encodeURIComponent("<%=lcl_feature%>");
  sParameter    += '&savetype=' + encodeURIComponent(p_field);
  sParameter    += '&' + lcl_param + '=' + encodeURIComponent(lcl_value);
  sParameter    += '&isAjaxRoutine=Y';

  doAjax('../communitylink/saveCommunityLinkOptions.asp', sParameter, 'displayScreenMsg', 'post', '0');
}

function enableDisableLabel() {
  lcl_value = document.getElementById("CL_postcomment_formid").value;

  if(lcl_value == "") {
     lcl_disabled = true;
  }else{
     lcl_disabled = false;
  }

  document.getElementById('CL_postcomment_label').disabled = lcl_disabled;

}

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
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0" onload="<%=lcl_onload%>">

<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"--> 

<div id="content">
 	<div id="centercontent">

<table border="0" cellpadding="6" cellspacing="0" class="start" width="100%">
  <tr>
      <td valign="top">
          <div style="margin-top:20px; margin-left:20px;">
            <table border="0" cellspacing="0" cellpadding="0" width="600">
              <tr>
                  <td><font size="+1"><strong><%=session("sOrgName")%>&nbsp;<%=lcl_pagetitle%>s</strong></font></td>
                  <td align="right"><span id="screenMsg" style="color:#ff0000; font-size:10pt; font-weight:bold;"></span></td>
              </tr>
            </table>
            <table border="0" cellspacing="0" cellpadding="0" style="margin-top:5px; margin-bottom:5px;">
              <tr>
                  <td>
                      <input type="button" class="button" value="New <%=lcl_pagetitle%>" onclick="window.location='manage_faq.asp?faqtype=<%=lcl_faqtype%>';" />&nbsp;
                      <input type="button" class="button" value="<%=lcl_pagetitle%> Categories" onclick="window.location='faq_categories.asp?faqtype=<%=lcl_faqtype%>';" />&nbsp;
                  </td>
                  <td align="right">
                  <%
                   'Ask a question/Submit a Rumor - Action Line form
                    if lcl_orghasfeature_action_line AND lcl_orghasfeature_postCommentRule then
                    'if lcl_orghasfeature_communitylink AND lcl_userhaspermission_communitylink then
                       lcl_feature_label   = GetFeatureName(lcl_feature)
                       lcl_comments_formid = getCommentsFormID(session("orgid"), "", lcl_feature)
                       lcl_comments_label  = getCommentsLabel(session("orgid"), "", lcl_feature)

                       response.write "<span style=""color:#800000"">" & vbcrlf
                       'response.write "CommunityLink: Link for """ & lcl_postCommentLabel & """ (" & lcl_feature_label & "):" & vbcrlf
                       response.write "Action Line Request to be used for """ & lcl_postCommentLabel & """" & vbcrlf
                       response.write "</span><br />" & vbcrlf

                      'If the user has the proper permission then allow then to maintain the list.
                      'If not then display the form selected.
                       if lcl_userhaspermission_postCommentRule then
                          response.write "<select name=""CL_postcomment_formid"" id=""CL_postcomment_formid"" onchange=""updatePostComment('FORMID');enableDisableLabel();"">" & vbcrlf
                                            displayActionLineForms session("orgid"), lcl_comments_formid, "Y"
                          response.write "</select>" & vbcrlf
                          response.write "<br />" & vbcrlf
                          response.write "<strong>Button Label: </strong>" & vbcrlf
                          response.write "<input type=""text"" name=""CL_postcomment_label"" id=""CL_postcomment_label"" size=""20"" maxlength=""50"" value=""" & lcl_comments_label & """ />" & vbcrlf
                          response.write "<input type=""button"" name=""saveLabelButton"" id=""saveLabelButton"" value=""Save Label"" class=""button"" onclick=""updatePostComment('LABEL');"" />" & vbcrlf
                       else
                          lcl_actionline_formname = getActionLineFormName(session("orgid"), lcl_comments_formid)


                          if lcl_actionline_formname <> "" then
                             response.write "[" & lcl_actionline_formname & "]<br />" & vbcrlf
                          end if

                          if lcl_comments_label <> "" then
                             response.write "<strong>Button Label: </strong>" & lcl_comments_label & vbcrlf
                          end if
                       end if
                    else
                       response.write "&nbsp;" & vbcrlf
                    end if
                  %>
                  </td>
              </tr>
            </table>
            <% subListFaqs lcl_faqtype %>
          </div>
      </td>
  </tr>
</table>

  </div>
</div>
	
<!--#Include file="../admin_footer.asp"--> 

</body>
</html>
<%
'------------------------------------------------------------------------------
sub subListFaqs(iFAQType)
	Dim iRowCount

	iOrgID = session("orgid")

	if iOrgID = "" then
  		iOrgID = -1
	end if

 if iFAQType = "" then
    iFAQType = "FAQ"
 end if

	iRowCount              = 0
 lcl_category_faq_count = 0

	sSQL = "SELECT FAQ.FaqID, FAQ.FaqQ, isnull(FAQ.faqcategoryid,0) as faqcategoryid, FAQ.sequence, "
 sSQL = sSQL & " isnull(faqcategoryname,'') as faqcategoryname, isnull(internalonly,0) as internalonly "
 sSQL = sSQL & " FROM FAQ LEFT OUTER JOIN faq_categories C ON C.faqcategoryid = faq.faqcategoryid "
	sSQL = sSQL & " WHERE faq.orgid = " & iOrgID
 sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & iFAQType & "' "
 sSQL = sSQL & " ORDER BY internalonly, displayorder, sequence"

	set oFaqList = Server.CreateObject("ADODB.Recordset")
	oFaqList.Open sSQL, Application("DSN"), 3, 1
	
	if not oFaqList.eof then
  		response.write "<div class=""shadow"">" & vbcrlf
		  response.write "<table cellspacing=""0"" cellpadding=""2"" class=""tablelist"" border=""0"">" & vbcrlf
  		response.write "  <tr>" & vbcrlf
    response.write "      <th align=""left"">&nbsp;Question</th>" & vbcrlf
    response.write "      <th align=""left"">Category</th>" & vbcrlf
    response.write "      <th>Sequence<br />Within<br />Category</th>" & vbcrlf
    response.write "      <th colspan=""2"">&nbsp;</th>" & vbcrlf

    if lcl_orghasfeature_rssfeeds_faqs AND lcl_userhaspermission_rssfeeds_faqs then
       response.write "      <th>Send to<br />RSS</th>" & vbcrlf
       response.write "      <th>RSS Send<br />Log</th>" & vbcrlf
    end if

    response.write "  </tr>" & vbcrlf

    lcl_bgcolor             = "#ffffff"
    lcl_original_categoryid = 0

    do while not oFaqList.eof
       lcl_bgcolor = changeBGColor(lcl_bgcolor,"#eeeeee","#ffffff")

       if isnull(oFaqList("FaqQ")) OR oFaqList("FaqQ")="" then
          iFaqQ = ""
       elseif len(oFaqList("FaqQ")) > 50 then
          iFaqQ = left(oFaqList("FaqQ"),40) & "..."
       else
          iFaqQ = oFaqList("FaqQ")
       end if

    			if oFaqList("internalonly") then
				      lcl_internal_label = "<br />(Internal)"
       else
          lcl_internal_label = ""
    			end if

    			iRowCount    = iRowCount + 1
       iMaxSequence = getMaxSequence(iOrgID,oFaqList("faqcategoryid"))

      'If the category of this row does NOT equal the category of the previous row then show a border line.
       if oFaqList("faqcategoryid") <> lcl_categoryid_original then
          lcl_lastrow_style = " style=""border-top:1pt solid #000000;"""
       else
          lcl_lastrow_style = ""
       end if

      'Check if we need to show the Move Up button
       if oFaqList("sequence") > 1 AND oFaqList("sequence") <= iMaxSequence AND oFaqList("faqcategoryid") = lcl_categoryid_original then
          lcl_move_up = "<img src=""../images/ieup.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move UP');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_faq.asp?direction=UP&ifaqid=" & oFaqList("FaqID") & "&orgid=" & iOrgID & "&faqcategoryid=" & oFaqList("faqcategoryid") & "&faqtype=" & iFAQType & "'"" /><br />" & vbcrlf
       else
          lcl_move_up = ""
       end if

      'Check if we need to show the Move Down button
       if oFaqList("sequence") < iMaxSequence then
          lcl_move_down = "<img src=""../images/iedown.gif"" align=""absmiddle"" border=""0"" class=""hotspot"" onmouseover=""tooltip.show('Move DOWN');"" onmouseout=""tooltip.hide();"" onclick=""location.href='order_faq.asp?direction=DOWN&ifaqid=" & oFaqList("FaqID") & "&orgid=" & iOrgID & "&faqcategoryid=" & oFaqList("faqcategoryid") & "&faqtype=" & iFAQType & "'"" />" & vbcrlf
       else
          lcl_move_down = ""
       end if

      'If neither button is shown we need to put in a "space" so that the border works properly if this record just happens to be
      'the last record for the category.
       if lcl_move_down = "" AND lcl_move_up = "" then
          lcl_move_down = "&nbsp;"
       end if

       response.write "  <tr id=""" & iRowCount & """ bgcolor=""" & lcl_bgcolor & """ onMouseOver=""mouseOverRow(this);"" onMouseOut=""mouseOutRow(this);"">" & vbcrlf
       response.write "      <td class=""formlist""" & lcl_lastrow_style & " title=""click to edit"" onClick=""location.href='manage_faq.asp?ifaqid=" & oFaqList("FaqID") & "&faqtype=" & iFAQType & "';"">&nbsp;<span id=""faq" & oFaqList("FaqID") & """>" & iFaqQ & "</span></td>" & vbcrlf
       response.write "      <td class=""formlist""" & lcl_lastrow_style & " title=""click to edit"" onClick=""location.href='manage_faq.asp?ifaqid=" & oFaqList("FaqID") & "&faqtype=" & iFAQType & "';"">" & oFaqList("faqcategoryname") & lcl_internal_label & "</td>" & vbcrlf
       response.write "      <td class=""formlist""" & lcl_lastrow_style & " title=""click to edit"" onClick=""location.href='manage_faq.asp?ifaqid=" & oFaqList("FaqID") & "&faqtype=" & iFAQType & "';"" align=""center"">(" & oFaqList("sequence") & ")</td>" & vbcrlf
       response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center"">" & lcl_move_up & lcl_move_down & "</td>" & vbcrlf
       response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center""><input type=""button"" name=""delete" & iRowCount & """ id=""delete"   & iRowCount & """ value=""Delete"" class=""button"" onclick=""confirm_delete('" & oFaqList("FaqID") & "');"" /></td>" & vbcrlf

       if lcl_orghasfeature_rssfeeds_faqs AND lcl_userhaspermission_rssfeeds_faqs then
          if oFaqList("internalonly") = 0 then
             response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center""><input type=""button"" name=""sendToRSS" & iRowCount & """ id=""sendToRSS"   & iRowCount & """ value=""Send"" class=""button"" onclick=""sendToRSS('" & oFaqList("FaqID") & "');"" /></td>" & vbcrlf

            'Check to see if a log exists for this row
             if checkRSSLogExists(session("orgid"),oFaqList("FaqID"),iFAQType) then
                response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center""><input type=""button"" name=""viewRSSLog" & iRowCount & """ id=""viewRSSLog" & iRowCount & """ value=""View"" class=""button"" onclick=""viewRSSLog('" & oFaqList("FaqID") & "');"" /></td>" & vbcrlf
             else
                response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center"">&nbsp;</td>" & vbcrlf
             end if
          else
             response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center"">&nbsp;</td>" & vbcrlf
             response.write "      <td class=""formlist""" & lcl_lastrow_style & " align=""center"">&nbsp;</td>" & vbcrlf
          end if
       end if

       response.write "  </tr>"  & vbcrlf

       lcl_categoryid_original = oFaqList("faqcategoryid")

       oFaqList.movenext
   loop

 		response.write "</table>" & vbcrlf
	  response.write "</div>" & vbcrlf

 else
  		response.write "<p style=""padding-top:10px; color:#ff0000; font-weight:bold;"">No " & lcl_pagetitle & "s have been created.</p>" & vbcrlf
	end if

	oFaqList.close
	set oFaqList = nothing 

End Sub

'------------------------------------------------------------------------------
function getMaxSequence(p_orgid,iCategoryID)
	 Dim sSql, oMax
  lcl_return = 0

 	sSQL = "SELECT MAX(FAQ.sequence) as MaxSequence "
  sSQL = sSQL & " FROM FAQ LEFT OUTER JOIN faq_categories C "
  sSQL = sSQL &      " ON C.faqcategoryid = faq.faqcategoryid "
  sSQL = sSQL &      " AND C.faqtype = faq.faqtype "
 	sSQL = sSQL & " WHERE faq.orgid = "    & p_orgid

  if iCategoryID = 0 then
     sSQL = sSQL & " AND faq.faqcategoryid is null "
  else
     sSQL = sSQL & " AND faq.faqcategoryid = " & iCategoryID
  end if

  sSQL = sSQL & " AND UPPER(faq.faqtype) = '" & lcl_faqtype & "'"

	 set oMax = Server.CreateObject("ADODB.Recordset")
	 oMax.Open sSQL, Application("DSN"), 3, 1

 	if isnull(oMax("MaxSequence")) then
	   	lcl_return = 0
 	else
	   	lcl_return = oMax("MaxSequence")
 	end if

 	oMax.close
	 set oMax = nothing

  getMaxSequence = lcl_return

end function
%>