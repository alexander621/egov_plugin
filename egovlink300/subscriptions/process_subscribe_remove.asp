<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/start_modules.asp" //-->
<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: subscribe_remove.asp
' AUTHOR: Steve Loar
' CREATED: 09/07/06
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Subscription removal.  Called from email link sent to subscriber.
'
' MODIFICATION HISTORY
' 1.0 09/07/06 Steve Loar - Initial version
' 1.1 10/05/09 David Boyer - Modified code to only remove a specific distribution list instead of all of them.
' 1.2  11/19/13  Terry Foster - CLng Bug Fix
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

	Dim iUserid, iDListID

	if request("u") <> "" AND request("dl") <> "" and isnumeric(replace(request("u"),"'","")) and isnumeric(replace(request("dl"),"'","")) then
	on error resume next
    iUserid  = CLng(replace(request("u"),"'",""))
    if err.number <> 0 then 
	    response.write "Sorry, something went wrong."
	    response.end
    end if
    on error goto 0

   'Check to see if there are multiple lists to unsubscribe from.
    iDListID = replace(request("dl"),"'","")

    if instr(iDListID,",") > 0 then 
       lcl_dl_array = split(iDListID,",")

       for each x in lcl_dl_array
           lcl_dlistid = CLng(x)

          'Remove list subscription from user.
           unSubscribe iUserid, lcl_dlistid
       next
    else
       iDListID = CLng(iDListID)

      'Remove list subscription from user.
       unSubscribe iUserid, iDListID
    end if

    'If request("d") <> "" Then 
       'iDistributionListId = request("d")
       'sListName = GetListName( iDistributionListId )

      'Remove list subscription from user.
       'unSubscribe iUserid
    'Else
       'response.redirect "subscribe.asp"
    'End If 
 else
    response.redirect "subscribe.asp"
 end if
%>
<html>
<head>
	<title>E-Gov Services <%=sOrgName%> - Subscription Removal</title>

	<link rel="stylesheet" type="text/css" href="../css/styles.css" />
	<link rel="stylesheet" type="text/css" href="../global.css" />
	<link rel="stylesheet" type="text/css" href="../css/style_<%=iorgid%>.css" />

</head>

<!--#Include file="../include_top.asp"-->

  <tr>
      <td valign="top">
		        <% RegisteredUserDisplay( "../" ) %>

          <div id="content">
            <div id="centercontent">
              <div class="box_header4"><%=sOrgName%> Subscriptions Removal</div>
                <div class="groupsmall2">
                  <p>You have been removed from this subscription list.</p> 
                  <p>If you wish to receive other subscriptions, you can do so here: 
                     <input type="button" name="manageSubscriptionsButton" id="manageSubscriptionsButton" value="Manage Subscriptions" class="button" onclick="location.href='<%=sEgovWebsiteURL%>/manage_mail_lists.asp'" />
                     <!-- <a href="subscribe.asp"><strong>here</strong></a>. -->
                  </p>
                </div>
                <br />  <br />  <br />
              </div>
            </div>
            <p>&nbsp;</p>

<!--#Include file="../include_bottom.asp"-->    
<!--#Include file="../includes/inc_dbfunction.asp"-->    
<%
'------------------------------------------------------------------------------
function GetListName( iDistributionListId )
	Dim sSQL, oList

	sSQL = "SELECT distributionlistname "
 sSQL = sSQL & " FROM egov_class_distributionlist "
 sSQL = sSQL & " WHERE distributionlistid = " & iDistributionListId

	set oList = Server.CreateObject("ADODB.Recordset")
	oList.Open sSQL, Application("DSN"), 3, 1

	if not oList.eof then
    GetListName = oList("distributionlistname")
 end if

	oList.close
	set oList = nothing 

end function

'------------------------------------------------------------------------------
sub unSubscribe(iUserid, iDLID)
	Dim sSQL, oDelList

 sSQL = "DELETE egov_class_distributionlist_to_user "
 sSQL = sSQL & " WHERE userid = " & iUserid
 sSQL = sSQL & " AND distributionlistid = " & iDLID

	set oDelList = Server.CreateObject("ADODB.Recordset")
	oDelList.Open sSQL, Application("DSN"), 3, 1

'Remove the subscription
	set oDelList = nothing

end sub


%>
