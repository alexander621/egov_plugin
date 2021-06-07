<%
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------
' FILENAME: unsubscribe.asp
' AUTHOR: David Boyer
' CREATED: 02/24/2010
' COPYRIGHT: Copyright 2006 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Generic page created to allow new subscribers to log on using a link generated in new subscription confirmation email
'
' MODIFICATION HISTORY
' 1.0 02/24/10 David Boyer - Initial version
'
'------------------------------------------------------------------------------
'
'------------------------------------------------------------------------------

 if request("c") <> "" then
    if isnumeric(request("c")) then
       response.redirect "http://www.egovlink.com/eclink/subscriptions/subscribe_action.asp?c=" & request("c")
    end if
 end if

 response.redirect "http://www2.egovlink.com"
%>
