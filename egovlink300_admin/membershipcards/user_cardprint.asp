<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="user_card_functions.asp" //-->
<%
'Check to see if the feature is offline
 if isFeatureOffline("registration") = "Y" then
    response.redirect "../admin/outage_feature_offline.asp"
 end if

 dim lcl_userid, lcl_status, lcl_onload, lcl_initPrint, lcl_OS
 dim lcl_userPermission

 sLevel             = "../"  'Override of value from common.asp
 lcl_onload         = ""
 lcl_initPrint      = "N"
 lcl_OS             = "XP"
 lcl_userPermission = "create_user_membershipcards"

 'Determine where the user is coming from so we can check for the proper permission.
  if request("os") <> "" then
     lcl_OS = request("os")
     lcl_OS = ucase(lcl_OS)
     lcl_OS = replace(lcl_OS, "'", "")
  end if

  if lcl_OS = "W7" then
     lcl_userPermission = "create_user_membershipcards_new"
  end if

 if not userhaspermission(session("userid"),lcl_userPermission) then
   	response.redirect sLevel & "permissiondenied.asp"
 end if
 
 if request("userid") = "" then
    lcl_userid = session("userid")
 else
    lcl_userid = request("userid")
 end if

 if request("STATUS") = "" then
    lcl_status = session("STATUS")
 else
    lcl_status = request("STATUS")
 end if

 if request("initPrint") = "Y" then
    lcl_initPrint = request("initPrint")
 end if

 lcl_reload_url = "user_cardprint.asp"
 lcl_reload_url = lcl_reload_url & "?userid=" & lcl_userid
 lcl_reload_url = lcl_reload_url & "&status=PRINT"
 lcl_reload_url = lcl_reload_url & "&card_layout=" & request("card_layout")
 lcl_reload_url = lcl_reload_url & "&initPrint=N"
 lcl_reload_url = lcl_reload_url & "&os=" & lcl_OS

 if lcl_initPrint = "Y" then
    lcl_onload = " onload=""location.href='" & lcl_reload_url & "';"""
 end if

 session("CARD_PRINT") = "Y"
 session("STATUS")     = lcl_status

'Determine the layout based on the printer assigned to the org
 'lcl_layout_id = getPrinter_CardLayout(session("orgid"))
%>
<html>
<head>
  <title>E-Gov Administration Console {Membership Card Print}</title>

  <link rel="stylesheet" href="../global.css" />
  <link rel="stylesheet" href="user_card.css" />
  <link rel="stylesheet" href="cardprint.css" media="print" />

<script>
function print_card()
{
  window.print();
}
</script>

<script defer>
function window.onload() {
<% 'if clng(lcl_layout_id) = clng(2) then %>
  factory.printing.printer           = "Zebra P330i USB Card Printer";
  factory.printing.header            = "";
  factory.printing.footer            = "";
  factory.printing.portrait          = false;
  factory.printing.leftMargin        = 0.1;
  factory.printing.rightMargin       = 0.15;
  factory.printing.topMargin         = 0.15;
  factory.printing.bottomMargin      = 0.1;
<% 'end if %>

  //enable control buttons
//  var templateSupported = factory.printing.IsTemplateSupported();
//  var controls = idControls.all.tags("input");
//  for ( i = 0; i < controls.length; i++ ) {
//     controls[i].disabled = false;
//     if (templateSupported && controls[i].className == "ie55" )
//         controls[i].style.display = "inline";
//     }
//}
</script>

<style>
  .print_margins {
     margin-top:    0px;
     margin-left:   0px;
     margin-right:  0px;
     margin-bottom: 0px;
  }
</style>

</head>
<body<%=lcl_onload%>>
<%
  response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
  'response.write "	<input type=""button"" disabled=""disabled"" value=""Print the page"" onclick=""factory.printing.Print(true)"" />&nbsp;&nbsp;" & vbcrlf
  'response.write "	<input type=""button"" disabled=""disabled"" value=""Print Preview..."" onclick=""factory.printing.Preview()"" class=""ie55"" />" & vbcrlf
  response.write "  <input type=""button"" value=""Print the page"" onclick=""print_card();"" />" & vbcrlf
  response.write "	<input type=""button"" value=""Close Window"" onclick=""window.close();"" />" & vbcrlf
  response.write "</div>" & vbcrlf

  response.write "<object id=""factory"" viewastext style=""display:none""" & vbcrlf
  response.write "  classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814""" & vbcrlf
  response.write "  codebase=""../includes/smsx.cab#Version=6,3,434,12"">" & vbcrlf
  response.write "</object>" & vbcrlf

  response.write "<div class=""print_margins"">" & vbcrlf
                    displayCard session("orgid"), lcl_userid, lcl_status
  response.write "</div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>