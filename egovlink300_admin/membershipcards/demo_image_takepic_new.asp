<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="../includes/time.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel = "../"  'Override of value from common.asp

lcl_member_id  = request("memberid")
lcl_poolpassid = request("poolpassid")

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo            = request("demo")
 lcl_demo_page_title = " (DEMO)"
 lcl_demo_url        = "&demo=" & lcl_demo

 lcl_image_filepath = "../images/MembershipCard_Photos/demo"
%>
<html>
<head><title> Membership Photo Taking System </title>
  <link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
  <link rel="stylesheet" type="text/css" href="../global.css" />

<script language="javascript">
function openWin(page) {
  OpenWin = window.open(page, "new", "toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes,width=700,height=600,screenX=0,screenY=0");
  if (document.images) {OpenWin.focus();}
}

function closeWin() {
  parent.opener.location.href="image_takepic_new.asp?memberid=<%=lcl_member_id%>&poolpassid=<%=lcl_poolpassid%>&demo=Y&step2=Y";
  parent.opener.check_instruction3('');
  parent.close();
}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<table border="0" cellspacing="0" cellpadding="0">
  <tr>
      <td><img src="<%=lcl_image_filepath%>/webcam_interface.jpg" width="487" height="324"></td>
  </tr>
  <tr>
      <td align="right" height="42" background="<%=lcl_image_filepath%>/webcam_interface_bottom_bg.jpg">
          <input type="button" value="Get Pictures" onclick="closeWin()">&nbsp;
          <input type="button" value="Cancel">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      </td>
  </tr>

</body>
</html>
