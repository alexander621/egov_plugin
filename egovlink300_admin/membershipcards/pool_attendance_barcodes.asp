<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!-- #include file="../includes/common.asp" //-->
<%
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
' FILENAME: pool_attendance_barcodes.asp
' AUTHOR:   David Boyer
' CREATED:  05/19/08
' COPYRIGHT: Copyright 2007 eclink, inc.
'			 All Rights Reserved.
'
' Description:  Displays all non-member attendance type barcodes.
'               Based on PoolPass Membership Rates (Attendance Type)
'
' MODIFICATION HISTORY
' 1.0  05/19/08  David Boyer - Created Code.
'
'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

sLevel     = "../" ' Override of value from common.asp
lcl_hidden = "hidden"  'Show/Hide all hidden fields.  TEXT=Show,HIDDEN=Hide

if NOT UserHasPermission( session("userid"), "print_barcodes" ) then
   response.redirect sLevel & "permissiondenied.asp"
end if
%>
<html>
<head>
  <title>E-GovLink {Custom Barcode List}</title>
  
	<link rel="stylesheet" type="text/css" href="../global.css">
	<link rel="stylesheet" type="text/css" href="../menu/menu_scripts/menu.css" />
	<link rel="stylesheet" type="text/css" href="cardprint.css" media="print" />

 <script src="../scripts/selectAll.js"></script>
 <script language="javascript" src="../scripts/modules.js"></script>

<script defer>
function window.onload() {
//  factory.printing.printer           = "Zebra P330i USB Card Printer";
  factory.printing.header            = "";
  factory.printing.footer            = "";
  factory.printing.portrait          = true;
//  factory.printing.leftMargin        = 0.15;
//  factory.printing.rightMargin       = 0.15;
//  factory.printing.topMargin         = 0.15;
//  factory.printing.bottomMargin      = 0.1;

  //enable control buttons
  var templateSupported = factory.printing.IsTemplateSupported();
  var controls = idControls.all.tags("input");
  for ( i = 0; i < controls.length; i++ ) {
     controls[i].disabled = false;
     if (templateSupported && controls[i].className == "ie55" )
         controls[i].style.display = "inline";
     }
}
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">

<object id="factory" viewastext style="display:none"
  classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
  codebase="../includes/smsx.cab#Version=6,3,434,12">
</object>

<div class="noprint">
<% ShowHeader sLevel %>
<!--#Include file="../menu/menu.asp"-->
</div>

<div id="content">
    <div id="centercontent">

<form action="pool_attendance_barcodes.asp" method="post" name="attendance_barcodes" id="attendance_barcodes">

<div>
  <font size="+1"><strong><%=session("sOrgName")%> - Custom Barcode List</strong></font><p>
</div>

<div id="idControls" class="noprint">
    <input disabled type="button" value="Print the page" onclick="factory.printing.Print(true)" />&nbsp;&nbsp;
    <input class="ie55" disabled type="button" value="Print Preview..." onclick="factory.printing.Preview()" />
</div>
<p>

  <%
    'Retrieve all attendance types
     sSQL = "SELECT attendancetypeid, attendancetype, isactive, isdefault "
     sSQL = sSQL & " FROM egov_pool_attendancetypes "
     sSQL = sSQL & " WHERE isactive = 1 "
     sSQL = sSQL & " AND attendancetypeid <> 1 "  'Member
     sSQL = sSQL & " ORDER BY UPPER(attendancetype) "

     set rs = Server.CreateObject("ADODB.Recordset")
     rs.Open sSQL, Application("DSN") , 3, 1

     if not rs.eof then
        lcl_typeid  = ""
        lcl_barcode = ""

        while not rs.eof
           lcl_typeid  = rs("attendancetypeid")
           lcl_barcode = ""

          'First get the length.
           lcl_length = len(lcl_typeid)

          'An X on the front tells the system, when the barcode is scanned, that this is a "custom" barcode.
          'If the length is LESS THAN 4 characters then concatenate zeros onto the front of the number AFTER the "X".
          'If not then simply show the number AFTER the "X".
           lcl_length = lcl_length + 1  '+1 to account for the "X"

           if lcl_length < 4 then
              for i = 1 to lcl_length
                  lcl_barcode = lcl_barcode & "0"
              next
           end if

          'Build the barcode.
           lcl_barcode = "X" & lcl_barcode & lcl_typeid

           BarCodeImg="barcode.aspx?FullAscii=1&X=2&Height=50&Value=" & lcl_barcode

          'Display the barcode
           response.write "  <strong class=""font-size: 16pt"">" & rs("attendancetype") & "</strong><br>" & vbcrlf
           response.write "  <img src=""" & BarCodeImg & """ height=""50px""><p>&nbsp;</p>" & vbcrlf

           rs.movenext
        wend
     end if
  %>
</form>

    </div>
</div>

<div class="noprint">
<!--#Include file="../admin_footer.asp"-->  
</div>
</body>
</html>