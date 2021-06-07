<!DOCTYPE HTML>
<!-- #include file="../includes/common.asp" //-->
<!-- #include file="membership_card_functions.asp" //-->
<%
'Check to see if the feature is offline
if isFeatureOffline("memberships") = "Y" then
   response.redirect "../admin/outage_feature_offline.asp"
end if

dim lcl_member_id, lcl_status

sLevel = "../"  'Override of value from common.asp

'Check for org features.
 lcl_orghasfeature_card_layout_multiplelayouts = orghasfeature("card_layout_multiplelayouts")

'Determine if this is a demo or not.  demo = Y means that these screens can function without the web camera attached
 lcl_demo            = request("demo")
 lcl_demo_page_title = ""
 lcl_demo_url        = ""
 lcl_initPrint       = "N"
 lcl_onload          = ""

'set up demo variables
 if lcl_demo = "Y" then
    lcl_demo_page_title = " (DEMO)"
    lcl_demo_url        = "&demo=" & lcl_demo
 end if

 if request("memberid") = "" then
    lcl_member_id = session("MEMBERID")
 else
    lcl_member_id = request("memberid")
 end if

 if request("STATUS") = "" then
    lcl_status = session("STATUS")
 else
    lcl_status = request("STATUS")
 end if

 if request("poolpassid") = "" then
    lcl_poolpassid = session("poolpassid")
 else
    lcl_poolpassid = request("poolpassid")
 end if

 lcl_rateid = getRateID(lcl_poolpassid)

' if lcl_status = "PRINT" then
'    save_card(lcl_member_id)
' end if

 if request("initPrint") = "Y" then
    lcl_initPrint = request("initPrint")
 end if

 lcl_reload_url = "card_print.asp"
 lcl_reload_url = lcl_reload_url & "?memberid="    & lcl_member_id
 lcl_reload_url = lcl_reload_url & "&poolpassid="  & lcl_poolpassid
 lcl_reload_url = lcl_reload_url & "&status=PRINT" & lcl_demo_url
 lcl_reload_url = lcl_reload_url & "&card_layout=" & request("card_layout")
 lcl_reload_url = lcl_reload_url & "&initPrint=N"

 if lcl_initPrint = "Y" then
    lcl_onload = " onload=""location.href='" & lcl_reload_url & "';"""
 end if

 session("CARD_PRINT") = "Y"
 session("STATUS")     = lcl_status

'Determine the layout based on the printer assigned to the org
 lcl_layout_id    = getPrinter_CardLayout(session("orgid"))
 lcl_layout_style = "layout" & lcl_layout_id & "_"
 lcl_print_side   = request("print_side")

'Retrieve the card layout attributes
 lcl_layout_maint = "N"

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

 if lcl_print_side <> "BACK" then
    if lcl_orghasfeature_card_layout_multiplelayouts then
       sSQL = sSQL & " LEFT OUTER JOIN egov_poolpassrates r ON m.cardid = r.cardid "
       sSQL = sSQL & " WHERE r.rateid = " & lcl_rateid
       sSQL = sSQL & " AND "
    else
       sSQL = sSQL & " WHERE "
    end if
 else
    sSQL = sSQL & " WHERE "
 end if

 sSQL = sSQL & " m.orgid = " & session("orgid")

 set oCardPrint = Server.CreateObject("ADODB.Recordset")
 oCardPrint.Open sSQL, Application("DSN"), 3, 1

 if not oCardPrint.eof then
    lcl_title            = oCardPrint("title")
    lcl_subtitle         = oCardPrint("subtitle")
    lcl_year_text        = oCardPrint("year_text")
    lcl_display_date     = oCardPrint("display_date")
    lcl_custom_image_url = oCardPrint("custom_image_url")
    lcl_quote            = oCardPrint("quote_text")
    lcl_color1           = oCardPrint("main_color")
    lcl_color2           = oCardPrint("secondary_color")
    lcl_text_color1      = oCardPrint("main_text_color")
    lcl_text_color2      = oCardPrint("secondary_text_color")
    lcl_back_text        = oCardPrint("back_text")
    lcl_back_text_color  = oCardPrint("back_text_color")
else
    lcl_title            = ""
    lcl_subtitle         = ""
    lcl_year_text        = ""
    lcl_display_date     = 0
    lcl_custom_image_url = ""
    lcl_quote            = ""
    lcl_color1           = "FFFFFF"
    lcl_color2           = "FFFF"
    lcl_text_color1      = "000000"
    lcl_text_color2      = "000000"
    lcl_back_text        = ""
    lcl_back_text_color  = "000000"
end if

oCardPrint.close
set oCardPrint = nothing
%>
<html>
<head>
  <title>E-Gov Administration Console {Membership Card Print}</title>

<script>
function print_card() {
  window.print();
}
</script>


  <link rel="stylesheet" href="../global.css" />
  <link rel="stylesheet" href="membership_card.css" />
  <link rel="stylesheet" href="cardprint.css" media="print" />

<style>
#backButtonDiv
{
  position: absolute;
  left:     0px;
  width:    300px;
  height:   176px;
  border:   1px solid #000000;
}

#backButtonTable
{
  font-size: 10pt;
}

#backButtonTable td
{
  margin:  0px;
  padding: 0px;
  border:  none;
}
</style>

</head>
<body<%=lcl_onload%>>
<%
  response.write "<div id=""idControls"" class=""noprint"">" & vbcrlf
  response.write "  <input type=""button"" value=""Print the page"" onclick=""factory.printing.Print(true)"" disabled=""disabled"" />&nbsp;&nbsp;" & vbcrlf
  response.write "  <input type=""button"" value=""Print Preview..."" onclick=""factory.printing.Preview()"" disabled=""disabled"" class=""ie55"" />" & vbcrlf
  response.write "  <input type=""button"" value=""Close Window"" onclick=""window.close();"" />" & vbcrlf
  response.write "</div>" & vbcrlf

  response.write "<object id=""factory"" viewastext style=""display:none"" classid=""clsid:1663ed61-23eb-11d2-b92f-008048fdd814"" codebase=""../includes/smsx.cab#Version=6,3,434,12""></object>" & vbcrlf

  response.write "<div class=""" & lcl_layout_style & "print_margins"">" & vbcrlf

  if UCASE(lcl_print_side) = "BACK" then
     response.write "<div id=""backButtonDiv"">" & vbcrlf
     response.write "  <table id=""backButtonTable"" width=""100%"" height=""100%"">" & vbcrlf
     response.write "    <tr>" & vbcrlf
     response.write "        <td id=""preview_card_back"" align=""center"" valign=""middle"" style=""color: #" & lcl_back_text_color & """>" & vbcrlf
     response.write              lcl_back_text & vbcrlf
     response.write "        </td>" & vbcrlf
     response.write "    </tr>" & vbcrlf
     response.write "  </table>" & vbcrlf
     response.write "</div>" & vbcrlf
  else
%>
    <!-- #include file="membership_card.asp" //-->
<%
  end if

  response.write "</div>" & vbcrlf
  response.write "</body>" & vbcrlf
  response.write "</html>" & vbcrlf
%>
<script defer>
window.onload = function () {
<% if CLng(lcl_layout_id) = CLng(2) then %>
  factory.printing.printer           = "Zebra P330i USB Card Printer";
  factory.printing.header            = "";
  factory.printing.footer            = "";
  factory.printing.portrait          = false;
  factory.printing.leftMargin        = 0.1;
  factory.printing.rightMargin       = 0.1;
  factory.printing.topMargin         = 0.1;
  factory.printing.bottomMargin      = 0.1;
<% end if %>

  //enable control buttons
  var templateSupported = factory.printing.IsTemplateSupported();
  //var controls = idControls.all.tags("input");
  var controls = document.getElementById("idControls").getElementsByTagName('input');

alert("HERE");
  for ( i = 0; i < controls.length; i++ ) {
     controls[i].disabled = false;

     if (templateSupported && controls[i].className == "ie55" )
         controls[i].style.display = "inline";
     }
}
</script>
