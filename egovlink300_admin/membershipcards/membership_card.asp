<style>
.barcode {
transform:rotate(270deg);
-ms-transform:rotate(270deg); /* IE 9 */
-webkit-transform:rotate(270deg); /* Opera, Chrome, and Safari */

}
</style>
<%

 Dim lcl_memberid, lcl_poolpassid, lcl_fname, lcl_lname, lcl_expiration_date, lcl_pathname

'If this is a REPRINT card then the image exists in the permanent folder
'If this is a NEW card then use the image in the TEMP folder
 if lcl_demo = "Y" then
    lcl_pathname = "../images/MembershipCard_Photos/demo/demo.jpg"
 elseif lcl_layout_maint = "Y" then
    lcl_pathname = "images/profile_image.jpg"
 else
    if lcl_status = "REPRINT" then
       lcl_pathname = "../images/MembershipCard_Photos/" & lcl_member_id & ".jpg"
    else
       set fs = Server.CreateObject("Scripting.FileSystemObject")

      'Build the variables to set up the image paths
       lcl_file_directory = Application("membershipcard_filedirectory")
       lcl_temp_folder    = "/temp"
       lcl_image          = "/" & lcl_member_id & ".jpg"

      'Set up the "temp" and "live" server image paths
       lcl_img_temp_check = lcl_file_directory & lcl_temp_folder & lcl_image
       lcl_img_live_check = lcl_file_directory & lcl_image

      'Set up the "temp" and "live" browser image paths
       lcl_img_temp = "../images/MembershipCard_Photos" & lcl_temp_folder & lcl_image
       lcl_img_live = "../images/MembershipCard_Photos" & lcl_image

       if (fs.FileExists(lcl_img_live_check)) = true then
           lcl_pathname = lcl_img_live
       else
           if (fs.FileExists(lcl_img_temp_check)) = true then
               lcl_pathname = lcl_img_temp
           else
               lcl_pathname = ""
           end if
       end if
    end if
 end if

'Get the member's info to display
 sSQL = "SELECT ppm.memberid, "
 sSQL = sSQL & " p.poolpassid, "
 sSQL = sSQL & " u.userfname, "
 sSQL = sSQL & " u.userlname, "
 sSQL = sSQL & " DateAdd(yy,1,P.paymentdate) as expiration_date, "
 sSQL = sSQL & " p.paymentdate as issue_date "
 sSQL = sSQL & " FROM egov_poolpassmembers ppm, "
 sSQL = sSQL &      " egov_familymembers fm, "
 sSQL = sSQL &      " egov_users u, "
 sSQL = sSQL &      " egov_users u2, "
 sSQL = sSQL &      " egov_poolpasspurchases p "
 sSQL = sSQL & " WHERE ppm.familymemberid = fm.familymemberid "
 sSQL = sSQL & " AND   fm.userid = u.userid "
 sSQL = sSQL & " AND   fm.belongstouserid = u2.userid "
 sSQL = sSQL & " AND   p.orgid = u.orgid "
 sSQL = sSQL & " AND   p.userid = fm.belongstouserid "
 sSQL = sSQL & " AND   ppm.memberid = " & clng(lcl_member_id)

 if UCASE(request("demo")) <> "Y" then
    sSQL = sSQL & " AND   p.poolpassid = " & clng(lcl_poolpassid)
 end if

 Set oMemberid = Server.CreateObject("ADODB.Recordset")
 oMemberid.Open sSQL, Application("DSN"), 3, 1

 if NOT oMemberid.EOF then
    lcl_memberid        = oMemberid("memberid") 
    lcl_fname           = oMemberid("userfname")
    lcl_lname           = oMemberid("userlname")
	   lcl_expiration_date = FormatDateTime(oMemberid("expiration_date"),vbshortdate)
    lcl_issue_date      = MonthName(month(oMemberid("issue_date"))) & " " & day(oMemberid("issue_date")) & ", " & year(oMemberid("issue_date"))
 else
    if lcl_layout_maint = "Y" then
       lcl_memberid        = lcl_layout_memberid
       lcl_fname           = lcl_layout_fname
       lcl_lname           = lcl_layout_lname
	      lcl_expiration_date = "06/01/2009"
       lcl_issue_date      = "June 1, 2008"
    else
       lcl_memberid        = lcl_member_id
       lcl_fname           = ""
       lcl_lname           = ""
   	   lcl_expiration_date = ""
       lcl_issue_date      = ""
    end if
 end if

 oMemberid.close
 Set oMemberid = Nothing 

'Build the barcode image
'*** REQUIRES THE CodeBehind2.dll LOCATED IN THE "../bin" FOLDER! ***
 Dim BarCodeImg, lcl_watermark_class, lcl_card_outline
 BarCodeImg="barcode.aspx?FullAscii=1&X=1&Height=50&Value=" & lcl_memberid

 if session("CARD_PRINT") = "Y" then
    lcl_watermark_class = "card_logo_print"
   	lcl_card_outline    = "card_outline_print"
 else
    lcl_watermark_class = "card_logo_display"
   	lcl_card_outline    = "card_outline_display"
 end if

'Get the name of the org
 sSQLo = "SELECT orgname "
 sSQLo = sSQLo & " FROM organizations "
 sSQLo = sSQLo & " WHERE orgid = " & session("orgid")

 Set rso = Server.CreateObject("ADODB.Recordset")
 rso.Open sSQLo, Application("DSN"), 3, 1

 if not rso.eof then
    lcl_org_name = rso("orgname")
 else
    lcl_org_name = "&nbsp;"
 end if

 rso.close
 set rso = nothing

'Determine the layout based on the printer assigned to the org
 lcl_layout_id = getPrinter_CardLayout(session("orgid"))

 if CLng(lcl_layout_id) = CLng(2) then  'Front and Back (option 1)
    if session("CARD_PRINT") = "Y" then
       lcl_left_position = "0"
    else
       if lcl_display_left <> "" then
          lcl_left_position = lcl_display_left
       else
          lcl_left_position = "0"
       end if
    end if

    lcl_border_size = "0"

   'Since Membership_card.asp is an INCLUDE FILE these fields are set in the file that calls Membership_card.asp
   'Resetting them here just lets us know all of the variables needed for the layout.
   'It also allows us to do any last minute modifications and setting of values (i.e. no color entered then we can default a value)
    lcl_year_text        = lcl_year_text
    lcl_display_date     = lcl_display_date
    lcl_title            = lcl_title
    lcl_subtitle         = lcl_subtitle
    lcl_quote            = lcl_quote
    lcl_custom_image_url = lcl_custom_image_url
    lcl_color1           = lcl_color1
    lcl_color2           = lcl_color2
    lcl_text_color1      = lcl_text_color1
    lcl_text_color2      = lcl_text_color2
    lcl_back_text        = lcl_back_text
    lcl_back_text_color  = lcl_back_text_color

    if lcl_color1 = "" then
       lcl_color1 = "FFFFFF"
    end if

    if lcl_color2 = "" then
       lcl_color2 = "FFFFFF"
    end if

    if lcl_text_color1 = "" then
       lcl_text_color1 = "000000"
    end if

    if lcl_text_color2 = "" then
       lcl_text_color2 = "000000"
    end if

    if lcl_back_text_color = "" then
       lcl_back_text_color = "000000"
    end if

    if CLng(lcl_display_date) = "1" then 'Show Expiration Date
       lcl_display_date_text = "Expires On: " & lcl_expiration_date
    elseif CLng(lcl_display_date) = "2" then 'Show Issued Date
       lcl_display_date_text = "Date Issued: " & lcl_issue_date
    else
       lcl_display_date_text = "&nbsp;"
    end if
%>
<!-- <div id="oFilter_image" style="position: absolute; left: 15px; width: 147px;"> -->
<!-- <div id="oFilter_orgname" style="position: absolute; left: 162px; filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=0, xray=0, mirror=0, invert=0, opacity=1, rotation=3); width: 176px; height: 30px; border: 1px solid #000000;" align="center" class="card_title">PARK CITY</div> -->
<!-- <div style="position: absolute; left: 193px; width: 70px; height: 176px; border: 1px solid #000000;" align="center" class="card_text">&nbsp;</div> -->
<!-- <div id="oFilter_barcode" style="position: absolute; left: 264px; filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=0, xray=0, mirror=0, invert=0, opacity=1, rotation=3); width: 176px; border: 1px solid #000000;" align="center"><img src="<%=BarCodeImg%>" style="height: 40;"></div> -->

<div style="position: absolute; left: <%=lcl_left_position%>px; width: 300px; height: 176px; border: 1px solid #000000;">
    <div style="position: absolute; width: 249px; height: 176px; border: <%=lcl_border_size%>px solid #000000;" class="layout<%=lcl_layout_id%>_member_name_text">
<!--    <div style="position: absolute; left: 1px; top: 1px; width: 249px; height: 176px; border: <% 'lcl_border_size%>px solid #000000;" class="layout<% 'lcl_layout_id%>_member_name_text"> -->
        <div id="member_name" style="position: absolute; left: 0px; top: 0px; width: 100%; height: 20px; border: <%=lcl_border_size%>px solid #000000; background-color: #<%=formatCardDisplayValue(lcl_color1)%>;" class="layout<%=lcl_layout_id%>_member_name_text" style="color: #<%=lcl_text_color1%>">
            &nbsp;<%=lcl_fname%>&nbsp;<%=lcl_lname%>
        </div>

        <div style="position: absolute; left: 0px; top: 20px; width: 132px; height: 140px; border: <%=lcl_border_size%>px solid #000000;">
            <div style="position: absolute; left: 0px; top: 0px; width: 100%; height: 15px; border: <%=lcl_border_size%>px solid #000000;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#<%=formatCardDisplayValue(lcl_color2)%>">
                  <tr>
                      <td class="layout<%=lcl_layout_id%>_year_card_text" style="color: #<%=lcl_text_color2%>" height="15px">&nbsp;<%=formatCardDisplayValue(lcl_year_text)%></td>
<!--                      <td width="30px"><img src="images/year_angle.jpg" width="30" height="15"></td> -->
                  </tr>
                </table>
            </div>
            <div style="position: absolute; left: 0px; top: 15px; width: 100%; height: 125px; border: <%=lcl_border_size%>px solid #000000;" class="layout<%=lcl_layout_id%>_card_text">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
                  <tr align="center" valign="middle">
                      <td><img src="<%=lcl_pathname%>" width="120" height="90" style="border: 1px solid #000000; border-radius: 6px;" /></td>
                  </tr>
                </table>
            </div>
        </div>

        <div style="position: absolute; left: 132px; top: 20px; width: 117px; height: 140px; border: <%=lcl_border_size%>px solid #000000;" align="center" class="layout<%=lcl_layout_id%>_card_text">
            <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
              <tr>
                  <td align="center">
                      <font style="font-size: 18px"><%=formatCardDisplayValue(lcl_title)%></font><br>
                      <font style="font-size: 14px"><%=formatCardDisplayValue(lcl_subtitle)%></font>
                  </td>
              </tr>
              <tr>
                  <td align="center" valign="top">
                   <% if lcl_custom_image_url <> "" then %>
                      <img src="<%=lcl_custom_image_url%>" width="100" height="50">
                   <% else %>
                      <p>&nbsp;</p><p>&nbsp;</p>
                   <% end if %>
                  </td>
              </tr>
              <tr>
                  <td align="center">
                      <font style="font-size: 14px"><%=formatCardDisplayValue(lcl_quote)%></font>
                  </td>
              </tr>
            </table>
        </div>

        <div style="position: absolute; left: 0px; top: 161px; width: 100%; height: 15px; border: <%=lcl_border_size%>px solid #000000; white-space:nowrap;">
            <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
              <tr>
<!--                  <td class="layout<%=lcl_layout_id%>_year_card_text" bgcolor="#0000C0" width="133px" height="15px">&nbsp;Date Issued: June 1, 2008<img src="images/date_issued_angle.jpg" width="30" height="15" style="position: absolute; left: 103px; top: 0px;"></td> -->
                  <td class="layout<%=lcl_layout_id%>_expiredate_text" style="color: #<%=lcl_text_color2%>" bgcolor="#<%=formatCardDisplayValue(lcl_color2)%>" width="133px" height="15px">&nbsp;<%=formatCardDisplayValue(lcl_display_date_text)%></td>
                  <td class="layout<%=lcl_layout_id%>_expiredate_text" style="color: #<%=lcl_text_color1%>" bgcolor="#<%=formatCardDisplayValue(lcl_color1)%>">&nbsp;</td>
              </tr>
            </table>              
        </div>
    </div>
<%
'ATTEMPTING TO FIND THE WIDTH AND HEIGHT OF THE BARCODE TO TRY AND AUTO-CENTER.
  'strFlnm = "http://www.egovlink.com/montgomery/admin/images/MembershipCard_Photos/13486.jpg"
  'if gfxSpex(strFlnm, lngWidth, lngHeight, lngDepth, strImageType) then
      'Set fs= CreateObject("Scripting.FileSystemObject") 
      'set myImg = loadpicture(server.mappath(session("egovclientwebsiteurl") &"/admin/membershipcards/" & BarCodeImg))
      'set myImg = loadpicture(server.mappath("http://www.egovlink.com/montgomery/admin/membershipcards/barcode.aspx?FullAscii=1&X=1&Height=40&Value=13486"))
      'set myImg = loadpicture("http://www.egovlink.com/montgomery/admin/images/MembershipCard_Photos/13486.jpg")
      'iWidth = myImg.width
  'end if

   'The "top" style parameter on this division (id="oFilter_barcode") is used to position (align center) the membership card.
   'The height of the card is 176px
      '*** The middle then is 88px ***
         '*** Half of that is 44px. ***
         'We do this because we want the MIDDLE of the barcode image to be positioned in the MIDDLE of the card.
         '44px is centers the barcode for (5) characters/numbers.
%>
    <div id="oFilter_barcode" style="position: absolute; left: 230px; top: 44px; filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=0, xray=0, mirror=0, invert=0, opacity=1, rotation=3); border: <%=lcl_border_size%>px solid #000000;" align="center" class="barcode"><img src="<%=BarCodeImg%>"></div>
</div>
<% elseif CLng(lcl_layout_id) = CLng(3) then  'Front and Back (option 2) %>
<div style="position: absolute; left: 0px; width: 300px; height: 176px; border: 1px solid #000000;">
    <div style="position: absolute; left: 1px; top: 1px; width: 249px; height: 176px; border: <%=lcl_border_size%>px solid #000000;" class="layout<%=lcl_layout_id%>_member_name_text">
        <div id="member_name" style="position: absolute; left: 0px; top: 0px; width: 100%; height: 20px; border: <%=lcl_border_size%>px solid #000000; background-color: #FFFF00;" class="layout<%=lcl_layout_id%>_member_name_text">
            &nbsp;<%=lcl_fname%>&nbsp;<%=lcl_lname%>
        </div>

        <div style="position: absolute; left: 0px; top: 20px; width: 132px; height: 140px; border: <%=lcl_border_size%>px solid #000000;">
            <div style="position: absolute; left: 0px; top: 0px; width: 100%; height: 15px; border: <%=lcl_border_size%>px solid #000000;">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" bgcolor="#0000C0">
                  <tr>
                      <td class="layout<%=lcl_layout_id%>_year_card_text">&nbsp;2008 Pool Member</td>
                      <td width="30px"><img src="images/year_angle.jpg" width="30" height="15"></td>
                  </tr>
                </table>
            </div>
            <div style="position: absolute; left: 0px; top: 15px; width: 100px; height: 125px; border: <%=lcl_border_size%>px solid #000000;" class="layout<%=lcl_layout_id%>_card_text"">
                <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
                  <tr align="center" valign="middle">
                      <td><img src="<%=lcl_pathname%>" width="93" height="75" style="border: 1px solid #000000"></td>
                  </tr>
                </table>
            </div>
            <div id="oFilter_image" style="position: absolute; left: 101px; top: 15px; filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=0, xray=0, mirror=0, invert=0, opacity=1, rotation=3); width: 124px; height: 30px; border: <%=lcl_border_size%>px solid #000000;" align="center" class="layout<%=lcl_layout_id%>_card_title">Go Makos!</div>
        </div>

        <div style="position: absolute; left: 132px; top: 20px; width: 117px; height: 140px; border: <%=lcl_border_size%>px solid #000000;" align="center" class="layout<%=lcl_layout_id%>_card_text">
            <table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
              <tr>
                  <td>
                      <font style="font-size: 18px">Montgomery</font><br>
                      <font style="font-size: 14px">Municipal Pool</font>
                  </td>
              </tr>
              <tr>
                  <td align="center" valign="top"><img src="images/logo_center.jpg" width="100" height="50"></td>
              </tr>
            </table>
        </div>

        <div style="position: absolute; left: 0px; top: 160px; width: 100%; height: 15px; border: <%=lcl_border_size%>px solid #000000;">
            <table border="0" cellspacing="0" cellpadding="0" width="100%">
              <tr>
                  <td class="layout<%=lcl_layout_id%>_year_card_text" bgcolor="#0000C0" width="133px" height="15px">&nbsp;Date Issued: June 1, 2008<img src="images/date_issued_angle.jpg" width="30" height="15" style="position: absolute; left: 103px; top: 0px;"></td>
                  <td class="layout<%=lcl_layout_id%>_year_card_text" bgcolor="#FFFF00">&nbsp;</td>
              </tr>
            </table>              
        </div>
    </div>
    <div id="oFilter_barcode" style="position: absolute; left: 249px; top: 1px; filter:progid:DXImageTransform.Microsoft.BasicImage(grayscale=0, xray=0, mirror=0, invert=0, opacity=1, rotation=3); width: 175px; border: <%=lcl_border_size%>px solid #000000;" align="center"><img src="<%=BarCodeImg%>" style="height: 40px;"></div>
</div>
<% elseif session("orgid") = "50000" then %>
<table border="0" cellspacing="0" cellpadding="0" width="300" height="180" class="layout<%=lcl_layout_id%>_<%=lcl_card_outline%>">
  <tr><td colspan="2" align="center" class="layout<%=lcl_layout_id%>_card_title">PETERS TOWNSHIP</td></tr>
  <tr valign="top">
      <td align="center" width="124" class="layout<%=lcl_layout_id%>_card_text">
          <table border="0" cellspacing="0" cellpadding="0" width="100%" class="layout<%=lcl_layout_id%>_img_outline">
            <tr><td><img src="<%=lcl_pathname%>" width="147" height="110" /></td></tr>
          </table>
      </td>
      <td><table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" class="layout<%=lcl_layout_id%>_card_text">
            <tr><td valign="top" align="center">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%" class="layout<%=lcl_layout_id%>_card_text">
                      <tr><td><img src="../images/poolpass_org_logos/logo_15.jpg" width="100" height="133" class="layout<%=lcl_layout_id%>_<%=lcl_watermark_class%>"></td></tr>
                      <tr align="center">
               					      <td><font class="layout<%=lcl_layout_id%>_member_name_text"><%=lcl_fname%><br><%=lcl_lname%></font><p>
                              <font class="layout<%=lcl_layout_id%>_card_label">EXPIRES ON</font><br>
                              <font class="layout<%=lcl_layout_id%>_expiredate_text"><%=lcl_expiration_date%></font>
                          </td></tr>
                    </table>
                </td></tr>
          </table>
      </td></tr>
  <tr><td colspan="2" align="center" height="40px"><img src="<%=BarCodeImg%>" style="Height:40"></td></tr>
</table>
<% else %>
<table border="0" cellspacing="0" cellpadding="0" width="300" height="180" class="layout<%=lcl_layout_id%>_<%=lcl_card_outline%>">
  <tr><td colspan="2" align="center" class="layout<%=lcl_layout_id%>_card_title"></td></tr>
  <tr valign="top">
      <td align="center" width="124" class="layout<%=lcl_layout_id%>_card_text">
          <table border="0" cellspacing="0" cellpadding="0" width="100%" class="layout<%=lcl_layout_id%>_img_outline">
            <tr><td><img src="<%=lcl_pathname%>" width="147" height="110" /></td></tr>
          </table>
      </td>
      <td><table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%" class="layout<%=lcl_layout_id%>_card_text">
            <tr><td valign="top" align="center">
                    <table border="0" cellspacing="0" cellpadding="0" width="100%" class="layout<%=lcl_layout_id%>_card_text">
                      <tr><td></td></tr>
                      <tr align="center">
               					      <td><font class="layout<%=lcl_layout_id%>_member_name_text"><%=lcl_fname%><br><%=lcl_lname%></font><p>
                              <font class="layout<%=lcl_layout_id%>_card_label">EXPIRES ON</font><br>
                              <font class="layout<%=lcl_layout_id%>_expiredate_text"><%=lcl_expiration_date%></font>
                          </td></tr>
                    </table>
                </td></tr>
          </table>
      </td></tr>
  <tr><td colspan="2" align="center" height="40px"><img src="<%=BarCodeImg%>" style="Height:40"></td></tr>
</table>
<% end if %>
