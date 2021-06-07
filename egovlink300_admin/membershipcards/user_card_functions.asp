<%
'------------------------------------------------------------------------------
function getCardPrintedTotal(iUserID)

  dim sSQL, oCardCount, sUserID

  lcl_return = 0
  sUserID    = 0

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  sSQL = "SELECT isnull(card_printed_count,0) as card_printed_count "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & lcl_userid

  set oCardCount = Server.CreateObject("ADODB.Recordset")
  oCardCount.Open sSQL, Application("DSN"), 0, 1

  if not oCardCount.eof then
     lcl_return = oCardCount("card_printed_count")
  end if

  oCardCount.close
  set oCardCount = nothing

  getCardPrintedTotal = lcl_return

end function

'-----------------------------------------------------------------------
sub save_card(p_orgid, p_UID)

'Set up the transfer the file from the TEMP folder to the permanent folder
  dim sSQL, fs, pathname, lcl_image_temp, lcl_image_new, lcl_from_file, lcl_to_file, lcl_orgid, lcl_uid, oCardSave

  lcl_pathname = Application("membershipcard_filedirectory")

  if p_orgid <> "" then
     lcl_orgid = clng(p_orgid)
  end if

  if p_UID <> "" then
     lcl_uid = clng(p_UID)
  end if

'DVLP
' lcl_pathname = "c:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\images\MembershipCard_Photos"

'TEST"
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\egovlink300_admin\images\MembershipCard_Photos"

'PROD
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink300_admin\images\MembershipCard_Photos"

  lcl_temp_folder = "\temp"
  lcl_image_temp  = "\" & lcl_uid & ".jpg"

  lcl_new_folder  = "\users"
  lcl_image_new   = "\" & lcl_orgid & "_" & lcl_uid & ".jpg"

  set fs = Server.CreateObject("Scripting.FileSystemObject")

 'Set the parameters for MoveFile
  lcl_from_file = lcl_pathname & lcl_temp_folder & lcl_image_temp
  lcl_to_file   = lcl_pathname & lcl_new_folder  & lcl_image_new
  lcl_overwrite = TRUE

 'Move the image
  if (fs.FileExists(lcl_from_file)) = true then
      fs.CopyFile lcl_from_file, lcl_to_file, lcl_overwrite

     'Delete the image from the temp directory.
      fs.DeleteFile lcl_from_file

     'Update e-gov_users to show that a pic has been uploaded.
      sSQL = "UPDATE egov_users "
      sSQL = sSQL & " SET card_pic_uploaded = 1 "
      sSQL = sSQL & " WHERE userid = " & lcl_uid

      set oCardSave = Server.CreateObject("ADODB.Recordset")
      oCardSave.Open sSQL, Application("DSN"), 3, 1

      set oCardSave = nothing

  end if

  set fs = nothing

end sub

'------------------------------------------------------------------------------
sub print_card(p_orgid, p_UID)

  dim lcl_orgid, lcl_uid, sSQL, sSQLu, oCardInfo, oUserCardUpdate
  dim lcl_card_pic_uploaded, lcl_card_printed_count

  lcl_orgid              = 0
  lcl_uid                = 0
  lcl_card_pic_uploaded  = false
  lcl_card_printed_count = 1

  if p_orgid <> "" then
     lcl_orgid = clng(p_orgid)
  end if

  if p_UID <> "" then
     lcl_uid = clng(p_UID)
  end if

 'Determine if the profile pic has been uploaded and the number of times the card has been printed.
  sSQL = "SELECT card_pic_uploaded, "
  sSQL = sSQL & " card_printed_count "
  sSQL = sSQL & " FROM egov_users "
  sSQL = sSQL & " WHERE userid = " & lcl_uid

  set oCardInfo = Server.CreateObject("ADODB.Recordset")
  oCardInfo.Open sSQL, Application("DSN"), 0, 1

  if not oCardInfo.eof then
     lcl_card_pic_uploaded  = oCardInfo("card_pic_uploaded")
     lcl_card_printed_count = oCardInfo("card_printed_count")
  end if

  if lcl_card_pic_uploaded then
     lcl_card_printed_count = lcl_card_printed_count + 1
  end if

  sSQLu = "UPDATE egov_users SET "
  sSQLu = sSQLu & " card_printed_count = " & lcl_card_printed_count
  sSQlu = sSQLu & " WHERE userid = " & lcl_uid

  set oUserCardUpdate = Server.CreateObject("ADODB.Recordset")
  oUserCardUpdate.Open sSQLu, Application("DSN"), 3, 1

  set oCardInfo       = nothing
  set oUserCardUpdate = nothing

end sub

'-----------------------------------------------------------------------
function remove_image(p_UID)
'remove the image from the TEMP folder on our server
 Dim fs, lcl_userid, lcl_file

'DVLP
' lcl_file = "c:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

'TEST"
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

'PROD
' lcl_file = "d:\wwwroot\www.cityegov.com\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

 lcl_userid = 0

 if p_UID <> "" then
    lcl_userid = clng(p_UID)
 end if

 lcl_file = Application("membershipcard_filedirectory")
 lcl_file = lcl_file & "\temp"
 lcl_file = lcl_file & "\" & lcl_userid & ".jpg"

 set fs = Server.CreateObject("Scripting.FileSystemObject")

 if (fs.FileExists(lcl_file)) = true then
      fs.deletefile(lcl_file)
 end if

 set fs = nothing

end function

'------------------------------------------------------------------------------
sub displayCard(iOrgID, iUserID, iStatus)

  dim sOrgID, lcl_orgname, sUserID, sStatus, sSQL
  dim lcl_pathname, lcl_file_directory, lcl_temp_folder, lcl_live_folder, lcl_new_folder
  dim lcl_temp_image, lcl_live_image, lcl_img_temp_check, lcl_img_live_check, lcl_img_temp, lcl_img_live
  dim lcl_fname, lcl_lname, lcl_layout_id
  dim BarCodeImg, lcl_watermark_class, lcl_card_outline

  sOrgID       = 0
  sUserID      = 0
  sStatus      = ""
  lcl_pathname = ""

  if iOrgID <> "" then
     sOrgID = clng(iOrgID)
  end if

  if iUserID <> "" then
     sUserID = clng(iUserID)
  end if

  if iStatus <> "" then
     sStatus = ucase(iStatus)
  end if

  lcl_orgname   = getOrgName(sOrgID)
  lcl_layout_id = getPrinter_CardLayout(sOrgID)

 'BEGIN: Set up folder path to image ------------------------------------------
 'If this is a REPRINT card then the image exists in the permanent folder
 'If this is a NEW card then use the image in the TEMP folder
  if sStatus = "REPRINT" OR sStatus = "PRINT" then
     lcl_pathname = "../images/MembershipCard_Photos/users/" & sOrgID & "_" & sUserID & ".jpg"
  else
     set fs = Server.CreateObject("Scripting.FileSystemObject")

    'Build the variables to set up the image paths
     lcl_file_directory = Application("membershipcard_filedirectory")
     lcl_temp_folder    = "/temp"
     lcl_temp_image     = "/" & sUserID & ".jpg"

     lcl_live_folder    = "/users"
     lcl_live_image     = "/" & sOrgID & "_" & sUserID & ".jpg"

    'Set up the "temp" and "live" server image paths
     lcl_img_temp_check = lcl_file_directory & lcl_temp_folder & lcl_temp_image
     lcl_img_live_check = lcl_file_directory & lcl_live_folder & lcl_live_image

    'Set up the "temp" and "live" browser image paths
     lcl_img_temp = "../images/MembershipCard_Photos" & lcl_temp_folder & lcl_temp_image
     lcl_img_live = "../images/MembershipCard_Photos" & lcl_live_folder & lcl_live_image

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
 'END: Set up folder path to image --------------------------------------------

 'BEGIN: Get user info --------------------------------------------------------
  sSQL = "SELECT u.userfname, "
  sSQL = sSQL & " u.userlname  "
  sSQL = sSQL & " FROM egov_users u "
  sSQL = sSQL & " WHERE u.userid = " & sUserID

  set oUserCardInfo = Server.CreateObject("ADODB.Recordset")
  oUserCardInfo.Open sSQL, Application("DSN"), 3, 1

  if NOT oUserCardInfo.EOF then
     lcl_fname = oUserCardInfo("userfname")
     lcl_lname = oUserCardInfo("userlname")
  end if

  oUserCardInfo.close
  set oUserCardInfo = nothing
 'END: Get user info ----------------------------------------------------------

'BEGIN: Build the barcode image -----------------------------------------------
'*** REQUIRES THE CodeBehind2.dll LOCATED IN THE "../bin" FOLDER! ***

 BarCodeImg = "barcode.aspx?FullAscii=1&X=1&Height=50&Value=" & sUserID

 if session("CARD_PRINT") = "Y" then
    lcl_watermark_class = "print"
   	lcl_card_outline    = "displayMembershipCard_print"
 else
    lcl_watermark_class = "display"
   	lcl_card_outline    = "displayMembershipCard"
 end if
'END: Build the barcode image -----------------------------------------------

'BEGIN: Display Card ----------------------------------------------------------
'NOTE: This layout uses the width/height of what is considered "layout2" and the look-n-feel of "layout1"
 response.write "<div class=""userlayout_print_margins"">" & vbcrlf
 response.write "<div id=""" & lcl_card_outline & """>" & vbcrlf
 response.write "<table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"">" & vbcrlf
 response.write "  <tr>" & vbcrlf
 response.write "      <td colspan=""2"" align=""center"" id=""card_title"">" & lcl_orgname & "</td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr valign=""top"">" & vbcrlf
 response.write "      <td class=""userlayout_card_text"">" & vbcrlf
 response.write "          <img name=""profilePic"" id=""profilePic"" src=""" & lcl_pathname & """ class=""profilePic"" />" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "      <td width=""100%"">" & vbcrlf
 response.write "          <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" height=""100%"" class=""layout1_card_text"">" & vbcrlf
 response.write "            <tr>" & vbcrlf
 response.write "                <td valign=""top"" align=""center"">" & vbcrlf
 response.write "                    <table border=""0"" cellspacing=""0"" cellpadding=""0"" width=""100%"" class=""layout1_card_text"">" & vbcrlf
 response.write "                      <tr><td><img name=""watermark_image"" id=""watermark_image"" src=""../images/poolpass_org_logos/watermark_" & sOrgID & ".jpg"" class=""watermark_" & lcl_watermark_class & """ /></td></tr>" & vbcrlf
 response.write "                      <tr align=""center"">" & vbcrlf
 response.write "               					      <td id=""member_name"">" & lcl_fname & "<br />" & lcl_lname & "</td>" & vbcrlf
 response.write "                      </tr>" & vbcrlf
 response.write "                    </table>" & vbcrlf
 response.write "                </td>" & vbcrlf
 response.write "            </tr>" & vbcrlf
 response.write "          </table>" & vbcrlf
 response.write "      </td>" & vbcrlf
 response.write "  </tr>" & vbcrlf
 response.write "  <tr><td colspan=""2"" align=""center"" height=""40px""><img src=""" & BarCodeImg & """ style=""height:40"" /></td></tr>" & vbcrlf
 response.write "</table>" & vbcrlf
 response.write "</div>" & vbcrlf
 response.write "</div>" & vbcrlf
'END: Display Card ------------------------------------------------------------
end sub
%>