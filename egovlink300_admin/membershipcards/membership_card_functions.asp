<%
function print_card(pMID)

 'Update print count ONLY when the card is printed.
  dim oCardExists, oCardPrint, lcl_card_printed, lcl_printed_count

 'determine if this is the first time that a card has been printed
  sSQLt = "SELECT card_printed, "
  sSQLt = sSQLt & " printed_count "
  sSQLt = sSQLt & " FROM egov_poolpassmembers "
  sSQLt = sSQLt & " WHERE memberid = " & clng(pMID)

  Set oCardExists = Server.CreateObject("ADODB.Recordset")
  oCardExists.Open sSqlt, Application("DSN"), 0, 1

  if oCardExists("card_printed") = "Y" then
   	 lcl_card_printed  = oCardExists("card_printed")
   	 lcl_printed_count = clng(oCardExists("printed_count")) + 1
  else
     lcl_card_printed  = "Y"
	    lcl_printed_count = 1
  end if

  'update the egov_poolpassmembers table for the member
   sSqlu = "UPDATE egov_poolpassmembers "
   sSqlu = sSqlu & " SET printed_count = " & lcl_printed_count & ", "
   sSqlu = sSqlu &     " card_printed = '" & lcl_card_printed  & "' "
   sSqlu = sSqlu & " WHERE memberid = " & clng(pMID)

   set oCardPrint = Server.CreateObject("ADODB.Recordset")
   oCardPrint.Open sSqlu, Application("DSN"), 0, 1

end function

'-----------------------------------------------------------------------
function save_card(pMID)

'Set up the transfer the file from the TEMP folder to the permanent folder
  dim fs, pathname, lcl_temp_image, lcl_from_file, lcl_to_file

  lcl_pathname = Application("membershipcard_filedirectory")

'DVLP
' lcl_pathname = "c:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\images\MembershipCard_Photos"

'TEST"
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\egovlink300_admin\images\MembershipCard_Photos"

'PROD
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink300_admin\images\MembershipCard_Photos"

  lcl_temp_folder = "\temp"
  lcl_image       = "\"&pMID&".jpg"

  set fs = Server.CreateObject("Scripting.FileSystemObject")

 'Set the parameters for MoveFile
  lcl_from_file = lcl_pathname&lcl_temp_folder&lcl_image
  lcl_to_file   = lcl_pathname&lcl_image
  lcl_overwrite = TRUE

 'Move the image
  if (fs.FileExists(lcl_from_file)) = true then
      fs.CopyFile lcl_from_file,lcl_to_file,TRUE
  end if

 'Update print count ONLY when the card is printed.
  dim oCardSave

 'update the egov_poolpassmembers table for the member
  sSqlm = "UPDATE egov_poolpassmembers "
  sSqlm = sSqlm & " SET card_printed = 'Y' "
  sSqlm = sSqlm & " WHERE memberid = " & clng(pMID)

  set oCardSave = Server.CreateObject("ADODB.Recordset")
  oCardSave.Open sSqlm, Application("DSN"), 0, 1

 'oCardSave.close
  set oCardSave = nothing

 'Delete the image from the temp directory.
  if (fs.FileExists(lcl_from_file)) = true then
      fs.DeleteFile lcl_from_file
  end if

  set fs = nothing

end function

'-----------------------------------------------------------------------
function remove_image(pMID)
'remove the image from the TEMP folder on our server
 Dim fs, lcl_file

'DVLP
' lcl_file = "c:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

'TEST"
' lcl_pathname = "d:\wwwroot\www.cityegov.com\egovlink release QA Test Environment 4.0.0\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

'PROD
' lcl_file = "d:\wwwroot\www.cityegov.com\egovlink300_admin\images\MembershipCard_Photos\temp\"&pMID&".jpg"

 lcl_file = Application("membershipcard_filedirectory")

 Set fs = Server.CreateObject("Scripting.FileSystemObject")
 if (fs.FileExists(lcl_file)) = true then
      fs.deletefile(lcl_file)
 end if

 set fs = nothing

end function

'------------------------------------------------------------------------------
function FormatPhone( Number )
	If Len(Number) = 10 Then
		FormatPhone = "(" & Left(Number,3) & ") " & Mid(Number, 4, 3) & "-" & Right(Number,4)
	Else
		FormatPhone = Number
	End If
end function

'------------------------------------------------------------------------------
sub createCardLayout(iTitle, _
                     iSubTitle, _
                     iYearText, _
                     iDisplayDate, _
                     iCustomImageURL, _
                     iQuote, _
                     iColor1, _
                     iColor2, _
                     iTextColor1, _
                     iTextColor2, _
                     iBackText, _
                     iBackTextColor, _
                     iLayoutName, _
                     iAction, _
                     iIsDisabled, _
                     iConfirmSoundFile)

 'Validate the columns that are to be inserted
  if iTitle <> "" then
     iTitle = "'" & dbsafe(iTitle) & "'"
  else
     iTitle = "NULL"
  end if

  if iSubTitle <> "" then
     iSubTitle = "'" & dbsafe(iSubTitle) & "'"
  else
     iSubTitle = "NULL"
  end if

  if iYearText <> "" then
     iYearText = "'" & dbsafe(iYearText) & "'"
  else
     iYearText = "NULL"
  end if

  if iDisplayDate = "" then
     iDisplayDate = 0
  end if

  if iCustomImageURL <> "" then
     iCustomImageURL = "'" & dbsafe(iCustomImageURL) & "'"
  else
     iCustomImageURL = "NULL"
  end if

  if iQuote <> "" then
     iQuote = "'" & dbsafe(iQuote) & "'"
  else
     iQuote = "NULL"
  end if

  if iColor1 <> "" then
     iColor1 = "'" & dbsafe(iColor1) & "'"
  else
     iColor1 = "NULL"
  end if

  if iColor2 <> "" then
     iColor2 = "'" & dbsafe(iColor2) & "'"
  else
     iColor2 = "NULL"
  end if

  if iTextColor1 <> "" then
     iTextColor1 = "'" & dbsafe(iTextColor1) & "'"
  else
     iTextColor1 = "NULL"
  end if

  if iTextColor2 <> "" then
     iTextColor2 = "'" & dbsafe(iTextColor2) & "'"
  else
     iTextColor2 = "NULL"
  end if

  if iBackText <> "" then
     iBackText = "'" & dbsafe(iBackText) & "'"
  else
     iBackText = "NULL"
  end if

  if iBackTextColor <> "" then
     iBackTextColor = "'" & dbsafe(iBackTextColor) & "'"
  else
     iBackTextColor = "NULL"
  end if

  if iLayoutName <> "" then
     iLayoutName = "'" & dbsafe(iLayoutName) & "'"
  end if

  if iIsDisabled = "on" then
     iIsDisabled = 0
  else
     iIsDisabled = 1
  end if

  if iConfirmSoundFile <> "" then
     iConfirmSoundFile = "'" & dbsafe(iConfirmSoundFile) & "'"
  else
     iConfirmSoundFile = "NULL"
  end if

 'Setup the insert statement
  sSQL = "INSERT INTO egov_membershipcard_layout ("
  sSQL = sSQL & "orgid,"
  sSQL = sSQL & "title,"
  sSQL = sSQL & "subtitle,"
  sSQL = sSQL & "year_text,"
  sSQL = sSQL & "display_date,"
  sSQL = sSQL & "custom_image_url,"
  sSQL = sSQL & "quote_text,"
  sSQL = sSQL & "main_color,"
  sSQL = sSQL & "secondary_color,"
  sSQL = sSQL & "main_text_color,"
  sSQL = sSQL & "secondary_text_color,"
  sSQL = sSQL & "back_text,"
  sSQL = sSQL & "back_text_color,"
  sSQL = sSQL & "layoutname,"
  sSQL = sSQL & "isDisabled,"
  sSQL = sSQL & "confirmsoundfile"
  sSQL = sSQL & ") VALUES ("
  sSQL = sSQL & session("orgid") & ", "
  sSQL = sSQL & iTitle           & ", "
  sSQL = sSQL & iSubTitle        & ", "
  sSQL = sSQL & iYearText        & ", "
  sSQL = sSQL & iDisplayDate     & ", "
  sSQL = sSQL & iCustomImageURL  & ", "
  sSQL = sSQL & iQuote           & ", "
  sSQL = sSQL & iColor1          & ", "
  sSQL = sSQL & iColor2          & ", "
  sSQL = sSQL & iTextColor1      & ", "
  sSQL = sSQL & iTextColor2      & ", "
  sSQL = sSQL & iBackText        & ", "
  sSQL = sSQL & iBackTextColor   & ", "
  sSQL = sSQL & iLayoutName      & ", "
  sSQL = sSQL & iIsDisabled      & ", "
  sSQL = sSQL & iConfirmSoundFile
  sSQL = sSQL & ")"

  set oInsertCard = Server.CreateObject("ADODB.Recordset")
  oInsertCard.Open sSQL, Application("DSN"), 3, 1

 'Retrieve the cardid that was just inserted
  sSQLid = "SELECT IDENT_CURRENT('egov_membershipcard_layout') as NewID"
  oInsertCard.Open sSQLid, Application("DSN"), 3, 1
  lcl_identity = oInsertCard.Fields("NewID").value

  set oInsertCard = nothing

  if iAction = "COPY LAYOUT" then
     lcl_success_msg = "SC"
  else
     lcl_success_msg = "SA"
  end if

  response.redirect "card_layout_maint.asp?cardid=" & lcl_identity & "&success=" & lcl_success_msg

end sub

'------------------------------------------------------------------------------
sub updateCardLayout(iCardID, _
                     iTitle, _
                     iSubTitle, _
                     iYearText, _
                     iDisplayDate, _
                     iCustomImageURL, _
                     iQuote, _
                     iColor1, _
                     iColor2, _
                     iTextColor1, _
                     iTextColor2, _
                     iBackText, _
                     iBackTextColor, _
                     iLayoutName, _
                     iIsDisabled, _
                     iConfirmSoundFile)

  if iTitle <> "" then
     iTitle = dbsafe(iTitle)
  end if

  if iSubTitle <> "" then
     iSubTitle = dbsafe(iSubTitle)
  end if

  if iYearText <> "" then
     iYearText = dbsafe(iYearText)
  end if

  if iCustomImageURL <> "" then
     iCustomImageURL = dbsafe(iCustomImageURL)
  end if

  if iQuote <> "" then
     iQuote = dbsafe(iQuote)
  end if

  if iColor1 <> "" then
     iColor1 = dbsafe(iColor1)
  end if

  if iColor2 <> "" then
     iColor2 = dbsafe(iColor2)
  end if

  if iTextColor1 <> "" then
     iTextColor1 = dbsafe(iTextColor1)
  end if

  if iTextColor2 <> "" then
     iTextColor2 = dbsafe(iTextColor2)
  end if

  if iBackText <> "" then
     iBackText = dbsafe(iBackText)
  end if

  if iBackTextColor <> "" then
     iBackTextColor = dbsafe(iBackTextColor)
  end if

  if iLayoutName <> "" then
     iLayoutName = dbsafe(iLayoutName)
  end if

  if iIsDisabled = "on" then
     iIsDisabled = 0
  else
     iIsDisabled = 1
  end if

  if iConfirmSoundFile <> "" then
     iConfirmSoundFile = "'" & dbsafe(iConfirmSoundFile) & "'"
  else
     iConfirmSoundFile = "NULL"
  end if

  sSQL = "UPDATE egov_membershipcard_layout SET "
  sSQL = sSQL & "title = '"                & iTitle            & "', "
  sSQL = sSQL & "subtitle = '"             & iSubTitle         & "', "
  sSQL = sSQL & "year_text = '"            & iYearText         & "', "
  sSQL = sSQL & "display_date = "          & iDisplayDate      & ", "
  sSQL = sSQL & "custom_image_url = '"     & iCustomImageUrl   & "', "
  sSQL = sSQL & "quote_text = '"           & iQuote            & "', "
  sSQL = sSQL & "main_color = '"           & iColor1           & "', "
  sSQL = sSQL & "secondary_color = '"      & iColor2           & "', "
  sSQL = sSQL & "main_text_color = '"      & iTextColor1       & "', "
  sSQL = sSQL & "secondary_text_color = '" & iTextColor2       & "', "
  sSQL = sSQL & "back_text = '"            & iBackText         & "', "
  sSQL = sSQL & "back_text_color = '"      & iBackTextColor    & "', "
  sSQL = sSQL & "layoutname = '"           & iLayoutName       & "', "
  sSQL = sSQL & "isDisabled = "            & iIsDisabled       & ", "
  sSQL = sSQL & "confirmsoundfile = "      & iConfirmSoundFile
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND cardid = " & iCardID

  set oUpdateCard = Server.CreateObject("ADODB.Recordset")
  oUpdateCard.Open sSQL, Application("DSN"), 3, 1

  set oUpdateCard = nothing

  response.redirect "card_layout_maint.asp?cardid=" & iCardID & "&success=SU"

end sub

'------------------------------------------------------------------------------
sub deleteCardLayout(iCardID)

  lcl_totalcnt = checkCardonRatePurchased(iCardID)

  if lcl_totalcnt = 0 then
     sSQL = "DELETE FROM egov_membershipcard_layout WHERE cardid = " & iCardID

     set oDelCard = Server.CreateObject("ADODB.Recordset")
     oDelCard.Open sSQL, Application("DSN"), 3, 1

  else
     response.redirect "card_layout_maint.asp?cardid=" & iCardID & "&success=NODEL"
  end if

  set oDelCard = nothing

  response.redirect "card_layout_maint.asp?success=SD"

end sub

'------------------------------------------------------------------------------
function checkCardonRatePurchased(iCardID)

  lcl_return = 0

  sSQL = "SELECT count(cardid) AS TotalCnt "
  sSQL = sSQL & " FROM egov_poolpassrates r, egov_poolpasspurchases p "
  sSQL = sSQL & " WHERE r.rateid = p.rateid "
  sSQL = sSQL & " AND r.orgid = " & session("orgid")
  sSQL = sSQL & " AND r.cardid = " & iCardID

  set oRateCheck = Server.CreateObject("ADODB.Recordset")
  oRateCheck.Open sSQL, Application("DSN"), 3, 1

  lcl_return = oRateCheck("TotalCnt")

  oRateCheck.close
  set oRateCheck = nothing

  checkCardonRatePurchased = lcl_return

end function

'------------------------------------------------------------------------------
sub alreadyExists(iCardID, _
                  iTitle, _
                  iSubTitle, _
                  iYearText, _
                  iDisplayDate, _
                  iCustomImageURL, _
                  iQuote, _
                  iColor1, _
                  iColor2, _
                  iTextColor1, _
                  iTextColor2, _
                  iBackText, _
                  iBackTextColor, _
                  iLayoutName, _
                  iIsDisabled, _
                  iConfirmSoundFile)

  response.redirect "card_layout_maint.asp?cardid="                 & iCardID           & _
                                         "&p_new_p_new_title="      & iTitle            & _
                                         "&p_new_subtitle="         & iSubTitle         & _
                                         "&p_new_year_text="        & iYearText         & _
                                         "&p_new_display_date="     & iDisplayDate      & _
                                         "&p_new_custom_image_url=" & iCustomImageURL   & _
                                         "&p_new_quote="            & iQuote            & _
                                         "&p_new_color1="           & iColor1           & _
                                         "&p_new_color2="           & iColor2           & _
                                         "&p_new_text_color1="      & iTextColor1       & _
                                         "&p_new_text_color2="      & iTextColor2       & _
                                         "&p_new_back_text="        & iBackText         & _
                                         "&p_new_back_text_color="  & iBackTextColor    & _
                                         "&p_new_layoutname="       & iLayoutName       & _
                                         "&p_new_isDisabled="       & iIsDisabled       & _
                                         "&p_new_confirmsoundfile=" & iConfirmSoundFile & _
                                         "&success=AE"

end sub

'------------------------------------------------------------------------------
function getMaxCardID()
  lcl_return = 0

  sSQL = "SELECT isnull(max(cardid),0) AS maxCardID "
  sSQL = sSQL & " FROM egov_membershipcard_layout "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")

  set oMaxCard = Server.CreateObject("ADODB.Recordset")
  oMaxCard.Open sSQL, Application("DSN"), 3, 1

  if not oMaxCard.eof then
     lcl_return = oMaxCard("maxCardID")
  end if

  oMaxCard.close
  set oMaxCard = nothing

  getMaxCardID = lcl_return

end function

'------------------------------------------------------------------------------
sub displayCardLayoutOptions(iCardID, _
                             iShowOnlyActive)

  sSQL = "SELECT cardid, isnull(layoutname,'[No Layout Name Available]') AS layoutname "
  sSQL = sSQL & " FROM egov_membershipcard_layout "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")

  if iShowOnlyActive = "Y" then
     sSQL = sSQL & " AND isDisabled = 0 "
  end if

  sSQL = sSQL & " ORDER BY UPPER(layoutname) "

  set oCardOptions = Server.CreateObject("ADODB.Recordset")
  oCardOptions.Open sSQL, Application("DSN"), 3, 1

  if not oCardOptions.eof then
     do while not oCardOptions.eof

        lcl_selected_cardlayout = ""

        if iCardID <> "" then
           if clng(iCardID) = oCardOptions("cardid") then
              lcl_selected_cardlayout = " selected=""selected"""
           end if
        end if

        response.write "  <option value=""" & oCardOptions("cardid") & """" & lcl_selected_cardlayout & ">" & oCardOptions("layoutname") & "</option>" & vbcrlf

        oCardOptions.movenext
     loop
  else
     response.write "  <option value=""0"">Default Layout</option>" & vbcrlf
  end if

  oCardOptions.close
  set oCardOptions = nothing

end sub

'------------------------------------------------------------------------------
function checkDuplicateLayoutName(iCardID, _
                                  iLayoutName)
  lcl_return = False

  sSQL = "SELECT cardid, layoutname "
  sSQL = sSQL & " FROM egov_membershipcard_layout "
  sSQL = sSQL & " WHERE orgid = " & session("orgid")
  sSQL = sSQL & " AND UPPER(layoutname) = '" & UCASE(iLayoutName) & "'"
  sSQL = sSQL & " AND cardid <> " & iCardID

  set oCardExists = Server.CreateObject("ADODB.Recordset")
  oCardExists.Open sSQL, Application("DSN"), 3, 1

  if not oCardExists.eof then
     lcl_return = True
  end if

  oCardExists.close
  set oCardExists = nothing

  checkDuplicateLayoutName = lcl_return

end function

'------------------------------------------------------------------------------
function getOriginalLayoutName(iCardID, _
                               iOriginalLayout)
  i          = ""
  lcl_return = iOriginalLayout

  do until not checkDuplicateLayoutName(iCardID,lcl_return)
     if i = "" then
        i = 0
     end if

     i = i + 1
     lcl_return = replace(lcl_return, lcl_return, iOriginalLayout & "_" & i)
  loop

  getOriginalLayoutName = lcl_return

end function

'------------------------------------------------------------------------------
function getRateID(iPoolPassID)
  lcl_return = ""

  if iPoolPassID <> "" then
     sSQL = "SELECT rateid "
     sSQL = sSQL & " FROM egov_poolpasspurchases "
     sSQL = sSQL & " WHERE orgid = " & session("orgid")
     sSQL = sSQL & " AND poolpassid = " & iPoolPassID

     set oRateID = Server.CreateObject("ADODB.Recordset")
     oRateID.Open sSQL, Application("DSN"), 3, 1

     if not oRateID.eof then

        lcl_return = oRateID("rateid")

        oRateID.close
        set oRateID = nothing
     end if

  end if

  getRateID = lcl_return

end function

'------------------------------------------------------------------------------
function checkMemberIDScanned(iMemberID, _
                              iRateID)
  lcl_return = False

  if iMemberID <> "" AND iRateID <> "" then
     sSQL = "SELECT count(attendancelogid) as TotalScans "
     sSQL = sSQL & " FROM egov_pool_attendance_log "
     sSQL = sSQL & " WHERE orgid = "  & session("orgid")
     sSQL = sSQL & " AND memberid = " & iMemberID
     sSQL = sSQL & " AND rateid = "   & iRateID

     set oScanExists = Server.CreateObject("ADODB.Recordset")
     oScanExists.Open sSQL, Application("DSN"), 3, 1

     if oScanExists("TotalScans") > 0 then
        lcl_return = True
     end if

     oScanExists.close
     set oScanExists = nothing

  end if

  checkMemberIDScanned = lcl_return

end function

'------------------------------------------------------------------------------
function formatUserNameForURL(iUserName)

  lcl_return = ""

  if iUserName <> "" then
     lcl_return = iUserName
     lcl_return = replace(lcl_return,"&","<<AMP>>")
     lcl_return = replace(lcl_return,"'","<<QUT>>")
     lcl_return = replace(lcl_return,"""","<<DBL>>")
  end if

  formatUserNameForURL = lcl_return

end function

'------------------------------------------------------------------------------
function formatUserNameForPage(iUserName)

  lcl_return = ""

  if iUserName <> "" then
     lcl_return = iUserName
     lcl_return = replace(lcl_return,"<<AMP>>","&")
     lcl_return = replace(lcl_return,"<<QUT>>","'")
     lcl_return = replace(lcl_return,"<<DBL>>","""")
  end if

  formatUserNameForPage = lcl_return

end function

'------------------------------------------------------------------------------
function dbsafe(p_value)

  lcl_return = ""

  if p_value <> "" then
     lcl_return = replace(p_value,"'","''")
  end if

  dbsafe = lcl_return

end function
%>
