function save_card(pMID)

'Set up the transfer the file from the TEMP folder to the permanent folder
 dim fs, pathname, lcl_temp_image, lcl_from_file, lcl_to_file
 lcl_pathname    = "c:\www_server_root\egovlink\egovlink release 4.0.0\egovlink300_admin\images\MembershipCard_Photos"
 lcl_temp_folder = "\temp"
 lcl_image       = "\"&pMID&".jpg"

 set fs = Server.CreateObject("Scripting.FileSystemObject")

'Set the parameters for MoveFile
 lcl_from_file = lcl_pathname&lcl_temp_folder&lcl_image
 lcl_to_file   = lcl_pathname&lcl_image
 lcl_overwrite = TRUE

 'Move the image
 fs.CopyFile lcl_from_file,lcl_to_file,TRUE

 'Update print count ONLY when the card is printed.
  dim oCardCheck, oCardSave, lcl_card_printed, lcl_printed_count

 'determine if this is the first time that a card has been printed
  sSqlc = "SELECT card_printed, printed_count FROM egov_poolpassmembers WHERE memberid = " & clng(pMID)

  Set oCardCheck = Server.CreateObject("ADODB.Recordset")
  oCardCheck.Open sSqlc, Application("DSN"), 0, 1

  if oCardCheck("card_printed") = "Y" then
	 lcl_card_printed         = oCardCheck("card_printed")
	 lcl_printed_count        = clng(oCardCheck("printed_count")) + 1
  else
     lcl_card_printed         = "Y"
	 lcl_printed_count        = 1
  end if

  'update the egov_poolpassmembers table for the member
   sSqlm = "UPDATE egov_poolpassmembers "
   sSqlm = sSqlm & " SET printed_count = " & lcl_printed_count & ", "
   sSqlm = sSqlm &     " card_printed = '" & lcl_card_printed  & "' "
   sSqlm = sSqlm & " WHERE memberid = " & clng(pMID)

   set oCardSave = Server.CreateObject("ADODB.Recordset")
   oCardSave.Open sSqlm, Application("DSN"), 0, 1

'   oCardCheck.close
'   set oCardCheck = nothing

'   oCardSave.close
'   set oCardSave = nothing

end function
