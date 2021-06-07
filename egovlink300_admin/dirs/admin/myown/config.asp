<%
dim conTABLEFIELDS()
dim l_length()
dim l_type()
'dim fields_description()
const cPageSize =10
const IDCol=0
const conTableWidth=500
const conTABLENAME="GroupsPermissions"
const DeleteTablename="RolesPermissions"
sConnForum =Application("DSN")
l_register=array(0,1,1)
l_display=array(1,1,1)
l_modify=array(0,1,1)
const langAdminTitle	= "Manage Committee"
const langAdminCloseWindow	= "Close Window"
Const langNewRecord		= "New Committee"
Const langRecordList	= "Committee List"
Const langSubmitNewRecord = "Submit New Record"
Const langUpdateRecord = "Update the Record"
Const langGoBack			= "Go Back"
Const langDelete			= "Delete"
const langResponseAfterPostSave="Successfully add thre record"
const langFirst ="First"
const langPrev ="Prev"
const langNext ="Next"
const langLast ="Last"
'=================================================
set conn = Server.CreateObject("ADODB.Connection")
conn.Open sConnForum
strSQL = "SELECT syscolumns.name, syscolumns.type, syscolumns.length, " & _
         "syscolumns.isnullable FROM sysobjects " & _
         "INNER JOIN syscolumns ON sysobjects.id = syscolumns.id " & _
         "where sysobjects.name = '"&conTABLENAME& _
         "' ORDER BY syscolumns.colid"
Set objRS = Server.CreateObject("ADODB.Recordset")
set objRS.ActiveConnection = conn
objRS.CursorLocation = 3 ' the CursorLocation and CursorType must be defined firstly, otherwise, the
objRS.CursorType = 3     ' Recordcount will be -1
objRS.Open strSQL
MaxFields=objRS.recordcount
redim conTABLEFIELDS(MaxFields-1)
redim l_length(MaxFields-1)
redim l_type(MaxFields-1)
'redim fields_description(MaxFields)
'response.write "<br>MaxFields="&MaxFields
'=============================
objRS.MoveFirst
for i=0 to MaxFields-1
conTABLEFIELDS(i)=objRS("name")
l_length(i)=objRS("length")
l_type(i)=objRS("type")
'---------------------
select case l_type(i)
case 56, 38
'tinyint,integer, 56 not null, 38, null
case 39
' varchar, char, not null, null
case 61,111
	l_length(i)=20
' date time, 61, not nulll, 111, null
case 35
	l_length(i)=1000 ' we are going to disply a textarea depending the l_length >150
' text not null, null
case 50
	l_length(i)=5
' bit, not null, null
case else
end select
'--------------------
'response.write "<br>Field="&conTABLEFIELDS(i)&" length="&l_length(i)&"  type="&l_type(i)
objRS.MoveNext
next
'=============================
fields_description= conTABLEFIELDS 
 ' a very strange thing is that, you cannot declare the fields_descriton
 ' before you set that, otherwise, it reports wrong.
'TableDescription=array( _ 
'"GroupID int PRIMARY KEY","OrgID int not null","groupname [varchar](50)", _
'"GroupDescription [varchar](100)","createtime datetime")  
'+++++++++++++++++++++++++++++++++++++++++++++++++
%>
<%
	function IIf(bCheck, sTrue, sFalse)
		if bCheck then IIf = sTrue Else IIf = sFalse
	end function
%>