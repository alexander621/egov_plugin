<!--#include file="config.asp"-->
<!--#include file="db.asp"-->
<%
	Class DBActionClass
		private cDB			'Database Class
        private tablefields 'table field columns
		public intColumns  'number of elements in array columns
        public tablename	'after done, I think it should be removed to config.asp 
		public currentpage
		public pagesize
		public TotalPages 
		public totalrecords
		public numstartid
		public numendid
		public strSQL
		private aSorted()
		private irec
		public sub Class_Initialize()
		    set cDB = new DBClass
			cDB.Open sConnForum
 	    tablename=conTABLENAME
		tablefields=conTABLEFIELDS
		intColumns=UBound(TableDescription)
		end sub
'---------------------------------------------------		
		public sub Class_Terminate()
			set cDB = Nothing
		end sub
'-----------------------------------------------------			
		'Opens a forum for a given ForumID and sets the class properties
public function createtable
 strSQL = "CREATE TABLE " & _
          tablename & " ("
 For intCounter = 0 To intColumns
'response.write "<br>counter/columns" & intCounter & intColumns &    tablefields(intCounter)
     strFields = strFields & _
         TableDescription(intCounter) 
       
If intCounter < intColumns Then
         strFields = strFields & ", "
End If
Next 
strSQL = strSQL & strFields & ")"
' response.write "strSQL=" &strSQL & intCounter & intColumns
' cDB.Execute(strSQL) 'it is too dangerout, so hidden it
 end function

'--------------------------------------------------------			
public function droptable
 strSQL = "drop table " & tablename
response.write "strSQL=" &strSQL & intCounter & intColumns
' cDB.Execute(strSQL) 'it is too dangerouts, do hidden it
 end function
'-------------------------------------------------
public function AddNew
strSQL = "select * from " & tablename & " where "&conTABLEFIELDS(0)&"=0"
set rsform=cDB.GetRS(strSQL)
rsform.AddNew
With request
	for each name in Request.Form
	rsform(name)=.form(name)
	Next 
end with
'	rsform(0)=getid 'member table can automatically add one in it
	rsform(intcolumns)=formatdatetime(now,2)&" "&formatdatetime(now,4)
rsform.Update  'No need to do here, since register is same for the administrator
destroy rsform
end function
'-----------------------------------------------------			
public function saveupdate(id)
strSQL = "select * from " & tablename & " where "&conTABLEFIELDS(0)&"="&id
set rsform=cDB.GetRS(strSQL)
for i =0 to rsform.fields.count-1
'response.write rsform.fields(i).name&"<br>"
next
With request
i=0
for each name in Request.Form
if name<>"ID" then 
rsform(name)=.form(name)
'response.write "<br>name="&name&" "&.form(name)
response.end

end if
i=i+1
Next 
end with
rsform.update
destroy rsform
end function

'-----------------------------------------------------			
public function getrecordarray
'strSQL has already been set in calling program before calling getrecordarray
set rs=cDB.GetRS(strSQL)
totalrecords = rs.RecordCount
if totalrecords=0 then
response.write "<br><FONT  COLOR=red>No property existing</FONT>"
response.end
end if
iFields = rs.Fields.Count
TotalPages = (totalrecords \ pagesize) + 1  '\means integer/integer
			if TotalPages < 1 then TotalPages = 1

if isNumeric(CurrentPage) then
				if CurrentPage < 1 then CurrentPage = 1
				if CurrentPage > TotalPages then CurrentPage = TotalPages
			else
				CurrentPage = 1
			end if
numstartid	= (CurrentPage-1) * PageSize
numendid	= IIf(numstartid + PageSize < totalrecords, numstartid+pagesize- 1, totalrecords - 1)
'response.write "rs:" & rs.fields("name_c") & "numendid:" &numendid&"<br>hello"
getrecordarray=rs.GetRows
destroy rs
end function
'-----------------------------------------------------		
public function deleterecord(id)
 strSQL = "delete from " & tablename &" where ID="&id
response.write "<br><B>Property deleted!!</B>"
cDB.Execute(strSQL)
 end function
end class
'-----------------------------------------------------	

function IIf(bCheck, sTrue, sFalse)
		if bCheck then IIf = sTrue Else IIf = sFalse
	end function
%>