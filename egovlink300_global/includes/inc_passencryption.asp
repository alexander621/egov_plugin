<%
function bytesToHex(aBytes)
    dim hexStr, x
    for x=1 to lenb(aBytes)
        hexStr= hex(ascb(midb( (aBytes),x,1)))
        if len(hexStr)=1 then hexStr="0" & hexStr
        bytesToHex=bytesToHex & hexStr
    next
end function

function sha256hashBytes(aBytes)
    Dim sha256
    set sha256 = CreateObject("System.Security.Cryptography.SHA256Managed")

    sha256.Initialize()
    'Note you MUST use computehash_2 to get the correct version of this method, and the bytes MUST be double wrapped in brackets to ensure they get passed in correctly.
    sha256hashBytes = sha256.ComputeHash_2( (aBytes) )
end function

function stringToUTFBytes(aString)
    Dim UTF8
    Set UTF8 = CreateObject("System.Text.UTF8Encoding")
    stringToUTFBytes = UTF8.GetBytes_4(aString)
end function

function GenerateRandomPassword()
dim intPWLength, intLoop, intCharType, strPwd
Const intMinPWLength = 6
Const intMaxPWLength = 10

' Generates a random number: 6, 7, 8, 9, or 10
' this number determines the length of the password. For instance, if
' the random number is 10 then, the password length will be 10
Randomize
intPWLength = int((intMaxPWLength - intMinPWLength + 1) * Rnd + intMinPWLength)
' now depending on the length of the password (dependent on the random
' number generated above), create random chracters between a-z, A-Z, or
' or 0-9 by using a for loop
for intLoop = 1 To intPWLength
' Generates a random number: 1, 2, or 3; where
' 1 gets a lowercase letter; 2 gets uppercase character, and
' 3 gets a number between 0 and 9
Randomize
intCharType = Int((3 * Rnd) + 1)

' now check if intCharType is 1, 2, or 3
select case intCharType
case 1
' get a lowercase letter a-z inclusive
Randomize
strPwd = strPwd & CHR(Int((25 * Rnd) + 97))
case 2
' get a uppercase letter A-Z inclusive
Randomize
strPwd = strPwd & CHR(Int((25 * Rnd) + 65))
case 3
' get a number between 0 and 9 inclusive
Randomize
strPwd = strPwd & CHR(Int((9 * Rnd) + 48))
end select
next

' return password
GenerateRandomPassword = strPwd
end function

function createHashedPassword(password)

	salt = bytesToHex(sha256hashBytes(stringToUTFBytes(GenerateRandomPassword())))
	hash = lcase(bytesToHex(sha256hashBytes(stringToUTFBytes(salt + password))))

	createHashedPassword = salt & hash
	

end function

Function ValidateUser(password, hashedUserPassword)

	if not isnull(hashedUserPassword) and hashedUserPassword <> "" then
		
        	salt = mid(hashedUserPassword, 1, 64)
        	validHashPw = mid(hashedUserPassword,65,64)
		'response.write salt & "<br />"
		'response.write validHashPw & "<br />"
		'response.end
	
        	passHash = lcase(bytesToHex(sha256hashBytes(stringToUTFBytes(salt + password))))
		'response.write "#" & passHash & "#<br />"
		'response.write "#" & validHashPw & "#<br />"
		'response.end

        	if StrComp(passHash,validHashPw) = 0 then
			ValidateUser = true
		end if
	end if
	    	
	if ValidateUser <> true then ValidateUser = false

End Function
%>
