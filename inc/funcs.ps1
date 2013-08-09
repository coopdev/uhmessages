function connectToDb()
{   
   $conn = New-Object MySql.Data.MySqlClient.MySqlConnection
   $conn.connectionString = $DbConnString   
   $conn.open()
   
   return ,$conn
}

# Fetch messages from database.
function fetchMessages($conn)
{   
   #$conn = connectToDb

   $stmt = New-Object MySql.Data.MySqlClient.MySqlCommand
   $stmt.connection = $conn
   $stmt.commandText = "SELECT * FROM messages WHERE is_processed = 0"

   $reader = $stmt.executeReader()
   
   # Must have comma before the return value in order to return
   # the actual object. Otherwise it returns an array for some
   # reason.
   return ,$reader
}


#######  PERSON MESSAGES #######

function addPerson($msg)
{  
   $person = @{}  
   $person.uhUuid = $msg.uhUuid
   
   # $uids = $msg.uids.childNodes
   # $uid  = $uids.item(0).InnerXml
   
   # First check if user already exists   
   $uhUuid = $person.uhUuid
   #$ADUser = Get-ADUser -Filter { employeeID -eq $uhUuid }
    
   # Generate random username since <addPerson> message does not include username.
   $person.username          = Get-Random
   $person.userPrincipalName = $person.username.toString() + $UpnSuffix
   $person.displayName       = checkEmptyVal($msg.displayName)   
   $person.firstName         = checkEmptyVal($msg.firstName)   
   $person.lastName          = checkEmptyVal($msg.lastName)
   $person.cn                = "$($person.firstName) $($person.lastName) ($($person.uhUuid))"
   $person.fullName          = checkEmptyVal($msg.fullName)
   
   
   # Create a random complex password.
   $pass = [System.Web.Security.Membership]::GeneratePassword(12,0)

   # Convert the generated password into a SecureString so 
   # that the -AccountPassword parameter in New-ADUser accepts the value.
   # -AsPlainText allows normal strings to be converted.
   # -Force will prevent a warning from showing.
   $person.securePass = ConvertTo-SecureString -AsPlainText $pass -Force 
   
   $date = Get-Date
   $year = $date.year
   $month = $date.month
   $day = $date.day   
   # Add the user to Active Directory.
   try {	             
      createADUser($person)  
      "$($person.lastName);$($person.firstName);$($person.uhUuid);$pass;$year-$month-$day" >> "accountInfo.txt"      
   }    
   catch {
      writeErrorsFor($person.uhUuid)
   }
   
   # Write account info to file.
   #"$($person.lastName);$($person.firstName);$($person.username);$pass" >> "$rootDir/accountInfo.txt"
   
}


function deletePerson($msg)
{
   $uhUuid = $msg.uhUuid
   
   # If user does not exists.
   if ( ($user = getADUser($uhUuid)) -eq $False ) {
      return
   }
   
   Disable-ADAccount -Identity $user
}


function retrofitPerson($msg)
{
   $uhUuid = $msg.uhUuid
   
   if ( $ADUser = getADUser($uhUuid) ) {
      write "User with uhuuid of $uhUuid already exists"      
      if ($ADUser.enabled -eq $False) {
         try {
            Set-ADUser -identity $ADUser -Enabled $True
         } catch {
            writeErrorsFor($person.uhUuid)
         }
      }
      return 
   } else {
      $actions = $msg.actions.childNodes   
      foreach ($a in $actions) {
         addPerson($a.messageData) 
      }   
   }   
}

#######  END PERSON MESSAGES #######


####### USERNAME MESSAGES #######

function addUsername($msg)
{
   $uhUuid = $msg.uhUuid
   
   # If user does not exists.
   if ( ($user = getADUser($uhUuid)) -eq $False ) {
      return
   }   
   
   try {
      Set-ADUser -Identity $user -samAccountName "$($msg.uid)"
      Set-ADUser -Identity $user -userPrincipalName "$($msg.uid)$($UpnSuffix)"
   } catch {
      writeErrorsFor($uhUuid)
   }
   
}

function retrofitUsername($msg)
{
   $uhUuid = $msg.uhUuid   
   $actions = $msg.actions.childNodes
   
   foreach ($a in $actions) {
      addUsername($a.messageData)
   }
}

function modifyUsernameUid($msg, $msgBefore)
{
   $newUsername = $msg.uid
   $oldUsername = $msgBefore.uid
   
   $user = Get-ADUser -Identity $oldUsername
   
   if ($user -eq $Null) {
      return
   }
   
   try {   
      Set-ADUser -Identity $user -samAccountName $newUsername
      Set-ADUser -Identity $user -userPrincipalName $newUsername
   } catch {
      writeErrorsFor($oldUsername)
   }
}


####### END USERNAME MESSAGES #######


####### AFFILIATION MESSAGES #######


# If user can be part of both Faculty and Staff, then change the code to:
#   store their <affID> in the employeeType attribute in the form of 
#   affID.group, e.g. 81823.faculty
function addAffiliation($msg)
{
   $uhUuid = $msg.uhUuid    
   
   $uids = $msg.uids.childNodes
   $uid  = $uids.item(0).InnerXml
   
   
   $aff = $msg.affiliation    
   $affId = $aff.affID
   $role = checkEmptyVal($aff.role)
   
   $group = getUsersGroup($aff)  
   
   if ( ($user = getADUser($uhUuid)) -ne $False ) {
      try {
         Add-ADGroupMember -Identity $group -Members $user
         $employeeType = "$affId.$group"
         if ($user.employeeType -eq $Null -or $user.employeeType.trim() -eq "") {      
            Set-ADUser -Identity $user -Add @{employeeType = $employeeType}
         } else {
            $employeeType = "$($user.employeeType), $employeeType"
            Set-ADUser -Identity $user -Replace @{employeeType = $employeeType}
         }   
      } catch  {
         writeErrorsFor($uhUuid)
      }      
   }     
   
   $group = getUsersUHGroup($aff) 
   
   if ( ($user = getUHADUser($uid)) -ne $False ) {
      try {
         Add-ADGroupMember -Identity $group -Members $user         
      } catch  {
         writeErrorsFor($uhUuid)
      }      
   }       
   
   # if ($user.memberOf.count -gt 0) {      
      # $currentGroup = $user.MemberOf[0]            
      # Remove-ADGroupMember -Identity $currentGroup -Members $user -Confirm:$false      
   # }   
   
}


function modifyAffiliation($msg)
{
   $uhUuid = $msg.uhUuid  
   
   $uids = $msg.uids.childNodes
   $uid  = $uids.item(0).InnerXml   
   
   if ( ($user = getADUser($uhUuid)) -eq $False ) {
      return
   }  

   $aff = $msg.affiliation    
   $newGroup = getUsersGroup($aff) 
   
   $affMap = $user.employeeType
   $affID  = $aff.affID

   $prevGroup = targetGroup $affMap $affID  

   $affMap = mapAffiliations $user.employeeType $aff.affID $newGroup 
   
   # It's wierd but you have to set this to a blank string first then concat the new
   # string onto it or it will say employeeType attribute can't be set.   
   $user.employeeType = ""
   $user.employeeType += $affMap      
  
   try {   
      Remove-ADGroupMember -Identity $prevGroup -Members $user -Confirm:$False
      Add-ADGroupMember -Identity $newGroup -Members $user       
      Set-ADUser -Identity $user -Replace @{employeeType = "$($user.employeeType)"}
   } catch {
      writeErrorsFor($uhUuid)
   }   
   
   
   if ( ($user = getUHADUser($uid)) -eq $False ) {
      return
   }  
   
   $newGroup = getUsersUHGroup($aff) 
   $prevUHGroup = targetUHGroup($prevGroup)

   try {   
      Remove-ADGroupMember -Identity $prevUHGroup -Members $user -Confirm:$False
      Add-ADGroupMember -Identity $newGroup -Members $user      
   } catch {
      writeErrorsFor($uhUuid)
   }   
   
   
}

function deleteAffiliation($msg)
{
   $uhUuid = $msg.uhUuid
   
   $uids = $msg.uids.childNodes
   $uid  = $uids.item(0).InnerXml
   
   if ( ($user = getADUser($uhUuid)) -eq $False ) {
      return
   }  
   
   $aff = $msg.affiliation   
   $group = getUsersGroup($aff)
   $affMap = $user.employeeType
   $affID  = $aff.affID
   
   $affMap = mapAffiliations $affMap $affID   
   

   if($affMap -eq $Null -or $affMap.trim() -eq "") { 
      $affMap = " "
   }     
   
   try {      
      Remove-ADGroupMember -Identity $group -Members $user -Confirm:$False
      Set-ADUser -Identity $user -Replace @{employeeType = "$affMap"}
   } catch {
      writeErrorsFor($uhUuid)
      #write $_.exception
   }  

   
   
   if ( ($user = getUHADUser($uid)) -eq $False ) {
      return
   }   
      
   $group = getUsersUHGroup($aff)         
   
   try {      
      Remove-ADGroupMember -Identity $group -Members $user -Confirm:$False      
   } catch {
      writeErrorsFor($uhUuid)
      #write $_.exception
   }   
}

function retrofitAffiliation($msg)
{
   $actions = $msg.actions.childNodes
   
   foreach ($a in $actions) {
      addAffiliation($a.messageData)
   }
}





function getUsersGroup($aff)
{   
   if ($aff.org -ne "hcc") {
      return $Others
   }
   
   $role = checkEmptyVal($aff.role)   
   $roleParts = $role.split(".")      
   
   switch ($roleParts[0]) 
   {
      "faculty"   { return $Faculty   }
      
      "staff"     { return $Staff     }

      "student"   { return $Students  }
  
      default     { return ""         }        
   }
}

function getUsersUHGroup($aff)
{   
   if ($aff.org -ne "hcc") {
      return $uhOthers
   }
   
   $role = checkEmptyVal($aff.role)   
   $roleParts = $role.split(".")      
   
   switch ($roleParts[0]) 
   {
      "faculty"   { return $uhFaculty   }
      
      "staff"     { return $uhStaff     }

      "student"   { return $uhStudents  }
  
      default     { return ""         }        
   }
}

function getADUser($employeeId)
{   
   $user = get-aduser -filter { employeeId -eq $employeeId } -Properties `
         "department", "title", "employeeType", "telephoneNumber", "otherTelephone", "facsimileTelephoneNumber", "otherFacsimileTelephoneNumber", `
         "physicalDeliveryOfficeName", "homePhone", "ipPhone", "otherHomePhone", "MemberOf", "otherIpPhone" `
         
   if ($user -eq $Null) {
      return $false
   }
   
   return $user
}

function getUHADUser($username)
{   
   try {
      $user = get-aduser -Identity $username -Server $uhADServer -Credential $uhADCreds -Properties "MemberOf", "employeeType"
   }
   catch {
      writeErrorsFor($username)
      return $false
   }   
   
   return ,$user
}

# Returns " " if $val is null or empty string.
# Returns $val otherwise.
function checkEmptyVal($val)
{
   if ($val -eq $Null -or $val -eq "") {
      return " "
   }
   
   return $val
}

function createADUser($user)
{
   $fname = $user.firstName
   $lname = $user.lastName
   
   New-ADUser -Name $user.cn -EmployeeId $user.uhUuid -SamAccountName $user.username `
      -UserPrincipalName $user.userPrincipalName -SurName $user.lastName -GivenName $user.firstName `
      -AccountPassword $user.securePass -Enabled $true -PasswordNeverExpires $true `
      -DisplayName "$fname $lname" -Path $Path
}

# $affMap represents the employeeType attribute which stores the map.
function mapAffiliations($affMap, $affID, $newGroup = $Null) {
   $currentAffMap = $affMap
   $affMap = "" # reset $affMap
   $mappings = $currentAffMap.split(",") # affID and group mappings, I.E 1234.HON CAMPUS Staff
   foreach ($m in $mappings) {
      $m = $m.trim()      
      $mParts = $m.split(".")      
      if ($mParts[0] -eq $affID) {
         $prevGroup = $mParts[1]                
      } else {
         # add back to map if the affID in question does not match the one being modified.
         $affMap += "$m, " 
      }      
   }
   
   if ($newGroup -ne $Null) {
      $affMap += "$($aff.affID).$newGroup, "
   }
   
   $affMap = $affMap -replace ', $' 
   
   return $affMap
}

# Returns the target group in the affiliation map.
function targetGroup($affMap,$affID)
{       
   $mappings = $affMap.split(",") # affID and group mappings, I.E 1234.HON CAMPUS Staff
   foreach ($m in $mappings) {
      $m = $m.trim()      
      $mParts = $m.split(".")      
      if ($mParts[0] -eq $affID) {
         $targetGroup = $mParts[1] 
         return $targetGroup         
      }
   }      
}

# Returns the appropriate UH group based on the passed in local group.
function targetUHGroup($localGroup)
{
   switch ($localGroup)
   {
      $Faculty  { return $uhFaculty  }
      
      $Staff    { return $uhStaff    }
      
      $Students { return $uhStudents }
      
      $Others   { return $uhOthers   }   
   }
}

function writeErrorsFor($uhuuid)
{
   
   $date = Get-Date
   $year = $date.year
   $month = $date.month
   $day = $date.day
   
   #$temp = $_ | select -expandproperty invocationinfo
   $errorInfo = $_ | select -expandproperty invocationinfo
   
   write "$year-$month-$day -- Error for $($uhuuid): $_ -- Line $($errorInfo.scriptLineNumber)" >> "errors.txt"  
   
}