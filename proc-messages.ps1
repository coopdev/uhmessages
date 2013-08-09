# Necessary for New-ADUser cmdlet.
Import-Module ActiveDirectory

# MySql connector.
[void][System.Reflection.Assembly]::LoadWithPartialName("MySql.Data")

# Load "System.Web" assembly in PowerShell console.
[Reflection.Assembly]::LoadWithPartialName("System.Web")

. ".\inc\config-vars.ps1"

# Init global variables
#$uhADCreds    = Get-Credential
$uhADServer   = ""

# UH AD groups.
$uhFaculty    = ""
$uhStaff      = ""
$uhStudents   = ""
$uhOthers     = ""

# Local groups.
$Faculty      = "Local Faculty"
$Staff        = "Local Staff"
$Students     = "Local Students"
$Others       = "Local Others"

# Database login is in config-vars.ps1

$Path         = 'OU=testUHMessages,DC=honolulu,DC=ads,DC=hawaii,DC=edu'
$UpnSuffix    = "@honolulu.ads.hawaii.edu"
# End global variables

. ".\inc\funcs.ps1"

$conn = connectToDb

$MysqlReader = fetchMessages($conn)

$msgIds = @()
while($MysqlReader.read()) {
   $msgIds += $MysqlReader.getString("id")
      
   $message = "<root>" + $MysqlReader.getString("message") + "</root>"
   $message = [xml]$message
   
   $message = $message.root.childNodes
   
   foreach ($m in $message) {
      $msg = $m.messageData   
      
      switch ($m.name) 
      {
         "addPerson"                  { addPerson $msg              }
         
         "retrofitPerson"             { retrofitPerson $msg         }          

         "deletePerson"               { deletePerson $msg           }         
         
         "addUsername"                { addUsername $msg            }
         
         "retrofitUsername"           { retrofitUsername $msg       }
         
         "modifyUsernameUid"          { modifyUsernameUid $msg $n.messageDataBefore  }
         
         "addAffiliation"             { addAffiliation $msg         }
         
         "retrofitAffiliation"        { retrofitAffiliation $msg    }
         
         "modifyAffiliation"          { modifyAffiliation $msg      } 
         
         "deleteAffiliation"          { deleteAffiliation $msg      }
      
      } # End switch.
   }
   
}

$MysqlReader.dispose()

$conn.close()


# Update messsages in DB to flag them as processed.


$conn = connectToDb

if ($msgIds.Count -gt 0) {
   $stmt = New-Object Mysql.Data.MySqlClient.MySqlCommand
   $stmt.connection = $conn
   
   
   $stmt.commandText = "UPDATE messages SET is_processed = 1 WHERE id in ("
   
   foreach ($id in $msgIds) {
      $stmt.commandText += "$id,"
   }
   
   # Get string up until final character (the last comma).
   $stmt.commandText = $stmt.commandText.subString(0, $stmt.commandText.length - 1)

   $stmt.commandText += ")"
   $stmt.executeNonQuery()
}

$conn.close()