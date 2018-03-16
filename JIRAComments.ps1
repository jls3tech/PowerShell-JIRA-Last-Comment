<#
@TechJLS3
JIRA Last Comment Extractor
3/16/2018
Call Script from Command Line:
e.g. .\JIRAComments.ps1 -TicketNo "JT-63,JT-253,jt-243,Jt-69,jT-169,BI-141,BM-1000"

#>

#Provide list of comma delimited JIRA Tickets
param (
    [string]$TicketNo
)

#Handles Authenication and API Call
function Get-Data([string]$username, [string]$password, [string]$url) {
      #Source: https://pallabpain.wordpress.com/2016/09/14/rest-api-call-with-basic-authentication-in-powershell/
      # Step 1. Create a username:password pair
      $credPair = "$($username):$($password)"
 
      # Step 2. Encode the pair to Base64 string
      $encodedCredentials = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($credPair))
 
      # Step 3. Form the header and add the Authorization attribute to it
      $headers = @{ Authorization = "Basic $encodedCredentials" }
 
      # Step 4. Make the GET request
      $responseData = Invoke-WebRequest -Uri $url -Method Get -Headers $headers -UseBasicParsing
 
      return $responseData
}

#Stuff
##User Name
$username = $env:UserName

##Password Handling
$SecurePassword = Read-Host "Enter Pass" -AsSecureString
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
$pass = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

#Splits Comma Delims
$ticketArray = $TicketNo -split ','

$joinedObject = foreach ($jid in $ticketArray){

    $error.clear()

    try { 
        $url = "https://jira/rest/api/2/issue/$jid/comment"
        $jira = Get-Data $username $pass $url
        $json = ($jira.Content | Out-String) | ConvertFrom-Json
        
        <#For Debugging
        Write-Host $jid
        $url
        #>
       
        #Get Comment Data
        $data = Write-Output $json.comments.body
        
        if ([string]::IsNullOrEmpty($data)){
            [pscustomobject]@{IssueID = $jid; Comment = "No Recorded Comment";} | Write-Output
            }
        else{
            [pscustomobject]@{IssueID = $jid; Comment = $data;} | Write-Output
            }
        }
    catch { 
       [pscustomobject]@{IssueID = $jid; Comment = "Error (Ticket likely does not exist)";} | Write-Output 
       }
}

#outputs file as CSV
$joinedObject|export-csv -Force "C:\Users\TechJLS3\Desktop\JIRAComments.csv"

<# For Debugging
$JoinedObject|Out-GridView
#>