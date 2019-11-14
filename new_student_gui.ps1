#Your XAML goes here :)
$inputXML = @"
<Window x:Class="newStudent.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:newStudent"
        mc:Ignorable="d"
        Title="NewStudentForm" Height="566.41" Width="359.259">
    <Grid Margin="0,0,2,0">
        <TextBlock HorizontalAlignment="Left" Margin="10,9,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="215" FontSize="20" Text="New Student Account"/>
        <TextBox x:Name="TextFirstName" HorizontalAlignment="Left" Height="23" Margin="72,43,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" VerticalContentAlignment="Bottom"/>
        <TextBox x:Name="TextLastName" HorizontalAlignment="Left" Height="23" Margin="72,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" VerticalContentAlignment="Bottom"/>
        <ComboBox x:Name="ComboCohort" HorizontalAlignment="Left" Margin="72,158,0,0" VerticalAlignment="Top" Width="120"/> 
        <ComboBox x:Name="ComboHouse" HorizontalAlignment="Left" Margin="72,185,0,0" VerticalAlignment="Top" Width="120"/>
        <TextBox x:Name="TextPassword" HorizontalAlignment="Left" Height="23" Margin="72,214,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" VerticalContentAlignment="Bottom"/>
        <TextBlock x:Name="DTextPasswordInstructions" HorizontalAlignment="Left" Margin="10,242,0,0" TextWrapping="Wrap" Text="Leave blank to generate a password" VerticalAlignment="Top"/>
        <TextBlock x:Name="DTextFirstName" HorizontalAlignment="Left" Margin="10,50,0,0" TextWrapping="Wrap" Text="First Name" VerticalAlignment="Top"/>
        <TextBlock x:Name="DTextLastName" HorizontalAlignment="Left" Margin="10,78,0,0" TextWrapping="Wrap" Text="Last Name" VerticalAlignment="Top"/>
        <TextBlock x:Name="DTextCohort" HorizontalAlignment="Left" Margin="10,164,0,0" TextWrapping="Wrap" Text="Cohort" VerticalAlignment="Top"/>
        <TextBlock x:Name="DTextHouse" HorizontalAlignment="Left" Margin="10,191,0,0" TextWrapping="Wrap" Text="House" VerticalAlignment="Top"/>
        <TextBlock x:Name="DTextPassword" HorizontalAlignment="Left" Margin="10,221,0,0" TextWrapping="Wrap" Text="Password" VerticalAlignment="Top"/>
        <Button x:Name="ButtonGenerateUsername" Content="Generate Username" HorizontalAlignment="Left" Margin="10,130,0,0" VerticalAlignment="Top" Width="120"/>
        <TextBox x:Name="TextUserName" HorizontalAlignment="Left" Height="23" Margin="72,99,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" VerticalContentAlignment="Bottom" IsEnabled="False"/>
        <TextBlock x:Name="DTextUsername" HorizontalAlignment="Left" Margin="10,106,0,0" TextWrapping="Wrap" Text="Username" VerticalAlignment="Top"/>
        <Button x:Name="ButtonGeneratePassword" Content="Generate Password" HorizontalAlignment="Left" Margin="10,263,0,0" VerticalAlignment="Top" Width="120"/>
        <Button x:Name="ButtonCreateUser" Content="Create User" HorizontalAlignment="Left" Margin="10,288,0,0" VerticalAlignment="Top" Width="120"/>
        <TextBox x:Name="TboxLog" HorizontalAlignment="Left" Height="148" Margin="10,345,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="316" VerticalScrollBarVisibility="Auto"/>
        <TextBlock x:Name="DTextLog" HorizontalAlignment="Left" Margin="10,324,0,0" TextWrapping="Wrap" Text="Log" VerticalAlignment="Top"/>

    </Grid>
</Window>
"@ 
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{
    $Form=[Windows.Markup.XamlReader]::Load( $reader )
}
catch{
    throw
}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{"trying item $($_.Name)";
    try {Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }
 
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
get-variable WPF*
}
 
Get-FormVariables
 
#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================

# Get some data and declare some variables
$i=0
$currentYear = get-date -f yyyy
$global:todayDate = get-date -f "dd-MM-yy hh:MM tt"
$global:cohortCount = [int]$currentYear
$maxYear = $global:cohortCount+5
############################
# Variables declared here for ease of planning and testing. Do not fill these with data
############################
$global:firstname = ""
$global:lastname = ""
$global:fullName = ""
$global:cohort = ""
$global:house = ""
############################
# These variables should all be filled. They will be used when creating the user
############################
$global:city  = ""
$global:company = ""
$global:country = ""
$global:department = ""
$global:description = ""
$global:streetAddress = ""
$global:state = ""
$global:poBox = ""
$global:scriptPath = ""
# Format for path is: "OU=OUName,DC=consoto,DC=local"
$global:path = "OU=Users,DC=contoso,DC=local"
$global:type = "User"
$global:emailTail = "@contoso.com"
$global:homedrive = "H"
############################
# Email notifcation configuration
############################
$global:emailCc = "gatesB@contoso.com", "cookt@apple.com"
$global:smtpServer = "10.0.1.254"
$global:emailFrom = "newuser@contoso.com"
############################

# Add cohort years from current year to 1 year before Kindy
Do {
$WPFComboCohort.AddChild("$global:cohortCount")   
$i++
$global:cohortCount++
Write-Host $i}
While($i -ne 14)                  

# Add School houses to combo dropdown box
$WPFComboHouse.AddChild("Blue") 
$WPFComboHouse.AddChild("Green") 
$WPFComboHouse.AddChild("Yellow") 
$WPFComboHouse.AddChild("Red") 
$WPFComboHouse.AddChild("No House")

Function SetUserName {
param($firstname, $lastname)
$global:fullName = "$($firstname) $($lastname)"
# Declare the username variable here
$global:username = ""
# Clean up the lastname for the username, remove any special characters or numbers as they aren't allowed in usernames
$cleanedLastName = $lastname -replace '[^a-zA-Z]', ''

# Clean up the firstname for the username, remove any special characters or numbers as they aren't allowed in usernames
$cleanedFirstName = $firstname -replace '[^a-zA-Z]', ''

# Ok, here we set the username. Our format is lastname + first initial of fisrt name. If that is taken then we include the next initial of the first name and so on. I should add a part here that then tries to add a number at the end of the username but I'll get to that in the next revision.
$global:username = $cleanedLastName+$cleanedFirstName.SubString(0,1); 
$testUsername = $global:username
$userExists = Get-ADUser -Filter {sAMAccountName -eq $testUsername}

If ($userExists -ne $Null) {
# User name is in list - stop here, start processing
$WPFTextUserName.Text = $global:username

for ($a=2; $attempt -le $firstname.Length; $a++) {

$testUsername = $cleanedLastName+$firstname.SubString(0,$a); 
$userExists = Get-ADUser -Filter {sAMAccountName -eq $testUsername}

$WPFTextUserName.Text = $testUsername
$global:username = $testUsername

# If we find a username we can use then we can jump out of the if statement
if ($userExists -eq $Null)
{break}
}
}
# Set the GUI to display the username, first, and last names
$WPFTextUserName.Text = $script:username
$global:firstname = $firstname
$global:lastname = $lastname
}

Function GeneratePassword {
# Select the two words from the list
# Make sure the password files are in the correct directory or update them here. 
# If they are missing for some reason then this is the format I use. List 1 is an adjective and list 2 is a noun. So you end up with Greenhouse or Hungrytable
$word1 = Get-Content ".\passwordlists\passwordlist1.txt" | Get-Random
$word2 = Get-Content ".\passwordlists\passwordlist2.txt" | Get-Random
# Now we put a random three digit number on the end and I don't want it to start with 00 or be 100 so we start at 101
$number = Get-Random -Minimum 101 -Maximum 999
# I'd like to add here a part that prevent repeating number - so 122, 222, wouldn't be allowed

# Update the GUI again
$global:password = $word1+$word2+$number
$WPFTextPassword.Text = $global:password
}
     
#Reference for when I forget how to do these things
 
#Adding items to a dropdown/combo box
    #$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
     
#Setting the text of a text box to the current PC name    
    #$WPFtextBox.Text = $env:COMPUTERNAME
     
#Adding code to a button, so that when clicked, it pings a system
# $WPFbutton.Add_Click({ Test-connection -count 1 -ComputerName $WPFtextBox.Text
# })

# Ok - here we actually do some real work and start to create the user account
Function CreateUser {
$currentYear = get-date -f yyyy
if ($WPFTextPassword.Text -eq "") {GeneratePassword}
if ($WPFTextUserName.Text -eq "") {SetUserName $WPFTextFirstName.Text $WPFTextLastName.Text}

# Read the combo boxes 
$global:cohort = $WPFComboCohort.Text
$global:house = $WPFComboHouse.Text

#################
# Groups to add student to - this is an array, make sure you edit it correctly
#$groupArray = @("$global:cohort Students", "All Students", "L$global:cohort")
$groupArray = @("$global:cohort Students", "All Students", "L$global:cohort")


if ($global:cohort -le $maxYear)
{
$global:houseAdd = $global:house+" Students"
$groupArray += "$global:houseAdd"
}

$WPFTBoxLog.AppendText("`nHouse is $global:house")
$WPFTBoxLog.AppendText("`nCohort is $global:cohort")
$WPFTBoxLog.AppendText("`nCurrent year is $currentYear")
$WPFTBoxLog.AppendText("`nMax year is $maxYear")

$global:fullPath = "OU=$global:cohort,$global:path"
$global:emailAddress ="$global:username$global:emailTail"
$global:homedir = "\\server\$global:username$"

$WPFTBoxLog.AppendText("`nFull path is $global:fullPath")
$WPFTBoxLog.AppendText("`nEmail address is $global:emailAddress")
$WPFTBoxLog.AppendText("`nHome directory is $global:homedir")

# Actually create the user
New-ADUser -Path $global:fullPath -Name $global:fullName -GivenName $global:firstname -Surname $global:lastname -DisplayName $global:fullName -SamAccountName $global:username -UserPrincipalName $global:username@contoso.com -PasswordNeverExpires 1 -EmailAddress $global:emailAddress -Description $global:description -Department $global:cohort -homedrive $global:homedrive -homedirectory $global:homedir -accountPassword (ConvertTo-SecureString -AsPlainText "$global:password" -Force) -PassThru  | Enable-ADAccount
$dateNow = get-date -f "dd-MM-yy hh:MM tt"
Set-ADUser $global:username -Replace @{info="Created on $dateNow by $env:username"}

# Create the homefolder now
New-Item -Path \\file-student\e$\Homedrive\$global:cohort\$global:username -type Directory -Force

############################
# Set sharing on homefolder - This block connects to the server and performs the file sharing
############################
Write-Host "Setting sharing on home folder START"
$server = "server"
$SrvPath = "E:\Homedrive\$global:cohort\$global:username"
$global:User = "$global:username"

$sb = {
param($User,$SrvPath)
  NET SHARE $User$=$SrvPath "/GRANT:Domain Admins,FULL" "/GRANT:$User,FULL" /REMARK:"Home folder for $User"
  } 

Invoke-Command -Computername "$server" -ScriptBlock $sb -ArgumentList $global:username,$SrvPath

Write-Host "Sharing on home folder completed"
############################
############################
# Set the permissions on homefolder
Write-Host "Setting permissions on home folder START"
$Acl = Get-Acl "\\server\e$\Homedrive\$global:cohort\$global:username"
$Acl.SetAccessRuleProtection($false, $true)
$ace = "$global:username","FullControl", "ContainerInherit,ObjectInherit","None","Allow"
$objACE = New-Object System.Security.AccessControl.FileSystemAccessRule($Ace)
$acl.AddAccessRule($objACE)
Set-ACL -Path "\\server\e$\Homedrive\$global:cohort\$global:username" -AclObject $Acl
Write-Host "Setting permissions on home folder END" -foregroundcolor "green"

# Add desktop path
New-Item -Path \\server\e$\Homedrive\$global:cohort\$global:username\Desktop -type Directory -Force
Write-Host "Created Desktop folder" -foregroundcolor "green"

# Add user to groups here
Write-Host "Starting to add $global:fullName to the array of AD groups" -foregroundcolor "green"

foreach ($group in $groupArray) {
	ADD-ADGroupMember “$group” –members $global:username
    $WPFTBoxLog.AppendText("`nAdded $global:fullName to $group")
}

}

Function SendEmail {
# This is where we compose the body of the email report
$emailBody = @"
<p>New User Account Details</p>
 <table width="600" border="0">
   <tbody>
     <tr>
       <td width="130"><strong>Username: </strong></td>
       <td width="454">$global:username</td>
     </tr>
     <tr>
       <td><strong>First Name:</strong></td>
       <td>$global:firstname</td>
     </tr>
     <tr>
       <td><strong>Last Name:</strong></td>
       <td>$global:lastname</td>
     </tr>
     <tr>
       <td><strong>Cohort:</strong></td>
       <td>$global:cohort</td>
     </tr>
     <tr>
       <td><strong>House:</strong></td>
       <td>$global:house</td>
     </tr>
     <tr>
       <td><strong>AD Path:</strong></td>
       <td>$global:fullpath</td>
     </tr>
     <tr>
       <td><strong>Home dir:</strong></td>
       <td>$global:homedir</td>
     </tr>
     <tr>
       <td><strong>Email Address:</strong></td>
       <td>$global:emailAddress</td>
     </tr>
   </tbody>
 </table>
 <p>Created by: $env:username</p>
"@

# Set emailTo to the user who created the account
$emailTo = "$env:username@contoso.com"

# Send the email
Send-MailMessage -To "$emailTo" -CC $emailCc -Subject "New user account for $global:fullName" -Body "$emailBody" -BodyAsHtml -Smtpserver "$global:smtpServer" -From "$global:emailFrom" 
}


$WPFButtonGenerateUsername.Add_Click({
if ($WPFTextFirstName.Text -ne "" -And $WPFTextLastName.Text -ne "")
{
SetUserName $WPFTextFirstName.Text $WPFTextLastName.Text
}
Else
{
$WPFTBoxLog.AppendText("`nMake sure you have entered a first and last name")
}
 })

$WPFButtonGeneratePassword.Add_Click({
GeneratePassword
})

$WPFButtonCreateUser.Add_Click({
# Check if cohort and house has been selected
if ($WPFComboCohort.Text -ne "" -And $WPFComboHouse.Text -ne "" -And $WPFTextFirstName.Text -ne "" -And $WPFTextLastName.Text -ne "")
{
CreateUser
SendEmail
$WPFButtonCreateUser.IsEnabled=$false
$WPFButtonGeneratePassword.IsEnabled=$false
$WPFButtonGenerateUsername.IsEnabled=$false
}
Else
{
$WPFTBoxLog.AppendText("`nMake sure you have completed all fields")
}

})
  
#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null