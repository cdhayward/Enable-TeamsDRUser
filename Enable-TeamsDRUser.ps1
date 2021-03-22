<#
Name:    Teams DR - Enable User
Author:  Chris Hayward - chrishayward.co.uk
Purpose: PowerShell GUI to manage Teams user settings for Direct Routing
Version: 1.0
Changes: DATE - Change
         19/03/2021 - Release 1.0

Known Errors:
Connect-MicrosoftTeams doesn't work in WinForm so script prompts for creds at launch.

#>

Connect-MicrosoftTeams

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$TeamsDREnableUsers              = New-Object system.Windows.Forms.Form
$TeamsDREnableUsers.ClientSize   = New-Object System.Drawing.Point(589,649)
$TeamsDREnableUsers.text         = "Teams Direct Routing - Enable Users"
$TeamsDREnableUsers.TopMost      = $false

$lbl_ScriptLog                   = New-Object system.Windows.Forms.Label
$lbl_ScriptLog.text              = "Script Status"
$lbl_ScriptLog.AutoSize          = $true
$lbl_ScriptLog.width             = 25
$lbl_ScriptLog.height            = 10
$lbl_ScriptLog.location          = New-Object System.Drawing.Point(67,40)
$lbl_ScriptLog.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtb_ScriptLog                  = New-Object system.Windows.Forms.TextBox
$txtb_ScriptLog.text             = "Click connect to get started"
$txtb_ScriptLog.multiline        = $True
$txtb_ScriptLog.width            = 393
$txtb_ScriptLog.height           = 60
$txtb_ScriptLog.location         = New-Object System.Drawing.Point(164,40)
$txtb_ScriptLog.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$txtb_ScriptLog.BackColor        = [System.Drawing.ColorTranslator]::FromHtml("#faf4cf")


$lstbx_SelectUser                = New-Object system.Windows.Forms.ListBox
$lstbx_SelectUser.text           = "listBox"
$lstbx_SelectUser.width          = 393
$lstbx_SelectUser.height         = 74
$lstbx_SelectUser.location       = New-Object System.Drawing.Point(160,238)

#btn_GetNumbers                  = New-Object system.Windows.Forms.Button
#$btn_GetNumbers.text             = "Get Numbers"
#$btn_GetNumbers.width            = 96
#$btn_GetNumbers.height           = 30
#$btn_GetNumbers.location         = New-Object System.Drawing.Point(400,555)
#$btn_GetNumbers.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#$txtb_NumberStart                = New-Object system.Windows.Forms.TextBox
#$txtb_NumberStart.multiline      = $false
#$txtb_NumberStart.width          = 214
#$txtb_NumberStart.height         = 20
#$txtb_NumberStart.location       = New-Object System.Drawing.Point(164,545)
#$txtb_NumberStart.Font           = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#$txtb_NumberEnd                  = New-Object system.Windows.Forms.TextBox
#$txtb_NumberEnd.multiline        = $false
#$txtb_NumberEnd.width            = 214
#$txtb_NumberEnd.height           = 20
#$txtb_NumberEnd.location         = New-Object System.Drawing.Point(163,578)
#$txtb_NumberEnd.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_VoiceRoutingPolicy          = New-Object system.Windows.Forms.Label
$lbl_VoiceRoutingPolicy.text     = "Voice Routing Policy"
$lbl_VoiceRoutingPolicy.AutoSize  = $true
$lbl_VoiceRoutingPolicy.width    = 25
$lbl_VoiceRoutingPolicy.height   = 10
$lbl_VoiceRoutingPolicy.location  = New-Object System.Drawing.Point(18,345)
$lbl_VoiceRoutingPolicy.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#lbl_EndNumber                   = New-Object system.Windows.Forms.Label
#$lbl_EndNumber.text              = "End Number"
#$lbl_EndNumber.AutoSize          = $true
#$lbl_EndNumber.width             = 25
#$lbl_EndNumber.height            = 10
#$lbl_EndNumber.location          = New-Object System.Drawing.Point(69,579)
#$lbl_EndNumber.Font              = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cbx_TenantDialPlan              = New-Object system.Windows.Forms.ComboBox
$cbx_TenantDialPlan.text         = "Global"
$cbx_TenantDialPlan.width        = 208
$cbx_TenantDialPlan.height       = 20
$cbx_TenantDialPlan.location     = New-Object System.Drawing.Point(163,384)
$cbx_TenantDialPlan.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cbx_VoiceRoutingPolicy          = New-Object system.Windows.Forms.ComboBox
$cbx_VoiceRoutingPolicy.text     = "Global"
$cbx_VoiceRoutingPolicy.width    = 208
$cbx_VoiceRoutingPolicy.height   = 20
$cbx_VoiceRoutingPolicy.location  = New-Object System.Drawing.Point(163,344)
$cbx_VoiceRoutingPolicy.Font     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

#$lbl_StartNumber                 = New-Object system.Windows.Forms.Label
#$lbl_StartNumber.text            = "Start Number"
#$lbl_StartNumber.AutoSize        = $true
#$lbl_StartNumber.width           = 25
#$lbl_StartNumber.height          = 10
#$lbl_StartNumber.location        = New-Object System.Drawing.Point(65,550)
#$lbl_StartNumber.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_TenantDialPlan              = New-Object system.Windows.Forms.Label
$lbl_TenantDialPlan.text         = "Tenant Dial Plan"
$lbl_TenantDialPlan.AutoSize     = $true
$lbl_TenantDialPlan.width        = 25
$lbl_TenantDialPlan.height       = 10
$lbl_TenantDialPlan.location     = New-Object System.Drawing.Point(42,384)
$lbl_TenantDialPlan.Font         = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btn_searchuser                  = New-Object system.Windows.Forms.Button
$btn_searchuser.text             = "Search"
$btn_searchuser.width            = 96
$btn_searchuser.height           = 30
$btn_searchuser.location         = New-Object System.Drawing.Point(385,188)
$btn_searchuser.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtb_SearchUser                 = New-Object system.Windows.Forms.TextBox
$txtb_SearchUser.multiline       = $false
$txtb_SearchUser.width           = 214
$txtb_SearchUser.height          = 20
$txtb_SearchUser.location        = New-Object System.Drawing.Point(160,198)
$txtb_SearchUser.Font            = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_SearchUser                  = New-Object system.Windows.Forms.Label
$lbl_SearchUser.text             = "Search User"
$lbl_SearchUser.AutoSize         = $true
$lbl_SearchUser.width            = 25
$lbl_SearchUser.height           = 10
$lbl_SearchUser.location         = New-Object System.Drawing.Point(67,199)
$lbl_SearchUser.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cbx_EnterpriseVoice             = New-Object system.Windows.Forms.CheckBox
$cbx_EnterpriseVoice.AutoSize    = $false
$cbx_EnterpriseVoice.width       = 187
$cbx_EnterpriseVoice.height      = 20
$cbx_EnterpriseVoice.location    = New-Object System.Drawing.Point(164,415)
$cbx_EnterpriseVoice.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_EnterpriseVoice             = New-Object system.Windows.Forms.Label
$lbl_EnterpriseVoice.text        = "Enterprise Voice"
$lbl_EnterpriseVoice.AutoSize    = $true
$lbl_EnterpriseVoice.width       = 25
$lbl_EnterpriseVoice.height      = 10
$lbl_EnterpriseVoice.location    = New-Object System.Drawing.Point(43,415)
$lbl_EnterpriseVoice.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$cbx_HostedVoicemail             = New-Object system.Windows.Forms.CheckBox
$cbx_HostedVoicemail.AutoSize    = $false
$cbx_HostedVoicemail.width       = 187
$cbx_HostedVoicemail.height      = 20
$cbx_HostedVoicemail.location    = New-Object System.Drawing.Point(164,440)
$cbx_HostedVoicemail.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_HostedVoicemail             = New-Object system.Windows.Forms.Label
$lbl_HostedVoicemail.text        = "Hosted Voicemail"
$lbl_HostedVoicemail.AutoSize    = $true
$lbl_HostedVoicemail.width       = 25
$lbl_HostedVoicemail.height      = 10
$lbl_HostedVoicemail.location    = New-Object System.Drawing.Point(43,440)
$lbl_HostedVoicemail.Font        = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btn_UpdateUser                  = New-Object system.Windows.Forms.Button
$btn_UpdateUser.text             = "Update"
$btn_UpdateUser.width            = 96
$btn_UpdateUser.height           = 30
$btn_UpdateUser.location         = New-Object System.Drawing.Point(163,495)
$btn_UpdateUser.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$txtb_LineURI                    = New-Object system.Windows.Forms.TextBox
$txtb_LineURI.multiline          = $false
$txtb_LineURI.text               = "tel:"
$txtb_LineURI.width              = 214
$txtb_LineURI.height             = 20
$txtb_LineURI.location           = New-Object System.Drawing.Point(163,463)
$txtb_LineURI.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$lbl_LineURI                     = New-Object system.Windows.Forms.Label
$lbl_LineURI.text                = "Line URI"
$lbl_LineURI.AutoSize            = $true
$lbl_LineURI.width               = 25
$lbl_LineURI.height              = 10
$lbl_LineURI.location            = New-Object System.Drawing.Point(87,463)
$lbl_LineURI.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btn_Connect                     = New-Object system.Windows.Forms.Button
$btn_Connect.text                = "Connect"
$btn_Connect.width               = 96
$btn_Connect.height              = 30
$btn_Connect.location            = New-Object System.Drawing.Point(164,114)
$btn_Connect.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$btn_disconnect                  = New-Object system.Windows.Forms.Button
$btn_disconnect.text             = "Disconnect"
$btn_disconnect.width            = 96
$btn_disconnect.height           = 30
$btn_disconnect.location         = New-Object System.Drawing.Point(280,109)
$btn_disconnect.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$gbx_Connect                     = New-Object system.Windows.Forms.Groupbox
$gbx_Connect.height              = 159
$gbx_Connect.width               = 573
$gbx_Connect.text                = "Connect to Teams"
$gbx_Connect.location            = New-Object System.Drawing.Point(0,5)

$gbx_UserSettings                = New-Object system.Windows.Forms.Groupbox
$gbx_UserSettings.height         = 240
$gbx_UserSettings.width          = 572
$gbx_UserSettings.text           = "User Settings"
$gbx_UserSettings.location       = New-Object System.Drawing.Point(1,330)

$gbx_SearchUser                  = New-Object system.Windows.Forms.Groupbox
$gbx_SearchUser.height           = 161
$gbx_SearchUser.width            = 575
$gbx_SearchUser.text             = "Search"
$gbx_SearchUser.location         = New-Object System.Drawing.Point(1,167)

$TeamsDREnableUsers.controls.AddRange(@($lstbx_SelectUser,$btn_GetNumbers,$txtb_ScriptLog,$lbl_ScriptLog,$txtb_NumberStart,$txtb_NumberEnd,$lbl_VoiceRoutingPolicy,$lbl_EndNumber,$cbx_TenantDialPlan,$cbx_VoiceRoutingPolicy,$lbl_StartNumber,$lbl_TenantDialPlan,$btn_searchuser,$txtb_SearchUser,$lbl_SearchUser,$cbx_EnterpriseVoice,$lbl_EnterpriseVoice,$cbx_HostedVoicemail,$lbl_HostedVoicemail,$btn_UpdateUser,$txtb_LineURI,$lbl_LineURI,$lbl_Username,$txtb_Username,$cbx_Override,$lbl_AdminDomain,$tbl_onmicrosoft,$btn_Connect,$gbx_Connect,$gbx_UserSettings,$gbx_SearchUser))
$gbx_Connect.controls.AddRange(@($lbl_OverrideAdminDomain,$txtb_AdminDomain,$btn_disconnect))

$btn_Connect.Add_Click({ Connect-Teams })
#$btn_GetNumbers.Add_Click({ Get-Numbers })
$btn_disconnect.Add_Click({ Disconnect-Teams })
$btn_searchuser.Add_Click({ SearchUser })
$btn_UpdateUser.Add_Click({ UpdateUser })
$lstbx_SelectUser.Add_SelectedIndexChanged({ SelectUser })




#Write your logic code here

Function Write-ScriptOutput ($Message){

$txtb_ScriptLog.Text = $Message
Write-Host $Message

}

function Connect-Teams {

    #Check MicrosoftTeams Module is v2 or higher

    Write-ScriptOutput -Message "Checking MS Teams Module version"
    if (!(Get-Module -ListAvailable | where {$_.Name -eq 'MicrosoftTeams' -and $_.Version -ge '2.0.0'})) {
        Write-ScriptOutput -Message "Please ensure Teams Module 2.0 or higher is installed"
    }

    #Check if we are already connected to Teams
    Write-ScriptOutput -Message "Checking if Teams is connected"
    #Get-CsHostedVoicemailPolicy -ErrorAction Ignore
    if (!(Get-PSSession | Where {$_.Name -like 'SfBPowerShellSessionViaTeamsModule*'})) {
        Write-ScriptOutput -Message "Importing Module MicrosoftTeams"
        Import-Module MicrosoftTeams
        Write-ScriptOutput -Message "Connecting to Microsoft Teams"
        Write-Host "Teams not connected"
        Disconnect-MicrosoftTeams
        Connect-MicrosoftTeams #-UseDeviceAuthentication
        #Write-ScriptOutput -Message "Check PowerShell Window and follow Device Authentication steps"
    }Else{
        Write-ScriptOutput -Message "Teams already connected - Search for a user"
      
    }

Load-TeamsPolicies

}



function Load-TeamsPolicies {
    
    # Clear comboboxes
    $cbx_VoiceRoutingPolicy.Items.Clear()
    $cbx_TenantDialPlan.Items.Clear()

    # Get OnlineVoiceRoutingPolicies
    Write-ScriptOutput -Message "Connecting - Getting list of OnlineVoiceRoutingPolicies"
    $OnlineVoiceRoutingPolicies = Get-CsOnlineVoiceRoutingPolicy | select identity
    ForEach ($vp in $OnlineVoiceRoutingPolicies){
        $cbx_VoiceRoutingPolicy.Items.Add("$($vp.identity)")
    }

    # Get TenantDialPlans
    Write-ScriptOutput -Message "Connecting - Getting list of TenantDialPlans"
    $TenantDialPlan = Get-CsTenantDialPlan | select identity
    ForEach ($dp in $TenantDialPlan){
        $cbx_TenantDialPlan.Items.Add("$($dp.identity)")
    }

Write-ScriptOutput -Message "Connected - Search for a user"

}


function Disconnect-Teams { 
    Write-ScriptOutput -Message "Disconnecting from Microsoft Teams"
    Disconnect-MicrosoftTeams
    Get-PSSession | Remove-PSSession
    Write-ScriptOutput -Message "Disconnected"
    #$cbx_VoicePolicy.Items.Add('Test3')
}

function Get-Numbers { 

    Get-CsOnlineUser | where {$_.Enabled -eq $True -and $_.EnterpriseVoiceEnabled -eq $True} | Select SipAddress, LineURI | Out-GridView

}

function SearchUser {
    

    Write-ScriptOutput -Message "Searching for $($txtb_SearchUser.text)" 

    If ($($txtb_SearchUser.text) -ne ''){
        $ReturnedUsers = Get-CsOnlineUser -Filter "Enabled -eq '$True' -and DisplayName -like '*$($txtb_SearchUser.text)*'" | Select DisplayName, SipAddress | Sort-Object -Property SipAddress

        $lstbx_SelectUser.Items.Clear()

        ForEach ($U in $ReturnedUsers){
            $lstbx_SelectUser.Items.Add("$($U.SipAddress) || ($($U.DisplayName))")
        }

        Write-ScriptOutput -Message "Search complete" 

    }Else{
        Write-ScriptOutput -Message "No user entered - Search for a user" 
    }
}

function SelectUser {

    #Write-Host "$($lstbx_SelectUser.SelectedItem)"
    Write-ScriptOutput -Message "Getting details for $($lstbx_SelectUser.SelectedItem)" 

    If ($($lstbx_SelectUser.SelectedItem) -ne $null){

    $SipAddress = $($lstbx_SelectUser.SelectedItem).Split(' || ')
    $SipAddress = $SipAddress[0]

    $U = Get-CsOnlineUser -Identity $SipAddress | select *Voice*, *LineURI*, *Dial*

    PopulateUser -OnlineVoiceRoutingPolicy $U.OnlineVoiceRoutingPolicy -TenantDialPlan $U.TenantDialPlan -EnterpriseVoiceEnabled $U.EnterpriseVoiceEnabled -OnPremLineURI $U.OnPremLineURI -HostedVoicemail $u.HostedVoicemail
    }Else{
        Write-ScriptOutput -Message "No user selected - Search for a user" 
    }
}

function PopulateUser ($OnlineVoiceRoutingPolicy, $TenantDialPlan, $EnterpriseVoiceEnabled, $OnPremLineURI, $HostedVoicemail){

    Write-Host "$OnlineVoiceRoutingPolicy | $TenantDialPlan | $EnterpriseVoiceEnabled | $OnPremLineURI"

    If ($OnlineVoiceRoutingPolicy -ne $null){
        $OnlineVoiceRoutingPolicy = "Tag:$OnlineVoiceRoutingPolicy"
        $cbx_VoiceRoutingPolicy.SelectedIndex = $cbx_VoiceRoutingPolicy.FindStringExact($OnlineVoiceRoutingPolicy)
    }else{
        $cbx_VoiceRoutingPolicy.SelectedIndex = $cbx_VoiceRoutingPolicy.FindStringExact('Global')
    }

    If ($TenantDialPlan -ne $null){
        $TenantDialPlan = "Tag:$TenantDialPlan"
        $cbx_TenantDialPlan.SelectedIndex = $cbx_TenantDialPlan.FindStringExact($TenantDialPlan)
    }else{
        $cbx_TenantDialPlan.SelectedIndex = $cbx_TenantDialPlan.FindStringExact('Global')
    }
    
    If ($EnterpriseVoiceEnabled -eq 'True'){
        #Write-Host "EV = True"
        $cbx_EnterpriseVoice.CheckState = 1
    }else{
        $cbx_EnterpriseVoice.CheckState = 0
    }

    If ($HostedVoicemail -eq 'True'){
        $cbx_HostedVoicemail.CheckState = 1
    }else{
        $cbx_HostedVoicemail.CheckState = 0
    }

    If ($OnPremLineURI -ne ''){
        $txtb_LineURI.Text = $OnPremLineURI
    }else{
        $txtb_LineURI.Text = 'tel:'
    }

    Write-ScriptOutput -Message "Now editing $($lstbx_SelectUser.SelectedItem)"
}

function UpdateUser (){
    
    Write-ScriptOutput -Message "Now updating $($lstbx_SelectUser.SelectedItem)"

    $SipAddress = $($lstbx_SelectUser.SelectedItem).Split(" || ")
    $SipAddress = $SipAddress[0]

    Write-Host "Updating:             $SipAddress"
    Write-Host "Voice Routing Policy: $($cbx_VoiceRoutingPolicy.Text)"
    Write-Host "Tenant Dial Plan:     $($cbx_TenantDialPlan.Text)"
    Write-Host "Enterprise Voice:     $($cbx_EnterpriseVoice.CheckState)"
    Write-Host "HostedVoicemail:      $($cbx_HostedVoicemail.CheckState)"
    Write-Host "LineURI:              $($txtb_LineURI.text)"

    
    If ($cbx_EnterpriseVoice.CheckState -eq 1){
        Write-ScriptOutput -Message "Enabling Enterprise Voice for $SipAddress"
        Set-CsUser -Identity $SipAddress -OnPremLineURI $txtb_LineURI.text -EnterpriseVoiceEnabled $true
    }Else{
        Write-ScriptOutput -Message "Disabling Enterprise Voice for $SipAddress"
        Set-CsUser -Identity $SipAddress -OnPremLineURI $txtb_LineURI.text -EnterpriseVoiceEnabled $false
    }

    If ($cbx_HostedVoicemail.CheckState -eq 1){
        Write-ScriptOutput -Message "Enabling Hosted Voicemail for $SipAddress"
        Set-CsUser -Identity $SipAddress -HostedVoiceMail $true
    }Else{
        Write-ScriptOutput -Message "Disabling Hosted Voicemail for $SipAddress"
        Set-CsUser -Identity $SipAddress -HostedVoiceMail $false
    }

    if ($cbx_VoiceRoutingPolicy.Text -eq 'Global'){
        Write-ScriptOutput -Message "Removing Voice Policy from $SipAddress"
        Grant-CsOnlineVoiceRoutingPolicy -Identity $SipAddress -PolicyName $null
    }Else{
        Write-ScriptOutput -Message "Granting Voice Policy $($cbx_VoiceRoutingPolicy.Text) to $SipAddress"
        Grant-CsOnlineVoiceRoutingPolicy -Identity $SipAddress -PolicyName $cbx_VoiceRoutingPolicy.Text
    }

    if ($cbx_TenantDialPlan.Text -eq 'Global'){
        Write-ScriptOutput -Message "Removing Dial Plan from $SipAddress"
        Grant-CsTenantDialPlan -Identity $SipAddress -PolicyName $null
    }Else{
        Write-ScriptOutput -Message "Granting Dial Plan $($cbx_TenantDialPlan.Text) to $SipAddress"
        Grant-CsTenantDialPlan -Identity $SipAddress -PolicyName $cbx_TenantDialPlan.Text
    }

    Write-ScriptOutput -Message "Update complete for $($lstbx_SelectUser.SelectedItem)"
}


#Connect-Teams
Load-TeamsPolicies

[void]$TeamsDREnableUsers.ShowDialog()


