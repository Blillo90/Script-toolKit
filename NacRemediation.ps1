<#
.SYNOPSIS
This Powershell script is used to add entries in the NAC ADLDS for Remediation purposes.
.DESCRIPTION
This script can be used to add entries into the NAC ADLDS for Remediation purposes. The entries are automatically deleted after a specific amount of time by an automatic background process in the NAC ADLDS servers.
A GUI is provided where the users can check current information and add new entries.
#>
#Libraries required for Forms and other functions in the script
Add-Type -AssemblyName System.Windows.Forms
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
$CurrentVersion = "NAC Remediation Interface v0.9"
#Container paths for global searches in the NAC ADLDS
$NACContainerPathList = @{
    "Getafe"  = "LDAP://nacget.intra.casa.corp:636/CN=NAC,DC=ds,DC=corp"
    "Sevilla" = "LDAP://nactab.intra.casa.corp:636/CN=NAC,DC=ds,DC=corp" }
#Container paths for searches in the NAC ADLDS Remediation OU
$NACRemediationContainerPathList = @{
    "Getafe"  = "LDAP://nacget.intra.casa.corp:636/CN=Remediation,CN=NAC,DC=ds,DC=corp"
    "Sevilla" = "LDAP://nactab.intra.casa.corp:636/CN=Remediation,CN=NAC,DC=ds,DC=corp" }

#Global definitions for the TextBox formats
$newline = [System.Environment]::NewLine
$green = [Drawing.Color]::Green
$red= [Drawing.Color]::Red
$black = [Drawing.Color]::Black
#This function is used to append text in a TextBox with a defined color
function AppendColor ([system.windows.Forms.RichTextBox] $tb, [String] $output, [Drawing.Color] $color = $black)
{
    $ss = $tb.TextLength
    $tb.AppendText($output)
    $sl = $tb.SelectionStart - $ss + 1
    $tb.Select($ss, $sl)
    $tb.SelectionColor = $color
    $tb.AppendText($newline)
}
#This function is used to clear the text from a TextBox
function ClearOutput ([system.windows.Forms.RichTextBox] $tb)
{
    $tb.Text = ""
}
#Creation of main form
$Form = New-Object system.Windows.Forms.Form
$Form.Text = $CurrentVersion
$Form.TopMost = $false
$Form.FormBorderStyle = 'FixedSingle'
$Form.MaximizeBox = $false
$Form.MinimizeBox = $false
$Form.Width = 800
$Form.Height = 520
#Label for the "MAC Address" field
$labelMAC = New-Object system.windows.Forms.Label
$labelMAC.Text = "MAC Address"
$labelMAC.AutoSize = $true
$labelMAC.Width = 25
$labelMAC.Height = 10
$labelMAC.location = new-object system.drawing.point(19,23)
$labelMAC.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelMAC)
#Field for the MAC Address
$inputMACAddress = New-Object system.windows.Forms.TextBox
$inputMACAddress.Width = 261
$inputMACAddress.Height = 20
$inputMACAddress.location = new-object system.drawing.point(124,23)
$inputMACAddress.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($inputMACAddress)
#Button to check if the MAC Address exists currently in the NAC ADLDS repository
$btnCheckMAC = New-Object system.windows.Forms.Button
$btnCheckMAC.Text = "Check MAC"
$btnCheckMAC.Width = 111
$btnCheckMAC.Height = 24
$btnCheckMAC.location = new-object system.drawing.point(410,22)
$btnCheckMAC.Font = "Microsoft Sans Serif,10"
$btnCheckMAC.Add_Click({
    # CHECK EXISTING DEVICES WITH THE GIVEN COMPUTER MAC ADDRESS
    $MACAddress = $inputMACAddress.Text
    #Select the corrent NAC ADLDS instance based on the selection of site by the user
    $NACContainerPath = $NACContainerPathList.($ddServer.SelectedItem)
    $NACRemediationContainerPath = $NACRemediationContainerPathList.($ddServer.SelectedItem)
    ClearOutput $tbOutput
    AppendColor $tbOutput ("Checking MAC: " + $MACAddress)
    #Check that the MAC address has the correct format
    if ($MACAddress -match '^([0-9a-f]{2}:){5}([0-9a-f]{2})$')
    {
        AppendColor $tbOutput "MAC Address format is correct" $green
        #deviceRemediationID is the field containing the MAC Address for objects that are in the Remediation OU,
        #for other objects the MAC Address is stored in the networkAddress field
        $strFilter = "(|(networkAddress=$MACAddress)(deviceRemediationID=$MACAddress))"

        $objDomain = New-Object System.DirectoryServices.DirectoryEntry($NACContainerPath)
        #We prepare the object that performs the search in the ADLDS
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $objDomain
        $objSearcher.PageSize = 1000
        $objSearcher.Filter = $strFilter
        $objSearcher.SearchScope = "Subtree"
        #Properties that should be returned in the search
        $colProplist = "cn","deviceType","deviceZone","networkAddress","deviceRemediationID","adminDisplayName","whenCreated","description",'devicemodel','device8021xcapable'
        foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}
        try
        {
                #Perform the search
                $colResults = $objSearcher.FindAll()
        }
        catch
        {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                AppendColor $tbOutput ("Error accesing LDAP: " + $ErrorMessage ) $red
                return
        }
        if ($colResults.Count -eq 0)
        {
                #No objects were found
                AppendColor $tbOutput "No object found with MAC=$MACAddress" $red
        }
        else
        {
                #One or more objects were found
                AppendColor $tbOutput ($colResults.Count.ToString() + " object(s) found: ") $green
                foreach ($objResult in $colResults)
                {
                    AppendColor $tbOutput ""
                    $objItem = $objResult.Properties;
                    #Depending on the OU were the object is located, we retrieve different properties of the objects
                    if ($objItem.adspath[0].Contains("CN=Remediation"))
                    {
                        $dt=$objItem.whencreated[0]
                        $objectStatus = "Object in REMEDIATION since " + $dt.ToString('g') + " by " + $objItem.admindisplayname
                        $objectMACAddress = $objItem.deviceremediationid
                    }
                    elseif ($objItem.adspath[0].Contains("CN=Migration"))
                    {
                        $dt=$objItem.whencreated[0]
                        $objectStatus = "Object in MIGRATION since " + $dt.ToString('g') + " by " + $objItem.admindisplayname
                        $objectMACAddress = $objItem.deviceremediationid
                    }
                    elseif ($objItem.adspath[0].Contains("CN=Exception"))
                    {
                        $dt=$objItem.whencreated[0]
                        $objectStatus = "Object in EXCEPTION since " + $dt.ToString('g') + " with description " + $objItem.admindisplayname
                        $objectMACAddress = $objItem.deviceremediationid
                    }
                    else
                    {
                        $objectStatus = "Object is NOT in remediation"
                        $objectMACAddress = $objItem.networkaddress
                    }

                    Write-Host $objItem.adspath
                    AppendColor $tbOutput ("Object Found: " + $objItem.adspath)
                    AppendColor $tbOutput ("`tName: " + $objItem.cn)
                    AppendColor $tbOutput ("`tType: " + $objItem.devicetype)
                    AppendColor $tbOutput ("`tZone: " + $objItem.devicezone)
                    AppendColor $tbOutput ("`tMAC: " + $objectMACAddress)
                    AppendColor $tbOutput ("`tStatus: " + $objectStatus)
                    if ($objItem.description[0].Equals("ESUnifiedCommsFlatFile"))
                    {
                        #For the Unified Communications objects there are two additional fields that we display
                        AppendColor $tbOutput ("`tPhone Model: " + $objItem.devicemodel)
                        AppendColor $tbOutput ("`t802.1x capable?: " + $objItem.device8021xcapable)
                    }
                }
        }
    }
    else
    {
        AppendColor $tbOutput "MAC Address format is incorrect" $red
    }

})
$Form.controls.Add($btnCheckMAC)
$labelCN = New-Object system.windows.Forms.Label
$labelCN.Text = "Computer Name"
$labelCN.AutoSize = $true
$labelCN.Width = 25
$labelCN.Height = 10
$labelCN.location = new-object system.drawing.point(19,63)
$labelCN.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelCN)
$inputCN = New-Object system.windows.Forms.TextBox
$inputCN.Width = 261
$inputCN.Height = 20
$inputCN.location = new-object system.drawing.point(124,63)
$inputCN.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($inputCN)
$btnCheckCN = New-Object system.windows.Forms.Button
$btnCheckCN.Text = "Check CN"
$btnCheckCN.Width = 111
$btnCheckCN.Height = 24
$btnCheckCN.location = new-object system.drawing.point(410,62)
$btnCheckCN.Font = "Microsoft Sans Serif,10"
$btnCheckCN.Add_Click({
    # CHECK EXISTING DEVICES WITH THE GIVEN COMPUTER NAME
    $CN = $inputCN.Text
    $NACContainerPath = $NACContainerPathList.($ddServer.SelectedItem)
    $NACRemediationContainerPath = $NACRemediationContainerPathList.($ddServer.SelectedItem)
    ClearOutput $tbOutput
    AppendColor $tbOutput ("Checking CN: " + $CN)
    $strFilter = "(cn=$CN)"

    $objDomain = New-Object System.DirectoryServices.DirectoryEntry($NACContainerPath)
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.PageSize = 1000
    $objSearcher.Filter = $strFilter
    $objSearcher.SearchScope = "Subtree"
    $colProplist = "cn","deviceType","deviceZone","networkAddress","deviceRemediationID","adminDisplayName","whenCreated","description",'devicemodel','device8021xcapable'
    foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}
    try
    {
            $colResults = $objSearcher.FindAll()
    }
    catch
    {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            AppendColor $tbOutput ("Error accesing LDAP: " + $ErrorMessage ) $red
            return
    }
    if ($colResults.Count -eq 0)
    {
            AppendColor $tbOutput "No object found with CN=$CN" $red
    }
    else
    {
            AppendColor $tbOutput ($colResults.Count.ToString() + " object(s) found: ") $green
            foreach ($objResult in $colResults)
            {
                AppendColor $tbOutput ""
                $objItem = $objResult.Properties;
                if ($objItem.adspath[0].Contains("CN=Remediation"))
                {
                    $dt=$objItem.whencreated[0]
                    $objectStatus = "Object in REMEDIATION since " + $dt.ToString('g') + " by " + $objItem.admindisplayname
                    $objectMACAddress = $objItem.deviceremediationid
                }
                elseif ($objItem.adspath[0].Contains("CN=Migration"))
                {
                    $dt=$objItem.whencreated[0]
                    $objectStatus = "Object in MIGRATION since " + $dt.ToString('g') + " by " + $objItem.admindisplayname
                    $objectMACAddress = $objItem.deviceremediationid
                }
                elseif ($objItem.adspath[0].Contains("CN=Exception"))
                {
                    $dt=$objItem.whencreated[0]
                    $objectStatus = "Object in EXCEPTION since " + $dt.ToString('g') + " with description " + $objItem.admindisplayname
                    $objectMACAddress = $objItem.deviceremediationid
                }
                else
                {
                    $objectStatus = "Object is NOT in remediation"
                    $objectMACAddress = $objItem.networkaddress
                }

                Write-Host $objItem.adspath
                AppendColor $tbOutput ("Object Found: " + $objItem.adspath)
                AppendColor $tbOutput ("`tName: " + $objItem.cn)
                AppendColor $tbOutput ("`tType: " + $objItem.devicetype)
                AppendColor $tbOutput ("`tZone: " + $objItem.devicezone)
                AppendColor $tbOutput ("`tMAC: " + $objectMACAddress)
                AppendColor $tbOutput ("`tStatus: " + $objectStatus)
                if ($objItem.description[0].Equals("ESUnifiedCommsFlatFile"))
                {
                    AppendColor $tbOutput ("`tPhone Model: " + $objItem.devicemodel)
                    AppendColor $tbOutput ("`t802.1x capable?: " + $objItem.device8021xcapable)
                }
            }
    }

})
$Form.controls.Add($btnCheckCN)
$btnAdd = New-Object system.windows.Forms.Button
$btnAdd.Text = "Add Device"
$btnAdd.Width = 111
$btnAdd.Height = 32
$btnAdd.location = new-object system.drawing.point(547,34)
$btnAdd.Font = "Microsoft Sans Serif,10"
$btnAdd.Add_Click({
    # ADD AN OBJECT TO THE REMEDIATION OU
    ClearOutput $tbOutput
    $NACContainerPath = $NACContainerPathList.($ddServer.SelectedItem)
    $NACRemediationContainerPath = $NACRemediationContainerPathList.($ddServer.SelectedItem)
    $MACAddress = $inputMACAddress.Text
    $CN = $inputCN.Text
    if ($CN -eq "")
    {
        AppendColor $tbOutput "You must fill in at least the Computer Name of the object you want to put into Remediation state" $red
        AppendColor $tbOutput "If you don't input the MAC Address, the CMDB will be checked to retrieve it" $red
        return
    }
    if ($MACAddress -eq "")
    {

        <#  USE THIS PROCESS TO LOOK FOR THE MAC IN THE ADLDS #>
        AppendColor $tbOutput "Retrieving MAC Address of the computer from the ADLDS"
        $strFilter = "(cn=$CN)"

        $objDomain = New-Object System.DirectoryServices.DirectoryEntry($NACContainerPath)
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $objDomain
        $objSearcher.PageSize = 1000
        $objSearcher.Filter = $strFilter
        $objSearcher.SearchScope = "Subtree"
        $colProplist = "cn","deviceType","deviceZone","networkAddress","deviceRemediationID","adminDisplayName","whenCreated"
        foreach ($i in $colPropList){$objSearcher.PropertiesToLoad.Add($i)}
        try
        {
                $colResults = $objSearcher.FindAll()
        }
        catch
        {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                AppendColor $tbOutput ("Error accesing LDAP: " + $ErrorMessage ) $red
                return
        }
        $ResultCount = $colResults.Count
        if ($ResultCount -eq 0)
        {
                AppendColor $tbOutput "No object found in ADLDS with CN=$CN. Cannot retrieve the MAC Address of the computer, please fill it in manually." $red
                return
        }
        if ($ResultCount -ne 1)
        {
                AppendColor $tbOutput "$ResultCount objects found in ADLDS with CN=$CN. Cannot retrieve the specific MAC Address of the computer, please fill it in manually." $red
                return
        }
        $MACAddress = $colResults[0].Properties.networkaddress[0].ToString()
        AppendColor $tbOutput "MAC Address $MACAddress found for $CN" $green
    }
    else
    {
        AppendColor $tbOutput ("Checking MAC format: " + $MACAddress)
        if ($MACAddress -match '^([0-9a-f]{2}:){5}([0-9a-f]{2})$')
        {
            AppendColor $tbOutput "MAC Address format is correct" $green
        }
        else
        {
            AppendColor $tbOutput "MAC Address format is incorrect" $red
            return
        }
    }
    $objOU = [ADSI]$NACRemediationContainerPath
    $objDevice=$objOU.Create("device", "CN=" + $CN)

    $objDevice.Put("deviceZone", "ES")
    $objDevice.Put("deviceType", "Remediation-PC")
    $objDevice.Put("deviceRemediationID", $MACAddress)
    $objDevice.Put("adminDisplayName", $env:UserDomain + "\" + $env:UserName)
    try
    {
            $objDevice.SetInfo()
    }
    catch
    {
            $ErrorMessage = $_.Exception.Message
            $FailedItem = $_.Exception.ItemName
            AppendColor $tbOutput ("Error: " + $ErrorMessage ) $red
            return
    }

    AppendColor $tbOutput ("Object " + $objDevice.distinguishedname + " was added correctly" ) $green

})
$Form.controls.Add($btnAdd)
$tbOutput = New-Object system.windows.Forms.RichTextBox
$tbOutput.Multiline = $true
$tbOutput.ReadOnly = $true
$tbOutput.Width = 766
$tbOutput.Height = 310
$tbOutput.location = new-object system.drawing.point(11,107)
$tbOutput.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($tbOutput)
$labelServer = New-Object system.windows.Forms.Label
$labelServer.Text = "Server"
$labelServer.AutoSize = $true
$labelServer.Width = 25
$labelServer.Height = 10
$labelServer.location = new-object system.drawing.point(19,440)
$labelServer.Font = "Microsoft Sans Serif,10"
$Form.controls.Add($labelServer)
$ddServer = new-object System.Windows.Forms.ComboBox
$ddServer.location = new-object System.Drawing.Size(124,440)
$ddServer.Size = new-object System.Drawing.Size(130,30)
[array]$DropDownArrayItems = "Getafe","Sevilla"
ForEach ($Item in $DropDownArrayItems) {
     [void] $ddServer.Items.Add($Item)
}
$ddServer.SelectedItem = $ddServer.Items[0]
$Form.controls.Add($ddServer)
[void]$Form.ShowDialog()
$Form.Dispose().Dispose()
