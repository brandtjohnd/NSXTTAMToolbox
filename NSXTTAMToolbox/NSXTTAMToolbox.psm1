#Load the Import-Excel module
ipmo ImportExcel

#  Install into the path below for greatest compatibility
#  C:\Windows\system32\WindowsPowerShell\v1.0\Modules\

#Load Functions into Memory
function Show-EULA {
    
    $eula = "Disclaimer and End User License Agreement `n `nDeveloped in 2021 by John Brandt (jbrandt@vmware.com) & Ralph Stoker (rstoker@vmware.com) `n`nThe Software is provided 'as is' without warranty of any kind, either express or implied. Use at your own risk. `n    `nThe use of the software and scripts is done at your own discretion and risk and with agreement that you will be `nsolely responsible for any damage to your computer system or loss of data that results from such activities. `nYou are solely responsible for adequate protection and backup of the data and equipment used in connection with `nany of the software, and we will not be liable for any damages that you may suffer in connection with using, `nmodifying or distributing any of this software. No advice or information, whether oral or written, `nobtained by you from us shall create any warranty for the software. `n    `nWe make no warranty that: `n    `nthe software will meet your requirements `nthe software will be uninterrupted, timely, secure or error-free `nthe results that may be obtained from the use of the software will be effective, accurate or reliable `nthe quality of the software will meet your expectations `nany errors in the software obtained from us will be corrected `n    `nThe Software: `n     `ncould include technical or other mistakes, inaccuracies or typographical errors.  `nmay be out of date, and we make no commitment to update such materials. `nWe assume no responsibility for errors or omissions in the software. `n    `nIn no event shall we be liable to you or any third parties for any special, punitive, incidental, `nindirect or consequential damages of any kind, or any damages whatsoever, including, without limitation, `nthose resulting from loss of use, data or profits, and on any theory of liability, `narising out of or in connection with the use of this software. `n"
    
    ####Construct User Prompt for proceeding with duplicates
    $Accept = New-Object System.Management.Automation.Host.ChoiceDescription '&Accept', 'Accept?'
    $decline = New-Object System.Management.Automation.Host.ChoiceDescription '&Decline', 'Decline?'
    #$skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip', 'Skip?'
    $options = [System.Management.Automation.Host.ChoiceDescription[]](<#$overwrite,#> $Accept, $Decline)
    $title = 'Do you accept the EULA?'
    $message = $eula
    ####
                    
    #Prompt user to decide how to proceed if a policy is duplicated
    $proceed = $host.ui.PromptForChoice($title, $message, $options, 0)
    return $proceed
    }
function Get-PoShVersion{

    #Get current Powershell version
    $powershellVersionMajor = $psversiontable.PSVersion.Major.ToString()
    $powershellVersionMinor = $psversiontable.PSVersion.Minor.ToString()
    $powershellVersion = $powershellVersionMajor + "." + $powershellVersionMinor

    <#
    #Adjust invoke-web request command based on PS Version
    if ($powershellVersion -lt 6){
        Write-Host "Your Version of Powershell is $powershellVersion. Adjusting Authentication Accordingly." -ForegroundColor Yellow
    }
    else
    {
        Write-Host "Your Version of Powershell is $powershellVersion. Adjusting Authentication Accordingly." -ForegroundColor Yellow
    }#>
    return $powershellVersion
}
function Get-NSXVersion{
    param(
        [Parameter(Mandatory=$False)]
        [string]
        $contentType = 'application/json',
        
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
        )

    #Compose URI for NSX Version
    $getBaseURI = "/api/v1/node"
    $proto = "https://"
    $versionGetUri = $proto + $nsxMgrFQDN + $getBaseURI
    
    #Login
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

    #Get NSX Version number
    $getNSXVersion = Invoke-WebRequest $versionGetUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
    $responseContent = $getNSXVersion.Content | ConvertFrom-Json
    $currentNSXVersion = $responseContent.product_version
    return $currentNSXVersion
}
function Show-Menu{
    param (
        [string]$menuTitle = 'NSX-T TAM Toolbox',

        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
    )
    $powershellVersion = Get-PoShVersion
    $currentNSXVersion = Get-NSXVersion -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
    Clear-Host
    Write-Host "================ $menuTitle ================" -ForegroundColor cyan
    Write-Host "Powershell Version:     " $powershellVersion
    Write-Host "NSX-T Version:          " $currentNSXVersion `n    
    Write-Host "Press '1' to Set VM Tags in NSX-T" -ForegroundColor cyan
    Write-Host "Press '2' to Create Security Groups in NSX-T" -ForegroundColor cyan
    Write-Host "Press '3' to Create Service Groups in NSX-T" -ForegroundColor cyan
    Write-Host "Press '4' to Create Distributed Firewall Policies & Rules in NSX-T" -ForegroundColor cyan
    Write-Host "Press '5' to Execute Functions 1-4 in Order"-ForegroundColor cyan
    Write-Host "Press '6' to Export specific Firewall Policy & Rules to Spreadsheet"-ForegroundColor cyan
    Write-Host "Press '7' to Migrate DFW Policy to a new Category" -ForegroundColor cyan
    Write-Host "Press '8' to Re-Import the input Spreadsheet" -ForegroundColor cyan
    Write-Host "Press 'H' to View Help" -ForegroundColor DarkYellow
    Write-Host "Press 'Q' to Exit this Script" `n `n -ForegroundColor DarkYellow
   
}
Function Get-NSXTAuth{


    param(
        [Parameter(Mandatory=$False)]
        [string]
        $contentType = 'application/json',
        
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
        )

    $proto = "https://"
    $baseUri = $proto+$nsxMgrFQDN
    $loginUri = $baseUri+'/api/session/create'
    $loginBody = @{
        j_username = $inCreds.UserName
        j_password = $inCreds.GetNetworkCredential().Password
        }
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Authenticate & Connect to NSX-T API
    try{
        #Get Rest Token
        $login = Invoke-WebRequest $loginUri -Method Post -Body $loginBody -headers @{'Content-Type' = 'application/x-www-form-urlencoded'} -SessionVariable session #-skipcertificatecheck
    }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response AUTH -ForegroundColor Red
        throw {'Failed to login. Please review the error and retry.' + $loginErr}
    }

    [string]$tokenString = $login.Headers['X-XSRF-TOKEN']
    $tokenHeaders = @{'X-XSRF-TOKEN' = $tokenString}
    $tokenHeaders.add('Content-Type',$contentType)
    
    return $session,$tokenHeaders
    }
Function Remove-NSXTAuth{

    param(
                
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory = $True)]
        [array]$tokenHeaders

        )
    $proto = "https://"
    $baseUri = $proto+$nsxMgrFQDN
    $destroyUri = $baseUri+'/api/session/destroy'
    
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Authenticate & Connect to NSX-T API
    try{
        #Get Rest Token
        $login = Invoke-WebRequest $destroyUri -Method Post -headers $tokenHeaders #-skipcertificatecheck
    }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to login. Please review the error and retry.' + $loginErr}
    }
}
Function Set-NSXTTags{
<#
.SYNOPSIS
                                             Set VM Tags in NSX-T 2.5

This was written with NSX-T 2.5 in mind. It will require re-work to be functional with 3.x and later

Developed by Ralph Stoker (rstoker@vmware.com) & John Brandt (jbrandt@vmware.com)
The purpose of this script is to create new Security Groups in NSX-T using a spreadsheet.

Please reach out with any feedback, suggestions, corrections, or additions. Thank you.

.DESCRIPTION

--------CAUTION------- --------CAUTION------- --------CAUTION-------
THIS PARTICULAR API CALL WILL OVERWRITE ANY AND ALL TAGS CURRENTLY APPLIED TO THE VM
ENSURE THE INPUT FILE HAS ALL TAGS NEEDED FOR THE INDIVIDUAL VMS

This function will use array input in the format that follows. You can import this information from CSV or Excel file, or you can construct it using Powershell objects. 

Spreadsheet Example:
The headers contain VMName, then a listing of Scope/Tags to be applied.  Scope is the header and the tag value is inline with the VM.  In this example we have 2 VMs.  They each will be tagged with Scope = agency & Tag = coc.

vmName,agency,environment,application,tier,loadbalanced,public,http-out
vm1,coc,dev,744sp,shrpnt,,,
vm2,coc,prod,bigapp,web,true,,

.EXAMPLE

Set-NSXTTags -nsxMgrFQDN nsxmgr01.corp.local -xlsVmTags $csv
$csv = import-csv $pathToCSV
Set-NSXTTags -nsxMgrFQDN nsxmgr01.corp.local -xlsVmTags $csv

This example uses a CSV input and will prompt for the credentials to the NSX Manager

.EXAMPLE

Set-NSXTTags -nsxMgrFQDN nsxmgr01.corp.local -xlsVmTags $xlsSheet -inCred $cred
$xlsSheet = Import-Excel -Path $xlsPath -WorksheetName 'Security Groups'
$cred = get-credential
Set-NSXTTags -nsxMgrFQDN nsxmgr01.corp.local -xlsVmTags $xlsSheet -inCred $cred

This example uses a third party module to import an XLS file for input and will a predefined variable for the credentials to the NSX Manager

.OUTPUTS
Array of Powershell Objects for each Service Group created or modified during the execution of this function

#>

[CmdletBinding(SupportsShouldProcess)]

    param(
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory = $True)]
        [Array]$xlsVmTags,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
        )

    #Define Session Variables
    #$nsxMgrFQDN = "nsxapp-01a.corp.local"
    #$csvPath = "C:\Users\Administrator\Desktop\CoC-CPD NSX Firewall worksheet.xlsx"
    $contentType = 'application/json'    
    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $baseUri = $proto+$nsxMgrFQDN
    $nsxGetVMPath = "/api/v1/fabric/virtual-machines"
    
    #Import-Data & Set Counters
    $csvTag = $xlsVmTags
    $vmCount = $csvTag.Count
    $loopcount = 0
    
    #Login
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds 
    
    #Loop through each line of CSV, Get NSX-T VM ID, Create JSON Body from CSV Data, Push new tags to VMs in NSX-T
    foreach ($vm in $csvTag){

    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds 
        Try{
            #Get VM NSX-T External ID Value
            $nsxVmFilter = "?display_name=" + $vm.vmName + "&included_fields=external_id"
            $nsxGetVMUri = $proto + $nsxMgrFQDN + $nsxGetVMPath + $nsxVmFilter
            $pushVmUri = $proto + $nsxMgrFQDN + "/api/v1/fabric/virtual-machines?action=update_tags"
            $getVmTags = Invoke-WebRequest $nsxGetVMUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
            $vmTagsObj = $getVmTags.Content | ConvertFrom-Json
            $vmExternalId = $vmtagsobj.results.external_id    
        }
        Catch
            {
            write-host "Catch Get"
            $e = $_.exception
            write-host $e.Response -ForegroundColor Red
        }

        #Alert user to missing VM and add the issue to the log.
        if($vmTagsObj.result_count -eq "0"){
                $pwd = pwd
                $dateTime = get-date
                $logPath = $pwd.Path + "\Log-vmTagging.txt"
                $message = $dateTime.ToString() + "---The VM you are trying to Tag, " + $vm.vmName + " does not exist, please check your Spreadsheet."
                $message | Out-File -FilePath $logPath -Append
                $loopcount += 1
                Write-Host Processing VM $loopcount of $vmCount 
                write-host $message -ForegroundColor Red 
                Write-Host "LogFile Location" $logPath
                Write-Host "Moving On..." `n 
                continue
            }
       
        #Create blank array to hold VM tags
        $tags = @()

        #Loop through tags and create a new object to add to the array
        foreach($vmProp in ($vm | gm -MemberType NoteProperty)){
            if ($vmProp.Name -ne "vmName"){
                $tagObj = '{"scope":"aeiou","tag":"aeiou"}' | ConvertFrom-Json
                $tagobj.scope = $vmProp.Name
                $tagObj.tag = $vm.($vmProp.Name)
                if(($tagObj.tag -ne $null) -and ( $tagObj.tag -ne "")) {
                    $tags += $tagObj
                }
            }            
        }

        #Create & Populate wrapper object containing VM External ID & required tags
        $tagPush = $null
        $tagPush = New-Object psobject -Property @{
            tags = $tags
            external_id = $vmExternalId
        }

        #Convert wrapper object to JSON for use in API Body
        $tagPushJSON = ConvertTo-Json $tagPush -Depth 5
          
        #Execute API call to Set NSX-T Tags
        try{
            $newtagResponse = Invoke-WebRequest $pushVmUri -Method Post -headers $tokenHeaders -WebSession $session -Body $tagPushJSON -ContentType $contentType -ErrorVariable err #-skipcertificatecheck
            $loopcount += 1
            write-host "push Try" $vm.vmname
            Write-Host Processing VM $loopcount of $vmCount
            Write-Host "NSX-T tags have been applied to" $vm.vmName -ForegroundColor Green
            Write-host "Status Code" $newTagResponse.StatusCode `n            
        }
        catch{
        write-host "Push Catch" $vm.vmname
            $e = $_.exception
            write-host $e.Response -ForegroundColor Red
        }
    }
}
Function New-NSXTSecGrp{
<#
.SYNOPSIS
                                             Create New Security Group in NSX-T 2.5

This was written with NSX-T 2.5 in mind. It will require re-work to be functional with 3.x and later

Developed by Ralph Stoker (rstoker@vmware.com) & John Brandt (jbrandt@vmware.com)
The purpose of this script is to create new Security Groups in NSX-T using a spreadsheet.

Please reach out with any feedback, suggestions, corrections, or additions. Thank you.

.DESCRIPTION
This function will use array input in the format that follows. You can import this information from CSV or Excel file, or you can construct it using Powershell objects. 
This function will return an array containing the Powershell objects of the Service Groups created or modified during the execution.

Spreadsheet Example:
The headers contain Group Name, Description, Scopes/Tags, Ip Sets & statically defined VM Members. Not all fields are needed. If left blank they will be ignored.

grpName,desc,IPaddr,agency,environment,application,tier,loadbalanced,public,http-out,membervmnames
coc_adca_prd,coc Active Directory Certificate Authorities,FALSE,coc,prod,adca,ad,,,,"master1,master2"
coc_addc_prd,coc Active Directory Domain Controllers,FALSE,coc,prod,addc,ad,,,,
coc_ip_grp,coc Active Directory Domain Controllers,"10.10.10.0/24,192.168.1.1",coc,prod,addc,ad,,,,"master1,master12"

.EXAMPLE

new-NSXTSecGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSecGrp $csv
$csv = import-csv $pathToCSV
new-NSXTSecGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSecGrp $csv

This example uses a CSV input and will prompt for the credentials to the NSX Manager

.EXAMPLE

new-NSXTSecGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSecGrp $xlsSheet -inCred $cred
$xlsSheet = Import-Excel -Path $xlsPath -WorksheetName 'Security Groups'
$cred = get-credential
new-NSXTSecGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $xlsSheet -inCred $cred

This example uses a third party module to import an XLS file for input and will a predefined variable for the credentials to the NSX Manager

.OUTPUTS
Array of Powershell Objects for each Service Group created or modified during the execution of this function

#>

    param(
    [Parameter(Mandatory = $True)]
    [String]$nsxMgrFQDN,

    [Parameter(Mandatory = $True)]
    [Array]$xlsSecGrp,

    [Parameter(Mandatory=$True)]
    [System.Management.Automation.PSCredential]
    [System.Management.Automation.Credential()]
    $inCreds
    )

    
    #Define Session Variables
    #$nsxMgrFQDN = "nsxapp-01a.corp.local"
    #$csvPath = "C:\Users\Administrator\Desktop\SecGroupsToCreate.csv"
    $contentType = 'application/json'
    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $baseUri = $proto+$nsxMgrFQDN
    $nsxGroupPath = "infra/domains/default/groups/"
    $nsxGetVMPath = "/api/v1/fabric/virtual-machines"
    $nsxGetSecGroup = "infra/domains/default/groups/"

    #Import-CSV,Set Counters & instantiate results array 
    $csvSecGroup = $xlsSecGrp
    $vmCount = $csvSecGroup.Count
    $loopcount = 0 
    $grpObjs = @()
    
    #Get existing SecGroups for later comparison

    #Login
            $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
    try{
        $getUri = $baseUri + $nsxPolicyAPIPath + $nsxGetSecGroup
        $getSecGrps = Invoke-WebRequest $getUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
        }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to Get existing Security Groups. Please review the error and retry.' + $loginErr}
        }

        $SecGrpsObj = ($getSecGrps.Content | ConvertFrom-Json).results

    #Loop through the Input Array to create or modify Service Groups
        
    #Loop through each line in CSV. Determine the Group details, criteria, ipsets & static members. Create PS object, Convert it to JSON and push it via REST Call
    foreach ($secGroup in $csvSecGroup){
        $proceed = $null

        if($SecGrpsObj | where {($_.display_name -eq $SecGroup.grpName)}){
            ####Construct User Prompt for proceeding with duplicates
            $overwrite = New-Object System.Management.Automation.Host.ChoiceDescription '&Overwrite', 'Overwrite?'
            $append = New-Object System.Management.Automation.Host.ChoiceDescription '&Append', 'Append?'
            $skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip', 'Skip?'
            $options = [System.Management.Automation.Host.ChoiceDescription[]](<#$overwrite,#> $append, $skip)
            $title = 'How to Proceed?'
            $message = "The Security Group $($SecGroup.grpName) already exists, how would you like to proceed?"
            ####
                    
            #Prompt user to decide how to proceed if a policy is duplicated
            $proceed = $host.ui.PromptForChoice($title, $message, $options, 0)
            $secGrpObjUri = ($SecGrpsObj | where {($_.display_name -eq $SecGroup[0].grpName)}).id
            }
            
        switch($proceed){
            #If the proceed value equals NULL, the policy is new and will be created
            $null{
                Write-Host $secGroup.grpName -ForegroundColor Cyan
                #Isolate Resource Type, ID & Display Name and other items in variables
                $ipObj,$tagObj,$vmObj,$memberVmNamesArray,$ipAddrArray,$extIdArray = $null
                $resourceType = "Group"
                $id = $SecGroup.grpName
                $displayName = $SecGroup.grpName
                $description = $secGroup.desc
                $tmpCriteria = @()

                #ensure that IPaddr is not null or empty before splitting into an array
                if(($secGroup.IPaddr -ne $null) -and ($secGroup.ipaddr -ne "")){
                    $ipAddrArray = $secGroup.IPaddr -split ","
                    }

                #ensure that membervmnames is not null or empty before splitting into an array
                if(($secGroup.membervmnames -ne $null) -and ($secGroup.membervmnames -ne "")){
                    $memberVmNamesArray = $secGroup.membervmnames -split ","
                    }             
        
                #create expression array - Root property for dynamic expressions
                $expression = @()

                #create expressions array - Dynamic expressions
                $expressions = @()

                #Populate arrays for membership criteria.  Tags, Ip Sets & Statically assigned VMs
                foreach($grpProp in ($secGroup | gm -MemberType NoteProperty)){
                    if (($grpProp.Name -ne "grpName") -and ($grpProp.Name -ne "desc") -and ($grpProp.Name -ne "IPaddr")-and ($grpProp.Name -ne "membervmnames") -and (($secGroup.($grpProp.Name) -ne $null) -and ($secGroup.($grpProp.Name) -ne ""))){
                        $tagObj = New-Object psobject -Property @{
                            value = $grpProp.Name + "|" + $secGroup.($grpProp.Name)
                            key = "Tag"
                            operator = "EQUALS"
                            resource_type = "Condition"
                            member_type = "VirtualMachine"
                        } 
                        $tmpCriteria += $tagObj
                    } 
                    elseif($grpProp.Name -eq "IPaddr"){

                        #wrapped in an IF to skip if $membervmnamesarray is empty 
                        if($ipAddrArray.Count -ne 0){
                        $ipObj = New-Object psobject -Property @{
                            ip_addresses = $ipAddrArray
                            resource_type = "IPAddressExpression"
                            }
                        }
                    }
                    elseif($grpProp.Name -eq "membervmnames"){

                    #Login
                    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                        $extIdArray = @()

                        #wrapped in an IF to skip if $membervmnamesarray is empty 
                        if($memberVmNamesArray.Count -ne 0){
                            foreach($vmName in $memberVmNamesArray){
                                #Get VM NSX-T External ID Value                        
                                $nsxVmFilter = "?display_name=" + $vmName + "&included_fields=external_id"
                                $nsxGetVMUri = $proto + $nsxMgrFQDN + $nsxGetVMPath + $nsxVmFilter  
                                $getVmId = $null 
                        
                                try{
                                    $getVmId = Invoke-WebRequest $nsxGetVMUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck 
                                    $vmId = $getVmId.Content | ConvertFrom-Json
                                    if($vmId.result_count -ne 0){
                                        $extIdArray += $vmId.results.external_id
                                    }           
                                }
                                catch{
                                    $evmID = $_.exception
                                    write-host $evmID.Response -ForegroundColor Red
                                }
                            }
                        }
                
                        if($extIdArray -ne 0){
                            $vmObj = New-Object psobject -Property @{
                                member_type = "VirtualMachine"
                                resource_type = "ExternalIDExpression"
                                external_ids = $extIdArray
                          }
                       }
                    }              
                }

                #Create Conjunction Operator Object
                $conjunctionObj = '{"conjunction_operator": "AND","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json

                #Loop through Temp Criteria to construct 'Expressions'. Build tag criteria with conjuntion operators inserted ensuring that we conform to the API requriements (single Tag to $expression, Multiple Tags to Expressions)
                for($i = $tmpCriteria.count - 1; $i -ge 0; $i = $i -1){
                    if($tmpCriteria.count -gt 1){
                        $expressions += $tmpCriteria[$i]
                        if($i -ne 0){
                            $expressions += $conjunctionObj
                        }
                    }
                    else{
                        $expression += $tmpCriteria[$i]
                    }
                }

                #Create Object to contain EXPRESSSIONS object               
                $expressionsObj = New-Object psobject -Property @{
                    expressions = $expressions
                    resource_type = 'NestedExpression'
                }
        
                #add Expressions object to expression array
                if($expressions.count -ne 0){
                    $expression += $expressionsObj
                }

                #add IP sets to Expression object if they exist
                if($ipObj -ne $null){
                    $ipConjunctionObj = '{"conjunction_operator": "OR","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json
                    $expression += $ipConjunctionObj
                    $expression += $ipObj
                }

                #add Static VM members to Expression object if they exist
                if($vmObj -ne $null){
                    $vmConjunctionObj = '{"conjunction_operator": "OR","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json
                    $expression += $vmConjunctionObj
                    $expression += $vmObj
                }
        
                #Create Final SecGroup creation Object containing Resource Type, ID, Display Name & Expression array
                $secGroupPush = $null
                $secGroupPush = New-Object psobject -Property @{
                    resource_type = $resourceType
                    display_name = $displayName
                    description = $description
                    expression = $expression
                }

                #Convert final SecGroup object to JSON for use in API Body
                $secGroupPushJSON = ConvertTo-Json $secGroupPush -Depth 5

                #Define Push URI for group creation
                $pushVmUri = $proto + $nsxMgrFQDN + $nsxPolicyAPIPath + $nsxGroupPath + $displayName

                #Login
                $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                #Execute API call to push the new group
                try{
                    $newtagResponse = Invoke-WebRequest $pushVmUri -Method Patch -headers $tokenHeaders -WebSession $session -Body $secGroupPushJSON -ContentType $contentType -ErrorVariable err #-skipcertificatecheck
                    $loopcount += 1
                    Write-Host Creating Group $loopcount of $vmCount
                    Write-Host "Security Group" $displayName "has been created" -ForegroundColor Green `n

                    #Get newly created group info, convert to Powershell Object and add it to an array for future use
                    $getSecGrpUri = $proto + $nsxMgrFQDN + $nsxPolicyAPIPath + $nsxGetSecGroup + $displayName
                    $grpObj = Invoke-WebRequest $getSecGrpUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
                    $grpObjPshell = $grpObj.Content | ConvertFrom-Json
                    $grpObjs += $grpObjPshell
                }
                catch{
                    $e = $_.exception
                    write-host $e.Response -ForegroundColor Red
                }
    }
            "0"{
            Write-Host $secGroup.grpName -ForegroundColor Cyan
                #Isolate Resource Type, ID & Display Name and other items in variables
                $ipObj,$tagObj,$vmObj,$memberVmNamesArray,$ipAddrArray,$extIdArray = $null
                $resourceType = "Group"
                $id = $SecGroup.grpName
                $displayName = $SecGroup.grpName
                $description = $secGroup.desc
                $tmpCriteria = @()

                #ensure that IPaddr is not null or empty before splitting into an array
                if(($secGroup.IPaddr -ne $null) -and ($secGroup.ipaddr -ne "")){
                    $ipAddrArray = $secGroup.IPaddr -split ","
                    }

                #ensure that membervmnames is not null or empty before splitting into an array
                if(($secGroup.membervmnames -ne $null) -and ($secGroup.membervmnames -ne "")){
                    $memberVmNamesArray = $secGroup.membervmnames -split ","
                    }             
        
                #create expression array - Root property for dynamic expressions
                $expression = @()

                #create expressions array - Dynamic expressions
                $expressions = @()

                #Populate arrays for membership criteria.  Tags, Ip Sets & Statically assigned VMs
                foreach($grpProp in ($secGroup | gm -MemberType NoteProperty)){
                    if (($grpProp.Name -ne "grpName") -and ($grpProp.Name -ne "desc") -and ($grpProp.Name -ne "IPaddr")-and ($grpProp.Name -ne "membervmnames") -and (($secGroup.($grpProp.Name) -ne $null) -and ($secGroup.($grpProp.Name) -ne ""))){
                        $tagObj = New-Object psobject -Property @{
                            value = $grpProp.Name + "|" + $secGroup.($grpProp.Name)
                            key = "Tag"
                            operator = "EQUALS"
                            resource_type = "Condition"
                            member_type = "VirtualMachine"
                        } 
                        $tmpCriteria += $tagObj
                    } 
                    elseif($grpProp.Name -eq "IPaddr"){

                        #wrapped in an IF to skip if $membervmnamesarray is empty 
                        if($ipAddrArray.Count -ne 0){
                        $ipObj = New-Object psobject -Property @{
                            ip_addresses = $ipAddrArray
                            resource_type = "IPAddressExpression"
                            }
                        }
                    }
                    elseif($grpProp.Name -eq "membervmnames"){

                    #Login
                    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                        $extIdArray = @()

                        #wrapped in an IF to skip if $membervmnamesarray is empty 
                        if($memberVmNamesArray.Count -ne 0){
                            foreach($vmName in $memberVmNamesArray){
                                #Get VM NSX-T External ID Value                        
                                $nsxVmFilter = "?display_name=" + $vmName + "&included_fields=external_id"
                                $nsxGetVMUri = $proto + $nsxMgrFQDN + $nsxGetVMPath + $nsxVmFilter  
                                $getVmId = $null 
                        
                                try{
                                    $getVmId = Invoke-WebRequest $nsxGetVMUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck 
                                    $vmId = $getVmId.Content | ConvertFrom-Json
                                    if($vmId.result_count -ne 0){
                                        $extIdArray += $vmId.results.external_id
                                    }           
                                }
                                catch{
                                    $evmID = $_.exception
                                    write-host $evmID.Response -ForegroundColor Red
                                }
                            }
                        }
                
                        if($extIdArray -ne 0){
                            $vmObj = New-Object psobject -Property @{
                                member_type = "VirtualMachine"
                                resource_type = "ExternalIDExpression"
                                external_ids = $extIdArray
                          }
                       }
                    }              
                }

                #Create Conjunction Operator Object
                $conjunctionObj = '{"conjunction_operator": "AND","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json

                #Loop through Temp Criteria to construct 'Expressions'. Build tag criteria with conjuntion operators inserted ensuring that we conform to the API requriements (single Tag to $expression, Multiple Tags to Expressions)
                for($i = $tmpCriteria.count - 1; $i -ge 0; $i = $i -1){
                    if($tmpCriteria.count -gt 1){
                        $expressions += $tmpCriteria[$i]
                        if($i -ne 0){
                            $expressions += $conjunctionObj
                        }
                    }
                    else{
                        $expression += $tmpCriteria[$i]
                    }
                }

                #Create Object to contain EXPRESSSIONS object               
                $expressionsObj = New-Object psobject -Property @{
                    expressions = $expressions
                    resource_type = 'NestedExpression'
                }
        
                #add Expressions object to expression array
                if($expressions.count -ne 0){
                    $expression += $expressionsObj
                }

                #add IP sets to Expression object if they exist
                if($ipObj -ne $null){
                    $ipConjunctionObj = '{"conjunction_operator": "OR","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json
                    $expression += $ipConjunctionObj
                    $expression += $ipObj
                }

                #add Static VM members to Expression object if they exist
                if($vmObj -ne $null){
                    $vmConjunctionObj = '{"conjunction_operator": "OR","resource_type": "ConjunctionOperator","marked_for_delete": false,"_protection": "NOT_PROTECTED"}' | ConvertFrom-Json
                    $expression += $vmConjunctionObj
                    $expression += $vmObj
                }
        
                #Create Final SecGroup creation Object containing Resource Type, ID, Display Name & Expression array
                $secGroupPush = $null
                $secGroupPush = New-Object psobject -Property @{
                    resource_type = $resourceType
                    display_name = $displayName
                    description = $description
                    expression = $expression
                }

                #Convert final SecGroup object to JSON for use in API Body
                $secGroupPushJSON = ConvertTo-Json $secGroupPush -Depth 5

                #Define Push URI for group creation
                $pushVmUri = $proto + $nsxMgrFQDN + $nsxPolicyAPIPath + $nsxGroupPath + $displayName

                #Login
                $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                #Execute API call to push the new group
                try{
                    $newtagResponse = Invoke-WebRequest $pushVmUri -Method Patch -headers $tokenHeaders -WebSession $session -Body $secGroupPushJSON -ContentType $contentType -ErrorVariable err #-skipcertificatecheck
                    $loopcount += 1
                    Write-Host Creating Group $loopcount of $vmCount
                    Write-Host "Security Group" $displayName "has been updated" -ForegroundColor Green `n

                    #Get newly created group info, convert to Powershell Object and add it to an array for future use
                    $getSecGrpUri = $proto + $nsxMgrFQDN + $nsxPolicyAPIPath + $nsxGetSecGroup + $displayName
                    $grpObj = Invoke-WebRequest $getSecGrpUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
                    $grpObjPshell = $grpObj.Content | ConvertFrom-Json
                    $grpObjs += $grpObjPshell
                }
                catch{
                    $e = $_.exception
                    write-host $e.Response -ForegroundColor Red
                }
            }
            "1"{
                write-host "You have chosen to SKIP the $($SecGroup.grpName) Policy" -ForegroundColor Red
                break
                }
            }


               
    }
    Write-Host " To access returned PSObject representations of newly created Security Groups view the grpObjs Variable in the script." `n "The variable will be displayed below for your convenience." -ForegroundColor Cyan
    return $grpObjs
}
Function New-NSXTSvcGrp{
<#
.SYNOPSIS
                                             Create New Service/Group in NSX-T 2.5

This was written with NSX-T 2.5 in mind. It will require re-work to be functional with 3.x and later

Developed by Ralph Stoker (rstoker@vmware.com) & John Brandt (jbrandt@vmware.com)
The purpose of this script is to create new Services/Groups in NSX-T using a spreadsheet.

Please reach out with any feedback, suggestions, corrections, or additions. Thank you.

.DESCRIPTION
This function will use array input in the format that follows. You can import this information from CSV or Excel file, or you can construct it using Powershell objects. 
This function will return an array containing the Powershell objects of the Service Groups created or modified during the execution.

Spreadsheet Example:
The headers contain Service Group Name, Description and the TCP and/or UDP port lists

svcGrpName,desc,TCP,UDP
testSvcGrp,Test Service Group,"10,20,30","40,50,60"
testSvcGrp2,Test Service Group 2,"11,21,31","41-45,61"

.EXAMPLE

new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $csv
$csv = import-csv $pathToCSV
new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $csv

This example uses a CSV input and will prompt for the credentials to the NSX Manager

.EXAMPLE

new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $xlsSheet -inCred $cred
$xlsSheet = Import-Excel -Path $xlsPath -WorksheetName 'Service Groups'
$cred = get-credential
new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $xlsSheet -inCred $cred

This example uses a third party module to import an XLS file for input and will a predefined variable for the credentials to the NSX Manager

.OUTPUTS
Array of Powershell Objects for each Service Group created or modified during the execution of this function

#>

    param(
        [Parameter(Mandatory=$False)]
        [string]
        $contentType = 'application/json',
        
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory = $True)]
        [Array]$xlsSvcGrp,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
        )

    #Define Session Variables
    $contentType = 'application/json'
    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $getUri = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/services'
    $baseUri = $proto+$nsxMgrFQDN

    #Old Authentication
    <#

    ##Requires -Version 6.0
    #The -skipcertificatecheck argument on invoke-webrequest ONLY exists in Powershell version 6.0+.
    #If you would like to run this script in a version prior to Powershell 6.0, Place an additional # at the beginning of line 38 AND remove the -skipcertificatecheck argument from ALL Invoke-WebRequest cmdlets in the script

    $loginUri = $baseUri+'/api/session/create'
    $loginBody = @{
      j_username = $inCreds.UserName
      j_password = $inCreds.GetNetworkCredential().Password
     }
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    #Authenticate & Connect to NSX-T API
    try{
        #Get Rest Token
        $login = Invoke-WebRequest $loginUri -Method Post -Body $loginBody -headers @{'Content-Type' = 'application/x-www-form-urlencoded'} -SessionVariable session #-skipcertificatecheck
    }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to login. Please review the error and retry.' + $loginErr}
    }

    #Construct and execute GET REST call for All Services
    
    [string]$tokenString = $login.Headers['X-XSRF-TOKEN']
    $tokenHeaders = @{'X-XSRF-TOKEN' = $tokenString}
    #>
 
    #Login
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

    #Get existing SecGroups for later comparison
    try{
        $getServices = Invoke-WebRequest $getUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
        }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to Get existing Services. Please review the error and retry.' + $loginErr}
        }

    #Create empty array to return new or modified Service Groups
    $grpObjs = @()

    #Loop through the Input Array to create or modify Service Groups
    foreach($svcGrp in $xlsSvcGrp){
        #Populate $proceed variable
        $proceed = 'y'
        
        #Loop through the current list of Service/Groups to determine if the group requested already exists
        foreach($row in ($getServices.Content | ConvertFrom-Json).results){
            if($row.display_name -eq $svcGrp.svcGrpName){
                
                #Confirm overwrite of existing Service Group
                $proceed = Read-Host -Prompt "The Service $($row.display_name) Already Exists. Would you like to overwrite? Y or N?"
                break
                }
            }

        #Create emtpy array to hold Service Entries
        $svcEntries = @()

        #Begin constructing JSON for object.
        foreach($svcProp in $svcGrp | gm -MemberType NoteProperty){
            #Create empty array to contain Service Ports
            $svcPortArr = $null
            
            #Skip Name and Description headers
            if(($svcProp.Name -ne 'svcGrpName') -and ($svcProp.Name -ne 'desc')){
                
                #Validate column contains data
                if(($svcGrp.($svcprop.Name) -ne "") -and ($svcGrp.($svcProp.Name) -ne $null)){
                    
                    #Split based on ',' to get array of Service Ports
                    $svcPortArr = ($svcGrp.($svcProp.Name) -replace ' ','').Split(',')
                    
                    #Create powershell object for Service Entry
                    $svcEntry = New-Object psobject -Property @{
                        resource_type = "L4PortSetServiceEntry"
                        display_name = $svcProp.Name
                        destination_ports = $svcPortArr
                        l4_protocol = $svcProp.Name.ToUpper()
                        }
                    
                    #add new Service Entry to array
                    $svcEntries += $svcEntry
                    }
                }
            }
        
        #Create powershell object for new Service Group
        $svcPSObj = New-Object psobject -Property @{
            display_name = $svcGrp.svcGrpName
            description = $svcGrp.desc
            service_entries = $svcEntries
            }
        
        #Check if existing Service Group should be overwritten
        if(($proceed -ne 'n') -and ($proceed -ne 'no')){
            #convert Powershell object to JSON String
            $svcGrpJson = $svcPSObj | ConvertTo-Json -Depth 10
            
            #Define Push URI for group creation
            $pushSvcGrpUri = $getUri + '/' + $svcPSObj.display_name -replace ' ','_'

            #Login
            $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

            #Execute API call to push the new group
            try{
                $newSvcGrpResponse = Invoke-WebRequest $pushSvcGrpUri -Method Patch -headers $tokenHeaders -WebSession $session -Body $svcGrpJson -ContentType $contentType -ErrorVariable err #-skipcertificatecheck
                Write-Host "Service Group" $svcPSObj.display_name "has been created" -ForegroundColor Green `n

                #Get newly created group info, convert to Powershell Object and add it to an array for future use
                $getSvcGrpUri = $pushSvcGrpUri
                $grpObj = Invoke-WebRequest $getSvcGrpUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
                $grpObjPshell = $grpObj.Content | ConvertFrom-Json
                $grpObjs += $grpObjPshell
                }
            catch{
                #Catch an error that may occur and continue looping through the doc
                $e = $_.exception

                #Print error to screen
                write-host "Unable to create/modify the Service/Group" $svcPSObj.display_name ", with the following error" $e.Response -ForegroundColor Red
                }
            }
        elseif($row.display_name -eq $svcPSObj.display_name){
            $grpObjs += $row
            }
        }
    return $grpObjs
}
Function New-NSXTDfwPolicyandRules{
    <#
    .SYNOPSIS
                                                 Create New Service/Group in NSX-T 2.5

    This was written with NSX-T 2.5 in mind. It will require re-work to be functional with 3.x and later

    Developed by Ralph Stoker (rstoker@vmware.com) & John Brandt (jbrandt@vmware.com)
    The purpose of this script is to create new Services/Groups in NSX-T using a spreadsheet.

    Please reach out with any feedback, suggestions, corrections, or additions. Thank you.

    .DESCRIPTION
    This function will use array input in the format that follows. You can import this information from CSV or Excel file, or you can construct it using Powershell objects. 
    This function will return an array containing the Powershell objects of the Service Groups created or modified during the execution.

    Spreadsheet Example:
    The headers contain Service Group Name, Description and the TCP and/or UDP port lists

    svcGrpName,desc,TCP,UDP
    testSvcGrp,Test Service Group,"10,20,30","40,50,60"
    testSvcGrp2,Test Service Group 2,"11,21,31","41-45,61"

    .EXAMPLE

    new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $csv
    $csv = import-csv $pathToCSV
    new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $csv

    This example uses a CSV input and will prompt for the credentials to the NSX Manager

    .EXAMPLE

    new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $xlsSheet -inCred $cred
    $xlsSheet = Import-Excel -Path $xlsPath -WorksheetName 'Service Groups'
    $cred = get-credential
    new-NSXTSvcGrp -nsxMgrFQDN nsxmgr01.corp.local -xlsSvcGrp $xlsSheet -inCred $cred

    This example uses a third party module to import an XLS file for input and will a predefined variable for the credentials to the NSX Manager

    .OUTPUTS
    Array of Powershell Objects for each Service Group created or modified during the execution of this function

    #>
    
    param(
            [Parameter(Mandatory=$False)]
            [string]
            $contentType = 'application/json',
        
            [Parameter(Mandatory = $True)]
            [String]$nsxMgrFQDN,

            [Parameter(Mandatory = $True)]
            [Array]$xlsSecPolicy,

            [Parameter(Mandatory = $True)]
            [Array]$xlsDfwRules,

            [Parameter(Mandatory=$True)]
            [System.Management.Automation.PSCredential]
            [System.Management.Automation.Credential()]
            $inCreds
            )

    ##Requires -Version 6.0
    #The -skipcertificatecheck argument on invoke-webrequest ONLY exists in Powershell version 6.0+.
    #If you would like to run this script in a version prior to Powershell 6.0, Place an additional # at the beginning of line 38 AND remove the -skipcertificatecheck argument from ALL Invoke-WebRequest cmdlets in the script


    #Define Session Variables
    $newPoliciesOut = @()
    $contentType = 'application/json'
    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $baseUri = $proto+$nsxMgrFQDN

    #Login
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

    #Construct and execute GET REST call for All Policies and Security Groups
    $getPoliciesUri = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/domains/default/security-policies'
    $getSecGrpsUri = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/domains/default/groups'
    $getSvcsUri = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/services'
   
    try{
        $getSecPolicies = (Invoke-WebRequest $getPoliciesUri -Method Get -headers $tokenHeaders -WebSession $session <#-skipcertificatecheck#> | ConvertFrom-Json).results
        }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to Get existing Security Policies. Please review the error and retry.' + $loginErr}
        }
    try{
        $getSecGrps = (Invoke-WebRequest $getSecGrpsUri -Method Get -headers $tokenHeaders -WebSession $session <#-skipcertificatecheck#> | ConvertFrom-Json).results
        }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to Get existing Security Groups. Please review the error and retry.' + $loginErr}
        }
    try{
        $getServices = (Invoke-WebRequest $getSvcsUri -Method Get -headers $tokenHeaders -WebSession $session <#-skipcertificatecheck#> | ConvertFrom-Json).results
        }
    catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        throw {'Failed to Get existing Services. Please review the error and retry.' + $loginErr}
        }
        
    #Create empty array to return new or modified Security Policies
    $SecPolicyObjs = @()
    
    #Loop through the Input Array to create or modify Service Groups
    foreach($SecPolicyIn in $xlsSecPolicy){
        #Nullify $proceed variable
        $proceed = $null
        $polObjUri = $null
            
        #Check for duplicates by Policy Name AND Category
        if($getSecPolicies | where {($_.display_name -eq $SecPolicyIn.policyName) -and ($_.category -eq $SecPolicyIn.category)}){
            ####Construct User Prompt for proceeding with duplicates
            $overwrite = New-Object System.Management.Automation.Host.ChoiceDescription '&Overwrite', 'Overwrite?'
            $append = New-Object System.Management.Automation.Host.ChoiceDescription '&Append', 'Append?'
            $skip = New-Object System.Management.Automation.Host.ChoiceDescription '&Skip', 'Skip?'
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($overwrite, <#$append,#> $skip)
            $title = 'How to Proceed?'
            $message = "The Policy $($SecPolicyIn.policyName) already exists, how would you like to proceed?"
            ####
                    
            #Prompt user to decide how to proceed if a policy is duplicated
            $proceed = $host.ui.PromptForChoice($title, $message, $options, 0)
            $polObjUri = ($getSecPolicies | where {($_.display_name -eq $SecPolicyIn.policyName) -and ($_.category -eq $SecPolicyIn.category)}).id
            }
            
        switch($proceed){
            #If the proceed value equals NULL, the policy is new and will be created
            $null{
                write-host "The new $($SecPolicyIn.policyName) policy will be created"
                
                #Login
                $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
                    
                #Instantiate Error Variable. If Variable is NOT null, DO NOT push Policy
                $polErr = $null

                #Create empty rules array and sequence number
                $rules = @()
                $ruleSeqNum = 0
                
                #Loop through rules for the new policy
                foreach($xlsDfwRule in $xlsDfwRules){
                    if($xlsDfwRule.policyName -eq $SecPolicyIn.policyName){
                        $ruleSeqNum = $ruleSeqNum + 10
                                                
                        #Error handling for missing Security Groups
                        if($xlsDfwRule.src -ne 'any'){
                            try{
                                $srcGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.src.Split(',')}) | ForEach-Object {$_.path}
                                if(!$srcGrps.GetType().isArray){
                                    $srcGrps = $srcGrps.split(',')
                                    }
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more SOURCE security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$srcGrps = @('ANY')}
                                                
                        #Error handling for missing Security Groups
                        if($xlsDfwRule.dst -ne 'any'){
                            try{
                                $dstGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.dst.Split(',')}) | ForEach-Object {$_.path}
                                if(!$dstGrps.GetType().isArray){
                                    $dstGrps = $dstGrps.Split(',')
                                    }
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more DESTINATION security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$dstGrps = @('ANY')}
                        
                        #Error handling for missing service(s)
                        if($xlsDfwRule.svc -ne 'any'){
                            Try{
                                $svcGrps = ($getServices | where {$_.display_name -in $xlsDfwRule.svc.Split(',')}) | ForEach-Object {$_.path}
                                if(!$svcGrps.GetType().isArray){
                                    $svcGrps = $svcGrps.Split(',')
                                    }
                                
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more service groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$svcGrps = @('ANY')}
                        
                        #Populate the applied to/Rule Scope array
                        if($xlsDfwRule.appliedTo -eq "any"){
                            $ruleScope = @('ANY')
                            }
                        else{
                            [array]$ruleScope = $srcGrps + $dstGrps | select -Unique
                            $ruleScope = $ruleScope.ForEach({if($_ -ne 'ANY'){$_}})
                            if(($ruleScope | Measure-Object).Count -eq 0){
                                $ruleScope = @('ANY')
                                }
                            }

                        $rule = New-Object PSObject -Property @{
                            display_name = $xlsDfwRule.ruleName
                            id = $xlsDfwRule.ruleName.Replace(' ','_')
                            description = $xlsDfwRule.desc
                            source_groups = $srcGrps
                            destination_groups = $dstGrps
                            services = $svcGrps
                            scope = $ruleScope
                            sequence_number = $ruleSeqNum
                            logged = $xlsDfwRule.logged.ToString().tolower()
                            action = $xlsDfwRule.action.ToUpper()
                            disabled = $xlsDfwRule.disabled
                            }
                        $rules +=$rule
                        }
                    }

                #Create new Policy Object
                
                try{
                    if($secPolicyIn.appliedTo -ne 'any'){
                        $polScope = ($getSecGrps | where {$_.display_name -in $SecPolicyIn.appliedTo.Split(',')}) | ForEach-Object {$_.path}
                        if(!$polScope.GetType().isArray){
                            $polScope = $polScope.Split(',')
                            }
                        }
                    else{$polScope = @('ANY')}
                    }
                catch{
                    $polErr = 'ERROR'
                    write-host 'ERROR - One or more APPLIED TO security groups for the policy' $SecPolicyIn.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                    }

                $newPolObj = New-Object PSObject -Property @{
                    rules = $rules
                    display_name = $SecPolicyIn.policyName
                    category = $SecPolicyIn.category
                    scope = $polScope
                    description = $SecPolicyIn.desc
                    }

                #Convert Policy Object to JSON for REST call
                $newPolJson = $newPolObj | ConvertTo-Json -Depth 10

                #Construct and execute PUT REST call to create Policy Object
                $polObjUri = $newPolObj.display_name -replace (' ','_')
                $putPolURI = $getPoliciesUri+ '/'+$polObjUri
               
                #Push new Firewall Policy
                if(!$polErr){
                    try{
                        $newPolicyResponse = Invoke-WebRequest $putPolURI -Method Put -headers $tokenHeaders -Body $newPolJson -WebSession $session -ErrorVariable err #-skipcertificatecheck
                        $newPol = $newPolicyResponse.Content | ConvertFrom-Json
                        
                        }
                    catch{
                        $e = $_.exception
                        write-host $e.Response -ForegroundColor Red
                        }
                    }
                else{Write-Host 'The policy' $newPolObj.display_name 'has NOT been created due to previously displayed errors to ensure accurate policy application! Please review the spreadsheet input and confirm all groups exist or have been included for creation' -ForegroundColor Yellow}
                    
                #Add the new policy to an array for output
                $newPoliciesOut += $newPol
                }
                
            #If the proceed value equals overwrite, the policy will be replaced
            0{
                write-host "You have chosen to OVERWRITE the $($SecPolicyIn.policyName) Policy" -ForegroundColor Yellow

                #Login
                $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
                    
                #Instantiate Error Variable. If Variable is NOT null, DO NOT push Policy
                $polErr = $null

                #Create empty rules array and sequence number
                $rules = @()
                $ruleSeqNum = 0

                #Loop through rules for the new policy
                foreach($xlsDfwRule in $xlsDfwRules){
                    if($xlsDfwRule.policyName -eq $SecPolicyIn.policyName){
                        $ruleSeqNum = $ruleSeqNum + 10
                                                
                        #Error handling for missing Security Groups
                        if($xlsDfwRule.src -ne 'any'){
                            try{
                                $srcGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.src.Split(',')}) | ForEach-Object {$_.path}
                                if(!$srcGrps.GetType().isArray){
                                    $srcGrps = $srcGrps.split(',')
                                    }
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more SOURCE security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$srcGrps = @('ANY')}
                                                
                        #Error handling for missing Security Groups
                        if($xlsDfwRule.dst -ne 'any'){
                            try{
                                $dstGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.dst.Split(',')}) | ForEach-Object {$_.path}
                                if(!$dstGrps.GetType().isArray){
                                    $dstGrps = $dstGrps.Split(',')
                                    }
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more DESTINATION security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$dstGrps = @('ANY')}
                        
                        #Error handling for missing service(s)
                        if($xlsDfwRule.svc -ne 'any'){
                            Try{
                                $svcGrps = ($getServices | where {$_.display_name -in $xlsDfwRule.svc.Split(',')}) | ForEach-Object {$_.path}
                                if(!$svcGrps.GetType().isArray){
                                    $svcGrps = $svcGrps.Split(',')
                                    }
                                
                                }
                            catch{
                                $polErr = 'ERROR'
                                write-host 'ERROR - One or more service groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                                }
                            }
                        else{$svcGrps = @('ANY')}
                        
                        #Populate the applied to/Rule Scope array
                        if($xlsDfwRule.appliedTo -eq "any"){
                            $ruleScope = @('ANY')
                            }
                        else{
                            [array]$ruleScope = $srcGrps + $dstGrps | select -Unique
                            $ruleScope = $ruleScope.ForEach({if($_ -ne 'ANY'){$_}})
                            if(($ruleScope | Measure-Object).Count -eq 0){
                                $ruleScope = @('ANY')
                                }
                            }

                        $rule = New-Object PSObject -Property @{
                            display_name = $xlsDfwRule.ruleName
                            id = $xlsDfwRule.ruleName.Replace(' ','_')
                            description = $xlsDfwRule.desc
                            source_groups = $srcGrps
                            destination_groups = $dstGrps
                            services = $svcGrps
                            scope = $ruleScope
                            sequence_number = $ruleSeqNum
                            logged = $xlsDfwRule.logged.ToString().tolower()
                            action = $xlsDfwRule.action.ToUpper()
                            disabled = $xlsDfwRule.disabled
                            }
                        $rules +=$rule
                        }
                    }

                #Create new Policy Object
                try{
                    if($secPolicyIn.appliedTo -ne 'any'){
                        $polScope = ($getSecGrps | where {$_.display_name -in $SecPolicyIn.appliedTo.Split(',')}) | ForEach-Object {$_.path}
                        if(!$polScope.GetType().isArray){
                            $polScope = $polScope.Split(',')
                            }
                        }
                    else{$polScope = @('ANY')}
                    }
                catch{
                    $polErr = 'ERROR'
                    write-host 'ERROR - One or more APPLIED TO security groups for the policy' $SecPolicyIn.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                    }

                $newPolObj = New-Object PSObject -Property @{
                    rules = $rules
                    display_name = $SecPolicyIn.policyName
                    category = $SecPolicyIn.category
                    scope = $polScope
                    description = $SecPolicyIn.desc
                    }

                #Convert Policy Object to JSON for REST call
                $newPolJson = $newPolObj | ConvertTo-Json -Depth 10

                #Construct and execute PUT REST call to create Policy Object
                $polObjUri = $newPolObj.display_name -replace (' ','_')
                $putPolURI = $getPoliciesUri+ '/'+$polObjUri                    

                #Push new Firewall Policy
                if(!$polErr){
                    #Delete the current Sec Policy
                    try{
                        #Login
                        $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
                        Invoke-WebRequest $putPolURI -Method Delete -headers $tokenHeaders -WebSession $session <#-skipcertificatecheck#>
                        }
                    catch{
                        $deletePolErr = $_.exception
                        Write-Host $loginErr.response -ForegroundColor Red
                        throw {'Failed to Delete existing Security Policy. Please review the error and retry.' + $deletePolErr}
                        }
                    
                    Start-Sleep -Milliseconds 3000

                    #Create the replacement policy
                    try{
                        #Login
                        $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                        $newPolicyResponse = Invoke-WebRequest $putPolURI -Method Put -headers $tokenHeaders -Body $newPolJson -WebSession $session -ErrorVariable err #-skipcertificatecheck
                        $newPol = $newPolicyResponse.Content | ConvertFrom-Json
                        
                        }
                    catch{
                        $e = $_.exception
                        write-host $e.Response -ForegroundColor Red
                        }
                    }
                else{Write-Host 'The policy' $newPolObj.display_name 'has NOT been OVERWRITTEN due to previously displayed errors to ensure accurate policy application! Please review the spreadsheet input and confirm all groups exist or have been included for creation' -ForegroundColor Yellow}
                    
                #Add the new policy to an array for output
                $newPoliciesOut += $newPol
                }
                
     
            #If the proceed value equals append, the policy will be appended
            <#
            1{
                write-host "You have chosen to APPEND the $($SecPolicyIn.policyName) Policy" -ForegroundColor Green
                
                #Login
                $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

                #Construct and execute PUT REST call to create Policy Object
                $putPolURI = $getPoliciesUri+ '/'+$polObjUri

                #Get the current Sec Policy
                try{
                    $getSecPolicy = Invoke-WebRequest $putPolURI -Method Get -headers $tokenHeaders -WebSession $session | ConvertFrom-Json #-skipcertificatecheck 
                    }
                catch{
                    $loginErr = $_.exception
                    Write-Host $loginErr.response -ForegroundColor Red
                    throw {'Failed to Get existing Security Policies. Please review the error and retry.' + $loginErr}
                    }
                #Instantiate Error Variable. If Variable is NOT null, DO NOT push Policy
                $polErr = $null

                #Create empty rules array and sequence number
                $rules = @()
                $ruleSeqNum = 0

                #Loop through rules for the new policy
                foreach($xlsDfwRule in $xlsDfwRules){
                    if($xlsDfwRule.policyName -eq $SecPolicyIn.policyName){
                        $ruleSeqNum = $ruleSeqNum + 10
                                                
                        #Error handling for missing Security Groups
                        try{
                            $srcGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.src.Split(',')}) | ForEach-Object {$_.path}
                            if(!$srcGrps.GetType().isArray){
                                $srcGrps = $srcGrps.split(',')
                                }
                            }
                        catch{
                            $polErr = 'ERROR'
                            write-host 'ERROR - One or more SOURCE security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                            }
                                                
                        #Error handling for missing Security Groups
                        try{
                            $dstGrps = ($getSecGrps | where {$_.display_name -in $xlsDfwRule.dst.Split(',')}) | ForEach-Object {$_.path}
                            if(!$dstGrps.GetType().isArray){
                                $dstGrps = $dstGrps.Split(',')
                                }
                            }
                        catch{
                            $polErr = 'ERROR'
                            write-host 'ERROR - One or more DESTINATION security groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                            }
                        
                        #Error handling for missing service(s)
                        Try{
                            $svcGrps = ($getServices | where {$_.display_name -in $xlsDfwRule.svc.Split(',')}) | ForEach-Object {$_.path}
                            if(!$svcGrps.GetType().isArray){
                                $svcGrps = $svcGrps.Split(',')
                                }
                                
                            }
                        catch{
                            $polErr = 'ERROR'
                            write-host 'ERROR - One or more service groups for the rule' $xlsDfwRule.ruleName 'in policy' $xlsDfwRule.policyName 'were not found! This policy will be SKIPPED to ensure accuracy!' -ForegroundColor Red
                            }
                        $rule = New-Object PSObject -Property @{
                            display_name = $xlsDfwRule.ruleName
                            id = $xlsDfwRule.ruleName.Replace(' ','_')
                            description = $xlsDfwRule.desc
                            source_groups = $srcGrps
                            destination_groups = $dstGrps
                            services = $svcGrps
                            scope = $srcGrps + $dstGrps | select -Unique
                            sequence_number = $ruleSeqNum
                            logged = $xlsDfwRule.logged.ToString().tolower()
                            action = $xlsDfwRule.action.ToUpper()
                            }
                        $rules +=$rule
                        }
                    }

                #Create new Policy Object
                $polScope = ($getSecGrps | where {$_.display_name -in $SecPolicyIn.appliedTo.Split(',')}) | ForEach-Object {$_.path}
                if(!$polScope.GetType().isArray){
                    $polScope = $polScope.Split(',')
                    }
                $newPolObj = New-Object PSObject -Property @{
                    rules = $rules
                    display_name = $SecPolicyIn.policyName
                    category = $SecPolicyIn.category
                    scope = $polScope
                    }

                #Convert Policy Object to JSON for REST call
                $newPolJson = $newPolObj | ConvertTo-Json -Depth 10

                #Construct and execute PUT REST call to create Policy Object
                $polObjUri = $newPolObj.display_name -replace (' ','_')
                $putPolURI = $getPoliciesUri+ '/'+$polObjUri
                    

                #Push new Firewall Policy
                if(!$polErr){
                    try{
                        $newPolicyResponse = Invoke-WebRequest $putPolURI -Method Patch -headers $tokenHeaders -Body $newPolJson -WebSession $session -ErrorVariable err #-skipcertificatecheck
                        $newPol = $newPolicyResponse.Content | ConvertFrom-Json
                        
                        }
                    catch{
                        $e = $_.exception
                        write-host $e.Response -ForegroundColor Red
                        }
                    }
                else{Write-Host 'The policy' $newPolObj.display_name 'has NOT been created due to previously displayed errors to ensure accurate policy application! Please review the spreadsheet input and confirm all groups exist or have been included for creation' -ForegroundColor Yellow}
                    
                #Add the new policy to an array for output
                $newPoliciesOut += $newPol
                }
     #>          

            #If the proceed value equals skip, the policy will be skipped
            2{
                write-host "You have chosen to SKIP the $($SecPolicyIn.policyName) Policy" -ForegroundColor Red
                break
                }
            }
        }
   # return $newPoliciesOut
    }
Function Show-Art{
#Image created at http://patorjk.com/software/taag/#p=display&f=Graffiti&t=Type%20Something%20
   Write-host "
    _   ________  __    ______   _________    __  ___   ______            ____              
   / | / / ___/ |/ /   /_  __/  /_  __/   |  /  |/  /  /_  __/___  ____  / / /_  ____  _  __
  /  |/ /\__ \|   /_____/ /      / / / /| | / /|_/ /    / / / __ \/ __ \/ / __ \/ __ \| |/_/
 / /|  /___/ /   /_____/ /      / / / ___ |/ /  / /    / / / /_/ / /_/ / / /_/ / /_/ />  <  
/_/ |_//____/_/|_|    /_/      /_/ /_/  |_/_/  /_/    /_/  \____/\____/_/_.___/\____/_/|_|  
                                                                                            
"
   }
Function Import-Spreadsheet{
    param(
        [Parameter(Mandatory = $True)]
        [string]$xlsPath
        )
    $xlsSvcGrp = Import-Excel -Path $xlsPath -WorksheetName 'Service Groups' -DataOnly
    $xlsVmTags = Import-Excel -Path $xlsPath -WorksheetName 'VM Tags' -DataOnly
    $xlsSecGrp = Import-Excel -Path $xlsPath -WorksheetName 'NSX-T Security Groups' -DataOnly
    $xlsSecPolicy = Import-Excel -Path $xlsPath -WorksheetName 'Firewall Policies' -DataOnly
    $xlsDfwRules = Import-Excel -Path $xlsPath -WorksheetName 'Firewall Rules' -DataOnly
    return $xlsVmTags,$xlsSecGrp,$xlsSvcGrp,$xlsSecPolicy,$xlsDfwRules
}
function Export-PolicyRules{
<#
.SYNOPSIS
                                             Export Policy by Name and all rules in NSX-T 2.5

This was written with NSX-T 2.5 in mind. It will require re-work to be functional with 3.x and later

Developed by Ralph Stoker (rstoker@vmware.com) & John Brandt (jbrandt@vmware.com)
The purpose of this script is to create new Security Groups in NSX-T using a spreadsheet.

Please reach out with any feedback, suggestions, corrections, or additions. Thank you.

.DESCRIPTION
This function will export an Excel file with all of the DFW rules in the user specified policy.

Spreadsheet Example:
The headers contain Group Name, Description, Scopes/Tags, Ip Sets & statically defined VM Members. Not all fields are needed. If left blank they will be ignored.

policyName	SequenceNumber	ruleName	desc	src	dst	svc	action	logged	appliedTo	disabled
agency_dev	2	coc_dev test1					ALLOW	FALSE		FALSE
agency_dev	15	dev-744-shrpnt-to-db	Communication between dev Sharepoint and DB	coc_adca_prd	coc_db_gen_dev	ResTestSvcGrp	ALLOW	FALSE	coc_adca_prd,coc_db_gen_dev	FALSE
agency_dev	25	dev-744-shrpnt-to-dev-744-shrpnt	Communication between dev Sharepoint servers	coc_addc_prd	coc_db_gen_dev,coc_adca_prd	ResTestSvcGrpApi	ALLOW	FALSE	coc_addc_prd,coc_adca_prd,coc_db_gen_dev	FALSE


.EXAMPLE

Export-PolicyRules -nsxMgrFQDN nsxmanager.domain.local -polName agency_dev -outPath c:\temp\

This example will prompt for the credentials to the NSX Manager. The resulting file will be placed where specified in -outPath.

.OUTPUTS
This function will export an Excel file with all of the DFW rules in the user specified policy.

#>
    param(
            
        [Parameter(Mandatory = $True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory = $True)]
        [String]$polName,

        #[Parameter(Mandatory = $True)]
        #[String]$outPath,

        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds
    )

    #Define Session Variables    
    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $baseUri = $proto+$nsxMgrFQDN    
    $nsxPolicyPath = "infra/domains/default/security-policies/" 
    $pwd = pwd
    $outPath = $pwd.Path + "\" + $policyName + ".xlsx"   
    #$polName = "coc_dev"
    #$outPath = "C:\temp\"
    #$OutFileName = $polName + "_RulesExport.xls"
    #$outfilePath = $outPath + $OutFileName
    $getPolUri =  $baseUri + $nsxPolicyAPIPath + $nsxPolicyPath+ $polName

    #authenticate to NSX-T API
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds

    #Get specified Policy and all included rules.
    try{
        $getPolicies = Invoke-WebRequest $getPolUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck 
        }
    Catch{
        $loginErr = $_.exception
        Write-Host $loginErr.response -ForegroundColor Red
        Write-Host 'Failed to Get user supplied policy. Likley the name was mis-typed.' -ForegroundColor Red
        continue
        #throw {'Failed to Get user supplied policy. Please review the error and retry.' + $loginErr}
        }

    #convert JSON Response to PowershellObject
    $PolObj = $getPolicies.Content | ConvertFrom-Json
    $rules = $PolObj.rules # |Export-Csv C:\Temp\Rules.csv -NoTypeInformation

    #Create blank array
    $ruleExport = @()

    #create Policy object for export
    $policyExport = [pscustomobject] @{
        policyName = $polobj.display_name
        desc = $polobj.description
        category = $PolObj.category
        appliedTo = $polobj.scope.ForEach({$_.split("/")[5]}) -join ","
        id = $PolObj.id
    }

    #Loop though rules in Policy and create custom object with the rule fields we are concerned with.
    foreach ($rule in $rules){
        $ruleObj =  [pscustomobject] @{
            policyName = $polName
            SequenceNumber = $rule.sequence_number
            ruleName = $rule.display_name
            desc = $rule.description
            src = $rule.source_groups.ForEach({$_.split("/")[5]}) -join ","
            dst = $rule.destination_groups.ForEach({$_.split("/")[5]}) -join ","
            svc = $rule.services.ForEach({$_.split("/")[3]}) -join ","
            action = $rule.action
            logged = $rule.logged
            appliedTo = $rule.scope.ForEach({$_.split("/")[5]}) -join ","
            disabled = $rule.disabled
        }
        #Add finished object to blank array
        $ruleExport += $ruleObj

        #Export populated array to excel for manual editing
        $ruleExport | Export-Excel $outPath -WorksheetName "Firewall Rules"
    }
    $policyExport | Export-Excel $outPath -WorksheetName "Firewall Policies"
    write-host "The export completed successfully. It can be found at" $outPath -ForegroundColor Green
}
function Copy-NSXTPolicy {
    param(
        [Parameter(Mandatory=$False)]
        [string]
        $contentType = 'application/json',
    
        [Parameter(Mandatory=$True)]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $inCreds,
    
        [Parameter(Mandatory=$True)]
        [String]$nsxMgrFQDN,

        [Parameter(Mandatory=$True)]
        [String]$nsxPolicyName,

        [Parameter(mandatory=$True)]
        [Validateset("Ethernet","Emergency","Infrastructure","Environment","Application")]
        [string]$dfwDestCategory
        )

    ##Requires -Version 6.0
    #The -skipcertificatecheck argument on invoke-webrequest ONLY exists in Powershell version 6.0+.
    #If you would like to run this script in a version prior to Powershell 6.0, Place an additional # at the beginning of line 22 AND remove the -skipcertificatecheck argument from ALL Invoke-WebRequest cmdlets in the script
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $True }
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12


    $proto = 'https://'
    $nsxPolicyAPIPath = '/policy/api/v1/'
    $baseUri = $proto+$nsxMgrFQDN
    
    #authenticate to NSX-T API
    $session,$tokenHeaders = get-NSXTAuth -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
    
    #Construct and execute GET REST call for Policy Object
    $getUri = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/domains/default/security-policies/'    

    try{
        $getPolicies = Invoke-WebRequest $getUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
        }
    catch{
        $e = $_.exception
        write-host $e.Response -ForegroundColor Red
        }

    #Convert GET REST call to Powershell Object
    $allPolObj = $getPolicies.Content | ConvertFrom-Json
    foreach($srcPol in $allpolObj.results){
    
        if($srcPol.display_name -eq $nsxPolicyName){
            break
            }
        }
    $getPolicyUri = $getUri+$srcPol.id
    
    try{
        $getPolicy = Invoke-WebRequest $getPolicyUri -Method Get -headers $tokenHeaders -WebSession $session #-skipcertificatecheck
        }
    catch{
        $e = $_.exception
        write-host $e.Response -ForegroundColor Red
        }

    $policyObj = $getPolicy.Content | ConvertFrom-Json
    #Create array to load rules
    $rules = @()

    #Loop through rules and create a new object to add to the array
    foreach($rule in $policyObj.rules){
        $ruleObj = New-Object psobject -Property @{
            action = $rule.action
            display_name = $rule.display_name
            id = $rule.display_name -Replace ' ','_'
            sequence_number = $rule.sequence_number
            source_groups = $rule.source_groups
            destination_groups = $rule.destination_groups
            services = $rule.services
            logged = $rule.logged
            scope = $rule.scope
            disabled = $rule.disabled
            }
        $rules += $ruleObj
        }

    #Create new Policy Object
    $newPolicyObj = New-Object psobject -Property @{
        rules = $rules
        display_name = $policyObj.display_name
        category = $dfwDestCategory
        scope = $policyObj.scope
        }

    #Convert Policy Object to JSON for REST call
    $newPolJson = $newPolicyObj | ConvertTo-Json -Depth 5

    #Construct and execute PUT REST call to create Policy Object
        #Confirm that ID of new Policy is unique
        foreach($srcPol in $allpolObj.results){
        
            if($srcPol.id -eq $newPolicyObj.display_name -Replace ' ','_'){
                if($srcPol.category -ne $dfwDestCategory){
                    $polObjUri = ($newPolicyObj.display_name -Replace ' ','_')+'_UPDATED'
                    }
                else{
                    $warnMessage = "The Policy: " + $newPolicyObj.display_name + " already exists in the category " + $dfwDestCategory + ". Would you like to continue?"
                    Write-Warning -Message $warnMessage -WarningAction Inquire
                    }
                break
                }
            else{
                $polObjUri = $newPolicyObj.display_name -Replace ' ','_'
                }
            }
    
    #Validate $polObjUri is unique
    if($allPolObj.results | where {$_.id -eq $polObjUri}){
        Write-Host "A policy with the ID" $polObjUri "already exists. Please review existing policies to ensure they are not duplicated." -ForegroundColor Red
        continue
        }


    $putURI = $proto+$nsxMgrFQDN+$nsxPolicyAPIPath+'infra/domains/default/security-policies/'+$polObjUri

    try{
        $newPolicyResponse = Invoke-WebRequest $putURI -Method Put -headers $tokenHeaders -Body $newPolJson -WebSession $session -ErrorVariable err #-skipcertificatecheck
        $newPol = $newPolicyResponse.Content | ConvertFrom-Json
        Write-Host "`nMigration Successful" `n -ForegroundColor Green
        return $newPol
        }
    catch{
        $e = $_.exception
        write-host $e.Response -ForegroundColor Red
        }
    }
Function New-NSXTObjects{
	#Wrapper to execute new NSX Firewall creation

	#Set user defined parameters for use later in the script
	param(
		[Parameter(Mandatory=$False)]
		[string]
		$contentType = 'application/json',
			
		[Parameter(Mandatory = $True)]
		[String]$nsxMgrFQDN,

		[Parameter(Mandatory = $True)]
		[string]$xlsPath,

		[Parameter(Mandatory=$True)]
		[System.Management.Automation.PSCredential]
		[System.Management.Automation.Credential()]
		$inCreds
		)
	
	#Import Excel worksheets
	$xlsVmTags,$xlsSecGrp,$xlsSvcGrp,$xlsSecPolicy,$xlsDfwRules = Import-Spreadsheet -xlsPath $xlsPath
	clear-host

	#Present Menu and execute selected option
	if((show-eula) -eq 0){
		show-Art
		Start-Sleep -Seconds 2
		do{
			Show-Menu -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds
			$menuOption = Read-Host "Please Select an Option"

			switch($menuOption){

				"1"{
					Set-NSXTTags -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsVmTags $xlsVmTags
					}

				"2"{
					new-NSXTSecGrp -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSecGrp $xlsSecGrp
					}

				"3"{
					new-NSXTSvcGrp -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSvcGrp $xlsSvcGrp
					}

				"4"{
					new-NSXTDfwPolicyandRules -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSecPolicy $xlsSecPolicy -xlsDfwRules $xlsDfwRules
					}

				"5"{
					Set-NSXTTags -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsVmTags $xlsVmTags
					new-NSXTSecGrp -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSecGrp $xlsSecGrp
					new-NSXTSvcGrp -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSvcGrp $xlsSvcGrp
					new-NSXTDfwPolicyandRules -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -xlsSecPolicy $xlsSecPolicy -xlsDfwRules $xlsDfwRules
					}
				"6"{
					$policyName = Read-Host -Prompt "Please enter the Firewall policy Name to export"
					
					Export-PolicyRules -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -polName $policyName
					}
				"7"{
					Write-Host "Migrating DFW Policy`n" -ForegroundColor Yellow
					$policyName = Read-Host -Prompt "Please enter the Distributed Firewall policy Name to migrate"
					write-host "`nDistributed Firewall Categories include: Ethernet, Emergency, Infrastructure, Environment, Application`n"
					$destCat = Read-Host -Prompt "Please enter the destination Distributed Firewall Category"
					
					Copy-NSXTPolicy -nsxMgrFQDN $nsxMgrFQDN -inCreds $inCreds -nsxPolicyName $policyName -dfwDestCategory $destCat
					
					}
				"8"{
					Write-Host "Importing New Spreadsheet"
					$xlsVmTags,$xlsSecGrp,$xlsSvcGrp,$xlsSecPolicy,$xlsDfwRules = Import-Spreadsheet -xlsPath $xlsPath
					Write-Host "Import Successful" `n -ForegroundColor Green                
					}
				}
			pause
			}
		until ($menuOption -eq "q")
		}
	else{Write-host "You have chosen to DECLINE the EULA. The script will now exit. Have a nice day" -ForegroundColor Yellow}
	
	}



Export-ModuleMember -Function Get-NSXTAuth,Remove-NSXTAuth,Set-NSXTTags,New-NSXTSecGrp,New-NSXTSvcGrp,New-NSXTDfwPolicyandRules,New-NSXTObjects,Export-PolicyRules,Copy-NSXTPolicy