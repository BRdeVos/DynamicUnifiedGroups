#This script can be used to create Office 365 groups (Unified groups), in bulk, with a dynamic membership. 
#The script is fed by a csv containing columns for the name (can be changed), the owners and the type. The latter defines the name and one of the dynamic membership rules.
#The Allowed Senders to the groups are set to be the group itself and IT (group: Information Technology)
#The creation of the groups is done with Exchange Online, the dynamic membership with AzureAD. To execute the script you need to connect to both.

#Connect-ExchangeOnline
#Connect-AzureAD

#Creating a collection to gather the information of the groups created in part 1 (Exchange) to transfer to part 2 (AzureAD)
$Groups = New-Object System.Collections.ArrayList

Write-Host Creating all groups from CSV

Import-CSV C:\Temp\LocationGroupsTest.csv |

#Part 1 (Exchange), creating the groups
Foreach-Object{ 
    
    #determine the last dynamic membership portion, All Users or Employees only
    if($_.Type -eq "AllUsers"){$DynType = '(user.extension_7ec91e710255474fb4c57ae77f2ad517_employeeType -eq "E" or user.extension_7ec91e710255474fb4c57ae77f2ad517_employeeType -eq "C")'}
        elseif($_.Type -eq "Employees") {$DynType = '(user.extension_7ec91e710255474fb4c57ae77f2ad517_employeeType -eq "E")'}
    
    #Concat the name
    $Name = $_.CountryCode  + "-" + $_.LocationCode + "-" + $_.Type
    #Concat the Dynamic membership rules
    $DynVar= '(user.accountEnabled -eq true) and (user.extension_7ec91e710255474fb4c57ae77f2ad517_iTLocationCode -eq "' + $_.LocationCode + '")' + " and " + $DynType
    
    #Creating the group, including all settings, owners and allowed senders
    New-UnifiedGroup –DisplayName $Name -Language $Language -AccessType Private
    Add-UnifiedGroupLinks -Identity $Name -LinkType Members -Links $_.OfficeMgr,$_.GeneralMgr
    Add-UnifiedGroupLinks -Identity $Name -LinkType Owners -Links $_.OfficeMgr,$_.GeneralMgr
    Set-UnifiedGroup -Identity $Name -AcceptMessagesOnlyFromSendersOrMembers "Information Technology",$Name -AutoSubscribeNewMembers:$true -UnifiedGroupWelcomeMessageEnabled:$false
        
    #Catching the created group in a collection for part 2
    $Group = New-Object System.Object
    $Group | Add-Member -MemberType NoteProperty -Name "Name" -Value $Name
    $Group | Add-Member -MemberType NoteProperty -Name "DynVariable" -Value $DynVar
    $Groups.Add($group) | Out-Null

    }


#Part 2 (AzureAD), setting the dynamic membership
Foreach($Group in $Groups) {
    
    #gather info
    $GroupId = (Get-AzureADMSGroup -SearchString $Group.Name).Id  
    $dynamicMembershipRule = $Group.DynVariable
    $dynamicGroupTypeString = "DynamicMembership"

    Write-Host Making $Group.Name dynamic
        
    [System.Collections.ArrayList]$groupTypes = (Get-AzureAdMsGroup -Id $GroupId).GroupTypes

    #add the dynamic group type to existing types
    $groupTypes.Add($dynamicGroupTypeString)

    #modify the group properties to i) change GroupTypes to add the dynamic type, ii) start execution of the rule, iii) set the rule
    Set-AzureAdMsGroup -Id $GroupId -GroupTypes $groupTypes.ToArray() -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
    
}

Write-Host All groups have been created.

