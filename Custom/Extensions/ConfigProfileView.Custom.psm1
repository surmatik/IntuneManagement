function Get-ConfigProfileViewSupported
{
    param($objectType)

    $objectType -and $objectType.Id -in @("SettingsCatalog", "DeviceConfiguration", "AdministrativeTemplates", "PowerShellScripts", "ConditionalAccess")
}

function New-ConfigProfileViewTab
{
    param(
        [string]$Header = "Configuration"
    )

    $xaml = @"
<TabItem Header="$Header" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0">
        <StackPanel Name="pnlConfigProfileView" Margin="10,8,10,10">
            <TextBlock Name="txtConfigProfileViewIntro"
                       Text="Loading policy settings..."
                       TextWrapping="Wrap"
                       Foreground="#444444"
                       Margin="0,0,0,12" />
            <Grid Name="grdConfigProfileScriptHost" Visibility="Collapsed">
                <TextBox Name="txtConfigProfileScriptText"
                         IsReadOnly="True"
                         AcceptsReturn="True"
                         FontFamily="Consolas"
                         FontSize="13"
                         Padding="12"
                         Background="#FBFBFB"
                         BorderBrush="#D8D8D8"
                         Foreground="#1F1F1F"
                         SelectionBrush="#CCE8FF"
                         SelectionOpacity="0.9"
                         TextWrapping="NoWrap"
                         MinHeight="520"
                         ScrollViewer.VerticalScrollBarVisibility="Auto"
                         ScrollViewer.HorizontalScrollBarVisibility="Auto" />
            </Grid>
        </StackPanel>
    </ScrollViewer>
</TabItem>
"@

    [Windows.Markup.XamlReader]::Parse($xaml)
}

function ConvertTo-ConfigProfileViewString
{
    param(
        $Value,
        [switch]$DecodeBase64Xml
    )

    if($null -eq $Value) { return "" }

    if($Value -is [Array])
    {
        return (($Value | ForEach-Object { ConvertTo-ConfigProfileViewString $_ }) -join [Environment]::NewLine)
    }

    if($Value -is [bool])
    {
        return $Value.ToString().ToLower()
    }

    if($DecodeBase64Xml)
    {
        try
        {
            $bytes = [Convert]::FromBase64String([string]$Value)
            return [System.Text.Encoding]::UTF8.GetString($bytes)
        }
        catch
        {
        }
    }

    if($Value -is [string] -or $Value -is [ValueType])
    {
        return [string]$Value
    }

    try
    {
        return ($Value | ConvertTo-Json -Depth 20 -Compress)
    }
    catch
    {
        return [string]$Value
    }
}

function New-ConfigProfileViewEntry
{
    param(
        [string]$Category,
        [string]$SubCategory,
        [string]$Name,
        [string]$Value,
        [string]$Description = "",
        [int]$Level = 0
    )

    [PSCustomObject]@{
        Category    = $Category
        SubCategory = $SubCategory
        Name        = $Name
        Value       = $Value
        Description = $Description
        Level       = $Level
    }
}

function Get-ConfigProfileViewScriptContent
{
    param($obj)

    if(-not $obj -or -not $obj.scriptContent)
    {
        return ""
    }

    ?? (Get-Base64ScriptContent $obj.scriptContent) ""
}

function Get-ConfigProfileViewFullObject
{
    param($selectedItem, $objectType)

    $obj = $selectedItem.Object
    if(-not $obj -or -not $obj.id -or -not $objectType.API)
    {
        return $obj
    }

    try
    {
        $fullObjInfo = Get-GraphObject $obj $objectType -SkipAssignments
        if($fullObjInfo -and $fullObjInfo.Object)
        {
            return $fullObjInfo.Object
        }
    }
    catch
    {
        Write-Log "Config profile view: failed to refresh full object for $($objectType.Id) / $($obj.id). $($_.Exception.Message)" 2
    }
    finally
    {
        Write-Status ""
    }

    $obj
}

function Get-ConfigProfileViewEntriesForDeviceConfiguration
{
    param($obj)

    $entries = @()
    if(@($obj.omaSettings).Count -gt 0)
    {
        foreach($setting in @($obj.omaSettings))
        {
            $value = $setting.value
            $decodeXml = $setting.'@odata.type' -eq '#microsoft.graph.omaSettingStringXml' -or $setting.'value@odata.type' -eq '#Binary'
            $entries += (New-ConfigProfileViewEntry `
                -Category "Configuration settings" `
                -SubCategory "" `
                -Name (?? $setting.displayName $setting.omaUri) `
                -Value (ConvertTo-ConfigProfileViewString $value -DecodeBase64Xml:$decodeXml) `
                -Description (?? $setting.description $setting.omaUri))
        }

        return $entries
    }

    $ignoreProperties = @(
        "@odata.context", "@odata.type", "@odata.id", "@odata.editLink", "id", "displayName", "description",
        "createdDateTime", "lastModifiedDateTime", "version", "roleScopeTagIds", "supportsScopeTags",
        "assignments", "Assignments"
    )

    foreach($prop in ($obj.PSObject.Properties | Where-Object MemberType -eq NoteProperty))
    {
        if($ignoreProperties -contains $prop.Name) { continue }
        if($prop.Name -like "@*" -or $prop.Name -like "#*") { continue }
        if($null -eq $prop.Value) { continue }
        if($prop.Value -isnot [string] -and $prop.Value -isnot [ValueType] -and $prop.Value -isnot [Array]) { continue }

        $entries += (New-ConfigProfileViewEntry `
            -Category "Configuration settings" `
            -SubCategory "" `
            -Name $prop.Name `
            -Value (ConvertTo-ConfigProfileViewString $prop.Value))
    }

    $entries
}

function Get-ConfigProfileViewSettingCatalogDefinition
{
    param($settingInstance, $settingDefinitions)

    $settingDefinitionId = $settingInstance.settingDefinitionId
    if(-not $settingDefinitionId) { return $null }

    $settingsDef = @($settingDefinitions | Where-Object id -eq $settingDefinitionId)[0]
    if($settingsDef)
    {
        return $settingsDef
    }

    if(-not $script:ConfigProfileViewCachedDefinitions)
    {
        $script:ConfigProfileViewCachedDefinitions = @{}
    }

    if(-not $script:ConfigProfileViewCachedDefinitions.ContainsKey($settingDefinitionId))
    {
        $script:ConfigProfileViewCachedDefinitions[$settingDefinitionId] = (Invoke-GraphRequest -Url "/deviceManagement/configurationSettings/$settingDefinitionId" -ODataMetadata "minimal" -NoError)
    }

    $script:ConfigProfileViewCachedDefinitions[$settingDefinitionId]
}

function Get-ConfigProfileViewSettingsCatalogCategories
{
    if(-not $script:ConfigProfileViewCategories)
    {
        $script:ConfigProfileViewCategories = @{}
        $categoryResponse = Invoke-GraphRequest -Url "/deviceManagement/configurationCategories?`$filter=platforms has 'windows10' and technologies has 'mdm'" -ODataMetadata "minimal" -NoError
        foreach($category in @($categoryResponse.value))
        {
            $script:ConfigProfileViewCategories[$category.id] = $category
        }
    }

    $script:ConfigProfileViewCategories
}

function ConvertTo-ConfigProfileViewSettingsCatalogEntries
{
    param(
        $settingInstance,
        $settingDefinitions,
        [int]$Level = 0
    )

    $entries = @()
    $settingsDef = Get-ConfigProfileViewSettingCatalogDefinition $settingInstance $settingDefinitions
    if(-not $settingsDef)
    {
        return $entries
    }

    $categories = Get-ConfigProfileViewSettingsCatalogCategories
    $categoryDef = $categories[$settingsDef.categoryId]
    $rootCategory = $categoryDef
    $subCategory = $null

    if($categoryDef -and $categoryDef.rootCategoryId -and $categoryDef.rootCategoryId -ne $categoryDef.id)
    {
        $rootCategory = $categories[$categoryDef.rootCategoryId]
        $subCategory = $categoryDef
    }

    $categoryName = if($rootCategory.displayName) { $rootCategory.displayName } else { "Configuration settings" }
    $subCategoryName = if($subCategory.displayName) { $subCategory.displayName } else { "" }
    $name = if($settingsDef.displayName) { $settingsDef.displayName.Trim() } else { $settingsDef.name }
    $description = if($settingsDef.description) { $settingsDef.description.Trim() } else { "" }

    switch ($settingInstance.'@odata.type')
    {
        '#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance'
        {
            $rawValue = $settingInstance.choiceSettingValue.value
            $displayValue = @($settingsDef.options | Where-Object itemId -eq $rawValue)[0].displayName
            if(-not $displayValue) { $displayValue = $rawValue }

            $entries += (New-ConfigProfileViewEntry -Category $categoryName -SubCategory $subCategoryName -Name $name -Value $displayValue -Description $description -Level $Level)

            foreach($childSetting in @($settingInstance.choiceSettingValue.children))
            {
                $entries += ConvertTo-ConfigProfileViewSettingsCatalogEntries -settingInstance $childSetting -settingDefinitions $settingDefinitions -Level ($Level + 1)
            }
        }
        '#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance'
        {
            $entries += (New-ConfigProfileViewEntry -Category $categoryName -SubCategory $subCategoryName -Name $name -Value (ConvertTo-ConfigProfileViewString $settingInstance.simpleSettingValue.value) -Description $description -Level $Level)
        }
        '#microsoft.graph.deviceManagementConfigurationChoiceSettingCollectionInstance'
        {
            $values = foreach($item in @($settingInstance.choiceSettingCollectionValue))
            {
                $resolvedValue = @($settingsDef.options | Where-Object itemId -eq $item.value)[0].displayName
                if($resolvedValue) { $resolvedValue } else { $item.value }
            }

            $entries += (New-ConfigProfileViewEntry -Category $categoryName -SubCategory $subCategoryName -Name $name -Value (($values | Where-Object { $_ }) -join [Environment]::NewLine) -Description $description -Level $Level)
        }
        '#microsoft.graph.deviceManagementConfigurationSimpleSettingCollectionInstance'
        {
            $values = @($settingInstance.simpleSettingCollectionValue | ForEach-Object { $_.value })
            $entries += (New-ConfigProfileViewEntry -Category $categoryName -SubCategory $subCategoryName -Name $name -Value (($values | Where-Object { $null -ne $_ }) -join [Environment]::NewLine) -Description $description -Level $Level)
        }
        '#microsoft.graph.deviceManagementConfigurationGroupSettingCollectionInstance'
        {
            foreach($groupValue in @($settingInstance.groupSettingCollectionValue))
            {
                foreach($childSetting in @($groupValue.children))
                {
                    $entries += ConvertTo-ConfigProfileViewSettingsCatalogEntries -settingInstance $childSetting -settingDefinitions $settingDefinitions -Level ($Level + 1)
                }
            }
        }
        '#microsoft.graph.deviceManagementConfigurationGroupSettingInstance'
        {
            foreach($groupValue in @($settingInstance.groupSettingValue))
            {
                foreach($childSetting in @($groupValue.children))
                {
                    $entries += ConvertTo-ConfigProfileViewSettingsCatalogEntries -settingInstance $childSetting -settingDefinitions $settingDefinitions -Level ($Level + 1)
                }
            }
        }
        default
        {
            $entries += (New-ConfigProfileViewEntry -Category $categoryName -SubCategory $subCategoryName -Name $name -Value (ConvertTo-ConfigProfileViewString $settingInstance) -Description $description -Level $Level)
        }
    }

    $entries
}

function Get-ConfigProfileViewEntriesForSettingsCatalog
{
    param($obj)

    $cfgSettings = (Invoke-GraphRequest -Url "/deviceManagement/configurationPolicies('$($obj.Id)')/settings?`$expand=settingDefinitions&top=1000" -ODataMetadata "minimal" -NoError).Value
    $entries = @()

    foreach($cfgSetting in @($cfgSettings))
    {
        if(-not $cfgSetting.settingInstance) { continue }
        $entries += ConvertTo-ConfigProfileViewSettingsCatalogEntries -settingInstance $cfgSetting.settingInstance -settingDefinitions $cfgSetting.settingDefinitions
    }

    $entries
}

function Get-ConfigProfileViewAdministrativeTemplatePresentationLabel
{
    param($presentationValue)

    if($presentationValue.'#Presentation_Label')
    {
        return $presentationValue.'#Presentation_Label'
    }

    if($presentationValue.presentation -and $presentationValue.presentation.label)
    {
        return $presentationValue.presentation.label
    }

    if($presentationValue.'presentation@odata.bind')
    {
        $presentation = Invoke-GraphRequest -Url $presentationValue.'presentation@odata.bind' -ODataMetadata "minimal" -NoError
        if($presentation -and $presentation.label)
        {
            return $presentation.label
        }
    }

    ""
}

function Get-ConfigProfileViewAdministrativeTemplatePresentationValue
{
    param($presentationValue)

    if($presentationValue.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationValueList')
    {
        return (@($presentationValue.values | ForEach-Object { "$($_.name): $($_.value)" }) -join [Environment]::NewLine)
    }

    if($presentationValue.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationValueMultiText')
    {
        return (@($presentationValue.values) -join [Environment]::NewLine)
    }

    $presentationObject = $presentationValue.presentation
    if(-not $presentationObject -and $presentationValue.'presentation@odata.bind')
    {
        $presentationObject = Invoke-GraphRequest -Url $presentationValue.'presentation@odata.bind' -ODataMetadata "minimal" -NoError
    }

    if($presentationObject -and $presentationObject.'@odata.type' -eq '#microsoft.graph.groupPolicyPresentationDropdownList')
    {
        $mappedValue = @($presentationObject.items | Where-Object value -eq $presentationValue.value)[0].displayName
        if($mappedValue)
        {
            return $mappedValue
        }
    }

    ConvertTo-ConfigProfileViewString $presentationValue.value
}

function Get-ConfigProfileViewEntriesForAdministrativeTemplates
{
    param($obj)

    $definitionValues = @($obj.definitionValues)
    if($definitionValues.Count -eq 0)
    {
        $definitionValues = @((Invoke-GraphRequest -Url "deviceManagement/groupPolicyConfigurations('$($obj.Id)')/definitionValues?`$expand=definition(`$select=id,classType,displayName,policyType,groupPolicyCategoryId)" -ODataMetadata "minimal" -NoError).value)
    }

    $entries = @()
    foreach($definitionValue in $definitionValues)
    {
        $settingName = ?? $definitionValue.'#Definition_displayName' $definitionValue.definition.displayName
        if(-not $settingName) { continue }

        $categoryPath = ?? $definitionValue.'#Definition_categoryPath' $definitionValue.definition.categoryPath
        $combinedValues = @()
        $status = if($definitionValue.enabled -eq $true) { "Enabled" } elseif($definitionValue.enabled -eq $false) { "Disabled" } else { "" }

        foreach($presentationValue in @($definitionValue.presentationValues))
        {
            $label = Get-ConfigProfileViewAdministrativeTemplatePresentationLabel $presentationValue
            $value = Get-ConfigProfileViewAdministrativeTemplatePresentationValue $presentationValue
            if($label -and $value)
            {
                $combinedValues += "${label}: $value"
            }
            elseif($value)
            {
                $combinedValues += $value
            }
        }

        if($status)
        {
            $combinedValues = @($status) + $combinedValues
        }

        $entries += (New-ConfigProfileViewEntry `
            -Category (?? $categoryPath "Administrative Templates") `
            -SubCategory "" `
            -Name $settingName `
            -Value ($combinedValues -join [Environment]::NewLine) `
            -Description (?? $definitionValue.definition.explainText ""))
    }

    $entries
}

function Get-ConfigProfileViewResolvedName
{
    param(
        [string]$Type,
        [string]$Id
    )

    if(-not $Id)
    {
        return ""
    }

    if(-not $script:ConfigProfileViewNameCache)
    {
        $script:ConfigProfileViewNameCache = @{}
    }

    $cacheKey = "$Type|$Id"
    if($script:ConfigProfileViewNameCache.ContainsKey($cacheKey))
    {
        return $script:ConfigProfileViewNameCache[$cacheKey]
    }

    $resolvedName = $Id
    try
    {
        switch ($Type)
        {
            'User'
            {
                $user = Invoke-GraphRequest -Url "/users/$Id?`$select=id,displayName,userPrincipalName" -ODataMetadata "minimal" -NoError
                if($user)
                {
                    $resolvedName = ?? $user.userPrincipalName $user.displayName $Id
                }
            }
            'Group'
            {
                $group = Invoke-GraphRequest -Url "/groups/$Id?`$select=id,displayName" -ODataMetadata "minimal" -NoError
                if($group.displayName)
                {
                    $resolvedName = $group.displayName
                }
            }
            'Application'
            {
                $appResponse = Invoke-GraphRequest -Url "/applications?`$filter=appId eq '$Id'&`$select=id,appId,displayName&`$top=1" -ODataMetadata "minimal" -NoError
                $app = @($appResponse.value)[0]
                if($app.displayName)
                {
                    $resolvedName = $app.displayName
                }
            }
            'NamedLocation'
            {
                $location = Invoke-GraphRequest -Url "/identity/conditionalAccess/namedLocations/$Id?`$select=id,displayName" -ODataMetadata "minimal" -NoError
                if($location.displayName)
                {
                    $resolvedName = $location.displayName
                }
            }
            'TermsOfUse'
            {
                $tou = Invoke-GraphRequest -Url "/identityGovernance/termsOfUse/agreements/$Id?`$select=id,displayName" -ODataMetadata "minimal" -NoError
                if($tou.displayName)
                {
                    $resolvedName = $tou.displayName
                }
            }
            'AuthenticationStrength'
            {
                $strength = Invoke-GraphRequest -Url "/identity/conditionalAccess/authenticationStrength/policies/$Id?`$select=id,displayName" -ODataMetadata "minimal" -NoError
                if($strength.displayName)
                {
                    $resolvedName = $strength.displayName
                }
            }
        }

        if($resolvedName -eq $Id)
        {
            $typeMap = @{
                User  = 'user'
                Group = 'group'
            }

            $body = @{
                ids = @($Id)
            }

            if($typeMap.ContainsKey($Type))
            {
                $body.types = @($typeMap[$Type])
            }

            $directoryObjectResponse = Invoke-GraphRequest -Url "/directoryObjects/getByIds" -HttpMethod "POST" -Content ($body | ConvertTo-Json -Depth 5) -ODataMetadata "minimal" -NoError
            $directoryObject = @($directoryObjectResponse.value)[0]
            if($directoryObject)
            {
                $resolvedName = ?? $directoryObject.userPrincipalName $directoryObject.displayName $directoryObject.appDisplayName $Id
            }
        }
    }
    catch
    {
        Write-Log "Config profile view: failed to resolve $Type '$Id'. $($_.Exception.Message)" 2
    }

    $script:ConfigProfileViewNameCache[$cacheKey] = $resolvedName
    $resolvedName
}

function Get-ConfigProfileViewRoleName
{
    param([string]$Id)

    if(-not $Id)
    {
        return ""
    }

    if(-not $script:ConfigProfileViewRoleCache)
    {
        $script:ConfigProfileViewRoleCache = @{}
        $roleTemplates = (Invoke-GraphRequest -Url "/directoryRoleTemplates?`$select=id,displayName" -ODataMetadata "minimal" -NoError).value
        foreach($roleTemplate in @($roleTemplates))
        {
            if($roleTemplate.id)
            {
                $script:ConfigProfileViewRoleCache[$roleTemplate.id] = ?? $roleTemplate.displayName $roleTemplate.id
            }
        }
    }

    ?? $script:ConfigProfileViewRoleCache[$Id] $Id
}

function ConvertTo-ConfigProfileViewJoinedLines
{
    param($Values)

    @($Values | Where-Object { $_ -ne $null -and $_ -ne "" }) -join [Environment]::NewLine
}

function ConvertTo-ConfigProfileViewConditionalAccessToken
{
    param(
        [string]$Token,
        [string]$Context
    )

    switch ($Token)
    {
        'All' { return 'All users' }
        'None' { return 'None' }
        'GuestsOrExternalUsers' { return 'Guests or external users' }
        'ServicePrincipalsInMyTenant' { return 'Service principals in this tenant' }
        'AllTrusted' { return 'All trusted locations' }
        'AllCompliantDevice' { return 'All compliant devices' }
        'All' { return 'All cloud apps' }
    }

    switch ($Context)
    {
        'State'
        {
            switch ($Token)
            {
                'enabled' { return 'On' }
                'disabled' { return 'Off' }
                'enabledForReportingButNotEnforced' { return 'Report-only' }
            }
        }
        'GrantControl'
        {
            switch ($Token)
            {
                'mfa' { return 'Require multifactor authentication' }
                'compliantDevice' { return 'Require device to be marked as compliant' }
                'domainJoinedDevice' { return 'Require Microsoft Entra hybrid joined device' }
                'approvedApplication' { return 'Require approved client app' }
                'compliantApplication' { return 'Require app protection policy' }
                'passwordChange' { return 'Require password change' }
                'block' { return 'Block access' }
            }
        }
        'ClientApp'
        {
            switch ($Token)
            {
                'all' { return 'All client app types' }
                'browser' { return 'Browser' }
                'mobileAppsAndDesktopClients' { return 'Mobile apps and desktop clients' }
                'exchangeActiveSync' { return 'Exchange ActiveSync clients' }
                'other' { return 'Other clients' }
            }
        }
    }

    $Token
}

function Get-ConfigProfileViewEntriesForConditionalAccess
{
    param($obj)

    $entries = @()

    $entries += (New-ConfigProfileViewEntry -Category 'General' -SubCategory '' -Name 'State' -Value (ConvertTo-ConfigProfileViewConditionalAccessToken $obj.state 'State'))

    $includeUsers = @()
    foreach($id in @($obj.conditions.users.includeUsers))
    {
        if($id -in @('All', 'None', 'GuestsOrExternalUsers'))
        {
            $includeUsers += (ConvertTo-ConfigProfileViewConditionalAccessToken $id 'Users')
        }
        else
        {
            $includeUsers += (Get-ConfigProfileViewResolvedName 'User' $id)
        }
    }
    foreach($id in @($obj.conditions.users.includeGroups))
    {
        $includeUsers += (Get-ConfigProfileViewResolvedName 'Group' $id)
    }
    if(@($obj.conditions.users.includeRoles).Count -gt 0)
    {
        $includeUsers += @($obj.conditions.users.includeRoles | ForEach-Object { "Role: $(Get-ConfigProfileViewRoleName $_)" })
    }
    if($includeUsers)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Assignments' -SubCategory 'Users or agents (Preview)' -Name 'Included users, groups or roles' -Value (ConvertTo-ConfigProfileViewJoinedLines $includeUsers))
    }

    $excludeUsers = @()
    foreach($id in @($obj.conditions.users.excludeUsers))
    {
        if($id -eq 'GuestsOrExternalUsers')
        {
            $excludeUsers += (ConvertTo-ConfigProfileViewConditionalAccessToken $id 'Users')
        }
        else
        {
            $excludeUsers += (Get-ConfigProfileViewResolvedName 'User' $id)
        }
    }
    foreach($id in @($obj.conditions.users.excludeGroups))
    {
        $excludeUsers += (Get-ConfigProfileViewResolvedName 'Group' $id)
    }
    if(@($obj.conditions.users.excludeRoles).Count -gt 0)
    {
        $excludeUsers += @($obj.conditions.users.excludeRoles | ForEach-Object { "Role: $(Get-ConfigProfileViewRoleName $_)" })
    }
    if($excludeUsers)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Assignments' -SubCategory 'Users or agents (Preview)' -Name 'Excluded users, groups or roles' -Value (ConvertTo-ConfigProfileViewJoinedLines $excludeUsers))
    }

    $includeApps = @()
    foreach($id in @($obj.conditions.applications.includeApplications))
    {
        if($id -eq 'All')
        {
            $includeApps += 'All cloud apps'
        }
        else
        {
            $includeApps += (Get-ConfigProfileViewResolvedName 'Application' $id)
        }
    }
    if($includeApps)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Target resources' -SubCategory 'Cloud apps or actions' -Name 'Included applications' -Value (ConvertTo-ConfigProfileViewJoinedLines $includeApps))
    }

    $excludeApps = @()
    foreach($id in @($obj.conditions.applications.excludeApplications))
    {
        if($id -eq 'None')
        {
            $excludeApps += 'None'
        }
        else
        {
            $excludeApps += (Get-ConfigProfileViewResolvedName 'Application' $id)
        }
    }
    if($excludeApps)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Target resources' -SubCategory 'Cloud apps or actions' -Name 'Excluded applications' -Value (ConvertTo-ConfigProfileViewJoinedLines $excludeApps))
    }

    if(@($obj.conditions.applications.includeUserActions).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Target resources' -SubCategory 'Cloud apps or actions' -Name 'User actions' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.applications.includeUserActions))
    }

    if(@($obj.conditions.applications.includeAuthenticationContextClassReferences).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Target resources' -SubCategory 'Cloud apps or actions' -Name 'Authentication context' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.applications.includeAuthenticationContextClassReferences))
    }

    if(@($obj.conditions.userRiskLevels).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'User risk' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.userRiskLevels))
    }

    if(@($obj.conditions.signInRiskLevels).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'Sign-in risk' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.signInRiskLevels))
    }

    if(@($obj.conditions.platforms.includePlatforms).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'Include platforms' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.platforms.includePlatforms))
    }

    if(@($obj.conditions.platforms.excludePlatforms).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'Exclude platforms' -Value (ConvertTo-ConfigProfileViewJoinedLines $obj.conditions.platforms.excludePlatforms))
    }

    $includeLocations = foreach($id in @($obj.conditions.locations.includeLocations))
    {
        if($id -eq 'All') { 'Any location' }
        elseif($id -eq 'AllTrusted') { 'All trusted locations' }
        else { Get-ConfigProfileViewResolvedName 'NamedLocation' $id }
    }
    if($includeLocations)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Network' -SubCategory 'Locations' -Name 'Included locations' -Value (ConvertTo-ConfigProfileViewJoinedLines $includeLocations))
    }

    $excludeLocations = foreach($id in @($obj.conditions.locations.excludeLocations))
    {
        if($id -eq 'AllTrusted') { 'All trusted locations' }
        else { Get-ConfigProfileViewResolvedName 'NamedLocation' $id }
    }
    if($excludeLocations)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Network' -SubCategory 'Locations' -Name 'Excluded locations' -Value (ConvertTo-ConfigProfileViewJoinedLines $excludeLocations))
    }

    if(@($obj.conditions.clientAppTypes).Count -gt 0)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'Client app types' -Value (ConvertTo-ConfigProfileViewJoinedLines (@($obj.conditions.clientAppTypes | ForEach-Object { ConvertTo-ConfigProfileViewConditionalAccessToken $_ 'ClientApp' }))))
    }

    if($obj.conditions.devices.deviceFilter)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Conditions' -SubCategory '' -Name 'Device filter' -Value "$($obj.conditions.devices.deviceFilter.mode): $($obj.conditions.devices.deviceFilter.rule)")
    }

    $grantControls = @()
    foreach($control in @($obj.grantControls.builtInControls))
    {
        $grantControls += (ConvertTo-ConfigProfileViewConditionalAccessToken $control 'GrantControl')
    }
    foreach($id in @($obj.grantControls.termsOfUse))
    {
        $grantControls += ("Terms of use: " + (Get-ConfigProfileViewResolvedName 'TermsOfUse' $id))
    }
    if($obj.grantControls.authenticationStrength)
    {
        $strengthId = ?? $obj.grantControls.authenticationStrength.id $obj.grantControls.authenticationStrength.policyId
        if($strengthId)
        {
            $grantControls += ("Authentication strength: " + (Get-ConfigProfileViewResolvedName 'AuthenticationStrength' $strengthId))
        }
        elseif($obj.grantControls.authenticationStrength.displayName)
        {
            $grantControls += ("Authentication strength: " + $obj.grantControls.authenticationStrength.displayName)
        }
    }
    if(@($obj.grantControls.customAuthenticationFactors).Count -gt 0)
    {
        $grantControls += ("Custom authentication factors: " + (ConvertTo-ConfigProfileViewJoinedLines $obj.grantControls.customAuthenticationFactors))
    }
    if($grantControls)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Grant' -Name 'Requirements' -Value (ConvertTo-ConfigProfileViewJoinedLines $grantControls))
    }
    if($obj.grantControls.operator)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Grant' -Name 'Operator' -Value $obj.grantControls.operator)
    }

    if($obj.sessionControls.applicationEnforcedRestrictions.isEnabled -eq $true)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Session' -Name 'Use app enforced restrictions' -Value 'Enabled')
    }
    if($obj.sessionControls.cloudAppSecurity.isEnabled -eq $true)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Session' -Name 'Use Conditional Access App Control' -Value $obj.sessionControls.cloudAppSecurity.cloudAppSecurityType)
    }
    if($obj.sessionControls.signInFrequency.isEnabled -eq $true)
    {
        $frequencyValue = "$($obj.sessionControls.signInFrequency.value) $($obj.sessionControls.signInFrequency.type)"
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Session' -Name 'Sign-in frequency' -Value $frequencyValue)
    }
    if($obj.sessionControls.continuousAccessEvaluation)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Session' -Name 'Continuous access evaluation' -Value $obj.sessionControls.continuousAccessEvaluation.mode)
    }
    if($obj.sessionControls.persistentBrowser.isEnabled -eq $true)
    {
        $entries += (New-ConfigProfileViewEntry -Category 'Access controls' -SubCategory 'Session' -Name 'Persistent browser session' -Value $obj.sessionControls.persistentBrowser.mode)
    }

    $entries
}

function Add-ConfigProfileViewSeparator
{
    param($panel)

    $separator = New-Object Windows.Shapes.Rectangle
    $separator.Height = 1
    $separator.Fill = '#D9D9D9'
    $separator.Margin = '0,8,0,10'
    [void]$panel.Children.Add($separator)
}

function Add-ConfigProfileViewSummaryText
{
    param(
        $panel,
        [string]$Title,
        [string]$Value,
        [switch]$Subtle
    )

    $titleBlock = New-Object Windows.Controls.TextBlock
    $titleBlock.Text = $Title
    $titleBlock.FontSize = 14
    $titleBlock.Margin = '0,0,0,3'
    [void]$panel.Children.Add($titleBlock)

    $valueBlock = New-Object Windows.Controls.TextBlock
    $valueBlock.Text = $Value
    $valueBlock.TextWrapping = 'Wrap'
    $valueBlock.FontSize = 15
    $valueBlock.Foreground = if($Subtle) { '#444444' } else { '#0F6CBD' }
    $valueBlock.Margin = '8,0,0,0'
    [void]$panel.Children.Add($valueBlock)
}

function Add-ConfigProfileViewConditionalAccessEntryBlock
{
    param(
        $panel,
        $entry
    )

    $nameBlock = New-Object Windows.Controls.TextBlock
    $nameBlock.Text = $entry.Name
    $nameBlock.FontSize = 13
    $nameBlock.Margin = '0,0,0,2'
    $nameBlock.TextWrapping = 'Wrap'
    [void]$panel.Children.Add($nameBlock)

    $valueBlock = New-Object Windows.Controls.TextBlock
    $valueBlock.Text = $entry.Value
    $valueBlock.FontSize = 15
    $valueBlock.Foreground = '#0F6CBD'
    $valueBlock.TextWrapping = 'Wrap'
    $valueBlock.Margin = '8,0,0,8'
    [void]$panel.Children.Add($valueBlock)
}

function Add-ConfigProfileViewConditionalAccessSection
{
    param(
        $panel,
        [string]$Title,
        [string]$Summary,
        $Entries
    )

    Add-ConfigProfileViewSeparator $panel

    $titleBlock = New-Object Windows.Controls.TextBlock
    $titleBlock.Text = $Title
    $titleBlock.FontSize = 14
    $titleBlock.Margin = '0,0,0,4'
    [void]$panel.Children.Add($titleBlock)

    if($Summary)
    {
        $summaryBlock = New-Object Windows.Controls.TextBlock
        $summaryBlock.Text = $Summary
        $summaryBlock.FontSize = 15
        $summaryBlock.Foreground = '#0F6CBD'
        $summaryBlock.TextWrapping = 'Wrap'
        $summaryBlock.Margin = '8,0,0,8'
        [void]$panel.Children.Add($summaryBlock)
    }

    $currentSubCategory = $null
    foreach($entry in @($Entries))
    {
        if($entry.SubCategory -and $entry.SubCategory -ne $currentSubCategory)
        {
            $subTitleBlock = New-Object Windows.Controls.TextBlock
            $subTitleBlock.Text = $entry.SubCategory
            $subTitleBlock.FontSize = 13
            $subTitleBlock.FontWeight = 'SemiBold'
            $subTitleBlock.Margin = '0,2,0,6'
            [void]$panel.Children.Add($subTitleBlock)
            $currentSubCategory = $entry.SubCategory
        }

        Add-ConfigProfileViewConditionalAccessEntryBlock -panel $panel -entry $entry
    }
}

function Get-ConfigProfileViewConditionalAccessAssignmentsSummary
{
    param($obj)

    if(@($obj.conditions.users.includeUsers | Where-Object { $_ -eq 'All' }).Count -gt 0)
    {
        return 'All users'
    }

    if(
        @($obj.conditions.users.includeUsers).Count -gt 0 -or
        @($obj.conditions.users.includeGroups).Count -gt 0 -or
        @($obj.conditions.users.includeRoles).Count -gt 0
    )
    {
        return 'Specific users included'
    }

    'Not configured'
}

function Get-ConfigProfileViewConditionalAccessResourcesSummary
{
    param($obj)

    if(@($obj.conditions.applications.includeApplications | Where-Object { $_ -eq 'All' }).Count -gt 0)
    {
        return "All resources (formerly 'All cloud apps')"
    }

    if(
        @($obj.conditions.applications.includeApplications).Count -gt 0 -or
        @($obj.conditions.applications.includeUserActions).Count -gt 0 -or
        @($obj.conditions.applications.includeAuthenticationContextClassReferences).Count -gt 0
    )
    {
        return 'Specific target resources selected'
    }

    'Not configured'
}

function Get-ConfigProfileViewConditionalAccessNetworkSummary
{
    param($obj)

    $includeLocations = @($obj.conditions.locations.includeLocations)
    $excludeLocations = @($obj.conditions.locations.excludeLocations)

    if(($includeLocations -contains 'All') -and ($excludeLocations -contains 'AllTrusted'))
    {
        return 'Any network or location and all trusted locations excluded'
    }

    if($includeLocations.Count -gt 0 -or $excludeLocations.Count -gt 0)
    {
        return 'Specific network locations selected'
    }

    'Any network or location'
}

function Get-ConfigProfileViewConditionalAccessConditionsSummary
{
    param($obj)

    $selectedConditions = 0
    if(@($obj.conditions.userRiskLevels).Count -gt 0) { $selectedConditions++ }
    if(@($obj.conditions.signInRiskLevels).Count -gt 0) { $selectedConditions++ }
    if(@($obj.conditions.platforms.includePlatforms).Count -gt 0 -or @($obj.conditions.platforms.excludePlatforms).Count -gt 0) { $selectedConditions++ }
    if(@($obj.conditions.clientAppTypes).Count -gt 0) { $selectedConditions++ }
    if($obj.conditions.devices.deviceFilter -and $obj.conditions.devices.deviceFilter.rule) { $selectedConditions++ }
    if(@($obj.conditions.devices.includeDevices).Count -gt 0 -or @($obj.conditions.devices.excludeDevices).Count -gt 0) { $selectedConditions++ }

    if($selectedConditions -eq 0)
    {
        return 'No extra conditions selected'
    }

    if($selectedConditions -eq 1)
    {
        return '1 condition selected'
    }

    "$selectedConditions conditions selected"
}

function Get-ConfigProfileViewConditionalAccessGrantSummary
{
    param($obj)

    $controlCount = @($obj.grantControls.builtInControls).Count
    $controlCount += @($obj.grantControls.termsOfUse).Count
    if($obj.grantControls.authenticationStrength) { $controlCount++ }
    $controlCount += @($obj.grantControls.customAuthenticationFactors).Count

    if($controlCount -eq 0)
    {
        return 'No grant controls selected'
    }

    if($controlCount -eq 1)
    {
        return '1 control selected'
    }

    "$controlCount controls selected"
}

function Get-ConfigProfileViewConditionalAccessSessionSummary
{
    param($obj)

    $controlCount = 0
    if($obj.sessionControls.applicationEnforcedRestrictions.isEnabled -eq $true) { $controlCount++ }
    if($obj.sessionControls.cloudAppSecurity.isEnabled -eq $true) { $controlCount++ }
    if($obj.sessionControls.signInFrequency.isEnabled -eq $true) { $controlCount++ }
    if($obj.sessionControls.continuousAccessEvaluation) { $controlCount++ }
    if($obj.sessionControls.persistentBrowser.isEnabled -eq $true) { $controlCount++ }

    if($controlCount -eq 0)
    {
        return 'No session controls selected'
    }

    if($controlCount -eq 1)
    {
        return '1 control selected'
    }

    "$controlCount controls selected"
}

function Set-ConfigProfileViewConditionalAccessSummary
{
    param(
        $tabItem,
        $obj
    )

    $panel = $tabItem.FindName('pnlConfigProfileView')
    $intro = $tabItem.FindName('txtConfigProfileViewIntro')
    $scriptHost = $tabItem.FindName('grdConfigProfileScriptHost')
    if(-not $panel)
    {
        return
    }

    if($scriptHost)
    {
        $scriptHost.Visibility = 'Collapsed'
    }

    if($intro)
    {
        $intro.Visibility = 'Collapsed'
        $intro.Text = ''
    }

    while($panel.Children.Count -gt 1)
    {
        $panel.Children.RemoveAt(1)
    }

    $nameLabel = New-Object Windows.Controls.TextBlock
    $nameLabel.Text = 'Name'
    $nameLabel.FontSize = 14
    $nameLabel.Margin = '0,0,0,4'
    [void]$panel.Children.Add($nameLabel)

    $nameBox = New-Object Windows.Controls.TextBox
    $nameBox.IsReadOnly = $true
    $nameBox.Text = ?? $obj.displayName ''
    $nameBox.Margin = '0,0,0,10'
    $nameBox.MaxWidth = 420
    [void]$panel.Children.Add($nameBox)

    $entries = @(Get-ConfigProfileViewEntriesForConditionalAccess $obj)

    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Assignments' -Summary (Get-ConfigProfileViewConditionalAccessAssignmentsSummary $obj) -Entries @($entries | Where-Object Category -eq 'Assignments')
    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Target resources' -Summary (Get-ConfigProfileViewConditionalAccessResourcesSummary $obj) -Entries @($entries | Where-Object Category -eq 'Target resources')
    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Network' -Summary (Get-ConfigProfileViewConditionalAccessNetworkSummary $obj) -Entries @($entries | Where-Object Category -eq 'Network')
    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Conditions' -Summary (Get-ConfigProfileViewConditionalAccessConditionsSummary $obj) -Entries @($entries | Where-Object Category -eq 'Conditions')
    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Grant' -Summary (Get-ConfigProfileViewConditionalAccessGrantSummary $obj) -Entries @($entries | Where-Object { $_.Category -eq 'Access controls' -and $_.SubCategory -eq 'Grant' })
    Add-ConfigProfileViewConditionalAccessSection -panel $panel -Title 'Session' -Summary (Get-ConfigProfileViewConditionalAccessSessionSummary $obj) -Entries @($entries | Where-Object { $_.Category -eq 'Access controls' -and $_.SubCategory -eq 'Session' })

    Add-ConfigProfileViewSeparator $panel
}

function Get-ConfigProfileViewEntries
{
    param($selectedItem, $objectType)

    $obj = Get-ConfigProfileViewFullObject $selectedItem $objectType
    if(-not $obj)
    {
        return @()
    }

    switch ($objectType.Id)
    {
        'DeviceConfiguration' { return @(Get-ConfigProfileViewEntriesForDeviceConfiguration $obj) }
        'SettingsCatalog' { return @(Get-ConfigProfileViewEntriesForSettingsCatalog $obj) }
        'AdministrativeTemplates' { return @(Get-ConfigProfileViewEntriesForAdministrativeTemplates $obj) }
        'ConditionalAccess' { return @(Get-ConfigProfileViewEntriesForConditionalAccess $obj) }
    }

    @()
}

function Set-ConfigProfileViewScriptContent
{
    param(
        $tabItem,
        $scriptText
    )

    $panel = $tabItem.FindName('pnlConfigProfileView')
    $intro = $tabItem.FindName('txtConfigProfileViewIntro')
    $scriptHost = $tabItem.FindName('grdConfigProfileScriptHost')
    $scriptTextBox = $tabItem.FindName('txtConfigProfileScriptText')
    if(-not $panel -or -not $intro -or -not $scriptHost -or -not $scriptTextBox)
    {
        return
    }

    while($panel.Children.Count -gt 2)
    {
        $panel.Children.RemoveAt(2)
    }

    $intro.Text = "This view shows the PowerShell script content stored in Intune."
    $scriptHost.Visibility = 'Visible'
    $scriptTextBox.Text = $scriptText
}

function Add-ConfigProfileViewEntryControl
{
    param(
        $panel,
        $entry
    )

    $border = New-Object Windows.Controls.Border
    $border.BorderBrush = [Windows.Media.Brushes]::LightGray
    $border.BorderThickness = '0,0,0,1'
    $border.Padding = '0,10,0,10'

    $grid = New-Object Windows.Controls.Grid
    $column1 = New-Object Windows.Controls.ColumnDefinition
    $column1.Width = '*'
    $column2 = New-Object Windows.Controls.ColumnDefinition
    $column2.Width = '250'
    [void]$grid.ColumnDefinitions.Add($column1)
    [void]$grid.ColumnDefinitions.Add($column2)

    $leftPanel = New-Object Windows.Controls.StackPanel
    $leftPanel.Margin = "$([Math]::Max(0, ($entry.Level * 18))),0,20,0"

    $nameBlock = New-Object Windows.Controls.TextBlock
    $nameBlock.Text = $entry.Name
    $nameBlock.FontSize = 14
    $nameBlock.TextWrapping = 'Wrap'
    [void]$leftPanel.Children.Add($nameBlock)

    if($entry.Description)
    {
        $descriptionBlock = New-Object Windows.Controls.TextBlock
        $descriptionBlock.Text = $entry.Description
        $descriptionBlock.Foreground = '#666666'
        $descriptionBlock.TextWrapping = 'Wrap'
        $descriptionBlock.Margin = '0,4,0,0'
        $descriptionBlock.FontSize = 11
        [void]$leftPanel.Children.Add($descriptionBlock)
    }

    $valueBlock = New-Object Windows.Controls.TextBlock
    $valueBlock.Text = $entry.Value
    $valueBlock.TextWrapping = 'Wrap'
    $valueBlock.FontSize = 14
    $valueBlock.VerticalAlignment = 'Center'
    $valueBlock.Margin = '10,0,0,0'

    [Windows.Controls.Grid]::SetColumn($leftPanel, 0)
    [Windows.Controls.Grid]::SetColumn($valueBlock, 1)
    [void]$grid.Children.Add($leftPanel)
    [void]$grid.Children.Add($valueBlock)

    $border.Child = $grid
    [void]$panel.Children.Add($border)
}

function Set-ConfigProfileViewEntries
{
    param(
        $tabItem,
        $entries
    )

    $panel = $tabItem.FindName('pnlConfigProfileView')
    $intro = $tabItem.FindName('txtConfigProfileViewIntro')
    $scriptHost = $tabItem.FindName('grdConfigProfileScriptHost')
    if(-not $panel -or -not $intro)
    {
        return
    }

    if($scriptHost)
    {
        $scriptHost.Visibility = 'Collapsed'
    }

    while($panel.Children.Count -gt 1)
    {
        $panel.Children.RemoveAt(1)
    }

    if(@($entries).Count -eq 0)
    {
        $intro.Text = "No readable configuration settings could be derived for this profile yet."
        return
    }

    $intro.Text = "This view summarizes the configured settings in a more readable layout, similar to Intune."

    $groupedCategories = $entries | Group-Object Category
    foreach($categoryGroup in $groupedCategories)
    {
        $expander = New-Object Windows.Controls.Expander
        $expander.Header = $categoryGroup.Name
        $expander.IsExpanded = $true
        $expander.Margin = '0,0,0,12'
        $expander.FontSize = 15
        $expander.FontWeight = 'SemiBold'

        $categoryPanel = New-Object Windows.Controls.StackPanel
        $categoryPanel.Margin = '8,10,8,0'

        $lastSubCategory = $null
        foreach($entry in $categoryGroup.Group)
        {
            if($entry.SubCategory -and $entry.SubCategory -ne $lastSubCategory)
            {
                $subCategoryBlock = New-Object Windows.Controls.TextBlock
                $subCategoryBlock.Text = $entry.SubCategory
                $subCategoryBlock.FontWeight = 'SemiBold'
                $subCategoryBlock.Margin = '0,4,0,8'
                $subCategoryBlock.FontSize = 13
                [void]$categoryPanel.Children.Add($subCategoryBlock)
                $lastSubCategory = $entry.SubCategory
            }

            Add-ConfigProfileViewEntryControl -panel $categoryPanel -entry $entry
        }

        $expander.Content = $categoryPanel
        [void]$panel.Children.Add($expander)
    }
}

function Invoke-AfterGraphObjectDetailsCreated
{
    param($detailsForm, $selectedItem, $objectType)

    if(-not (Get-ConfigProfileViewSupported $objectType))
    {
        return
    }

    try
    {
        $existingTab = @($detailsForm.Items | Where-Object { $_.Header -in @('Configuration', 'Script') })[0]
        if($existingTab)
        {
            return
        }

        $tabHeader = if($objectType.Id -eq 'PowerShellScripts') { 'Script' } else { 'Configuration' }
        $configTab = New-ConfigProfileViewTab -Header $tabHeader
        [void]$detailsForm.Items.Insert(0, $configTab)

        if($objectType.Id -eq 'PowerShellScripts')
        {
            $fullObj = Get-ConfigProfileViewFullObject -selectedItem $selectedItem -objectType $objectType
            Set-ConfigProfileViewScriptContent -tabItem $configTab -scriptText (Get-ConfigProfileViewScriptContent $fullObj)
        }
        elseif($objectType.Id -eq 'ConditionalAccess')
        {
            $fullObj = Get-ConfigProfileViewFullObject -selectedItem $selectedItem -objectType $objectType
            Set-ConfigProfileViewConditionalAccessSummary -tabItem $configTab -obj $fullObj
        }
        else
        {
            $entries = Get-ConfigProfileViewEntries -selectedItem $selectedItem -objectType $objectType
            Set-ConfigProfileViewEntries -tabItem $configTab -entries $entries
        }

        $detailsForm.SelectedIndex = 0
    }
    catch
    {
        Write-Log "Config profile view tab failed to load. $($_.Exception.Message)" 3
    }
}
