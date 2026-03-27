function Get-ConfigProfileViewSupported
{
    param($objectType)

    $objectType -and $objectType.Id -in @("SettingsCatalog", "DeviceConfiguration", "AdministrativeTemplates", "PowerShellScripts")
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
