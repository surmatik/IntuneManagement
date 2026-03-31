function Get-ConfigProfileAssignmentSupported
{
    param($objectType)

    $objectType -and $objectType.Id -in @("SettingsCatalog","DeviceConfiguration","AdministrativeTemplates")
}

function New-ConfigProfileAssignmentPanel
{
    $xaml = @"
<Grid xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Name="grdAssignmentsPanel" Margin="10,8,10,5">
    <Grid Margin="10,8,10,5">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <StackPanel>
            <TextBlock Text="Assignments" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,2" />
            <TextBlock Name="txtAssignmentInfo" Text="Open this tab to load and manage direct assignments for this configuration profile." TextWrapping="Wrap" Foreground="#444444" />
        </StackPanel>
        <Grid Grid.Row="1" Margin="0,14,0,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="Auto" />
                <ColumnDefinition Width="Auto" />
            </Grid.ColumnDefinitions>
            <TextBox Name="txtAssignmentGroupSearch" Margin="0,0,8,0" MinWidth="260" ToolTip="Search Entra ID groups by name prefix" />
            <Button Name="btnAssignmentSearch" Grid.Column="1" Content="Add groups" MinWidth="95" Margin="0,0,8,0" />
            <Button Name="btnAssignmentAddAllUsers" Grid.Column="2" Content="Add all users" MinWidth="105" Margin="0,0,8,0" />
            <Button Name="btnAssignmentAddAllDevices" Grid.Column="3" Content="Add all devices" MinWidth="110" />
        </Grid>
        <Border Grid.Row="2" BorderBrush="#D9D9D9" BorderThickness="1" Background="#FAFAFA" Padding="10" Margin="0,0,0,12">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="*" />
                </Grid.RowDefinitions>
                <StackPanel Orientation="Horizontal">
                    <TextBlock Text="Group search results" FontSize="15" FontWeight="SemiBold" />
                    <Button Name="btnAssignmentAddInclude" Content="Add to included" MinWidth="120" Margin="18,0,8,0" />
                    <Button Name="btnAssignmentAddExclude" Content="Add to excluded" MinWidth="120" />
                </StackPanel>
                <ListBox Name="lstAssignmentGroupResults"
                         Grid.Row="1"
                         DisplayMemberPath="DisplayName"
                         MinHeight="90"
                         MaxHeight="180"
                         Margin="0,8,0,0"
                         ScrollViewer.VerticalScrollBarVisibility="Auto"
                         ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                         ScrollViewer.CanContentScroll="True" />
            </Grid>
        </Border>
        <Grid Grid.Row="3">
            <Grid.RowDefinitions>
                <RowDefinition Height="*" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
            </Grid.RowDefinitions>
            <Border BorderBrush="#D9D9D9" BorderThickness="1" Background="White" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>
                    <DockPanel LastChildFill="False">
                        <TextBlock Name="txtIncludedGroupsHeader" Text="Included groups (0)" FontSize="15" FontWeight="SemiBold" />
                        <Button Name="btnAssignmentRemoveInclude" Content="Remove selected" MinWidth="120" DockPanel.Dock="Right" />
                    </DockPanel>
                    <DataGrid Name="dgProfileAssignmentsInclude" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Single" SelectionUnit="FullRow" Margin="0,8,0,0" MinHeight="120">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Groups" Binding="{Binding DisplayName}" Width="*" />
                            <DataGridTextColumn Header="Type" Binding="{Binding TargetType}" Width="120" />
                            <DataGridTextColumn Header="Filter" Binding="{Binding FilterDisplay}" Width="220" />
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
            <Border Grid.Row="1" Background="#EEF5FD" BorderBrush="#D5E6F7" BorderThickness="1" Padding="10" Margin="0,10,0,10">
                <TextBlock Text="When excluding groups, keep in mind that Intune does not allow every mix of user and device targets across include and exclude assignments." TextWrapping="Wrap" />
            </Border>
            <Border Grid.Row="2" BorderBrush="#D9D9D9" BorderThickness="1" Background="White" Padding="10">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto" />
                        <RowDefinition Height="*" />
                    </Grid.RowDefinitions>
                    <DockPanel LastChildFill="False">
                        <TextBlock Name="txtExcludedGroupsHeader" Text="Excluded groups (0)" FontSize="15" FontWeight="SemiBold" />
                        <Button Name="btnAssignmentRemoveExclude" Content="Remove selected" MinWidth="120" DockPanel.Dock="Right" />
                    </DockPanel>
                    <DataGrid Name="dgProfileAssignmentsExclude" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" SelectionMode="Single" SelectionUnit="FullRow" Margin="0,8,0,0" MinHeight="120">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Groups" Binding="{Binding DisplayName}" Width="*" />
                            <DataGridTextColumn Header="Type" Binding="{Binding TargetType}" Width="120" />
                            <DataGridTextColumn Header="Filter" Binding="{Binding FilterDisplay}" Width="220" />
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </Border>
        </Grid>
        <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,12,0,0">
            <Button Name="btnAssignmentRefresh" Content="Refresh" MinWidth="100" Height="30" Padding="14,4" Margin="0,0,8,0" VerticalAlignment="Center" />
            <Button Name="btnAssignmentSave" Content="Review + save" MinWidth="125" Height="30" Padding="14,4" VerticalAlignment="Center" />
        </StackPanel>
    </Grid>
</Grid>
"@

    [Windows.Markup.XamlReader]::Parse($xaml)
}

function Get-ConfigProfileAssignmentGroupName
{
    param($groupId)

    if(-not $groupId) { return "" }
    if(-not $script:ConfigProfileAssignmentGroupCache) { $script:ConfigProfileAssignmentGroupCache = @{} }
    if($script:ConfigProfileAssignmentGroupCache.ContainsKey($groupId)) { return $script:ConfigProfileAssignmentGroupCache[$groupId] }

    $groupName = $null

    $groupObj = Invoke-GraphRequest -Url "/groups/$groupId?`$select=id,displayName" -ODataMetadata "None" -GraphVersion "beta" -NoError
    $groupName = $groupObj.displayName

    if(-not $groupName)
    {
        $filterResponse = Invoke-GraphRequest -Url "/groups?`$select=id,displayName&`$filter=id eq '$groupId'&`$top=1" -ODataMetadata "None" -GraphVersion "beta" -NoError
        $groupName = @($filterResponse.value)[0].displayName
    }

    if(-not $groupName)
    {
        $body = @{
            ids   = @($groupId)
            types = @("group")
        } | ConvertTo-Json -Depth 5

        $byIdsResponse = Invoke-GraphRequest -Url "/directoryObjects/getByIds" -HttpMethod "POST" -Content $body -ODataMetadata "None" -GraphVersion "beta" -NoError
        $groupName = @($byIdsResponse.value)[0].displayName
    }

    if(-not $groupName)
    {
        Write-Log "Could not resolve display name for assignment group '$groupId'. Falling back to id." 2
        $groupName = $groupId
    }

    $script:ConfigProfileAssignmentGroupCache[$groupId] = $groupName
    $groupName
}

function ConvertTo-ConfigProfileAssignmentEntry
{
    param($assignment)

    $assignmentType = ?? $assignment.'@odata.type' ""
    $targetType = ?? $assignment.target.'@odata.type' ""
    $targetMode = "Include"
    $targetLabel = ""
    $entryType = "Other"
    $groupId = ""
    $filterId = ""
    $filterType = ""

    if($assignment.target)
    {
        $groupId = ?? $assignment.target.groupId ""
        $filterId = ?? $assignment.target.deviceAndAppManagementAssignmentFilterId ""
        $filterType = ?? $assignment.target.deviceAndAppManagementAssignmentFilterType ""
    }

    if(-not $groupId -and $assignment.targetGroupId)
    {
        $groupId = $assignment.targetGroupId
    }

    if(-not $targetType -and $groupId)
    {
        if($assignment.excludeGroup -eq $true)
        {
            $targetType = "#legacy.exclusionGroupAssignmentTarget"
        }
        else
        {
            $targetType = "#legacy.groupAssignmentTarget"
        }
    }

    switch ($targetType)
    {
        "#microsoft.graph.groupAssignmentTarget"
        {
            $entryType = "Group"
            $targetLabel = Get-ConfigProfileAssignmentGroupName $groupId
        }
        "#microsoft.graph.exclusionGroupAssignmentTarget"
        {
            $entryType = "Group"
            $targetMode = "Exclude"
            $targetLabel = Get-ConfigProfileAssignmentGroupName $groupId
        }
        "#legacy.groupAssignmentTarget"
        {
            $entryType = "Group"
            $targetLabel = Get-ConfigProfileAssignmentGroupName $groupId
        }
        "#legacy.exclusionGroupAssignmentTarget"
        {
            $entryType = "Group"
            $targetMode = "Exclude"
            $targetLabel = Get-ConfigProfileAssignmentGroupName $groupId
        }
        "#microsoft.graph.allLicensedUsersAssignmentTarget"
        {
            $entryType = "All Users"
            $targetLabel = "All Users"
        }
        "#microsoft.graph.allDevicesAssignmentTarget"
        {
            $entryType = "All Devices"
            $targetLabel = "All Devices"
        }
        default
        {
            $targetMode = "Other"
            $targetLabel = ?? $targetType $assignmentType
        }
    }

    $entry = [PSCustomObject]@{
        DisplayName   = $targetLabel
        TargetMode    = $targetMode
        TargetType    = $entryType
        GroupId       = $groupId
        FilterId      = $filterId
        FilterType    = $filterType
        FilterDisplay = if($filterId) { "$filterType`: $filterId" } else { "" }
        RawAssignment = $assignment
    }

    Write-Log "Converted assignment: assignmentType='$assignmentType', targetType='$targetType', groupId='$groupId', mode='$targetMode', entryType='$entryType', displayName='$targetLabel'"
    $entry
}

function Get-ConfigProfileAssignmentEditorEntries
{
    param($assignments)

    $entries = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    foreach($assignment in @($assignments))
    {
        try
        {
            Write-Log "Raw assignment json: $($assignment | ConvertTo-Json -Depth 20 -Compress)"
        }
        catch
        {
            Write-Log "Raw assignment json could not be serialized" 2
        }
        $entries.Add((ConvertTo-ConfigProfileAssignmentEntry $assignment))
    }
    Write-Log "Editor entries created: $($entries.Count)"
    return ,$entries
}

function Get-NormalizedConfigProfileAssignments
{
    param($value)

    if($null -eq $value) { return @() }
    if($value -is [string]) { return @() }

    if($value.PSObject -and $value.PSObject.Properties["Value"])
    {
        return @(Get-NormalizedConfigProfileAssignments $value.Value)
    }

    if($value.PSObject -and $value.PSObject.Properties["value"])
    {
        return @(Get-NormalizedConfigProfileAssignments $value.value)
    }

    if($value -is [System.Collections.IEnumerable] -and $value -isnot [hashtable] -and -not ($value.PSObject -and $value.PSObject.Properties["target"]))
    {
        return @($value)
    }

    return @($value)
}

function Get-ConfigProfileAssignmentsFromObject
{
    param($objectType, $obj)

    if($obj.Assignments)
    {
        $existingAssignments = @(Get-NormalizedConfigProfileAssignments $obj.Assignments)
        Write-Log "Config profile assignments for $($objectType.Id) / $($obj.Id): using object.Assignments count $($existingAssignments.Count)"
        return $existingAssignments
    }

    if($obj.assignments)
    {
        $existingAssignments = @(Get-NormalizedConfigProfileAssignments $obj.assignments)
        Write-Log "Config profile assignments for $($objectType.Id) / $($obj.Id): using object.assignments count $($existingAssignments.Count)"
        return $existingAssignments
    }

    $assignmentUrl = "$($objectType.API)/$($obj.Id)/assignments"
    $response = Invoke-GraphRequest -Url $assignmentUrl -ODataMetadata "Minimal" -GraphVersion "beta" -NoError
    $assignments = @(Get-NormalizedConfigProfileAssignments $response)
    Write-Log "Config profile assignments for $($objectType.Id) / $($obj.Id): endpoint $assignmentUrl returned $($assignments.Count) item(s)"

    if($assignments)
    {
        if(-not $obj.PSObject.Properties["Assignments"])
        {
            $obj | Add-Member -MemberType NoteProperty -Name "Assignments" -Value $assignments -Force
        }
        else
        {
            $obj.Assignments = $assignments
        }
    }

    @($assignments)
}

function Search-ConfigProfileAssignmentGroups
{
    param($searchText)

    $searchText = "$searchText".Trim()
    $groups = @()

    if(-not $searchText)
    {
        $url = "/groups?`$select=id,displayName&`$top=50&`$orderby=displayName"
        $groups = @((Invoke-GraphRequest -Url $url -ODataMetadata "None" -GraphVersion "beta" -NoError).Value)
    }
    else
    {
        $escapedFilterText = $searchText.Replace("'","''")
        $escapedSearchText = $searchText.Replace('"','\"')

        $searchUrl = "/groups?`$select=id,displayName&`$top=50&`$orderby=displayName&`$search=`"displayName:$escapedSearchText`""
        $groups = @((Invoke-GraphRequest -Url $searchUrl -ODataMetadata "None" -GraphVersion "beta" -AdditionalHeaders @{ ConsistencyLevel = "eventual" } -NoError).Value)

        if(-not $groups)
        {
            $filterUrl = "/groups?`$select=id,displayName&`$top=50&`$filter=startswith(displayName,'$escapedFilterText')"
            $groups = @((Invoke-GraphRequest -Url $filterUrl -ODataMetadata "None" -GraphVersion "beta" -NoError).Value)
        }

        if(-not $groups)
        {
            $fallbackUrl = "/groups?`$select=id,displayName&`$top=250"
            $fallbackGroups = @((Invoke-GraphRequest -Url $fallbackUrl -ODataMetadata "None" -GraphVersion "beta" -AllPages -NoError).Value)
            $groups = @($fallbackGroups | Where-Object { $_.displayName -like "*$searchText*" } | Select-Object -First 50)
        }
    }

    @($groups | Where-Object { $_ -and $_.displayName } | Sort-Object -Property displayName -Unique | ForEach-Object {
        $script:ConfigProfileAssignmentGroupCache[$_.id] = $_.displayName
        [PSCustomObject]@{
            DisplayName = $_.displayName
            Id = $_.id
        }
    })
}

function Add-ConfigProfileAssignmentEntry
{
    param($entry)

    if(-not $entry)
    {
        Write-Log "Assignments UI: no assignment entry was created" 2
        Set-ConfigProfileAssignmentInfo "No assignment entry was created."
        [System.Windows.MessageBox]::Show("No assignment entry was created.", "Assignments", "OK", "Warning") | Out-Null
        return $false
    }
    if($null -eq $script:ConfigProfileAssignmentEntries)
    {
        Write-Log "Assignments UI: entries collection not loaded yet" 2
        Set-ConfigProfileAssignmentInfo "Assignments are not loaded yet. Click Refresh if needed."
        [System.Windows.MessageBox]::Show("Assignments are not loaded yet. Open the Assignments tab first or click Refresh.", "Assignments", "OK", "Warning") | Out-Null
        return $false
    }

    $exists = $script:ConfigProfileAssignmentEntries | Where-Object {
        $_.TargetType -eq $entry.TargetType -and
        $_.TargetMode -eq $entry.TargetMode -and
        $_.GroupId -eq $entry.GroupId
    } | Select-Object -First 1

    if($exists)
    {
        Write-Log "Assignments UI: assignment already exists for groupId '$($entry.GroupId)' mode '$($entry.TargetMode)'"
        Set-ConfigProfileAssignmentInfo "This assignment already exists."
        [System.Windows.MessageBox]::Show("This assignment already exists.", "Assignments", "OK", "Information") | Out-Null
        return $false
    }

    $script:ConfigProfileAssignmentEntries.Add($entry)
    Sync-ConfigProfileAssignmentLists
    Set-ConfigProfileAssignmentInfo "Added '$($entry.DisplayName)' to $($entry.TargetMode.ToLower()) assignments. Included: $($script:ConfigProfileAssignmentIncludeEntries.Count), Excluded: $($script:ConfigProfileAssignmentExcludeEntries.Count)."
    return $true
}

function Sync-ConfigProfileAssignmentLists
{
    if($null -eq $script:ConfigProfileAssignmentEntries) { return }
    if($null -eq $script:ConfigProfileAssignmentIncludeEntries) { return }
    if($null -eq $script:ConfigProfileAssignmentExcludeEntries) { return }

    $script:ConfigProfileAssignmentIncludeEntries.Clear()
    $script:ConfigProfileAssignmentExcludeEntries.Clear()

    foreach($entry in @($script:ConfigProfileAssignmentEntries))
    {
        if($entry.TargetMode -eq "Exclude")
        {
            $script:ConfigProfileAssignmentExcludeEntries.Add($entry)
            continue
        }

        $script:ConfigProfileAssignmentIncludeEntries.Add($entry)
    }

    Update-ConfigProfileAssignmentCounters
    Refresh-ConfigProfileAssignmentViews
    Write-Log "Assignment lists synced. Included=$($script:ConfigProfileAssignmentIncludeEntries.Count), Excluded=$($script:ConfigProfileAssignmentExcludeEntries.Count)"
}

function Update-ConfigProfileAssignmentCounters
{
    if(-not $script:ConfigProfileAssignmentTab) { return }

    Set-XamlProperty $script:ConfigProfileAssignmentTab "txtIncludedGroupsHeader" "Text" "Included groups ($($script:ConfigProfileAssignmentIncludeEntries.Count))"
    Set-XamlProperty $script:ConfigProfileAssignmentTab "txtExcludedGroupsHeader" "Text" "Excluded groups ($($script:ConfigProfileAssignmentExcludeEntries.Count))"
}

function Refresh-ConfigProfileAssignmentViews
{
    if(-not $script:ConfigProfileAssignmentTab) { return }

    $dgInclude = $script:ConfigProfileAssignmentTab.FindName("dgProfileAssignmentsInclude")
    $dgExclude = $script:ConfigProfileAssignmentTab.FindName("dgProfileAssignmentsExclude")

    if($dgInclude)
    {
        $dgInclude.ItemsSource = $null
        $dgInclude.ItemsSource = $script:ConfigProfileAssignmentIncludeEntries
        $dgInclude.Items.Refresh()
    }

    if($dgExclude)
    {
        $dgExclude.ItemsSource = $null
        $dgExclude.ItemsSource = $script:ConfigProfileAssignmentExcludeEntries
        $dgExclude.Items.Refresh()
    }
}

function Select-ConfigProfileAssignmentEntry
{
    param(
        [string]$GridName,
        $Entry
    )

    if(-not $script:ConfigProfileAssignmentTab -or -not $Entry -or -not $GridName) { return }

    try
    {
        $grid = $script:ConfigProfileAssignmentTab.FindName($GridName)
        if(-not $grid) { return }

        $grid.SelectedItem = $Entry
        $grid.ScrollIntoView($Entry)
    }
    catch
    {
        Write-LogError "Failed to select assignment entry '$($Entry.DisplayName)' in grid '$GridName'." $_.Exception
    }
}

function Set-ConfigProfileAssignmentInfo
{
    param([string]$Message)

    if(-not $script:ConfigProfileAssignmentTab) { return }
    Set-XamlProperty $script:ConfigProfileAssignmentTab "txtAssignmentInfo" "Text" $Message
}

function Remove-ConfigProfileAssignmentEntry
{
    param(
        $entry,
        [switch]$SkipConfirmation
    )

    if(-not $entry) { return }
    if($null -eq $script:ConfigProfileAssignmentEntries) { return }

    if($SkipConfirmation -ne $true)
    {
        $message = "Do you want to remove this assignment?`n`nTarget:`n$($entry.DisplayName)`n`nType: $($entry.TargetType)`nMode: $($entry.TargetMode)"
        if(([System.Windows.MessageBox]::Show($message, "Remove assignment?", "YesNo", "Warning")) -ne "Yes")
        {
            Set-ConfigProfileAssignmentInfo "Removal canceled for '$($entry.DisplayName)'."
            return $false
        }
    }

    [void]$script:ConfigProfileAssignmentEntries.Remove($entry)
    Sync-ConfigProfileAssignmentLists
    Set-ConfigProfileAssignmentInfo "Removed '$($entry.DisplayName)'. Included: $($script:ConfigProfileAssignmentIncludeEntries.Count), Excluded: $($script:ConfigProfileAssignmentExcludeEntries.Count)."
    return $true
}

function New-ConfigProfileAssignmentEntry
{
    param(
        [string]$TargetType,
        [string]$TargetMode,
        [string]$DisplayName,
        [string]$GroupId = ""
    )

    [PSCustomObject]@{
        DisplayName   = $DisplayName
        TargetMode    = $TargetMode
        TargetType    = $TargetType
        GroupId       = $GroupId
        FilterId      = ""
        FilterType    = ""
        FilterDisplay = ""
        RawAssignment = $null
    }
}

function ConvertTo-ConfigProfileGraphAssignment
{
    param($entry)

    if($entry.TargetMode -eq "Other" -and $entry.RawAssignment)
    {
        return ($entry.RawAssignment | ConvertTo-Json -Depth 50 | ConvertFrom-Json)
    }

    $target = [ordered]@{}
    switch ("$($entry.TargetType)|$($entry.TargetMode)")
    {
        "Group|Include"
        {
            $target["@odata.type"] = "#microsoft.graph.groupAssignmentTarget"
            $target["groupId"] = $entry.GroupId
        }
        "Group|Exclude"
        {
            $target["@odata.type"] = "#microsoft.graph.exclusionGroupAssignmentTarget"
            $target["groupId"] = $entry.GroupId
        }
        "All Users|Include"
        {
            $target["@odata.type"] = "#microsoft.graph.allLicensedUsersAssignmentTarget"
        }
        "All Devices|Include"
        {
            $target["@odata.type"] = "#microsoft.graph.allDevicesAssignmentTarget"
        }
        default
        {
            return $null
        }
    }

    if($entry.FilterId)
    {
        $target["deviceAndAppManagementAssignmentFilterId"] = $entry.FilterId
        $target["deviceAndAppManagementAssignmentFilterType"] = $entry.FilterType
    }

    [PSCustomObject]@{
        target = [PSCustomObject]$target
    }
}

function Invoke-ConfigProfileAssignmentAutoSave
{
    param([string]$Reason = "")

    Write-Log "Assignments UI: auto-save requested. Reason='$Reason'"

    try
    {
        $saved = Save-ConfigProfileAssignments -SkipConfirmation -Silent
        Write-Log "Assignments UI: auto-save completed. Success=$saved"
        return [bool]$saved
    }
    catch
    {
        Write-LogError "Assignments UI: auto-save failed." $_.Exception
        Set-ConfigProfileAssignmentInfo "Assignments could not be saved. Check the log for details."
        [System.Windows.MessageBox]::Show("Assignments could not be saved. Check the log for details.", "Assignments", "OK", "Warning") | Out-Null
        return $false
    }
}

function Update-ConfigProfileAssignmentJsonForEnvironment
{
    param([string]$Json)

    if([string]::IsNullOrWhiteSpace($Json)) { return $Json }

    $importPath = ""
    try
    {
        if($global:txtImportPath)
        {
            $importPath = "$($global:txtImportPath.Text)"
        }
    }
    catch
    {
        $importPath = ""
    }

    if([string]::IsNullOrWhiteSpace($importPath))
    {
        return $Json
    }

    Update-JsonForEnvironment $Json
}

function Save-ConfigProfileAssignments
{
    param(
        [switch]$SkipConfirmation,
        [switch]$Silent
    )

    Write-Log "Save-ConfigProfileAssignments entered. Silent=$Silent SkipConfirmation=$SkipConfirmation"

    if(-not $script:ConfigProfileAssignmentObjectType -or -not $script:ConfigProfileAssignmentCurrentObject)
    {
        Write-Log "Save-ConfigProfileAssignments aborted because assignment context is missing." 2
        Set-ConfigProfileAssignmentInfo "Assignments could not be saved because the current object is missing."
        return $false
    }

    $currentObjectName = Get-GraphObjectName $script:ConfigProfileAssignmentCurrentObject $script:ConfigProfileAssignmentObjectType
    if($SkipConfirmation -ne $true -and ([System.Windows.MessageBox]::Show("Are you sure you want to update assignments?`n`nObject:`n$currentObjectName", "Save Assignments?", "YesNo", "Warning")) -ne "Yes")
    {
        Write-Log "Save-ConfigProfileAssignments canceled by user."
        return $false
    }

    Write-Status "Saving assignments for $currentObjectName"
    Set-ConfigProfileAssignmentInfo "Saving assignments to Intune..."

    try
    {
        $graphAssignments = @()
        foreach($entry in @($script:ConfigProfileAssignmentEntries))
        {
            $assignment = ConvertTo-ConfigProfileGraphAssignment $entry
            if($assignment) { $graphAssignments += $assignment }
        }

        $api = "$($script:ConfigProfileAssignmentObjectType.API)/$($script:ConfigProfileAssignmentCurrentObject.Id)/assign"
        $assignmentType = ?? $script:ConfigProfileAssignmentObjectType.AssignmentsType "assignments"
        $body = @{ $assignmentType = @($graphAssignments) } | ConvertTo-Json -Depth 30
        $body = Update-ConfigProfileAssignmentJsonForEnvironment $body

        Write-Log "Saving config profile assignments using $api"
        Write-Log "Assignments payload: $body"
        $result = Invoke-GraphRequest -Url $api -HttpMethod "POST" -Content $body -GraphVersion "beta"
        $resultText = "<no content>"
        if($null -ne $result)
        {
            try
            {
                $resultText = $result | ConvertTo-Json -Depth 10 -Compress
            }
            catch
            {
                $resultText = "<non-serializable response>"
            }
        }
        Write-Log "Assignments save result: $resultText"

        Load-ConfigProfileAssignments -ForceReloadAssignments
        if($Silent -ne $true)
        {
            [System.Windows.MessageBox]::Show("Assignments updated successfully.", "Assignments", "OK", "Information") | Out-Null
        }
        else
        {
            Set-ConfigProfileAssignmentInfo "Assignments updated successfully in Intune."
        }

        return $true
    }
    catch
    {
        Write-LogError "Save-ConfigProfileAssignments failed." $_.Exception
        Set-ConfigProfileAssignmentInfo "Assignments could not be saved. Check the log for details."
        if($Silent -ne $true)
        {
            [System.Windows.MessageBox]::Show("Assignments could not be saved. Check the log for details.", "Assignments", "OK", "Warning") | Out-Null
        }
        return $false
    }
    finally
    {
        Write-Status ""
    }
}

function Load-ConfigProfileAssignments
{
    param(
        [switch]$ReloadObject,
        [switch]$ForceReloadAssignments
    )

    if($script:ConfigProfileAssignmentsLoading -eq $true)
    {
        Write-Log "Load-ConfigProfileAssignments skipped because a previous load is still running."
        return
    }
    $script:ConfigProfileAssignmentsLoading = $true

    try
    {
        Write-Log "Load-ConfigProfileAssignments started. ReloadObject=$ReloadObject ForceReloadAssignments=$ForceReloadAssignments"
        Set-ConfigProfileAssignmentInfo "Loading assignments from Intune..."

        if($ReloadObject -or -not $script:ConfigProfileAssignmentCurrentObject)
        {
            $script:ConfigProfileAssignmentCurrentObject = (Get-GraphObject $script:ConfigProfileAssignmentSelectedItem.Object $script:ConfigProfileAssignmentObjectType).Object
            if($script:ConfigProfileAssignmentCurrentObject)
            {
                $script:ConfigProfileAssignmentSelectedItem.Object = $script:ConfigProfileAssignmentCurrentObject
                if($script:ConfigProfileAssignmentDetailsForm)
                {
                    Set-XamlProperty $script:ConfigProfileAssignmentDetailsForm "txtValue" "Text" (ConvertTo-Json $script:ConfigProfileAssignmentCurrentObject -Depth 50)
                }
            }
        }

        if($ForceReloadAssignments -eq $true -and $script:ConfigProfileAssignmentCurrentObject)
        {
            if($script:ConfigProfileAssignmentCurrentObject.PSObject.Properties["Assignments"])
            {
                $script:ConfigProfileAssignmentCurrentObject.Assignments = $null
            }
            if($script:ConfigProfileAssignmentCurrentObject.PSObject.Properties["assignments"])
            {
                $script:ConfigProfileAssignmentCurrentObject.assignments = $null
            }
        }

        $assignments = Get-ConfigProfileAssignmentsFromObject $script:ConfigProfileAssignmentObjectType $script:ConfigProfileAssignmentCurrentObject
        $script:ConfigProfileAssignmentEntries = Get-ConfigProfileAssignmentEditorEntries $assignments
        Sync-ConfigProfileAssignmentLists
        Set-ConfigProfileAssignmentInfo "Manage direct assignments for this configuration profile. Existing assignment filters are preserved, but filters can't be edited here. Raw assignments: $(@($assignments).Count), Included: $($script:ConfigProfileAssignmentIncludeEntries.Count), Excluded: $($script:ConfigProfileAssignmentExcludeEntries.Count)."
        $script:ConfigProfileAssignmentsLoaded = $true
    }
    finally
    {
        $script:ConfigProfileAssignmentsLoading = $false
    }
}

function Initialize-ConfigProfileAssignmentsForm
{
    param($root, $selectedItem, $objectType, $detailsForm = $null)

    if(-not (Get-ConfigProfileAssignmentSupported $objectType)) { return }
    if(-not $root -or -not $selectedItem -or -not $selectedItem.Object) { return }

    $script:ConfigProfileAssignmentTab = $root
    $script:ConfigProfileAssignmentDetailsForm = $detailsForm
    $script:ConfigProfileAssignmentSelectedItem = $selectedItem
    $script:ConfigProfileAssignmentObjectType = $objectType
    $script:ConfigProfileAssignmentCurrentObject = $null
    $script:ConfigProfileAssignmentGroupCache = @{}
    $script:ConfigProfileAssignmentSearchResults = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:ConfigProfileAssignmentIncludeEntries = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:ConfigProfileAssignmentExcludeEntries = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
    $script:ConfigProfileAssignmentsLoaded = $false
    $script:ConfigProfileAssignmentsLoading = $false

    Write-Log "Initializing assignments form for objectType '$($objectType.Id)' and object '$($selectedItem.Object.Id)'"

    Set-XamlProperty $root "lstAssignmentGroupResults" "ItemsSource" $script:ConfigProfileAssignmentSearchResults
    Set-XamlProperty $root "dgProfileAssignmentsInclude" "ItemsSource" $script:ConfigProfileAssignmentIncludeEntries
    Set-XamlProperty $root "dgProfileAssignmentsExclude" "ItemsSource" $script:ConfigProfileAssignmentExcludeEntries
    Set-XamlProperty $root "txtAssignmentInfo" "Text" "Load and manage direct assignments for this configuration profile."

    Add-XamlEvent $root "btnAssignmentRefresh" "Add_Click" ({
        Write-Log "Assignments UI: refresh clicked"
        Load-ConfigProfileAssignments -ForceReloadAssignments
    })

    Add-XamlEvent $root "btnAssignmentSearch" "Add_Click" ({
        $searchText = (Get-XamlProperty $script:ConfigProfileAssignmentTab "txtAssignmentGroupSearch" "Text").Trim()
        Write-Log "Assignments UI: search clicked with text '$searchText'"
        $script:ConfigProfileAssignmentSearchResults.Clear()
        foreach($group in @(Search-ConfigProfileAssignmentGroups $searchText))
        {
            $script:ConfigProfileAssignmentSearchResults.Add($group)
        }
        Write-Log "Assignments UI: search results count $($script:ConfigProfileAssignmentSearchResults.Count)"
    })

    Add-XamlEvent $root "txtAssignmentGroupSearch" "Add_KeyDown" ({
        if($_.Key -ne "Return") { return }
        $searchText = (Get-XamlProperty $script:ConfigProfileAssignmentTab "txtAssignmentGroupSearch" "Text").Trim()
        Write-Log "Assignments UI: search enter pressed with text '$searchText'"
        $script:ConfigProfileAssignmentSearchResults.Clear()
        foreach($group in @(Search-ConfigProfileAssignmentGroups $searchText))
        {
            $script:ConfigProfileAssignmentSearchResults.Add($group)
        }
        Write-Log "Assignments UI: search results count $($script:ConfigProfileAssignmentSearchResults.Count)"
    })

    Add-XamlEvent $root "btnAssignmentAddInclude" "Add_Click" ({
        $group = Get-XamlProperty $script:ConfigProfileAssignmentTab "lstAssignmentGroupResults" "SelectedItem"
        Write-Log "Assignments UI: add include clicked. Selected group '$($group.DisplayName)' / '$($group.Id)'"
        if(-not $group)
        {
            Set-ConfigProfileAssignmentInfo "Select a group from the search results first."
            [System.Windows.MessageBox]::Show("Select a group from the search results first.", "Assignments", "OK", "Information") | Out-Null
            return
        }
        $wasAdded = Add-ConfigProfileAssignmentEntry (New-ConfigProfileAssignmentEntry -TargetType "Group" -TargetMode "Include" -DisplayName $group.DisplayName -GroupId $group.Id)
        if($wasAdded)
        {
            $newEntry = $script:ConfigProfileAssignmentIncludeEntries | Select-Object -Last 1
            Select-ConfigProfileAssignmentEntry "dgProfileAssignmentsInclude" $newEntry
            Invoke-ConfigProfileAssignmentAutoSave "Add include"
        }
    })

    Add-XamlEvent $root "btnAssignmentAddExclude" "Add_Click" ({
        $group = Get-XamlProperty $script:ConfigProfileAssignmentTab "lstAssignmentGroupResults" "SelectedItem"
        Write-Log "Assignments UI: add exclude clicked. Selected group '$($group.DisplayName)' / '$($group.Id)'"
        if(-not $group)
        {
            Set-ConfigProfileAssignmentInfo "Select a group from the search results first."
            [System.Windows.MessageBox]::Show("Select a group from the search results first.", "Assignments", "OK", "Information") | Out-Null
            return
        }
        $wasAdded = Add-ConfigProfileAssignmentEntry (New-ConfigProfileAssignmentEntry -TargetType "Group" -TargetMode "Exclude" -DisplayName $group.DisplayName -GroupId $group.Id)
        if($wasAdded)
        {
            $newEntry = $script:ConfigProfileAssignmentExcludeEntries | Select-Object -Last 1
            Select-ConfigProfileAssignmentEntry "dgProfileAssignmentsExclude" $newEntry
            Invoke-ConfigProfileAssignmentAutoSave "Add exclude"
        }
    })

    Add-XamlEvent $root "btnAssignmentAddAllUsers" "Add_Click" ({
        Write-Log "Assignments UI: add all users clicked"
        $wasAdded = Add-ConfigProfileAssignmentEntry (New-ConfigProfileAssignmentEntry -TargetType "All Users" -TargetMode "Include" -DisplayName "All Users")
        if($wasAdded)
        {
            $newEntry = $script:ConfigProfileAssignmentIncludeEntries | Select-Object -Last 1
            Select-ConfigProfileAssignmentEntry "dgProfileAssignmentsInclude" $newEntry
            Invoke-ConfigProfileAssignmentAutoSave "Add all users"
        }
    })

    Add-XamlEvent $root "btnAssignmentAddAllDevices" "Add_Click" ({
        Write-Log "Assignments UI: add all devices clicked"
        $wasAdded = Add-ConfigProfileAssignmentEntry (New-ConfigProfileAssignmentEntry -TargetType "All Devices" -TargetMode "Include" -DisplayName "All Devices")
        if($wasAdded)
        {
            $newEntry = $script:ConfigProfileAssignmentIncludeEntries | Select-Object -Last 1
            Select-ConfigProfileAssignmentEntry "dgProfileAssignmentsInclude" $newEntry
            Invoke-ConfigProfileAssignmentAutoSave "Add all devices"
        }
    })

    Add-XamlEvent $root "btnAssignmentRemoveInclude" "Add_Click" ({
        $selectedAssignment = Get-XamlProperty $script:ConfigProfileAssignmentTab "dgProfileAssignmentsInclude" "SelectedItem"
        Write-Log "Assignments UI: remove include clicked. Selected '$($selectedAssignment.DisplayName)'"
        if(-not $selectedAssignment) { return }
        if((Remove-ConfigProfileAssignmentEntry $selectedAssignment) -eq $true)
        {
            Invoke-ConfigProfileAssignmentAutoSave "Remove include"
        }
    })

    Add-XamlEvent $root "btnAssignmentRemoveExclude" "Add_Click" ({
        $selectedAssignment = Get-XamlProperty $script:ConfigProfileAssignmentTab "dgProfileAssignmentsExclude" "SelectedItem"
        Write-Log "Assignments UI: remove exclude clicked. Selected '$($selectedAssignment.DisplayName)'"
        if(-not $selectedAssignment) { return }
        if((Remove-ConfigProfileAssignmentEntry $selectedAssignment) -eq $true)
        {
            Invoke-ConfigProfileAssignmentAutoSave "Remove exclude"
        }
    })

    Add-XamlEvent $root "btnAssignmentSave" "Add_Click" ({
        Write-Log "Assignments UI: save clicked with $($script:ConfigProfileAssignmentEntries.Count) entry/entries"
        Save-ConfigProfileAssignments
    })

    Write-Status "Loading assignments"
    Load-ConfigProfileAssignments -ReloadObject
    Write-Status ""
}

function Show-ConfigProfileAssignmentsForm
{
    if(-not (Get-ConfigProfileAssignmentSupported $global:curObjectType)) { return }
    if(-not $global:dgObjects.SelectedItem -or -not $global:dgObjects.SelectedItem.Object) { return }

    $panel = New-ConfigProfileAssignmentPanel
    if(-not $panel) { return }

    Initialize-ConfigProfileAssignmentsForm $panel $global:dgObjects.SelectedItem $global:curObjectType

    $objName = Get-GraphObjectName $global:dgObjects.SelectedItem.Object $global:curObjectType
    $title = "Assignments"
    if($objName) { $title = "$title - $objName" }

    Show-ModalForm $title $panel
}

function Update-ConfigProfileAssignmentsButtonState
{
    if(-not $global:dgObjects) { return }
    if(-not $global:dgObjects.Parent) { return }

    $isSupported = Get-ConfigProfileAssignmentSupported $global:curObjectType
    $hasSelectedItems = ($global:dgObjects.ItemsSource | Where IsSelected -eq $true) -or ($null -ne $global:dgObjects.SelectedItem)

    $button = $global:dgObjects.Parent.FindName("btnAssignments")
    if(-not $button) { return }

    $button.Visibility = if($isSupported) { "Visible" } else { "Collapsed" }
    $button.IsEnabled = $isSupported -and $hasSelectedItems
}

function Invoke-ShowMainWindow
{
    $button = [System.Windows.Controls.Button]::new()
    $button.Content = "Assignments"
    $button.Name = "btnAssignments"
    $button.MinWidth = 100
    $button.Margin = "0,0,5,0"
    $button.IsEnabled = $false
    $button.Visibility = "Collapsed"
    $button.ToolTip = "Manage profile assignments"

    $button.Add_Click({
        Show-ConfigProfileAssignmentsForm
    })

    $global:spSubMenu.RegisterName($button.Name, $button)
    $global:spSubMenu.Children.Insert(1, $button)
    Update-ConfigProfileAssignmentsButtonState
}

function Invoke-EMSelectedItemsChanged
{
    Update-ConfigProfileAssignmentsButtonState
}

function Invoke-ViewActivated
{
    Update-ConfigProfileAssignmentsButtonState
}

function Invoke-AfterGraphObjectDetailsCreated
{
    param($detailsForm, $selectedItem, $objectType)
    return

}
