function Get-MenuGroupTitle
{
    param($viewItem)

    if(-not $viewItem) { return "Other" }

    switch ($viewItem.GroupId)
    {
        "Apps" { return "Apps" }
        "AppConfiguration" { return "Apps" }
        "AppProtection" { return "Apps" }
        "CompliancePolicies" { return "Compliance" }
        "ConditionalAccess" { return "Identity" }
        "EndpointAnalytics" { return "Analytics" }
        "EndpointSecurity" { return "Security" }
        "EnrollmentRestrictions" { return "Enrollment" }
        "WinEnrollment" { return "Enrollment" }
        "AppleEnrollment" { return "Enrollment" }
        "DeviceConfiguration" { return "Configuration" }
        "CustomAttributes" { return "Configuration" }
        "PolicySets" { return "Configuration" }
        "Scripts" { return "Scripts" }
        "TenantAdmin" { return "Tenant" }
        "Azure" { return "Tenant" }
        "WinFeatureUpdates" { return "Updates" }
        "WinQualityUpdates" { return "Updates" }
        "WinUpdatePolicies" { return "Updates" }
        "WinDriverUpdatePolicies" { return "Updates" }
        default
        {
            if($viewItem.GroupId)
            {
                return $viewItem.GroupId
            }

            return "Other"
        }
    }
}

function Set-MenuGroupTitle
{
    param($viewItem)

    if(-not $viewItem) { return }

    $groupTitle = Get-MenuGroupTitle $viewItem
    if(($viewItem | Get-Member -MemberType NoteProperty -Name "MenuGroupTitle"))
    {
        $viewItem.MenuGroupTitle = $groupTitle
    }
    else
    {
        $viewItem | Add-Member -NotePropertyName "MenuGroupTitle" -NotePropertyValue $groupTitle
    }
}

function Set-MenuGroupTitles
{
    param($viewItems)

    foreach($viewItem in @($viewItems))
    {
        Set-MenuGroupTitle $viewItem
    }
}

function Set-MenuGrouping
{
    if(-not $global:lstMenuItems) { return }
    if(-not $global:lstMenuItems.ItemsSource) { return }

    $menuView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($global:lstMenuItems.ItemsSource)
    if(-not $menuView) { return }
    if($menuView -isnot [System.Windows.Data.ListCollectionView]) { return }

    $menuView.SortDescriptions.Clear()
    [void]$menuView.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('MenuGroupTitle', [System.ComponentModel.ListSortDirection]::Ascending))
    [void]$menuView.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('Title', [System.ComponentModel.ListSortDirection]::Ascending))

    $menuView.GroupDescriptions.Clear()
    [void]$menuView.GroupDescriptions.Add([System.Windows.Data.PropertyGroupDescription]::new("MenuGroupTitle"))
    $menuView.Refresh()
}

function Invoke-UpdateViewMenuItems
{
    if(-not $global:ViewMenuItems) { return }

    $viewItems = @($global:ViewMenuItems)
    Set-MenuGroupTitles $viewItems
    $global:ViewMenuItems = @($viewItems)
}

function Invoke-AfterShowViewMenu
{
    Set-MenuGrouping
}

function Invoke-AfterMainWindowCreated
{
    Set-MenuGrouping
}
