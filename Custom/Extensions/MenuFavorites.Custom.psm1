function Get-MenuFavoriteEntries
{
    $favoriteValue = Get-Setting "General" "MenuFavorites" ""
    if(-not $favoriteValue) { return @() }

    @($favoriteValue.ToString().Split('|') | Where-Object { $_.Trim() -ne "" })
}

function Get-MenuFavoriteKey
{
    param($viewItem, $viewId = $global:currentViewObject.ViewInfo.Id)

    if(-not $viewItem -or -not $viewItem.Id -or -not $viewId) { return $null }

    "$viewId`:$($viewItem.Id)"
}

function Set-MenuItemFavoriteState
{
    param($viewItem, [bool]$isFavorite)

    if(($viewItem | Get-Member -MemberType NoteProperty -Name "IsFavorite"))
    {
        $viewItem.IsFavorite = $isFavorite
    }
    else
    {
        $viewItem | Add-Member -NotePropertyName "IsFavorite" -NotePropertyValue $isFavorite
    }
}

function Set-MenuItemsFavoriteState
{
    param($viewItems)

    $favoriteEntries = @(Get-MenuFavoriteEntries)

    foreach($viewItem in @($viewItems))
    {
        $favoriteKey = Get-MenuFavoriteKey $viewItem
        Set-MenuItemFavoriteState $viewItem ($favoriteEntries -contains $favoriteKey)
    }
}

function Save-MenuFavoriteState
{
    param($viewItem, [bool]$isFavorite)

    $favoriteKey = Get-MenuFavoriteKey $viewItem
    if(-not $favoriteKey) { return }

    $favoriteEntries = [System.Collections.Generic.List[string]]::new()
    foreach($entry in @(Get-MenuFavoriteEntries))
    {
        if($entry -and $favoriteEntries.Contains($entry) -eq $false)
        {
            $favoriteEntries.Add($entry)
        }
    }

    while($favoriteEntries.Contains($favoriteKey))
    {
        $favoriteEntries.Remove($favoriteKey) | Out-Null
    }

    if($isFavorite)
    {
        $favoriteEntries.Add($favoriteKey)
    }

    Save-Setting "General" "MenuFavorites" (($favoriteEntries | Sort-Object -Unique) -join "|")
}

function Update-MenuFavoriteMenuState
{
    $selectedItem = Get-SelectedMenuMenuItem
    $isEnabled = $null -ne $selectedItem
    $header = if($selectedItem -and $selectedItem.IsFavorite -eq $true) { "Remove from favorites" } else { "Add to favorites" }

    if($global:mnuToggleFavorite)
    {
        $global:mnuToggleFavorite.IsEnabled = $isEnabled
        $global:mnuToggleFavorite.Header = $header
    }
    if($global:mnuToggleFavoriteSecondary)
    {
        $global:mnuToggleFavoriteSecondary.IsEnabled = $isEnabled
        $global:mnuToggleFavoriteSecondary.Header = $header
    }
}

function Get-SelectedMenuMenuItem
{
    if($global:lstFavoriteMenuItems -and $global:lstFavoriteMenuItems.SelectedItem)
    {
        return $global:lstFavoriteMenuItems.SelectedItem
    }

    if($global:lstMenuItems)
    {
        return $global:lstMenuItems.SelectedItem
    }

    $null
}

function Set-SelectedMenuMenuItem
{
    param($selectedItem, [string]$Source = "")

    if($script:MenuFavoriteSelectionSync -eq $true) { return }
    $script:MenuFavoriteSelectionSync = $true

    try
    {
        if($global:lstFavoriteMenuItems)
        {
            if($selectedItem -and $selectedItem.IsFavorite -eq $true)
            {
                $global:lstFavoriteMenuItems.SelectedItem = $selectedItem
            }
            elseif($Source -ne "Favorites")
            {
                $global:lstFavoriteMenuItems.SelectedItem = $null
            }
        }

        if($global:lstMenuItems)
        {
            $global:lstMenuItems.SelectedItem = $selectedItem
        }
    }
    finally
    {
        $script:MenuFavoriteSelectionSync = $false
    }
}

function Refresh-FavoriteMenuItems
{
    if(-not $global:lstFavoriteMenuItems) { return }

    $currentSelection = Get-SelectedMenuMenuItem
    $favoriteItems = @($global:ViewMenuItems | Where-Object IsFavorite -eq $true)
    foreach($favoriteItem in $favoriteItems)
    {
        $favoriteIcon = $null
        if($favoriteItem.Icon -or [IO.File]::Exists(($global:AppRootFolder + "\Xaml\Icons\$($favoriteItem.Id).xaml")))
        {
            $favoriteIcon = Get-XamlObject ($global:AppRootFolder + "\Xaml\Icons\$((?? $favoriteItem.Icon $favoriteItem.Id)).xaml")
        }

        if(($favoriteItem | Get-Member -MemberType NoteProperty -Name "FavoriteIconImage"))
        {
            $favoriteItem.FavoriteIconImage = $favoriteIcon
        }
        else
        {
            $favoriteItem | Add-Member -NotePropertyName "FavoriteIconImage" -NotePropertyValue $favoriteIcon
        }
    }
    $global:lstFavoriteMenuItems.ItemsSource = $favoriteItems

    $hasFavorites = $favoriteItems.Count -gt 0
    if($global:txtFavoritesHeader)
    {
        $global:txtFavoritesHeader.Visibility = if($hasFavorites) { "Visible" } else { "Collapsed" }
    }
    $global:lstFavoriteMenuItems.Visibility = if($hasFavorites) { "Visible" } else { "Collapsed" }

    if($currentSelection -and $currentSelection.IsFavorite -eq $true)
    {
        $updatedFavoriteSelection = $favoriteItems | Where-Object Id -eq $currentSelection.Id | Select-Object -First 1
        if($updatedFavoriteSelection)
        {
            Set-SelectedMenuMenuItem $updatedFavoriteSelection
        }
    }
}

function Toggle-SelectedMenuFavorite
{
    $selectedItem = Get-SelectedMenuMenuItem
    if(-not $selectedItem) { return }

    $selectedItemId = $selectedItem.Id
    $newFavoriteState = -not ($selectedItem.IsFavorite -eq $true)

    Save-MenuFavoriteState $selectedItem $newFavoriteState
    Set-MenuItemFavoriteState $selectedItem $newFavoriteState

    Show-ViewMenu

    $menuItems = @($global:ViewMenuItems)
    $updatedSelection = $menuItems | Where-Object Id -eq $selectedItemId | Select-Object -First 1
    if($updatedSelection)
    {
        Set-SelectedMenuMenuItem $updatedSelection
    }

    Update-MenuFavoriteMenuState
}

function Invoke-UpdateViewMenuItems
{
    if(-not $global:ViewMenuItems) { return }

    $viewItems = @($global:ViewMenuItems)
    Set-MenuItemsFavoriteState $viewItems
    $global:ViewMenuItems = @($viewItems)
}

function Invoke-AfterShowViewMenu
{
    Refresh-FavoriteMenuItems
    Update-MenuFavoriteMenuState
}

function Invoke-AfterMainWindowCreated
{
    if(-not $global:lstMenuItems) { return }

    if(-not $global:mnuMenuItems)
    {
        $global:mnuMenuItems = $global:lstMenuItems.ContextMenu
    }
    if(-not $global:mnuFavoriteMenuItems -and $global:lstFavoriteMenuItems)
    {
        $global:mnuFavoriteMenuItems = $global:lstFavoriteMenuItems.ContextMenu
    }
    if(-not $global:mnuToggleFavorite -and $global:mnuMenuItems)
    {
        $global:mnuToggleFavorite = $global:mnuMenuItems.Items | Where-Object Name -eq "mnuToggleFavorite" | Select-Object -First 1
    }
    if(-not $global:mnuToggleFavoriteSecondary -and $global:mnuFavoriteMenuItems)
    {
        $global:mnuToggleFavoriteSecondary = $global:mnuFavoriteMenuItems.Items | Where-Object Name -eq "mnuToggleFavoriteSecondary" | Select-Object -First 1
    }

    $global:lstMenuItems.Add_PreviewMouseRightButtonDown({
        param($sender, $e)

        $depObj = $e.OriginalSource
        while($depObj -and -not ($depObj -is [System.Windows.Controls.ListBoxItem]))
        {
            $depObj = [System.Windows.Media.VisualTreeHelper]::GetParent($depObj)
        }

        if($depObj -is [System.Windows.Controls.ListBoxItem])
        {
            $depObj.IsSelected = $true
            $depObj.Focus()
        }
    })

    if($global:lstFavoriteMenuItems)
    {
        $global:lstFavoriteMenuItems.Add_PreviewMouseRightButtonDown({
            param($sender, $e)

            $depObj = $e.OriginalSource
            while($depObj -and -not ($depObj -is [System.Windows.Controls.ListBoxItem]))
            {
                $depObj = [System.Windows.Media.VisualTreeHelper]::GetParent($depObj)
            }

            if($depObj -is [System.Windows.Controls.ListBoxItem])
            {
                $depObj.IsSelected = $true
                $depObj.Focus()
            }
        })
    }

    if($global:mnuMenuItems)
    {
        $global:mnuMenuItems.Add_Opened({
            Update-MenuFavoriteMenuState
        })
    }
    if($global:mnuFavoriteMenuItems)
    {
        $global:mnuFavoriteMenuItems.Add_Opened({
            Update-MenuFavoriteMenuState
        })
    }

    if($global:mnuToggleFavorite)
    {
        $global:mnuToggleFavorite.Add_Click({
            Toggle-SelectedMenuFavorite
        })
    }
    if($global:mnuToggleFavoriteSecondary)
    {
        $global:mnuToggleFavoriteSecondary.Add_Click({
            Toggle-SelectedMenuFavorite
        })
    }

    $global:lstMenuItems.Add_SelectionChanged({
        if($script:MenuFavoriteSelectionSync -eq $true) { return }
        if(-not $global:lstMenuItems.SelectedItem) { return }
        if($global:lstMenuItems.SelectedItem.IsFavorite -ne $true -and $global:lstFavoriteMenuItems)
        {
            $script:MenuFavoriteSelectionSync = $true
            try
            {
                $global:lstFavoriteMenuItems.SelectedItem = $null
            }
            finally
            {
                $script:MenuFavoriteSelectionSync = $false
            }
        }
        Update-MenuFavoriteMenuState
    })

    if($global:lstFavoriteMenuItems)
    {
        $global:lstFavoriteMenuItems.Add_SelectionChanged({
            if($script:MenuFavoriteSelectionSync -eq $true) { return }
            if(-not $global:lstFavoriteMenuItems.SelectedItem) { return }

            Set-SelectedMenuMenuItem $global:lstFavoriteMenuItems.SelectedItem "Favorites"
            Update-MenuFavoriteMenuState
        })
    }
}
