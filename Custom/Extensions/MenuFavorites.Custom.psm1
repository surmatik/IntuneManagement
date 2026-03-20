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

    $menuGroupTitle = if($isFavorite) { "Favoriten" } else { "Alle Menupunkte" }

    if(($viewItem | Get-Member -MemberType NoteProperty -Name "IsFavorite"))
    {
        $viewItem.IsFavorite = $isFavorite
    }
    else
    {
        $viewItem | Add-Member -NotePropertyName "IsFavorite" -NotePropertyValue $isFavorite
    }

    if(($viewItem | Get-Member -MemberType NoteProperty -Name "MenuGroupTitle"))
    {
        $viewItem.MenuGroupTitle = $menuGroupTitle
    }
    else
    {
        $viewItem | Add-Member -NotePropertyName "MenuGroupTitle" -NotePropertyValue $menuGroupTitle
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
    if(-not $global:mnuToggleFavorite) { return }

    $selectedItem = $global:lstMenuItems.SelectedItem
    $global:mnuToggleFavorite.IsEnabled = $null -ne $selectedItem

    if($selectedItem -and $selectedItem.IsFavorite -eq $true)
    {
        $global:mnuToggleFavorite.Header = "Remove from favorites"
    }
    else
    {
        $global:mnuToggleFavorite.Header = "Add to favorites"
    }
}

function Toggle-SelectedMenuFavorite
{
    $selectedItem = $global:lstMenuItems.SelectedItem
    if(-not $selectedItem) { return }

    $selectedItemId = $selectedItem.Id
    $newFavoriteState = -not ($selectedItem.IsFavorite -eq $true)

    Save-MenuFavoriteState $selectedItem $newFavoriteState
    Set-MenuItemFavoriteState $selectedItem $newFavoriteState

    Show-ViewMenu

    $menuItems = @($global:lstMenuItems.ItemsSource)
    $updatedSelection = $menuItems | Where-Object Id -eq $selectedItemId | Select-Object -First 1
    if($updatedSelection)
    {
        $global:lstMenuItems.SelectedItem = $updatedSelection
    }

    Update-MenuFavoriteMenuState
}

function Invoke-UpdateViewMenuItems
{
    if(-not $global:ViewMenuItems) { return }

    $viewItems = @($global:ViewMenuItems)
    Set-MenuItemsFavoriteState $viewItems

    $favoriteItems = @($viewItems | Where-Object IsFavorite -eq $true)
    $otherItems = @($viewItems | Where-Object IsFavorite -ne $true)
    $global:ViewMenuItems = @($favoriteItems + $otherItems)
}

function Invoke-AfterShowViewMenu
{
    $menuView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($global:lstMenuItems.ItemsSource)
    if($menuView -and $menuView -is [System.Windows.Data.ListCollectionView])
    {
        $menuView.GroupDescriptions.Clear()
        $menuView.GroupDescriptions.Add([System.Windows.Data.PropertyGroupDescription]::new("MenuGroupTitle"))
    }

    Update-MenuFavoriteMenuState
}

function Invoke-AfterMainWindowCreated
{
    if(-not $global:lstMenuItems) { return }

    if(-not $global:mnuMenuItems)
    {
        $global:mnuMenuItems = $global:lstMenuItems.ContextMenu
    }
    if(-not $global:mnuToggleFavorite -and $global:mnuMenuItems)
    {
        $global:mnuToggleFavorite = $global:mnuMenuItems.Items | Where-Object Name -eq "mnuToggleFavorite" | Select-Object -First 1
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

    if($global:mnuMenuItems)
    {
        $global:mnuMenuItems.Add_Opened({
            Update-MenuFavoriteMenuState
        })
    }

    if($global:mnuToggleFavorite)
    {
        $global:mnuToggleFavorite.Add_Click({
            Toggle-SelectedMenuFavorite
        })
    }

    $global:lstMenuItems.Add_SelectionChanged({
        Update-MenuFavoriteMenuState
    })
}
