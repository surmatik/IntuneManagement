function Invoke-MenuFilterBoxChanged
{
    param($txtBox)

    if(-not $txtBox -or -not $global:lstMenuItems) { return }

    $menuView = [System.Windows.Data.CollectionViewSource]::GetDefaultView($global:lstMenuItems.ItemsSource)
    if(-not $menuView) { return }

    $filterText = ""

    if($txtBox.Text.Trim() -eq "" -and $txtBox.IsFocused -eq $false)
    {
        $txtBox.FontStyle = "Italic"
        $txtBox.Tag = 1
        $txtBox.Text = "Search"
        $txtBox.Foreground = "LightGray"
    }
    elseif($txtBox.Tag -eq "1" -and $txtBox.Text -eq "Search" -and $txtBox.IsFocused -eq $false)
    {
        $filterText = ""
    }
    else
    {
        $txtBox.FontStyle = "Normal"
        $txtBox.Tag = $null
        $txtBox.Foreground = "Black"
        $txtBox.Background = "White"
        $filterText = $txtBox.Text.Trim()
    }

    if([string]::IsNullOrWhiteSpace($filterText))
    {
        $menuView.Filter = $null
    }
    else
    {
        $escapedFilterText = [regex]::Escape($filterText)
        $menuView.Filter = {
            param($item)

            if(-not $item) { return $false }

            return ($item.Title -match $escapedFilterText) -or ($item.Id -match $escapedFilterText)
        }
    }

    $menuView.Refresh()

    if($global:lstFavoriteMenuItems -and $global:lstFavoriteMenuItems.SelectedItem)
    {
        return
    }

    if($global:lstMenuItems.SelectedItem -and $menuView.Contains($global:lstMenuItems.SelectedItem))
    {
        return
    }

    $firstVisibleItem = $null
    foreach($menuItem in $menuView)
    {
        $firstVisibleItem = $menuItem
        break
    }

    $global:lstMenuItems.SelectedItem = $firstVisibleItem
}

function Invoke-AfterShowViewMenu
{
    if($global:txtMenuFilter)
    {
        Invoke-MenuFilterBoxChanged $global:txtMenuFilter
    }
}

function Invoke-AfterMainWindowCreated
{
    if(-not $global:txtMenuFilter -or -not $global:lstMenuItems) { return }

    Add-XamlEvent $global:window "txtMenuFilter" "Add_GotFocus" ({
        if($this.Tag -eq "1" -and $this.Text -eq "Search")
        {
            $this.Text = ""
        }
        Invoke-MenuFilterBoxChanged $this
    })

    Add-XamlEvent $global:window "txtMenuFilter" "Add_LostFocus" ({
        Invoke-MenuFilterBoxChanged $this
    })

    Add-XamlEvent $global:window "txtMenuFilter" "Add_TextChanged" ({
        Invoke-MenuFilterBoxChanged $this
    })

    Invoke-MenuFilterBoxChanged $global:txtMenuFilter
}
