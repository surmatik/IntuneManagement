function Invoke-InitializeModule
{
    if(-not $global:EMViewObject) { return }

    Add-ViewItem (New-Object PSObject -Property @{
        Title = "Passkey Report"
        Id = "UserPasskeyFidoStatus"
        ViewID = "IntuneGraphAPI"
        API = "/reports/authenticationMethods/userRegistrationDetails"
        QUERYLIST = "`$filter=userType eq 'member'&`$orderby=userPrincipalName"
        NameProperty = "userPrincipalName"
        ViewProperties = @("PasskeyOrFido","userDisplayName","userPrincipalName")
        Permissons = @("AuditLog.Read.All")
        ShowButtons = @("Export","View")
        GroupId = "EndpointAnalytics"
        Icon = "Report"
        LoadAllPages = $true
        ExpandAssignmentsList = $false
        PostListCommand = { Start-PostListTenantPasskeyAnalytics @args }
    })
}

function Set-TenantPasskeyAnalyticsProperty
{
    param(
        $Object,
        [string]$Name,
        $Value
    )

    if(-not $Object -or -not $Name) { return }

    if($Object.PSObject.Properties[$Name])
    {
        $Object.$Name = $Value
    }
    else
    {
        $Object | Add-Member -MemberType NoteProperty -Name $Name -Value $Value
    }
}

function Test-IsTenantPasskeyAnalyticsView
{
    $global:curObjectType -and $global:curObjectType.Id -eq "UserPasskeyFidoStatus"
}

function Get-TenantPasskeyAnalyticsItems
{
    if(-not $global:dgObjects) { return @() }
    if(-not $global:dgObjects.ItemsSource) { return @() }

    @($global:dgObjects.ItemsSource)
}

function Hide-TenantPasskeySummary
{
    if($global:brdPasskeySummary)
    {
        $global:brdPasskeySummary.Visibility = "Collapsed"
    }
}
function Get-TenantPasskeyAnalyticsTenantKey
{
    $orgId = ?? $global:Organization.id ""
    if($orgId) { return "$orgId" }

    $orgName = ?? $global:Organization.displayName ([Environment]::GetEnvironmentVariable("Organization", [System.EnvironmentVariableTarget]::Process)) ""
    "$orgName"
}

function Get-TenantPasskeyAnalyticsLicenseLookup
{
    $tenantKey = Get-TenantPasskeyAnalyticsTenantKey
    if($script:TenantPasskeyAnalyticsLicenseLookup -and $script:TenantPasskeyAnalyticsLicenseLookupTenant -eq $tenantKey)
    {
        return $script:TenantPasskeyAnalyticsLicenseLookup
    }

    $lookup = @{}
    try
    {
        $usersResponse = Invoke-GraphRequest -Url "/users?`$select=id,assignedLicenses&`$filter=userType eq 'Member'" -ODataMetadata "None" -GraphVersion "v1.0" -AllPages
        foreach($user in @($usersResponse.value))
        {
            $userId = "$($user.id)".Trim().ToLowerInvariant()
            if(-not $userId) { continue }

            $isLicensed = @($user.assignedLicenses).Count -gt 0
            $lookup[$userId] = $isLicensed
        }
    }
    catch
    {
        Write-Log "Passkey report: failed to resolve user license data. $($_.Exception.Message)" 2
    }

    $script:TenantPasskeyAnalyticsLicenseLookup = $lookup
    $script:TenantPasskeyAnalyticsLicenseLookupTenant = $tenantKey
    $lookup
}

function Set-TenantPasskeyAnalyticsGridFilter
{
    if(-not (Test-IsTenantPasskeyAnalyticsView)) { return }
    if(-not $global:dgObjects -or -not $global:dgObjects.ItemsSource) { return }
    if($global:dgObjects.ItemsSource -isnot [System.Windows.Data.ListCollectionView]) { return }

    $hideWithoutLicense = ($global:chkPasskeyHideWithoutLicense -and $global:chkPasskeyHideWithoutLicense.IsChecked -eq $true)

    $filterText = ""
    if($global:txtFilter -and $global:txtFilter.Tag -ne "1" -and $global:txtFilter.Text -and $global:txtFilter.Text -ne "Filter")
    {
        $filterText = $global:txtFilter.Text.Trim()
    }

    $global:dgObjects.ItemsSource.Filter = {
        param($item)

        if($hideWithoutLicense -and $item.HasLicense -ne $true)
        {
            return $false
        }

        if([string]::IsNullOrWhiteSpace($filterText))
        {
            return $true
        }

        return ($null -ne ($item.PSObject.Properties | Where-Object { $_.Name -notin @("IsSelected","Object", "ObjectType") -and $_.Value -match [regex]::Escape($filterText) }))
    }

    if($global:txtFilter)
    {
        Invoke-FilterBoxChanged $global:txtFilter -ForceUpdate
    }
}


function ConvertTo-TenantPasskeyAnalyticsMethodTokens
{
    param([AllowNull()][Object]$Value)

    $tokens = @()
    foreach($entry in @($Value))
    {
        $rawEntry = "$entry"
        if(-not $rawEntry) { continue }

        foreach($part in @($rawEntry -split ','))
        {
            $candidate = "$part"
            if(-not $candidate) { continue }

            # Remove zero-width/control characters that can break rendering in WPF/Excel.
            $candidate = $candidate -replace '[\u0000-\u001F\u007F\u200B\u200C\u200D\u2060\uFEFF]', ''
            $candidate = $candidate.Trim()
            if(-not $candidate) { continue }

            # Keep only characters that are meaningful for method keys.
            $candidate = $candidate -replace '[^A-Za-z0-9_-]', ''
            if(-not $candidate) { continue }

            $tokens += $candidate.ToLowerInvariant()
        }
    }

    @($tokens | Where-Object { $_ } | Select-Object -Unique)
}

function Get-TenantPasskeyAnalyticsState
{
    param($Object)

    $registeredMethods = @(ConvertTo-TenantPasskeyAnalyticsMethodTokens $Object.methodsRegistered)
    $systemPreferredMethods = @(ConvertTo-TenantPasskeyAnalyticsMethodTokens $Object.systemPreferredAuthenticationMethods)

    $hasPasskey = @($registeredMethods | Where-Object { $_ -match '(?i)passkey' }).Count -gt 0
    $hasFidoKey = (@($registeredMethods | Where-Object { $_ -match '(?i)fido' }).Count -gt 0) -or (@($systemPreferredMethods | Where-Object { $_ -match '(?i)fido' }).Count -gt 0)
    $hasPasskeyOrFido = $hasPasskey -or $hasFidoKey

    [PSCustomObject]@{
        RegisteredMethods = $registeredMethods
        SystemPreferredMethods = $systemPreferredMethods
        HasPasskey = $hasPasskey
        HasFidoKey = $hasFidoKey
        HasPasskeyOrFido = $hasPasskeyOrFido
        PasskeyOrFidoValue = $(if($hasPasskeyOrFido) { 'Yes' } else { 'No' })
        PasskeyValue = $(if($hasPasskey) { 'Yes' } else { 'No' })
        FidoValue = $(if($hasFidoKey) { 'Yes' } else { 'No' })
    }
}
function Ensure-TenantPasskeyAnalyticsColumns
{
    if(-not (Test-IsTenantPasskeyAnalyticsView)) { return }
    if(-not $global:dgObjects) { return }

    $passkeyColumns = @()
    $columnsToRemove = @()
    foreach($column in @($global:dgObjects.Columns))
    {
        $header = "{0}" -f $column.Header
        switch ($header)
        {
            'PasskeyOrFido'
            {
                $column.Header = 'Passkey or FIDO'
                $passkeyColumns += $column
                continue
            }
            'Passkey or FIDO'
            {
                $passkeyColumns += $column
                continue
            }
            'userDisplayName' { $column.Header = 'Display Name'; continue }
            'userPrincipalName' { $column.Header = 'User Principal Name'; continue }
            'HasLicense'
            {
                # Internal bool used for filtering only.
                $columnsToRemove += $column
                continue
            }
            'License'
            {
                $column.Header = 'Licensed'
                continue
            }
        }
    }

    foreach($removeColumn in $columnsToRemove)
    {
        [void]$global:dgObjects.Columns.Remove($removeColumn)
    }

    if($passkeyColumns.Count -eq 0)
    {
        $column = [System.Windows.Controls.DataGridTextColumn]::new()
        $column.Header = 'Passkey or FIDO'
        $column.IsReadOnly = $true
        $column.Binding = [System.Windows.Data.Binding]::new('PasskeyOrFido')
        $global:dgObjects.Columns.Insert(0, $column)
        return
    }

    for($i = ($passkeyColumns.Count - 1); $i -ge 1; $i--)
    {
        [void]$global:dgObjects.Columns.Remove($passkeyColumns[$i])
    }
}
function Update-TenantPasskeySummary
{
    if(-not (Test-IsTenantPasskeyAnalyticsView))
    {
        Hide-TenantPasskeySummary
        return
    }

    if(-not $global:brdPasskeySummary) { return }

    $items = @(Get-TenantPasskeyAnalyticsItems)
    $totalUsers = $items.Count
    $configuredUsers = @($items | Where-Object { $_.PasskeyOrFido -eq "Yes" }).Count
    $missingUsers = [Math]::Max(0, ($totalUsers - $configuredUsers))
    $coverage = if($totalUsers -gt 0) { [Math]::Round((($configuredUsers / $totalUsers) * 100), 1) } else { 0 }

    $global:txtPasskeySummaryTotalValue.Text = ("{0:N0}" -f $totalUsers)
    $global:txtPasskeySummaryConfiguredValue.Text = ("{0:N0}" -f $configuredUsers)
    $global:txtPasskeySummaryMissingValue.Text = ("{0:N0}" -f $missingUsers)
    $global:txtPasskeySummaryCoverageValue.Text = ("{0:N1}%" -f $coverage)
    $global:txtPasskeySummaryText.Text = "$configuredUsers of $totalUsers member users have Passkey or FIDO configured"
    $global:pbPasskeySummaryCoverage.Value = [double]$coverage
    $global:brdPasskeySummary.Visibility = "Visible"
}

function Register-TenantPasskeyAnalyticsCollectionWatcher
{
    $itemsSource = ?? $global:dgObjects.ItemsSource $null
    if(-not $itemsSource) { return }

    $sourceCollection = $null
    if($itemsSource -is [System.Windows.Data.CollectionView])
    {
        $sourceCollection = $itemsSource.SourceCollection
    }

    if(-not $sourceCollection) { return }
    if($script:TenantPasskeyAnalyticsObservedCollection -eq $sourceCollection) { return }

    $script:TenantPasskeyAnalyticsObservedCollection = $sourceCollection
    if($sourceCollection -is [System.Collections.Specialized.INotifyCollectionChanged])
    {
        $sourceCollection.add_CollectionChanged({ Update-TenantPasskeySummary })
    }
}
function Get-TenantPasskeyAnalyticsExportRows
{
    $items = @(Get-TenantPasskeyAnalyticsItems)
    $rows = @()

    foreach($item in $items)
    {
        if(-not $item.Object) { continue }

        $state = Get-TenantPasskeyAnalyticsState $item.Object
        $registeredMethodsExport = if($state.RegisteredMethods.Count -gt 0) { (($state.RegisteredMethods | ForEach-Object { ConvertTo-TenantPasskeyAnalyticsMethodLabel $_ }) -join ', ') } else { '' }
        $systemPreferredExport = if($state.SystemPreferredMethods.Count -gt 0) { (($state.SystemPreferredMethods | ForEach-Object { ConvertTo-TenantPasskeyAnalyticsMethodLabel $_ }) -join ', ') } else { '' }

        $rows += [PSCustomObject]@{
            'Passkey or FIDO' = $state.PasskeyOrFidoValue
            'Display Name' = (?? $item.userDisplayName $item.Object.userDisplayName)
            'User Principal Name' = (?? $item.userPrincipalName $item.Object.userPrincipalName)
            'Passkey' = $state.PasskeyValue
            'FIDO' = $state.FidoValue
            'Registered Methods' = $registeredMethodsExport
            'System Preferred' = $systemPreferredExport
        }
    }

    @($rows)
}

function ConvertTo-TenantPasskeyAnalyticsExcelColumnName
{
    param([int]$Index)

    $name = ""
    $current = $Index
    while($current -gt 0)
    {
        $remainder = ($current - 1) % 26
        $name = [char](65 + $remainder) + $name
        $current = [int](($current - 1) / 26)
    }

    $name
}

function ConvertTo-TenantPasskeyAnalyticsExcelXml
{
    param([AllowNull()][string]$Value)

    $text = if($null -eq $Value) { "" } else { [string]$Value }
    $escaped = [System.Security.SecurityElement]::Escape($text)
    if($null -eq $escaped) { $escaped = "" }
    $escaped = $escaped.Replace("`r`n", "&#10;").Replace("`r", "&#10;").Replace("`n", "&#10;")
    $escaped
}

function New-TenantPasskeyAnalyticsWorksheetXml
{
    param(
        [string[]]$Headers,
        [object[]]$Rows
    )

    $rowCount = ?? $Rows.Count 0
    $columnCount = [Math]::Max((?? $Headers.Count 0), 1)
    $lastColumnName = ConvertTo-TenantPasskeyAnalyticsExcelColumnName $columnCount
    $lastRowNumber = [Math]::Max(($rowCount + 1), 1)
    $dimension = "A1:$lastColumnName$lastRowNumber"

    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    [void]$sb.Append('<dimension ref="').Append($dimension).Append('" />')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0" /></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15" />')
    [void]$sb.Append('<sheetData>')

    [void]$sb.Append('<row r="1">')
    for($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++)
    {
        $headerValue = ""
        if($columnIndex -lt $Headers.Count) { $headerValue = [string]$Headers[$columnIndex] }

        $columnName = ConvertTo-TenantPasskeyAnalyticsExcelColumnName ($columnIndex + 1)
        $cellRef = "$columnName" + "1"
        [void]$sb.Append('<c r="').Append($cellRef).Append('" s="1" t="inlineStr"><is><t xml:space="preserve">')
        [void]$sb.Append((ConvertTo-TenantPasskeyAnalyticsExcelXml $headerValue))
        [void]$sb.Append('</t></is></c>')
    }
    [void]$sb.Append('</row>')

    $rowNumber = 2
    foreach($row in (?? $Rows @()))
    {
        [void]$sb.Append('<row r="').Append($rowNumber).Append('">')
        for($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++)
        {
            $columnName = ConvertTo-TenantPasskeyAnalyticsExcelColumnName ($columnIndex + 1)
            $cellRef = "$columnName$rowNumber"

            $headerName = ""
            if($columnIndex -lt $Headers.Count) { $headerName = $Headers[$columnIndex] }

            $cellValue = ""
            if($row -and $headerName -and $row.PSObject.Properties[$headerName])
            {
                $cellValue = [string]$row.$headerName
            }

            [void]$sb.Append('<c r="').Append($cellRef).Append('" t="inlineStr"><is><t xml:space="preserve">')
            [void]$sb.Append((ConvertTo-TenantPasskeyAnalyticsExcelXml $cellValue))
            [void]$sb.Append('</t></is></c>')
        }
        [void]$sb.Append('</row>')
        $rowNumber++
    }

    [void]$sb.Append('</sheetData>')
    [void]$sb.Append('</worksheet>')
    $sb.ToString()
}

function Add-TenantPasskeyAnalyticsZipEntry
{
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryPath,
        [string]$Content
    )

    $entry = $Archive.CreateEntry($EntryPath, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = [System.IO.StreamWriter]::new($stream, [System.Text.UTF8Encoding]::new($false))
    try
    {
        $writer.Write($Content)
    }
    finally
    {
        $writer.Dispose()
        $stream.Dispose()
    }
}

function Get-TenantPasskeyAnalyticsDefaultExportFileName
{
    $tenant = ?? $global:Organization.displayName ([Environment]::GetEnvironmentVariable("Organization", [System.EnvironmentVariableTarget]::Process)) ""
    $invalidChars = [Regex]::Escape((-join [System.IO.Path]::GetInvalidFileNameChars()))
    $tenantSafe = ([string]$tenant -replace "[$invalidChars]", "_").Trim()
    if(-not $tenantSafe) { $tenantSafe = "Tenant" }

    $timestamp = (Get-Date).ToString("yyyy-MM-dd_HHmm")
    "PasskeyReport_{0}_{1}.xlsx" -f $tenantSafe, $timestamp
}

function Save-TenantPasskeyAnalyticsExcelWorkbook
{
    param(
        [string]$FilePath,
        [object[]]$Rows,
        [int]$TotalUsers,
        [int]$ConfiguredUsers,
        [int]$MissingUsers,
        [double]$Coverage
    )

    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue

    if(Test-Path -LiteralPath $FilePath)
    {
        Remove-Item -LiteralPath $FilePath -Force
    }

    $summaryRows = @(
        [PSCustomObject]@{ Metric = "Users"; Value = "$TotalUsers" },
        [PSCustomObject]@{ Metric = "Configured"; Value = "$ConfiguredUsers" },
        [PSCustomObject]@{ Metric = "Missing"; Value = "$MissingUsers" },
        [PSCustomObject]@{ Metric = "Coverage %"; Value = ("{0:N1}" -f $Coverage) }
    )

    $summaryXml = New-TenantPasskeyAnalyticsWorksheetXml -Headers @("Metric", "Value") -Rows $summaryRows
    $userHeaders = if($Rows.Count -gt 0) { @($Rows[0].PSObject.Properties.Name) } else { @("Passkey or FIDO", "Display Name", "User Principal Name", "Passkey", "FIDO", "Registered Methods", "System Preferred") }
    $usersXml = New-TenantPasskeyAnalyticsWorksheetXml -Headers $userHeaders -Rows $Rows

    $contentTypesXml = [string]::Join("`n", @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />',
        '  <Default Extension="xml" ContentType="application/xml" />',
        '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />',
        '  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />',
        '  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />',
        '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />',
        '</Types>'
    ))

    $rootRelsXml = [string]::Join("`n", @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />',
        '</Relationships>'
    ))

    $workbookXml = [string]::Join("`n", @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        '  <sheets>',
        '    <sheet name="Summary" sheetId="1" r:id="rId1" />',
        '    <sheet name="Users" sheetId="2" r:id="rId2" />',
        '  </sheets>',
        '</workbook>'
    ))

    $workbookRelsXml = [string]::Join("`n", @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
        '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />',
        '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml" />',
        '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />',
        '</Relationships>'
    ))

    $stylesXml = [string]::Join("`n", @(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '  <fonts count="2">',
        '    <font><sz val="11" /><name val="Calibri" /></font>',
        '    <font><b /><sz val="11" /><name val="Calibri" /></font>',
        '  </fonts>',
        '  <fills count="2">',
        '    <fill><patternFill patternType="none" /></fill>',
        '    <fill><patternFill patternType="gray125" /></fill>',
        '  </fills>',
        '  <borders count="1">',
        '    <border><left /><right /><top /><bottom /><diagonal /></border>',
        '  </borders>',
        '  <cellStyleXfs count="1">',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />',
        '  </cellStyleXfs>',
        '  <cellXfs count="2">',
        '    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />',
        '    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" />',
        '  </cellXfs>',
        '  <cellStyles count="1">',
        '    <cellStyle name="Normal" xfId="0" builtinId="0" />',
        '  </cellStyles>',
        '</styleSheet>'
    ))

    $zip = $null
    try
    {
        $zip = [System.IO.Compression.ZipFile]::Open($FilePath, [System.IO.Compression.ZipArchiveMode]::Create)
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "[Content_Types].xml" -Content $contentTypesXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "_rels/.rels" -Content $rootRelsXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "xl/workbook.xml" -Content $workbookXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "xl/_rels/workbook.xml.rels" -Content $workbookRelsXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "xl/styles.xml" -Content $stylesXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "xl/worksheets/sheet1.xml" -Content $summaryXml
        Add-TenantPasskeyAnalyticsZipEntry -Archive $zip -EntryPath "xl/worksheets/sheet2.xml" -Content $usersXml
    }
    finally
    {
        if($zip) { $zip.Dispose() }
    }
}

function Export-TenantPasskeyAnalyticsExcel
{
    if(-not (Test-IsTenantPasskeyAnalyticsView)) { return $false }

    $rows = @(Get-TenantPasskeyAnalyticsExportRows)
    if($rows.Count -eq 0)
    {
        [System.Windows.MessageBox]::Show('There are no analytics rows to export.', 'Passkey Report', 'OK', 'Information') | Out-Null
        return $true
    }

    $dlgSave = [System.Windows.Forms.SaveFileDialog]::new()
    $dlgSave.InitialDirectory = ?? (Get-Setting '' 'LastUsedRoot') $env:USERPROFILE
    $dlgSave.FileName = Get-TenantPasskeyAnalyticsDefaultExportFileName
    $dlgSave.DefaultExt = 'xlsx'
    $dlgSave.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
    if($dlgSave.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK -or -not $dlgSave.FileName)
    {
        return $true
    }

    try
    {
        $totalUsers = $rows.Count
        $configuredUsers = @($rows | Where-Object { $_.'Passkey or FIDO' -eq 'Yes' }).Count
        $missingUsers = $totalUsers - $configuredUsers
        $coverage = if($totalUsers -gt 0) { [Math]::Round((($configuredUsers / $totalUsers) * 100), 1) } else { 0 }

        Save-TenantPasskeyAnalyticsExcelWorkbook -FilePath $dlgSave.FileName -Rows $rows -TotalUsers $totalUsers -ConfiguredUsers $configuredUsers -MissingUsers $missingUsers -Coverage $coverage

        Save-Setting '' 'LastUsedRoot' ([System.IO.Path]::GetDirectoryName($dlgSave.FileName))
        [System.Windows.MessageBox]::Show("Excel export saved to:`n$($dlgSave.FileName)", 'Passkey Report', 'OK', 'Information') | Out-Null
    }
    catch
    {
        Write-LogError 'Failed to export Passkey Report to Excel' $_.Exception
        [System.Windows.MessageBox]::Show('Excel export failed. Check file path permissions and try again.', 'Passkey Report', 'OK', 'Error') | Out-Null
    }

    $true
}
function New-TenantPasskeyAnalyticsViewTab
{
    $xaml = @"
<TabItem Header="Overview" xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0">
        <StackPanel Margin="12,10,12,12">
            <Border Background="#F7FAFC" BorderBrush="#D7E1EA" BorderThickness="1" CornerRadius="10" Padding="14" Margin="0,0,0,12">
                <StackPanel>
                    <TextBlock Name="txtPasskeyDetailTitle" Text="Passkey and FIDO overview" FontSize="18" FontWeight="Bold" Foreground="#102A43" />
                    <TextBlock Name="txtPasskeyDetailSubtitle" Text="User authentication summary" Margin="0,4,0,0" Foreground="#52606D" TextWrapping="Wrap" />
                </StackPanel>
            </Border>

            <Grid Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto" />
                    <ColumnDefinition Width="*" />
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                    <RowDefinition Height="Auto" />
                </Grid.RowDefinitions>

                <TextBlock Text="Display Name" FontWeight="SemiBold" Foreground="#334E68" Margin="0,0,16,8" />
                <TextBlock Name="txtPasskeyDetailDisplayName" Grid.Column="1" Foreground="#102A43" Margin="0,0,0,8" TextWrapping="Wrap" />

                <TextBlock Text="User Principal Name" Grid.Row="1" FontWeight="SemiBold" Foreground="#334E68" Margin="0,0,16,8" />
                <TextBlock Name="txtPasskeyDetailUserPrincipalName" Grid.Row="1" Grid.Column="1" Foreground="#102A43" Margin="0,0,0,8" TextWrapping="Wrap" />

                <TextBlock Text="Registered Methods" Grid.Row="2" FontWeight="SemiBold" Foreground="#334E68" Margin="0,0,16,8" VerticalAlignment="Top" />
                <StackPanel Grid.Row="2" Grid.Column="1" Margin="0,0,0,8">
                    <WrapPanel Name="pnlPasskeyDetailRegisteredMethods" />
                    <TextBlock Name="txtPasskeyDetailRegisteredMethodsEmpty" Text="No registered methods reported" Foreground="#52606D" Visibility="Collapsed" />
                </StackPanel>

                <TextBlock Text="System Preferred" Grid.Row="3" FontWeight="SemiBold" Foreground="#334E68" Margin="0,0,16,0" VerticalAlignment="Top" />
                <StackPanel Grid.Row="3" Grid.Column="1">
                    <WrapPanel Name="pnlPasskeyDetailSystemPreferred" />
                    <TextBlock Name="txtPasskeyDetailSystemPreferredEmpty" Text="No system-preferred methods reported" Foreground="#52606D" Visibility="Collapsed" />
                </StackPanel>
            </Grid>

            <Border Background="#FFFFFF" BorderBrush="#D7E1EA" BorderThickness="1" CornerRadius="8" Padding="12">
                <StackPanel>
                    <TextBlock Text="Interpretation" FontWeight="SemiBold" Foreground="#334E68" Margin="0,0,0,6" />
                    <TextBlock Name="txtPasskeyDetailInterpretation" TextWrapping="Wrap" Foreground="#102A43" />
                </StackPanel>
            </Border>
        </StackPanel>
    </ScrollViewer>
</TabItem>
"@

    [Windows.Markup.XamlReader]::Parse($xaml)
}

function ConvertTo-TenantPasskeyAnalyticsMethodLabel
{
    param([string]$Method)

    if(-not $Method) { return '' }

    $raw = "$Method".Trim()
    if(-not $raw) { return '' }

    $compact = ($raw -replace '\s+', '')
    $compact = $compact -replace '[\u0000-\u001F\u007F\u200B\u200C\u200D\u2060\uFEFF]', ''
    $compact = $compact -replace '[^A-Za-z0-9_-]', ''
    $key = $compact.ToLowerInvariant()

    switch($key)
    {
        'passkeydeviceboundauthenticator' { return 'Passkey Device-Bound' }
        'windowshelloforbusiness' { return 'Windows Hello for Business' }
        'microsoftauthenticatorpush' { return 'Microsoft Authenticator Push' }
        'microsoftauthenticatorpasswordless' { return 'Microsoft Authenticator Passwordless' }
        'softwareonetimepasscode' { return 'Software OTP' }
        'phoneappnotification' { return 'Phone App Notification' }
        'mobilephone' { return 'Mobile Phone' }
        'email' { return 'Email' }
        'sms' { return 'SMS' }
        'voice' { return 'Voice' }
        'fido2' { return 'FIDO2' }
        default
        {
            # fallback: split by separators and title-case words, avoid per-letter spacing artifacts
            $label = $key -replace '[_-]+', ' '
            $ti = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
            return $ti.ToTitleCase($label).Trim()
        }
    }
}
function Add-TenantPasskeyAnalyticsMethodBadges
{
    param(
        $TabItem,
        [string]$PanelName,
        [string]$EmptyTextName,
        [Array]$Methods
    )

    if(-not $TabItem) { return }

    $panel = $TabItem.FindName($PanelName)
    $emptyText = $TabItem.FindName($EmptyTextName)
    if(-not $panel -or -not $emptyText) { return }

    $panel.Children.Clear()
    $normalizedMethods = @(
        @($Methods) |
            Where-Object { $null -ne $_ -and "$_".Trim() -ne "" } |
            ForEach-Object { ConvertTo-TenantPasskeyAnalyticsMethodLabel "$_" }
    )

    if($normalizedMethods.Count -eq 0)
    {
        $emptyText.Visibility = 'Visible'
        return
    }

    $emptyText.Visibility = 'Collapsed'
    foreach($method in $normalizedMethods)
    {
        $border = [System.Windows.Controls.Border]::new()
        $border.Margin = '0,0,8,8'
        $border.Padding = '10,5,10,5'
        $border.CornerRadius = [System.Windows.CornerRadius]::new(14)
        $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#EEF2F6')
        $border.BorderBrush = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#D7E1EA')
        $border.BorderThickness = [System.Windows.Thickness]::new(1)

        $textBlock = [System.Windows.Controls.TextBlock]::new()
        $textBlock.Text = $method
        $textBlock.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#243B53')
        $textBlock.FontWeight = 'SemiBold'

        $border.Child = $textBlock
        [void]$panel.Children.Add($border)
    }
}

function Set-TenantPasskeyAnalyticsViewTab
{
    param($tabItem, $selectedItem)

    if(-not $tabItem -or -not $selectedItem -or -not $selectedItem.Object) { return }

    $obj = $selectedItem.Object
    $state = Get-TenantPasskeyAnalyticsState $obj

    Set-XamlProperty $tabItem 'txtPasskeyDetailTitle' 'Text' (?? $obj.userDisplayName 'User overview')
    Set-XamlProperty $tabItem 'txtPasskeyDetailSubtitle' 'Text' ('Passkey and FIDO registration for ' + (?? $obj.userPrincipalName 'this user'))
    Set-XamlProperty $tabItem 'txtPasskeyDetailDisplayName' 'Text' (?? $obj.userDisplayName '')
    Set-XamlProperty $tabItem 'txtPasskeyDetailUserPrincipalName' 'Text' (?? $obj.userPrincipalName '')
    Add-TenantPasskeyAnalyticsMethodBadges -TabItem $tabItem -PanelName 'pnlPasskeyDetailRegisteredMethods' -EmptyTextName 'txtPasskeyDetailRegisteredMethodsEmpty' -Methods $state.RegisteredMethods
    Add-TenantPasskeyAnalyticsMethodBadges -TabItem $tabItem -PanelName 'pnlPasskeyDetailSystemPreferred' -EmptyTextName 'txtPasskeyDetailSystemPreferredEmpty' -Methods $state.SystemPreferredMethods

    $interpretation = if($state.HasPasskeyOrFido)
    {
        'This user has at least one Passkey or FIDO-related authentication method configured.'
    }
    else
    {
        'This user currently has no Passkey or FIDO method visible in the report.'
    }
    Set-XamlProperty $tabItem 'txtPasskeyDetailInterpretation' 'Text' $interpretation
}

function Start-PostListTenantPasskeyAnalytics
{
    param($objList, $objectType)

    $licenseLookup = Get-TenantPasskeyAnalyticsLicenseLookup

    foreach($obj in @($objList))
    {
        $state = Get-TenantPasskeyAnalyticsState $obj.Object
        Set-TenantPasskeyAnalyticsProperty $obj 'PasskeyOrFido' $state.PasskeyOrFidoValue

        $objectId = "$(?? $obj.Object.id $obj.id)".Trim().ToLowerInvariant()
        $hasLicense = $true
        if($objectId -and $licenseLookup.ContainsKey($objectId))
        {
            $hasLicense = [bool]$licenseLookup[$objectId]
        }
        elseif($obj.Object -and $obj.Object.PSObject.Properties['isLicensed'])
        {
            $hasLicense = [bool]$obj.Object.isLicensed
        }

        Set-TenantPasskeyAnalyticsProperty $obj 'HasLicense' $hasLicense
        Set-TenantPasskeyAnalyticsProperty $obj 'License' $(if($hasLicense) { 'Yes' } else { 'No' })
    }

    @($objList | Sort-Object -Property userPrincipalName)
}

function Invoke-GraphObjectsChanged
{
    Ensure-TenantPasskeyAnalyticsColumns
    Register-TenantPasskeyAnalyticsCollectionWatcher
    Set-TenantPasskeyAnalyticsGridFilter
    Update-TenantPasskeySummary
}

function Invoke-AfterMainWindowCreated
{
    if(-not $global:dgObjects) { return }
    if($script:TenantPasskeyAnalyticsItemsSourceHooked -eq $true) { return }

    $dpd = [System.ComponentModel.DependencyPropertyDescriptor]::FromProperty([System.Windows.Controls.ItemsControl]::ItemsSourceProperty, [System.Windows.Controls.DataGrid])
    if($dpd)
    {
        $dpd.AddValueChanged($global:dgObjects, {
            Ensure-TenantPasskeyAnalyticsColumns
            Register-TenantPasskeyAnalyticsCollectionWatcher
            Set-TenantPasskeyAnalyticsGridFilter
            Update-TenantPasskeySummary
        })
        $script:TenantPasskeyAnalyticsItemsSourceHooked = $true
    }

    if($global:btnExport -and $script:TenantPasskeyAnalyticsExportHooked -ne $true)
    {
        $global:btnExport.Add_PreviewMouseLeftButtonDown({
            if(Test-IsTenantPasskeyAnalyticsView)
            {
                $ret = Export-TenantPasskeyAnalyticsExcel
                if($ret -eq $true)
                {
                    $_.Handled = $true
                }
            }
        })
        $script:TenantPasskeyAnalyticsExportHooked = $true
    }

    if($global:chkPasskeyHideWithoutLicense -and $script:TenantPasskeyAnalyticsLicenseFilterHooked -ne $true)
    {
        $global:chkPasskeyHideWithoutLicense.Add_Checked({
            Set-TenantPasskeyAnalyticsGridFilter
            Update-TenantPasskeySummary
        })
        $global:chkPasskeyHideWithoutLicense.Add_Unchecked({
            Set-TenantPasskeyAnalyticsGridFilter
            Update-TenantPasskeySummary
        })
        $script:TenantPasskeyAnalyticsLicenseFilterHooked = $true
    }

    if($global:txtFilter -and $script:TenantPasskeyAnalyticsTextFilterHooked -ne $true)
    {
        $global:txtFilter.Add_TextChanged({
            if(Test-IsTenantPasskeyAnalyticsView)
            {
                Set-TenantPasskeyAnalyticsGridFilter
                Update-TenantPasskeySummary
            }
        })
        $script:TenantPasskeyAnalyticsTextFilterHooked = $true
    }

    Hide-TenantPasskeySummary
}
function Invoke-AfterGraphObjectDetailsCreated
{
    param($detailsForm, $selectedItem, $objectType)

    if(-not $objectType -or $objectType.Id -ne 'UserPasskeyFidoStatus')
    {
        return
    }

    try
    {
        $existingTab = @($detailsForm.Items | Where-Object { $_.Header -eq 'Overview' })[0]
        if(-not $existingTab)
        {
            $existingTab = New-TenantPasskeyAnalyticsViewTab
            [void]$detailsForm.Items.Insert(0, $existingTab)
        }

        Set-TenantPasskeyAnalyticsViewTab -tabItem $existingTab -selectedItem $selectedItem
        $detailsForm.SelectedIndex = 0
    }
    catch
    {
        Write-Log "Passkey report detail view failed to load. $($_.Exception.Message)" 3
    }
}







