Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

[System.Windows.Forms.Application]::EnableVisualStyles()

function Normalize-Newlines([string]$Text) {
    if ($null -eq $Text) { return "" }
    $t = $Text
    $t = $t.Replace([char]0x2028, "`n") # Unicode line separator
    $t = $t.Replace([char]0x2029, "`n") # Unicode paragraph separator
    $t = $t.Replace([char]0x0085, "`n") # NEL
    $t = $t -replace "`r`n", "`n"
    $t = $t -replace "`r", "`n"
    return $t
}

function Escape-MarkdownCell([string]$Cell) {
    if ($null -eq $Cell) { return "" }
    $v = $Cell.Trim()
    $v = $v -replace '\|', '\|'
    $v = $v -replace '\r?\n', '<br>'
    return $v
}

function Is-NumericCell([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $true }
    $v = $Value.Trim()
    return $v -match '^[\s~≈+\-]?\d[\d\s.,]*(?:\s*[%₽$€])?$'
}

function Convert-TsvToMarkdown([string]$Text) {
    $Text = Normalize-Newlines $Text
    $lines = @($Text -split "`n" | Where-Object { $_ -ne "" })
    if ($lines.Count -eq 0) { throw "The clipboard is empty." }

    $rawRows = @()
    $maxRawCols = 0

    foreach ($line in $lines) {
        $cells = @($line -split "`t", -1)
        if ($cells.Count -gt $maxRawCols) { $maxRawCols = $cells.Count }
        $rawRows += ,$cells
    }

    if ($maxRawCols -lt 2) {
        throw "No tabular data found. Copy a range of cells from Excel."
    }

    # Excel ranges with merged cells often contain completely empty
    # physical TSV columns between real data columns. These columns
    # are meaningless in Markdown, so remove columns that are empty in every row.
    # Important: individual empty cells inside otherwise non-empty columns are preserved.
    $usedColumns = New-Object System.Collections.Generic.List[int]

    for ($c = 0; $c -lt $maxRawCols; $c++) {
        $hasValue = $false

        foreach ($row in $rawRows) {
            if ($c -lt $row.Count) {
                $v = [string]$row[$c]
                if (-not [string]::IsNullOrWhiteSpace($v)) {
                    $hasValue = $true
                    break
                }
            }
        }

        if ($hasValue) {
            [void]$usedColumns.Add($c)
        }
    }

    if ($usedColumns.Count -lt 2) {
        throw "The copied range contains fewer than two non-empty columns."
    }

    $maxCols = $usedColumns.Count

    $normalized = @()
    foreach ($row in $rawRows) {
        $arr = New-Object System.Collections.Generic.List[string]

        foreach ($sourceCol in $usedColumns) {
            if ($sourceCol -lt $row.Count) {
                [void]$arr.Add([string]$row[$sourceCol])
            } else {
                [void]$arr.Add("")
            }
        }

        $normalized += ,$arr.ToArray()
    }

    $numericColumns = @()
    for ($c = 0; $c -lt $maxCols; $c++) {
        $isNumeric = $true
        $nonEmpty = 0

        for ($r = 1; $r -lt $normalized.Count; $r++) {
            $val = [string]$normalized[$r][$c]
            if (-not [string]::IsNullOrWhiteSpace($val)) {
                $nonEmpty++
                if (-not (Is-NumericCell $val)) {
                    $isNumeric = $false
                    break
                }
            }
        }

        $numericColumns += ($isNumeric -and $nonEmpty -gt 0)
    }

    $out = New-Object System.Collections.Generic.List[string]

    $header = @()
    for ($c = 0; $c -lt $maxCols; $c++) {
        $header += (Escape-MarkdownCell ([string]$normalized[0][$c]))
    }
    [void]$out.Add("| " + ($header -join " | ") + " |")

    $sep = @()
    for ($c = 0; $c -lt $maxCols; $c++) {
        if ($numericColumns[$c]) { $sep += "---:" }
        else { $sep += "---" }
    }
    [void]$out.Add("|" + (($sep | ForEach-Object { " $_ " }) -join "|") + "|")

    for ($r = 1; $r -lt $normalized.Count; $r++) {
        $cells = @()
        for ($c = 0; $c -lt $maxCols; $c++) {
            $cells += (Escape-MarkdownCell ([string]$normalized[$r][$c]))
        }
        [void]$out.Add("| " + ($cells -join " | ") + " |")
    }

    return ($out -join "`r`n")
}

function Split-MarkdownRow([string]$Line) {
    $s = $Line.Trim()
    if ($s.StartsWith("|")) { $s = $s.Substring(1) }
    if ($s.EndsWith("|")) { $s = $s.Substring(0, $s.Length - 1) }

    $placeholder = [char]0xE000
    $s = $s -replace '\\\|', [string]$placeholder
    $parts = @($s -split '\|', -1)
    $result = @()

    foreach ($p in $parts) {
        $v = $p.Trim()
        $v = $v.Replace([string]$placeholder, "|")
        $v = $v -replace '<br\s*/?>', ' / '
        if ($v -match '^\*\*(.*)\*\*$') { $v = $Matches[1] }
        elseif ($v -match '^__(.*)__$') { $v = $Matches[1] }
        elseif ($v -match '^`(.*)`$') { $v = $Matches[1] }
        $result += $v
    }
    return ,$result
}

function Is-MarkdownSeparator([string]$Line) {
    $cells = @(Split-MarkdownRow $Line)
    if ($cells.Count -lt 2) { return $false }
    foreach ($cell in $cells) {
        $v = $cell.Trim()
        if ($v -notmatch '^:?-{3,}:?$') { return $false }
    }
    return $true
}

function Extract-MarkdownTable([string]$Text) {
    $Text = Normalize-Newlines $Text
    $lines = @($Text -split "`n")

    for ($i = 0; $i -lt ($lines.Count - 1); $i++) {
        if (($lines[$i] -match '\|') -and (Is-MarkdownSeparator $lines[$i + 1])) {
            $table = New-Object System.Collections.Generic.List[string]
            [void]$table.Add($lines[$i])
            [void]$table.Add($lines[$i + 1])

            for ($j = $i + 2; $j -lt $lines.Count; $j++) {
                $line = $lines[$j]
                if ([string]::IsNullOrWhiteSpace($line)) { break }
                if ($line -notmatch '\|') { break }
                [void]$table.Add($line)
            }
            return ,$table.ToArray()
        }
    }
    return $null
}

function Recover-FlattenedMarkdown([string]$Text) {
    # Some applications remove line breaks when copying Markdown.
    # The end "|" of one row and the beginning "|" of the next then become "||".
    $s = Normalize-Newlines $Text
    $s = [regex]::Replace($s, '(?m)^\s*```(?:md|markdown|text)?\s*$', '')
    $s = $s.Trim()

    # Main v4 fix: reconstruct rows from "||" / "|   |".
    $candidate = [regex]::Replace($s, '\|\s*\|', "|`n|")
    $table = @(Extract-MarkdownTable $candidate)
    if ($null -ne $table -and $table.Count -ge 2) {
        return $candidate
    }

    return $null
}

function Convert-FlattenedMarkdownToTsv([string]$Text) {
    # First try to reconstruct real rows from "||".
    $recovered = Recover-FlattenedMarkdown $Text
    if (-not [string]::IsNullOrWhiteSpace($recovered)) {
        $table = @(Extract-MarkdownTable $recovered)
        $out = New-Object System.Collections.Generic.List[string]

        $header = @(Split-MarkdownRow $table[0])
        [void]$out.Add(($header -join "`t"))

        for ($i = 2; $i -lt $table.Count; $i++) {
            $cells = @(Split-MarkdownRow $table[$i])
            [void]$out.Add(($cells -join "`t"))
        }

        return ($out -join "`r`n")
    }

    # Fallback mode: determine the number of columns from the separator row |---|---:|.
    $s = Normalize-Newlines $Text
    $s = [regex]::Replace($s, '(?m)^\s*```(?:md|markdown|text)?\s*$', '')
    $s = ($s -replace '[\r\n]+', ' ').Trim()

    if ($s -notmatch '\|' -or $s -notmatch '-{3,}') {
        throw "The flattened Markdown table could not be recognized."
    }

    if ($s.StartsWith("|")) { $s = $s.Substring(1) }
    if ($s.EndsWith("|")) { $s = $s.Substring(0, $s.Length - 1) }

    $raw0 = @($s -split '\|', -1)
    $raw = @()
    foreach ($item in $raw0) { $raw += ([string]$item).Trim() }

    $bestStart = -1
    $bestLen = 0
    $runStart = -1
    $runLen = 0

    for ($i = 0; $i -lt $raw.Count; $i++) {
        if ($raw[$i] -match '^:?-{3,}:?$') {
            if ($runLen -eq 0) { $runStart = $i }
            $runLen++
            if ($runLen -gt $bestLen) {
                $bestLen = $runLen
                $bestStart = $runStart
            }
        } else {
            $runLen = 0
            $runStart = -1
        }
    }

    if ($bestLen -lt 2) {
        throw "Could not determine the number of columns from the |---|---| separator row."
    }

    $cols = $bestLen

    # The last N non-empty values before the separator are treated as headers.
    $headerRev = New-Object System.Collections.Generic.List[string]
    for ($i = $bestStart - 1; $i -ge 0 -and $headerRev.Count -lt $cols; $i--) {
        if (-not [string]::IsNullOrWhiteSpace($raw[$i])) {
            [void]$headerRev.Add($raw[$i])
        }
    }

    if ($headerRev.Count -ne $cols) {
        throw "Could not reconstruct the table header."
    }

    $header = @()
    for ($i = $headerRev.Count - 1; $i -ge 0; $i--) {
        $header += $headerRev[$i]
    }

    $out = New-Object System.Collections.Generic.List[string]
    [void]$out.Add(($header -join "`t"))

    # After the separator, collect non-empty values and group them by column count.
    $values = New-Object System.Collections.Generic.List[string]
    for ($i = $bestStart + $bestLen; $i -lt $raw.Count; $i++) {
        if (-not [string]::IsNullOrWhiteSpace($raw[$i])) {
            [void]$values.Add($raw[$i])
        }
    }

    for ($i = 0; $i -lt $values.Count; $i += $cols) {
        $row = @()
        for ($c = 0; $c -lt $cols; $c++) {
            $idx = $i + $c
            if ($idx -lt $values.Count) { $row += $values[$idx] }
            else { $row += "" }
        }
        if (($row -join "").Trim().Length -gt 0) {
            [void]$out.Add(($row -join "`t"))
        }
    }

    if ($out.Count -lt 2) {
        throw "No data rows were found."
    }

    return ($out -join "`r`n")
}

function Convert-MarkdownToTsv([string]$Text) {
    $Text = Normalize-Newlines $Text

    if (($Text -match "`t") -and ($Text -notmatch '\|')) {
        return $Text.Trim()
    }

    # 1) Normal multiline Markdown table.
    $table = @(Extract-MarkdownTable $Text)
    if ($null -ne $table -and $table.Count -ge 2) {
        $out = New-Object System.Collections.Generic.List[string]
        $header = @(Split-MarkdownRow $table[0])
        [void]$out.Add(($header -join "`t"))

        for ($i = 2; $i -lt $table.Count; $i++) {
            $cells = @(Split-MarkdownRow $table[$i])
            [void]$out.Add(($cells -join "`t"))
        }
        return ($out -join "`r`n")
    }

    # 2) Flattened Markdown.
    # v2 checked whether there was any \n at all; Windows often leaves one trailing
    # newline in the clipboard, so recovery was skipped. Since v3, always try recovery.
    if (($Text -match '\|') -and ($Text -match '-{3,}')) {
        return Convert-FlattenedMarkdownToTsv $Text
    }

    throw "No Markdown table found. A header and a separator row such as |---|---:| are required."
}

function Repair-Utf8Mojibake([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }

    # Typical case: UTF-8 bytes containing Cyrillic were misinterpreted
    # as Windows-1251 and became mojibake such as "РјР°Р»РёРЅР°".
    if ($Text -match '[РС][\x80-\xBFА-Яа-яЁё]') {
        try {
            $cp1251 = [System.Text.Encoding]::GetEncoding(1251)
            $bytes = $cp1251.GetBytes($Text)
            $fixed = [System.Text.Encoding]::UTF8.GetString($bytes)

            # Accept the repair only if characteristic mojibake noise is reduced.
            $beforeNoise = ([regex]::Matches($Text, 'Р.|С.')).Count
            $afterNoise  = ([regex]::Matches($fixed, 'Р.|С.')).Count

            if ($afterNoise -lt $beforeNoise) {
                return $fixed
            }
        } catch {
            # If repair fails, return the original text.
        }
    }

    return $Text
}

function Strip-Html([string]$Html) {
    $v = $Html
    $v = $v -replace '(?is)<br\s*/?>', ' / '
    $v = $v -replace '(?is)</p\s*>', ' '
    $v = $v -replace '(?is)<[^>]+>', ''
    $v = [System.Web.HttpUtility]::HtmlDecode($v)
    $v = $v -replace '[\r\n\t]+', ' '
    $v = $v -replace '\s{2,}', ' '
    $v = Repair-Utf8Mojibake $v
    return $v.Trim()
}

function Convert-HtmlTableToTsv([string]$Html) {
    if ([string]::IsNullOrWhiteSpace($Html) -or $Html -notmatch '(?is)<table\b') {
        throw "No HTML table was found in the clipboard."
    }

    $tableMatch = [regex]::Match($Html, '(?is)<table\b[^>]*>(.*?)</table>')
    if (-not $tableMatch.Success) { throw "Could not parse the HTML table." }

    $rows = [regex]::Matches($tableMatch.Value, '(?is)<tr\b[^>]*>(.*?)</tr>')
    $out = New-Object System.Collections.Generic.List[string]

    foreach ($row in $rows) {
        $cells = [regex]::Matches($row.Groups[1].Value, '(?is)<t[hd]\b[^>]*>(.*?)</t[hd]>')
        if ($cells.Count -eq 0) { continue }

        $vals = @()
        foreach ($cell in $cells) {
            $vals += (Strip-Html $cell.Groups[1].Value)
        }
        [void]$out.Add(($vals -join "`t"))
    }

    if ($out.Count -lt 1) { throw "No cells were found in the HTML table." }
    return ($out -join "`r`n")
}

function Get-ClipboardSnapshot {
    $obj = [System.Windows.Forms.Clipboard]::GetDataObject()
    $text = ""
    $html = ""

    if ($null -ne $obj) {
        if ($obj.GetDataPresent([System.Windows.Forms.DataFormats]::UnicodeText)) {
            $text = [string]$obj.GetData([System.Windows.Forms.DataFormats]::UnicodeText)
        } elseif ($obj.GetDataPresent([System.Windows.Forms.DataFormats]::Text)) {
            $text = [string]$obj.GetData([System.Windows.Forms.DataFormats]::Text)
        }

        if ($obj.GetDataPresent([System.Windows.Forms.DataFormats]::Html)) {
            $html = [string]$obj.GetData([System.Windows.Forms.DataFormats]::Html)
        }
    }

    return [PSCustomObject]@{
        Text = $text
        Html = $html
    }
}

function Set-ClipboardText([string]$Text) {
    if (-not [string]::IsNullOrEmpty($Text)) {
        [System.Windows.Forms.Clipboard]::SetText($Text)
    }
}


$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MD Table Converter"
    Width="1180"
    Height="780"
    MinWidth="920"
    MinHeight="620"
    WindowStartupLocation="CenterScreen"
    Background="#0B0D12"
    FontFamily="Segoe UI"
    Foreground="#F5F7FB">

    <Window.Resources>
        <SolidColorBrush x:Key="BgBrush" Color="#0B0D12"/>
        <SolidColorBrush x:Key="PanelBrush" Color="#131720"/>
        <SolidColorBrush x:Key="Panel2Brush" Color="#171C27"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#272E3C"/>
        <SolidColorBrush x:Key="TextBrush" Color="#F5F7FB"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#8E98A8"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#8B5CF6"/>
        <SolidColorBrush x:Key="AccentHoverBrush" Color="#9D73FA"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#22C55E"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#EF4444"/>

        <Style x:Key="PrimaryButton" TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,11"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Root"
                                Background="{TemplateBinding Background}"
                                CornerRadius="10"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Root" Property="Background" Value="{StaticResource AccentHoverBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Root" Property="Opacity" Value="0.84"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Root" Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="#1B2130"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="Padding" Value="14,9"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Root"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="9"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter
                                HorizontalAlignment="Center"
                                VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Root" Property="Background" Value="#252C3B"/>
                                <Setter TargetName="Root" Property="BorderBrush" Value="#3A4355"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="Root" Property="Opacity" Value="0.82"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource SecondaryButton}">
            <Setter Property="Foreground" Value="#FCA5A5"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Foreground" Value="#E8ECF4"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontFamily" Value="Cascadia Mono, Consolas"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="White"/>
            <Setter Property="SelectionBrush" Value="#6D4AC7"/>
            <Setter Property="Padding" Value="0"/>
        </Style>

        <Style TargetType="ScrollBar">
            <Setter Property="Width" Value="10"/>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="18"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="18"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="18"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0"
                Background="{StaticResource PanelBrush}"
                BorderBrush="{StaticResource BorderBrush}"
                BorderThickness="1"
                CornerRadius="18"
                Padding="20">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="16"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Border Width="48" Height="48"
                        CornerRadius="14"
                        Background="#241A3A"
                        BorderBrush="#4C3580"
                        BorderThickness="1">
                    <TextBlock Text="↔"
                               FontSize="26"
                               FontWeight="SemiBold"
                               Foreground="#C4B5FD"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"/>
                </Border>

                <StackPanel Grid.Column="2" VerticalAlignment="Center">
                    <TextBlock Text="MD ↔ Excel"
                               FontSize="24"
                               FontWeight="SemiBold"
                               Foreground="{StaticResource TextBrush}"/>
                    <TextBlock Text="Smart Excel ↔ Markdown converter with full Unicode support"
                               Margin="0,4,0,0"
                               FontSize="13"
                               Foreground="{StaticResource MutedBrush}"/>
                </StackPanel>

                <Border Grid.Column="3"
                        CornerRadius="999"
                        Background="#161B25"
                        BorderBrush="{StaticResource BorderBrush}"
                        BorderThickness="1"
                        Padding="12,7"
                        VerticalAlignment="Center">
                    <StackPanel Orientation="Horizontal">
                        <Ellipse Width="7" Height="7"
                                 Fill="{StaticResource SuccessBrush}"
                                 Margin="0,0,7,0"
                                 VerticalAlignment="Center"/>
                        <TextBlock Text="Local"
                                   Foreground="#B8C1CF"
                                   FontSize="12"
                                   FontWeight="SemiBold"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <!-- Actions -->
        <Border Grid.Row="2"
                Background="{StaticResource PanelBrush}"
                BorderBrush="{StaticResource BorderBrush}"
                BorderThickness="1"
                CornerRadius="16"
                Padding="14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <Button x:Name="BtnPaste"
                        Grid.Column="0"
                        Style="{StaticResource SecondaryButton}"
                        Content="Paste from Clipboard"
                        ToolTip="Paste an Excel table or Markdown from the clipboard"/>

                <Button x:Name="BtnOpen"
                        Grid.Column="2"
                        Style="{StaticResource SecondaryButton}"
                        Content="Open .md"
                        ToolTip="Open a Markdown file"/>

                <Button x:Name="BtnClear"
                        Grid.Column="4"
                        Style="{StaticResource DangerButton}"
                        Content="Clear"/>

                <Button x:Name="BtnExcelToMd"
                        Grid.Column="6"
                        Style="{StaticResource PrimaryButton}"
                        Content="Excel → Markdown"
                        MinWidth="170"
                        ToolTip="Convert TSV / Excel data to a Markdown table"/>

                <Button x:Name="BtnMdToExcel"
                        Grid.Column="8"
                        Style="{StaticResource PrimaryButton}"
                        Content="Markdown → Excel"
                        MinWidth="170"
                        ToolTip="Convert a Markdown table to data ready to paste into Excel"/>
            </Grid>
        </Border>

        <!-- Editors -->
        <Grid Grid.Row="4">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="16"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <Border Grid.Column="0"
                    Background="{StaticResource PanelBrush}"
                    BorderBrush="{StaticResource BorderBrush}"
                    BorderThickness="1"
                    CornerRadius="16"
                    Padding="16">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="12"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel>
                            <TextBlock Text="Source Data"
                                       FontSize="15"
                                       FontWeight="SemiBold"/>
                            <TextBlock Text="Excel / TSV / Markdown"
                                       Margin="0,3,0,0"
                                       FontSize="12"
                                       Foreground="{StaticResource MutedBrush}"/>
                        </StackPanel>
                        <Border Grid.Column="1"
                                CornerRadius="999"
                                Background="#1A202C"
                                Padding="9,5"
                                VerticalAlignment="Top">
                            <TextBlock Text="INPUT"
                                       FontSize="10"
                                       FontWeight="Bold"
                                       Foreground="#7DD3FC"/>
                        </Border>
                    </Grid>

                    <Border Grid.Row="2"
                            Background="#0F131B"
                            BorderBrush="#222A38"
                            BorderThickness="1"
                            CornerRadius="12"
                            Padding="14">
                        <TextBox x:Name="InputText"
                                 AcceptsReturn="True"
                                 AcceptsTab="True"
                                 TextWrapping="NoWrap"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"/>
                    </Border>
                </Grid>
            </Border>

            <Border Grid.Column="2"
                    Background="{StaticResource PanelBrush}"
                    BorderBrush="{StaticResource BorderBrush}"
                    BorderThickness="1"
                    CornerRadius="16"
                    Padding="16">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="12"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel>
                            <TextBlock Text="Result"
                                       FontSize="15"
                                       FontWeight="SemiBold"/>
                            <TextBlock Text="Automatically copied to the clipboard"
                                       Margin="0,3,0,0"
                                       FontSize="12"
                                       Foreground="{StaticResource MutedBrush}"/>
                        </StackPanel>
                        <Border Grid.Column="1"
                                CornerRadius="999"
                                Background="#1A202C"
                                Padding="9,5"
                                VerticalAlignment="Top">
                            <TextBlock Text="OUTPUT"
                                       FontSize="10"
                                       FontWeight="Bold"
                                       Foreground="#86EFAC"/>
                        </Border>
                    </Grid>

                    <Border Grid.Row="2"
                            Background="#0F131B"
                            BorderBrush="#222A38"
                            BorderThickness="1"
                            CornerRadius="12"
                            Padding="14">
                        <TextBox x:Name="OutputText"
                                 AcceptsReturn="True"
                                 AcceptsTab="True"
                                 TextWrapping="NoWrap"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"/>
                    </Border>
                </Grid>
            </Border>
        </Grid>

        <!-- Footer -->
        <Border Grid.Row="6"
                Background="{StaticResource PanelBrush}"
                BorderBrush="{StaticResource BorderBrush}"
                BorderThickness="1"
                CornerRadius="14"
                Padding="14,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="10"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="StatusDot"
                             Width="8" Height="8"
                             Fill="{StaticResource SuccessBrush}"
                             Margin="0,0,9,0"/>
                    <TextBlock x:Name="StatusText"
                               Text="Ready"
                               Foreground="#AAB4C3"
                               FontSize="12"
                               VerticalAlignment="Center"/>
                </StackPanel>

                <Button x:Name="BtnCopy"
                        Grid.Column="1"
                        Style="{StaticResource SecondaryButton}"
                        Content="Copy Result"/>

                <Button x:Name="BtnSave"
                        Grid.Column="3"
                        Style="{StaticResource SecondaryButton}"
                        Content="Save…"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$InputText    = $window.FindName("InputText")
$OutputText   = $window.FindName("OutputText")
$BtnPaste     = $window.FindName("BtnPaste")
$BtnOpen      = $window.FindName("BtnOpen")
$BtnClear     = $window.FindName("BtnClear")
$BtnExcelToMd = $window.FindName("BtnExcelToMd")
$BtnMdToExcel = $window.FindName("BtnMdToExcel")
$BtnCopy      = $window.FindName("BtnCopy")
$BtnSave      = $window.FindName("BtnSave")
$StatusText   = $window.FindName("StatusText")
$StatusDot    = $window.FindName("StatusDot")

function Set-Status([string]$Message, [string]$Kind = "ok") {
    $StatusText.Text = $Message
    switch ($Kind) {
        "error" { $StatusDot.Fill = [Windows.Media.Brushes]::Tomato }
        "info"  { $StatusDot.Fill = (New-Object Windows.Media.SolidColorBrush ([Windows.Media.Color]::FromRgb(139,92,246))) }
        default { $StatusDot.Fill = (New-Object Windows.Media.SolidColorBrush ([Windows.Media.Color]::FromRgb(34,197,94))) }
    }
}

function Show-AppError([string]$Title, [string]$Message) {
    Set-Status $Message "error"
    [System.Windows.MessageBox]::Show(
        $window,
        $Message,
        $Title,
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    ) | Out-Null
}

$BtnPaste.Add_Click({
    try {
        $clip = Get-ClipboardSnapshot

        if (-not [string]::IsNullOrWhiteSpace($clip.Text) -and $clip.Text -match "`t") {
            $InputText.Text = $clip.Text
            Set-Status "Unicode data pasted from Excel" "info"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($clip.Html) -and $clip.Html -match '(?is)<table\b') {
            $tsv = Convert-HtmlTableToTsv $clip.Html
            $InputText.Text = $tsv
            Set-Status "HTML table detected in the clipboard" "info"
        } else {
            $plain = $clip.Text
            $recovered = $null

            if (($plain -match '\|') -and ($plain -match '-{3,}')) {
                $recovered = Recover-FlattenedMarkdown $plain
            }

            if (-not [string]::IsNullOrWhiteSpace($recovered)) {
                $InputText.Text = $recovered
                Set-Status "Flattened Markdown restored into rows" "info"
            } else {
                $InputText.Text = $plain
                Set-Status "Data pasted from the clipboard"
            }
        }
    } catch {
        Show-AppError "Clipboard Error" $_.Exception.Message
    }
})

$BtnExcelToMd.Add_Click({
    try {
        $clip = Get-ClipboardSnapshot
        $source = $null

        # IMPORTANT v8:
        # For Excel, prefer UnicodeText from the clipboard because it reliably
        # preserves Unicode text. Extra physical columns caused by merged cells
        # are already removed by Convert-TsvToMarkdown when completely empty.
        if (-not [string]::IsNullOrWhiteSpace($clip.Text) -and $clip.Text -match "`t") {
            $source = $clip.Text
            $InputText.Text = $source
            Set-Status "Excel table read as Unicode" "info"
        }
        elseif (-not [string]::IsNullOrWhiteSpace($InputText.Text)) {
            $source = $InputText.Text
        }
        elseif (-not [string]::IsNullOrWhiteSpace($clip.Html) -and $clip.Html -match '(?is)<table\b') {
            # Fallback path for rare cases where UnicodeText is unavailable.
            $source = Convert-HtmlTableToTsv $clip.Html
            $InputText.Text = $source
            Set-Status "Fallback HTML parser used" "info"
        }
        else {
            $source = $clip.Text
        }

        $result = Convert-TsvToMarkdown $source
        $OutputText.Text = $result
        Set-ClipboardText $result
        Set-Status "Markdown created and copied to the clipboard"
    } catch {
        Show-AppError "Excel → Markdown" $_.Exception.Message
    }
})

$BtnMdToExcel.Add_Click({
    try {
        $result = $null

        if (-not [string]::IsNullOrWhiteSpace($InputText.Text)) {
            $result = Convert-MarkdownToTsv $InputText.Text
        } else {
            $clip = Get-ClipboardSnapshot

            if (-not [string]::IsNullOrWhiteSpace($clip.Html) -and $clip.Html -match '(?is)<table\b') {
                $result = Convert-HtmlTableToTsv $clip.Html
            } else {
                $result = Convert-MarkdownToTsv $clip.Text
            }
        }

        $OutputText.Text = $result
        Set-ClipboardText $result
        Set-Status "Ready to paste into Excel — press Ctrl+V"
    } catch {
        Show-AppError "Markdown → Excel" $_.Exception.Message
    }
})

$BtnOpen.Add_Click({
    try {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Markdown (*.md)|*.md|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $InputText.Text = [System.IO.File]::ReadAllText($dialog.FileName, [System.Text.Encoding]::UTF8)
            Set-Status ("Opened: " + [System.IO.Path]::GetFileName($dialog.FileName))
        }
    } catch {
        Show-AppError "Open Error" $_.Exception.Message
    }
})

$BtnClear.Add_Click({
    $InputText.Clear()
    $OutputText.Clear()
    Set-Status "Cleared"
    $InputText.Focus()
})

$BtnCopy.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($OutputText.Text)) {
            throw "The result is empty."
        }
        Set-ClipboardText $OutputText.Text
        Set-Status "Result copied to the clipboard"
    } catch {
        Show-AppError "Copy Error" $_.Exception.Message
    }
})

$BtnSave.Add_Click({
    try {
        if ([string]::IsNullOrWhiteSpace($OutputText.Text)) {
            throw "The result is empty."
        }

        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Filter = "Markdown (*.md)|*.md|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $dialog.DefaultExt = "md"
        $dialog.FileName = "table.md"

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            [System.IO.File]::WriteAllText(
                $dialog.FileName,
                $OutputText.Text,
                (New-Object System.Text.UTF8Encoding($false))
            )
            Set-Status ("Saved: " + [System.IO.Path]::GetFileName($dialog.FileName))
        }
    } catch {
        Show-AppError "Save Error" $_.Exception.Message
    }
})

# Keyboard shortcuts:
# Ctrl+V — paste into the input field using the standard behavior.
# Ctrl+Shift+M — Excel → Markdown
# Ctrl+Shift+E — Markdown → Excel
$window.Add_KeyDown({
    param($sender, $e)

    if ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) {
        if ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift) {
            if ($e.Key -eq [System.Windows.Input.Key]::M) {
                $BtnExcelToMd.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                $e.Handled = $true
            }
            elseif ($e.Key -eq [System.Windows.Input.Key]::E) {
                $BtnMdToExcel.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
                $e.Handled = $true
            }
        }
    }
})

$InputText.Focus()
[void]$window.ShowDialog()
