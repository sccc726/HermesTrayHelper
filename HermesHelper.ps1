param(
    [switch]$SelfTest
)

$ErrorActionPreference = "Stop"

$script:WslDistro = "Ubuntu"
$script:ActionsPath = Join-Path $PSScriptRoot "actions.json"

function ConvertTo-PowerShellSingleQuotedLiteral {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ($Value -replace "'", "''") + "'"
}

function ConvertTo-BashSingleQuotedLiteral {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) {
        return "''"
    }

    return "'" + ($Value -replace "'", "'\''") + "'"
}

function Get-ShortText {
    param(
        [string]$Text,
        [int]$MaxLength = 63
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ""
    }

    $singleLine = ($Text -replace "\s+", " ").Trim()
    if ($singleLine.Length -le $MaxLength) {
        return $singleLine
    }

    if ($MaxLength -le 3) {
        return $singleLine.Substring(0, $MaxLength)
    }

    return $singleLine.Substring(0, $MaxLength - 3) + "..."
}

function Get-HermesActionValue {
    param(
        [Parameter(Mandatory = $true)]$Action,
        [Parameter(Mandatory = $true)][string]$Name
    )

    foreach ($property in $Action.PSObject.Properties) {
        if ($property.Name -eq $Name) {
            return $property.Value
        }
    }

    return $null
}

function Test-HermesActionGroup {
    param($Action)

    if ($null -eq $Action) {
        return $false
    }

    return ([string](Get-HermesActionValue -Action $Action -Name "type")) -eq "group"
}

function Assert-HermesAction {
    param(
        [Parameter(Mandatory = $true)]$Action,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $label = [string](Get-HermesActionValue -Action $Action -Name "label")
    $type = [string](Get-HermesActionValue -Action $Action -Name "type")
    $command = [string](Get-HermesActionValue -Action $Action -Name "command")

    if ([string]::IsNullOrWhiteSpace($label)) {
        throw "$Path 缺少 label。"
    }

    if (Test-HermesActionGroup -Action $Action) {
        $module = [string](Get-HermesActionValue -Action $Action -Name "module")
        if ($module -eq "sessions") {
            return
        }

        $children = @(Get-HermesActionValue -Action $Action -Name "children")
        if ($children.Count -eq 0) {
            throw "分组动作 '$label' 缺少 children。"
        }

        for ($i = 0; $i -lt $children.Count; $i++) {
            Assert-HermesAction -Action $children[$i] -Path "$Path.children[$i]"
        }

        return
    }

    if (-not [string]::IsNullOrWhiteSpace($type)) {
        throw "$Path 的 type '$type' 不受支持。"
    }

    if ([string]::IsNullOrWhiteSpace($command)) {
        throw "动作 '$label' 缺少 command。"
    }
}

function Get-HermesExecutableActionCount {
    param($Actions)

    if ($null -eq $Actions) {
        return 0
    }

    $count = 0
    foreach ($action in @($Actions)) {
        if ($null -eq $action) {
            continue
        }
        if (Test-HermesActionGroup -Action $action) {
            $count += Get-HermesExecutableActionCount -Actions (Get-HermesActionValue -Action $action -Name "children")
        }
        else {
            $count++
        }
    }

    return $count
}

function Get-FirstExecutableAction {
    param($Actions)

    if ($null -eq $Actions) {
        return $null
    }

    foreach ($action in @($Actions)) {
        if ($null -eq $action) {
            continue
        }
        if (Test-HermesActionGroup -Action $action) {
            $result = Get-FirstExecutableAction -Actions (Get-HermesActionValue -Action $action -Name "children")
            if ($null -ne $result) {
                return $result
            }
        }
        else {
            return $action
        }
    }

    return $null
}

function Load-HermesActions {
    if (-not (Test-Path -LiteralPath $script:ActionsPath)) {
        throw "找不到动作配置文件：$script:ActionsPath"
    }

    $raw = Get-Content -LiteralPath $script:ActionsPath -Raw -Encoding UTF8
    $parsed = ConvertFrom-Json -InputObject $raw
    $actions = @($parsed)

    if ($actions.Count -eq 0) {
        throw "动作配置文件为空：$script:ActionsPath"
    }

    for ($i = 0; $i -lt $actions.Count; $i++) {
        Assert-HermesAction -Action $actions[$i] -Path "actions.json[$i]"
    }

    return $actions
}

function Invoke-WslTextCommand {
    param(
        [Parameter(Mandatory = $true)][string]$Command
    )

    $output = @()
    $exitCode = 1

    try {
        $output = & wsl.exe -d $script:WslDistro -e bash -lc $Command 2>&1
        $exitCode = $LASTEXITCODE
    }
    catch {
        $output = @($_.Exception.Message)
        $exitCode = 1
    }

    $text = ($output | ForEach-Object { [string]$_ }) -join [Environment]::NewLine

    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = $text.Trim()
    }
}

function Test-HermesHealth {
    $wslCommand = Get-Command wsl.exe -ErrorAction SilentlyContinue
    if ($null -eq $wslCommand) {
        return [pscustomobject]@{
            State = "Error"
            Summary = "找不到 wsl.exe"
            Details = "Windows 无法找到 wsl.exe。"
            ToolTip = "Hermes Helper: 找不到 WSL"
        }
    }

    $probe = Invoke-WslTextCommand -Command "printf 'PATH='; command -v hermes; hermes --version"
    if ($probe.ExitCode -ne 0) {
        $details = if ($probe.Output) { $probe.Output } else { "WSL 命令执行失败。" }
        return [pscustomobject]@{
            State = "Error"
            Summary = "Hermes 不可用"
            Details = $details
            ToolTip = Get-ShortText -Text ("Hermes Helper: " + $details)
        }
    }

    $lines = @($probe.Output -split "`r?`n" | Where-Object { $_.Trim() })
    $pathLine = $lines | Where-Object { $_ -like "PATH=*" } | Select-Object -First 1
    $versionLine = $lines | Where-Object { $_ -like "Hermes Agent*" } | Select-Object -First 1

    $path = if ($pathLine) { $pathLine.Substring(5).Trim() } else { "hermes" }
    $version = if ($versionLine) { $versionLine.Trim() } else { "Hermes 可用" }

    return [pscustomobject]@{
        State = "Ok"
        Summary = $version
        Details = "WSL: $script:WslDistro`nHermes: $path`n$version"
        ToolTip = Get-ShortText -Text ("Hermes Helper: " + $version)
    }
}

function Get-GatewayStatus {
    $probe = Invoke-WslTextCommand -Command "hermes gateway status"
    $output = [string]$probe.Output
    $normalized = $output.ToLowerInvariant()
    $state = "unknown"
    $label = "状态: 未知"

    if ($normalized -match "active:\s+active\s+\(running\)" -or $normalized -match "gateway service is running") {
        $state = "running"
        $label = "状态: 运行中"
    }
    elseif (
        $normalized -match "active:\s+inactive" -or
        $normalized -match "active:\s+failed" -or
        $normalized -match "not running" -or
        $normalized -match "inactive\s+\(dead\)" -or
        $normalized -match "could not be found" -or
        $normalized -match "not loaded"
    ) {
        $state = "stopped"
        $label = "状态: 未运行"
    }

    if ($probe.ExitCode -ne 0 -and $state -eq "unknown") {
        $label = "状态: 未知"
    }

    return [pscustomobject]@{
        State = $state
        Label = $label
        Details = if ($output) { $output } else { "无法读取 Gateway 状态。" }
        ExitCode = $probe.ExitCode
    }
}

function Remove-AnsiEscape {
    param([AllowNull()][string]$Text)

    if ($null -eq $Text) {
        return ""
    }

    return $Text -replace "`e\[[0-9;?]*[ -/]*[@-~]", ""
}

function ConvertFrom-HermesSessionsList {
    param([AllowNull()][string]$Text)

    $sessions = New-Object System.Collections.ArrayList
    $lines = @((Remove-AnsiEscape -Text $Text) -split "`r?`n")

    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) {
            continue
        }
        if ($trimmed -match "^Title\s+Preview\s+Last Active\s+ID$") {
            continue
        }
        if ($trimmed -match "^[─\-\s]+$") {
            continue
        }
        if ($trimmed -eq "No sessions found.") {
            continue
        }

        if ($trimmed -notmatch "^(?<before>.+?)\s+(?<id>[A-Za-z0-9_][A-Za-z0-9_.:-]*)$") {
            continue
        }

        $sessionId = $Matches["id"]
        $beforeId = $Matches["before"].Trim()
        $parts = @($beforeId -split "\s{2,}" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })

        if ($parts.Count -lt 2) {
            continue
        }

        $title = $parts[0].Trim()
        $lastActive = $parts[$parts.Count - 1].Trim()
        $preview = ""
        if ($parts.Count -gt 2) {
            $preview = ($parts[1..($parts.Count - 2)] -join " ").Trim()
        }

        if ($title -eq "—" -or [string]::IsNullOrWhiteSpace($title)) {
            $displayName = $preview
        }
        else {
            $displayName = $title
        }

        if ([string]::IsNullOrWhiteSpace($displayName)) {
            $displayName = $sessionId
        }

        $label = Get-ShortText -Text $displayName -MaxLength 42
        if (-not [string]::IsNullOrWhiteSpace($lastActive)) {
            $label = "$label - $lastActive"
        }

        [void]$sessions.Add([pscustomobject]@{
            Id = $sessionId
            Title = $title
            Preview = $preview
            LastActive = $lastActive
            Label = $label
            ToolTip = "ID: $sessionId`nTitle: $title`nPreview: $preview"
        })
    }

    return @($sessions.ToArray())
}

function Get-HermesSessions {
    param(
        [int]$Limit = 10,
        [string]$Source = "cli"
    )

    if ($Limit -lt 1) {
        $Limit = 10
    }

    $command = "hermes sessions list --limit $Limit"
    if (-not [string]::IsNullOrWhiteSpace($Source)) {
        $command = $command + " --source " + (ConvertTo-BashSingleQuotedLiteral -Value $Source)
    }

    $probe = Invoke-WslTextCommand -Command $command
    $sessions = if ($probe.ExitCode -eq 0) {
        ConvertFrom-HermesSessionsList -Text $probe.Output
    }
    else {
        @()
    }

    return [pscustomobject]@{
        ExitCode = $probe.ExitCode
        Output = $probe.Output
        Sessions = @($sessions)
    }
}

function Show-InputDialog {
    param(
        [string]$Title,
        [string]$Prompt,
        [AllowNull()][string]$DefaultText
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($Title) { $Title } else { "输入" }
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(520, 260)
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = if ($Prompt) { $Prompt } else { "请输入内容：" }
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(16, 16)
    [void]$form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.AcceptsReturn = $true
    $textBox.AcceptsTab = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBox.Location = New-Object System.Drawing.Point(16, 44)
    $textBox.Size = New-Object System.Drawing.Size(472, 112)
    if ($null -ne $DefaultText) {
        $textBox.Text = $DefaultText
    }
    [void]$form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "确定"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(312, 174)
    $okButton.Size = New-Object System.Drawing.Size(82, 30)
    [void]$form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "取消"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(406, 174)
    $cancelButton.Size = New-Object System.Drawing.Size(82, 30)
    [void]$form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton
    $form.Add_Shown({
        $textBox.Focus()
        $textBox.SelectAll()
    })

    $result = $form.ShowDialog()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }

    return $textBox.Text
}

function Show-StatusDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][string]$Text
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = "CenterScreen"
    $form.Size = New-Object System.Drawing.Size(760, 520)
    $form.MinimumSize = New-Object System.Drawing.Size(560, 360)
    $form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 9)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true
    $textBox.ReadOnly = $true
    $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $textBox.WordWrap = $false
    $textBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $textBox.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Top -bor
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Left -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $textBox.Location = New-Object System.Drawing.Point(14, 14)
    $textBox.Size = New-Object System.Drawing.Size(716, 404)
    $textBox.Text = $Text
    [void]$form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "确定"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Anchor = (
        [System.Windows.Forms.AnchorStyles]::Bottom -bor
        [System.Windows.Forms.AnchorStyles]::Right
    )
    $okButton.Location = New-Object System.Drawing.Point(648, 434)
    $okButton.Size = New-Object System.Drawing.Size(82, 30)
    [void]$form.Controls.Add($okButton)

    $form.AcceptButton = $okButton
    [void]$form.ShowDialog()
}

function Start-HermesWslTerminal {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$Command
    )

    $titleLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value ("Hermes Helper - " + $DisplayName)
    $distroLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value $script:WslDistro
    $commandLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value $Command

    $launcher = @"
`$Host.UI.RawUI.WindowTitle = $titleLiteral
Write-Host 'Hermes Helper'
Write-Host ('WSL:     ' + $distroLiteral)
Write-Host ('Command: ' + $commandLiteral)
Write-Host ''
& wsl.exe -d $distroLiteral -e bash -lc $commandLiteral
`$exitCode = `$LASTEXITCODE
Write-Host ''
Write-Host ('Command exited with code {0}' -f `$exitCode)
"@

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launcher))
    Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoExit",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        $encoded
    )
}

function Start-HermesContinueSessionTerminal {
    param(
        [Parameter(Mandatory = $true)][string]$DisplayName,
        [Parameter(Mandatory = $true)][string]$Command,
        [Parameter(Mandatory = $true)][string]$FallbackCommand
    )

    $titleLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value ("Hermes Helper - " + $DisplayName)
    $distroLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value $script:WslDistro
    $commandLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value $Command
    $fallbackLiteral = ConvertTo-PowerShellSingleQuotedLiteral -Value $FallbackCommand

    $launcher = @"
`$Host.UI.RawUI.WindowTitle = $titleLiteral
Write-Host 'Hermes Helper'
Write-Host ('WSL:     ' + $distroLiteral)
Write-Host ('Command: ' + $commandLiteral)
Write-Host ''
& wsl.exe -d $distroLiteral -e bash -lc $commandLiteral
`$exitCode = `$LASTEXITCODE
if (`$exitCode -ne 0) {
    Write-Host ''
    Write-Host 'No previous CLI session found. Opening session browser...'
    Write-Host ('Command: ' + $fallbackLiteral)
    Write-Host ''
    & wsl.exe -d $distroLiteral -e bash -lc $fallbackLiteral
    `$exitCode = `$LASTEXITCODE
}
Write-Host ''
Write-Host ('Command exited with code {0}' -f `$exitCode)
"@

    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($launcher))
    Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoExit",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-EncodedCommand",
        $encoded
    )
}

function Show-HermesCommandResult {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)]$Result
    )

    $success = $Result.ExitCode -eq 0
    $title = if ($success) { "$Label 完成" } else { "$Label 失败" }
    $icon = if ($success) {
        [System.Windows.Forms.MessageBoxIcon]::Information
    }
    else {
        [System.Windows.Forms.MessageBoxIcon]::Error
    }

    $output = if ($Result.Output) { $Result.Output } else { "(无输出)" }
    $message = if ($success) {
        "$Label 已完成。"
    }
    else {
        "$Label 执行失败，退出码：$($Result.ExitCode)"
    }

    $message = $message + [Environment]::NewLine + [Environment]::NewLine + (Get-ShortText -Text $output -MaxLength 700)

    [System.Windows.Forms.MessageBox]::Show(
        $message,
        $title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $icon
    ) | Out-Null
}

function Show-HermesErrorDialog {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)]$ErrorRecord
    )

    [System.Windows.Forms.MessageBox]::Show(
        [string]$ErrorRecord,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

function Start-HermesAction {
    param(
        [Parameter(Mandatory = $true)]$Action
    )

    $command = [string](Get-HermesActionValue -Action $Action -Name "command")
    $label = [string](Get-HermesActionValue -Action $Action -Name "label")
    $requiresInput = Get-HermesActionValue -Action $Action -Name "requiresInput"
    $executionMode = [string](Get-HermesActionValue -Action $Action -Name "executionMode")
    $fallbackCommand = [string](Get-HermesActionValue -Action $Action -Name "fallbackCommand")

    if ([string]::IsNullOrWhiteSpace($executionMode)) {
        $executionMode = "terminalHold"
    }

    if ($null -ne $requiresInput -and [bool]$requiresInput) {
        $inputText = Show-InputDialog `
            -Title ([string](Get-HermesActionValue -Action $Action -Name "inputTitle")) `
            -Prompt ([string](Get-HermesActionValue -Action $Action -Name "inputPrompt"))
        if ($null -eq $inputText) {
            return
        }

        if ([string]::IsNullOrWhiteSpace($inputText)) {
            [System.Windows.Forms.MessageBox]::Show(
                "输入为空，已取消执行。",
                "Hermes Helper",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return
        }

        $command = $command.TrimEnd() + " " + (ConvertTo-BashSingleQuotedLiteral -Value $inputText)
    }

    switch ($executionMode) {
        "terminalHold" {
            Start-HermesWslTerminal -DisplayName $label -Command $command
        }
        "continueSession" {
            if ([string]::IsNullOrWhiteSpace($fallbackCommand)) {
                $fallbackCommand = "hermes sessions browse"
            }
            Start-HermesContinueSessionTerminal -DisplayName $label -Command $command -FallbackCommand $fallbackCommand
        }
        "silentMessage" {
            $result = Invoke-WslTextCommand -Command $command
            Show-HermesCommandResult -Label $label -Result $result
        }
        "statusDialog" {
            $result = Invoke-WslTextCommand -Command $command
            $header = "Command: $command`r`nExitCode: $($result.ExitCode)`r`n`r`n"
            $body = if ($result.Output) { $result.Output } else { "(无输出)" }
            Show-StatusDialog -Title $label -Text ($header + $body)
        }
        default {
            Start-HermesWslTerminal -DisplayName $label -Command $command
        }
    }
}

function Update-GatewayMenu {
    param(
        [Parameter(Mandatory = $true)]$StatusItem,
        [Parameter(Mandatory = $true)][hashtable]$ActionItems
    )

    $StatusItem.Text = "状态: 正在刷新..."
    $StatusItem.ToolTipText = "正在静默执行 hermes gateway status"
    [System.Windows.Forms.Application]::DoEvents()

    $status = Get-GatewayStatus
    $StatusItem.Text = $status.Label
    $StatusItem.ToolTipText = Get-ShortText -Text $status.Details -MaxLength 500

    foreach ($key in $ActionItems.Keys) {
        $ActionItems[$key].Visible = $true
    }

    switch ($status.State) {
        "running" {
            if ($ActionItems.ContainsKey("start")) { $ActionItems["start"].Visible = $false }
        }
        "stopped" {
            if ($ActionItems.ContainsKey("restart")) { $ActionItems["restart"].Visible = $false }
            if ($ActionItems.ContainsKey("stop")) { $ActionItems["stop"].Visible = $false }
        }
        default {
            if ($ActionItems.ContainsKey("stop")) { $ActionItems["stop"].Visible = $false }
        }
    }
}

function Remove-HermesSession {
    param(
        [Parameter(Mandatory = $true)]$Session
    )

    $message = "确定删除这个会话吗？`n`n$($Session.Label)`nID: $($Session.Id)"
    $choice = [System.Windows.Forms.MessageBox]::Show(
        $message,
        "删除会话",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        return $false
    }

    $sessionLiteral = ConvertTo-BashSingleQuotedLiteral -Value $Session.Id
    $result = Invoke-WslTextCommand -Command "hermes sessions delete --yes $sessionLiteral"
    Show-HermesCommandResult -Label "删除会话" -Result $result

    return ($result.ExitCode -eq 0)
}

function Rename-HermesSession {
    param(
        [Parameter(Mandatory = $true)]$Session
    )

    $defaultTitle = [string]$Session.Title
    if ([string]::IsNullOrWhiteSpace($defaultTitle) -or $defaultTitle -eq "—") {
        $defaultTitle = [string]$Session.Preview
    }

    $inputTitle = Show-InputDialog `
        -Title "重命名会话" `
        -Prompt "输入新的会话标题：" `
        -DefaultText $defaultTitle

    if ($null -eq $inputTitle) {
        return $false
    }

    $newTitle = $inputTitle.Trim()
    if ([string]::IsNullOrWhiteSpace($newTitle)) {
        [System.Windows.Forms.MessageBox]::Show(
            "标题不能为空。",
            "重命名会话",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return $false
    }

    $backtickChar = [string][char]96
    $blockedCharacters = @('\', '"', $backtickChar, '$')
    foreach ($blockedCharacter in $blockedCharacters) {
        if ($newTitle.Contains($blockedCharacter)) {
            [System.Windows.Forms.MessageBox]::Show(
                ('标题不能包含以下字符：\ " ' + $backtickChar + ' $'),
                "重命名会话",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            return $false
        }
    }

    $sessionLiteral = ConvertTo-BashSingleQuotedLiteral -Value $Session.Id
    $quotedTitle = "\`"$newTitle\`""
    $result = Invoke-WslTextCommand -Command "hermes sessions rename $sessionLiteral $quotedTitle"
    $output = if ($result.Output) { $result.Output } else { "(无输出)" }
    $renamed = ([string]$result.Output).IndexOf("renamed to", [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    $success = $result.ExitCode -eq 0 -and $renamed

    if ($success) {
        [System.Windows.Forms.MessageBox]::Show(
            ("重命名会话已完成。`n`n" + (Get-ShortText -Text $output -MaxLength 700)),
            "重命名会话完成",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return $true
    }

    [System.Windows.Forms.MessageBox]::Show(
        ("重命名会话失败，退出码：$($result.ExitCode)`n`n" + (Get-ShortText -Text $output -MaxLength 700)),
        "重命名会话失败",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return $false
}

function Update-SessionsMenu {
    param(
        [Parameter(Mandatory = $true)]$MenuItem,
        [int]$Limit = 10,
        [string]$Source = "cli",
        [string]$FallbackCommand = "hermes sessions browse"
    )

    $MenuItem.DropDownItems.Clear()

    $loadingItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $loadingItem.Text = "正在加载会话..."
    $loadingItem.Enabled = $false
    [void]$MenuItem.DropDownItems.Add($loadingItem)
    [System.Windows.Forms.Application]::DoEvents()

    $result = Get-HermesSessions -Limit $Limit -Source $Source
    $MenuItem.DropDownItems.Clear()

    if ($result.ExitCode -ne 0) {
        $errorItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $errorItem.Text = "加载会话失败"
        $errorItem.Enabled = $false
        $errorItem.ToolTipText = Get-ShortText -Text $result.Output -MaxLength 500
        [void]$MenuItem.DropDownItems.Add($errorItem)
    }
    elseif ($result.Sessions.Count -eq 0) {
        $emptyItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $emptyItem.Text = "没有会话"
        $emptyItem.Enabled = $false
        [void]$MenuItem.DropDownItems.Add($emptyItem)
    }
    else {
        foreach ($session in @($result.Sessions)) {
            $localSession = $session
            $localMenuItem = $MenuItem
            $localLimit = $Limit
            $localSource = $Source
            $localFallbackCommand = $FallbackCommand

            $sessionItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $sessionItem.Text = [string]$localSession.Label
            $sessionItem.ToolTipText = [string]$localSession.ToolTip

            $resumeItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $resumeItem.Text = "启动"
            $resumeItem.Add_Click({
                try {
                    $sessionLiteral = ConvertTo-BashSingleQuotedLiteral -Value $localSession.Id
                    Start-HermesWslTerminal -DisplayName ("恢复会话 - " + $localSession.Id) -Command "hermes --resume $sessionLiteral"
                }
                catch {
                    Show-HermesErrorDialog -Title "启动会话失败" -ErrorRecord $_
                }
            }.GetNewClosure())
            [void]$sessionItem.DropDownItems.Add($resumeItem)

            $renameItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $renameItem.Text = "重命名会话"
            $renameItem.Add_Click({
                try {
                    if (Rename-HermesSession -Session $localSession) {
                        Update-SessionsMenu -MenuItem $localMenuItem -Limit $localLimit -Source $localSource -FallbackCommand $localFallbackCommand
                    }
                }
                catch {
                    Show-HermesErrorDialog -Title "重命名会话失败" -ErrorRecord $_
                }
            }.GetNewClosure())
            [void]$sessionItem.DropDownItems.Add($renameItem)

            $deleteItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $deleteItem.Text = "删除会话"
            $deleteItem.Add_Click({
                try {
                    if (Remove-HermesSession -Session $localSession) {
                        Update-SessionsMenu -MenuItem $localMenuItem -Limit $localLimit -Source $localSource -FallbackCommand $localFallbackCommand
                    }
                }
                catch {
                    Show-HermesErrorDialog -Title "删除会话失败" -ErrorRecord $_
                }
            }.GetNewClosure())
            [void]$sessionItem.DropDownItems.Add($deleteItem)

            [void]$MenuItem.DropDownItems.Add($sessionItem)
        }
    }

    [void]$MenuItem.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $fallbackAction = [pscustomobject]@{
        label = "打开完整会话浏览器"
        command = $FallbackCommand
        requiresInput = $false
    }
    [void]$MenuItem.DropDownItems.Add((New-HermesMenuItem -Action $fallbackAction))
}

function New-SessionsMenuItem {
    param(
        [Parameter(Mandatory = $true)]$Action
    )

    $sessionsItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $sessionsItem.Text = [string](Get-HermesActionValue -Action $Action -Name "label")

    $limit = Get-HermesActionValue -Action $Action -Name "limit"
    if ($null -eq $limit -or [int]$limit -lt 1) {
        $limit = 10
    }
    else {
        $limit = [int]$limit
    }

    $source = [string](Get-HermesActionValue -Action $Action -Name "source")
    if ([string]::IsNullOrWhiteSpace($source)) {
        $source = "cli"
    }

    $fallbackCommand = [string](Get-HermesActionValue -Action $Action -Name "fallbackCommand")
    if ([string]::IsNullOrWhiteSpace($fallbackCommand)) {
        $fallbackCommand = "hermes sessions browse"
    }

    $localSessionsItem = $sessionsItem
    $localLimit = $limit
    $localSource = $source
    $localFallbackCommand = $fallbackCommand

    $placeholder = New-Object System.Windows.Forms.ToolStripMenuItem
    $placeholder.Text = "点击展开加载"
    $placeholder.Enabled = $false
    [void]$sessionsItem.DropDownItems.Add($placeholder)

    $sessionsItem.Add_DropDownOpening({
        try {
            Update-SessionsMenu -MenuItem $localSessionsItem -Limit $localLimit -Source $localSource -FallbackCommand $localFallbackCommand
        }
        catch {
            $localSessionsItem.DropDownItems.Clear()
            $errorItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $errorItem.Text = "加载会话失败"
            $errorItem.Enabled = $false
            $errorItem.ToolTipText = [string]$_
            [void]$localSessionsItem.DropDownItems.Add($errorItem)
        }
    }.GetNewClosure())

    return $sessionsItem
}

function New-HermesMenuItem {
    param(
        [Parameter(Mandatory = $true)]$Action,
        [scriptblock]$AfterAction
    )

    if (Test-HermesActionGroup -Action $Action) {
        if ([string](Get-HermesActionValue -Action $Action -Name "module") -eq "gateway") {
            return New-GatewayMenuItem -Action $Action
        }
        if ([string](Get-HermesActionValue -Action $Action -Name "module") -eq "sessions") {
            return New-SessionsMenuItem -Action $Action
        }

        $groupItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $groupItem.Text = [string](Get-HermesActionValue -Action $Action -Name "label")

        foreach ($child in @(Get-HermesActionValue -Action $Action -Name "children")) {
            [void]$groupItem.DropDownItems.Add((New-HermesMenuItem -Action $child))
        }

        return $groupItem
    }

    $localAction = $Action
    $item = New-Object System.Windows.Forms.ToolStripMenuItem
    $item.Text = [string](Get-HermesActionValue -Action $localAction -Name "label")
    $localAfterAction = $AfterAction
    $item.Add_Click({
        try {
            Start-HermesAction -Action $localAction
            if ($null -ne $localAfterAction) {
                & $localAfterAction
            }
        }
        catch {
            Show-HermesErrorDialog -Title "Hermes Helper 执行失败" -ErrorRecord $_
        }
    }.GetNewClosure())

    return $item
}

function New-GatewayMenuItem {
    param(
        [Parameter(Mandatory = $true)]$Action
    )

    $gatewayItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $gatewayItem.Text = [string](Get-HermesActionValue -Action $Action -Name "label")

    $statusItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $statusItem.Text = "状态: 点击展开刷新"
    $statusItem.Enabled = $false
    [void]$gatewayItem.DropDownItems.Add($statusItem)
    [void]$gatewayItem.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $actionItems = @{}
    $localStatusItem = $statusItem
    $localActionItems = $actionItems
    $refreshGatewayMenu = {
        if ($null -ne $localStatusItem -and $null -ne $localActionItems) {
            Update-GatewayMenu -StatusItem $localStatusItem -ActionItems $localActionItems
        }
    }.GetNewClosure()

    foreach ($child in @(Get-HermesActionValue -Action $Action -Name "children")) {
        $childItem = New-HermesMenuItem -Action $child -AfterAction $refreshGatewayMenu
        $key = [string](Get-HermesActionValue -Action $child -Name "gatewayAction")

        if ([string]::IsNullOrWhiteSpace($key)) {
            if ([string](Get-HermesActionValue -Action $child -Name "command") -match "hermes\s+gateway\s+(\S+)") {
                $key = $Matches[1]
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $actionItems[$key] = $childItem
        }

        [void]$gatewayItem.DropDownItems.Add($childItem)
    }

    $gatewayItem.Add_DropDownOpening({
        Update-GatewayMenu -StatusItem $localStatusItem -ActionItems $localActionItems
    }.GetNewClosure())

    return $gatewayItem
}

function Initialize-HermesTray {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $actions = Load-HermesActions
    $health = Test-HermesHealth

    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Text = $health.ToolTip
    $notifyIcon.Visible = $true

    if ($health.State -eq "Ok") {
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
        $balloonIcon = [System.Windows.Forms.ToolTipIcon]::Info
    }
    else {
        $notifyIcon.Icon = [System.Drawing.SystemIcons]::Warning
        $balloonIcon = [System.Windows.Forms.ToolTipIcon]::Warning
    }

    $menu = New-Object System.Windows.Forms.ContextMenuStrip

    $statusItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $statusItem.Text = Get-ShortText -Text ("状态: " + $health.Summary) -MaxLength 80
    $statusItem.Enabled = $false
    [void]$menu.Items.Add($statusItem)
    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    foreach ($action in $actions) {
        [void]$menu.Items.Add((New-HermesMenuItem -Action $action))
    }

    [void]$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))

    $exitItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $exitItem.Text = "退出工具"
    $exitItem.Add_Click({
        $notifyIcon.Visible = $false
        $notifyIcon.Dispose()
        [System.Windows.Forms.Application]::Exit()
    }.GetNewClosure())
    [void]$menu.Items.Add($exitItem)

    $firstAction = Get-FirstExecutableAction -Actions $actions
    $notifyIcon.Add_DoubleClick({
        if ($firstAction) {
            Start-HermesAction -Action $firstAction
        }
    }.GetNewClosure())

    $notifyIcon.ContextMenuStrip = $menu
    $notifyIcon.ShowBalloonTip(2500, "Hermes Helper", $health.Details, $balloonIcon)

    [System.Windows.Forms.Application]::Run()
}

function Invoke-SelfTest {
    $actions = Load-HermesActions
    $health = Test-HermesHealth
    $gatewayStatus = Get-GatewayStatus
    $sessionsResult = Get-HermesSessions -Limit 10 -Source "cli"

    Write-Host "Actions: $((Get-HermesExecutableActionCount -Actions $actions)) executable, $($actions.Count) top-level"
    Write-Host "Health:  $($health.State) - $($health.Summary)"
    Write-Host "Gateway: $($gatewayStatus.State) - $($gatewayStatus.Label)"
    Write-Host "Sessions: $($sessionsResult.Sessions.Count) parsed"

    if ($health.State -ne "Ok") {
        Write-Host $health.Details
        exit 1
    }
}

if ($SelfTest) {
    Invoke-SelfTest
    exit 0
}

Initialize-HermesTray
