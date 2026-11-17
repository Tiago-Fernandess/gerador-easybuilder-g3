$root = Split-Path $PSScriptRoot -Parent
$target = Join-Path $root 'Gerador_EasyBuilder.bat'
$icon = Join-Path $PSScriptRoot 'Button-G3-cor-padrão-05.ico'
$desktop = [Environment]::GetFolderPath('Desktop')
if ([string]::IsNullOrWhiteSpace($desktop) -or -not (Test-Path -LiteralPath $desktop)) {
    $desktop = $root
}

$shortcutPath = Join-Path $desktop 'Gerador IHM 2.0.lnk'
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $env:ComSpec
$shortcut.Arguments = '/c ""{0}""' -f $target
$shortcut.WorkingDirectory = $root
if (Test-Path -LiteralPath $icon) {
    $shortcut.IconLocation = "$icon,0"
}
$shortcut.Description = 'Gerador IHM 2.0'
$shortcut.Save()
