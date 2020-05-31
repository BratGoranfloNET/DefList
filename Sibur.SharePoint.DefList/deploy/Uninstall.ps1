if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Чтение файла конфигурации
[xml]$configFile = Get-Content -Path "$PSScriptRoot\config.xml" -Encoding UTF8
$config = $configFile.Config

# Параметры 
$solutionFileName = $config.SolutionName

# Получить решение
$solution = Get-SPSolution -Identity $solutionFileName -ErrorAction:SilentlyContinue

if ($solution -ne $null) {
    # Деинсталяция решения
    Write-Host "Uninstalling Sibur.SharePoint.VZL.wsp"
    Uninstall-SPSolution -Identity $solutionFileName -Confirm:$False -Language 0 -ErrorAction:SilentlyContinue
    Write-Host "Started solution retraction..." 
    $deployed = $solution.Deployed
    while ($deployed -eq $True) {
        Write-Host " > Uninstall in progress..."
        Start-Sleep -s 10
        $solution = Get-SPSolution -Identity $solutionFileName
        if ($solution.Deployed -eq $False -And $solution.JobExists -eq $False) {
            $deployed = $False
        }
    }

    # Удаление решения
    Remove-SPSolution -Identity $solutionFileName -Confirm:$False
    Write-Host "Solution $solutionFileName removed."
}
else {
    Write-Host "Solution not found." -ForegroundColor Red
}