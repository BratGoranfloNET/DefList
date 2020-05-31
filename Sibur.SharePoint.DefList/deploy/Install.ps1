if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin "Microsoft.SharePoint.PowerShell"
}

$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

# Чтение файла конфигурации
[xml]$configFile = Get-Content -Path "$PSScriptRoot\config.xml" -Encoding UTF8
$config = $configFile.Config

# Параметры 
$solutionFileName = $config.SolutionName
$wspLiteralPath = (Get-Location).Path + "\$solutionFileName"
$FarmScanFeature = $config.FarmScanFeature
$WebScanFeature = $config.WebScanFeature




# Удалить решение при наличии
$solution = Get-SPSolution -Identity $solutionFileName -ErrorAction:SilentlyContinue
if ($solution -ne $null) {
  .\Uninstall.ps1   
}

# Загрузить решение
Add-SPSolution -LiteralPath $wspLiteralPath | Out-Null
Write-Host "Solution $solutionFileName added." 

# Получить решение
$solution = Get-SPSolution -Identity $solutionFileName

if ($solution -ne $null) {
    # Установка решения
    Write-Host "Installing Sibur.SharePoint.DefList.wsp"
    Install-SPSolution -Identity $solutionFileName -GACDeployment -Force    
    Write-Host "Started solution installation..." 
    $deployed = $False
    while ($deployed -eq $False) {
        Write-Host " > Install in progress..."
        Start-Sleep -s 10
        $solution = Get-SPSolution -Identity $solutionFileName
        if ($solution.Deployed -eq $True -And $solution.JobExists -eq $False) {
            $deployed = $True
        }

    }
    Write-Host "Solution $solutionFileName installed." 
    
 


}