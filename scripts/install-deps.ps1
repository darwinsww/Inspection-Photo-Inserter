$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
$pipConfigPath = Join-Path $projectRoot ".pip\pip.ini"
$requirementsPath = Join-Path $projectRoot "requirements.txt"

if (-not (Test-Path $pipConfigPath)) {
    throw "Missing project pip config: $pipConfigPath"
}

if (-not (Test-Path $requirementsPath)) {
    throw "Missing requirements file: $requirementsPath"
}

$previousPipConfig = $env:PIP_CONFIG_FILE
$env:PIP_CONFIG_FILE = $pipConfigPath

try {
    pip install -r $requirementsPath
}
finally {
    if ([string]::IsNullOrWhiteSpace($previousPipConfig)) {
        Remove-Item Env:PIP_CONFIG_FILE -ErrorAction SilentlyContinue
    }
    else {
        $env:PIP_CONFIG_FILE = $previousPipConfig
    }
}
