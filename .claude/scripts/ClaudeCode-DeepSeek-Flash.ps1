param(
    [string]$ProjectPath,
    [switch]$ResetApiKey,
    [switch]$NoLaunch,
    [string]$ApiKey,
    [switch]$SaveApiKey,
    [switch]$UseSavedApiKey,
    [switch]$OnlyResetApiKey
)

$ErrorActionPreference = "Stop"
$env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")

$rootDir = (Get-Item $PSScriptRoot).Parent.Parent.FullName
$configDir = Join-Path $rootDir ".claude-code-deepseek"
$keyFile = Join-Path $configDir "deepseek_api_key.dat"

function U {
    param([string]$Text)
    return [regex]::Unescape($Text)
}

function Test-CommandExists {
    param([string]$Name)
    return $null -ne (Get-Command $Name -ErrorAction SilentlyContinue)
}

function Convert-SecureStringToPlainText {
    param([Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-SavedApiKey {
    if (-not (Test-Path -LiteralPath $keyFile)) {
        return $null
    }

    try {
        $encrypted = Get-Content -LiteralPath $keyFile -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($encrypted)) {
            return $null
        }

        $secure = ConvertTo-SecureString -String $encrypted
        return Convert-SecureStringToPlainText -SecureString $secure
    }
    catch {
        Write-Warning (U '\u5df2\u4fdd\u5b58\u7684 API Key \u8bfb\u53d6\u5931\u8d25\uff0c\u5c06\u6539\u4e3a\u624b\u52a8\u8f93\u5165\u3002')
        return $null
    }
}

function Save-ApiKey {
    param([string]$ApiKey)

    if (-not (Test-Path -LiteralPath $configDir)) {
        New-Item -ItemType Directory -Path $configDir -Force | Out-Null
    }

    $secure = ConvertTo-SecureString -String $ApiKey -AsPlainText -Force
    $encrypted = ConvertFrom-SecureString -SecureString $secure
    Set-Content -LiteralPath $keyFile -Value $encrypted -Encoding ASCII
}

function Remove-SavedApiKey {
    if (Test-Path -LiteralPath $keyFile) {
        Remove-Item -LiteralPath $keyFile -Force
    }
}

function Get-ClaudeBinaryPath {
    $candidates = @(
        "C:\Users\$env:USERNAME\AppData\Roaming\npm\node_modules\@anthropic-ai\claude-code\node_modules\@anthropic-ai\claude-code-win32-x64\claude.exe",
        "C:\Users\$env:USERNAME\AppData\Roaming\npm\node_modules\@anthropic-ai\claude-code-win32-x64\claude.exe",
        "C:\Users\$env:USERNAME\AppData\Roaming\npm\node_modules\@anthropic-ai\claude-code\bin\claude.exe"
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            $bytes = [System.IO.File]::ReadAllBytes($candidate)
            $headerText = [System.Text.Encoding]::ASCII.GetString($bytes[0..200])
            if ($headerText -notmatch 'Bun is a fast JavaScript') {
                return $candidate
            }
        }
    }
    return $null
}

$script:ClaudeBinaryPath = Get-ClaudeBinaryPath

if (-not (Test-CommandExists "claude") -and -not $script:ClaudeBinaryPath) {
    Write-Error (U '\u672a\u68c0\u6d4b\u5230 claude \u547d\u4ee4\uff0c\u8bf7\u5148\u786e\u8ba4 Claude Code \u5df2\u6b63\u786e\u5b89\u88c5\u3002')
}

if (-not (Test-CommandExists "git")) {
    Write-Warning (U '\u672a\u68c0\u6d4b\u5230 git \u547d\u4ee4\uff0c\u90e8\u5206\u4ed3\u5e93\u529f\u80fd\u53ef\u80fd\u53d7\u9650\u3002')
}

if ($ResetApiKey) {
    Remove-SavedApiKey
    Write-Host (U '\u5df2\u6e05\u9664\u672c\u673a\u4fdd\u5b58\u7684 DeepSeek API Key\u3002') -ForegroundColor Yellow
    if ($OnlyResetApiKey) {
        exit 0
    }
}

if ([string]::IsNullOrWhiteSpace($ProjectPath)) {
    $defaultPath = $rootDir
    $inputPath = Read-Host ((U '\u8bf7\u8f93\u5165\u9879\u76ee\u76ee\u5f55\uff08\u76f4\u63a5\u56de\u8f66\u4f7f\u7528\u5f53\u524d\u76ee\u5f55\uff1a') + $defaultPath + (U '\uff09'))
    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        $ProjectPath = $defaultPath
    }
    else {
        $ProjectPath = $inputPath.Trim([char]34)
    }
}

if (-not (Test-Path -LiteralPath $ProjectPath)) {
    Write-Error ((U '\u9879\u76ee\u76ee\u5f55\u4e0d\u5b58\u5728\uff1a') + $ProjectPath)
}

$apiKey = $ApiKey
$savedApiKey = Get-SavedApiKey

if ([string]::IsNullOrWhiteSpace($apiKey) -and -not [string]::IsNullOrWhiteSpace($savedApiKey)) {
    if ($UseSavedApiKey) {
        $useSavedKeyAnswer = "y"
    }
    else {
        $useSavedKeyAnswer = Read-Host (U '\u68c0\u6d4b\u5230\u5df2\u4fdd\u5b58\u7684 DeepSeek API Key\uff0c\u662f\u5426\u76f4\u63a5\u4f7f\u7528\uff1f(Y/n)')
    }

    if ([string]::IsNullOrWhiteSpace($useSavedKeyAnswer) -or $useSavedKeyAnswer -imatch '^(y|yes)$') {
        $apiKey = $savedApiKey
        Write-Host (U '\u5df2\u4f7f\u7528\u672c\u673a\u4fdd\u5b58\u7684 API Key\u3002') -ForegroundColor Green
    }
}

if ([string]::IsNullOrWhiteSpace($apiKey)) {
    $apiKey = Read-Host (U '\u8bf7\u8f93\u5165 DeepSeek API Key')
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        Write-Error (U 'API Key \u4e0d\u80fd\u4e3a\u7a7a\u3002')
    }

    $saveApiKeyAnswer = Read-Host (U '\u662f\u5426\u4fdd\u5b58\u5230\u672c\u673a\uff0c\u4f9b\u4e0b\u6b21\u514d\u8f93\u4f7f\u7528\uff1f(Y/n)')
    if ([string]::IsNullOrWhiteSpace($saveApiKeyAnswer) -or $saveApiKeyAnswer -imatch '^(y|yes)$') {
        Save-ApiKey -ApiKey $apiKey
        Write-Host (U 'API Key \u5df2\u52a0\u5bc6\u4fdd\u5b58\u5230\u811a\u672c\u540c\u76ee\u5f55\u7684\u672c\u5730\u914d\u7f6e\u4e2d\u3002') -ForegroundColor Green
    }
}
elseif ($SaveApiKey) {
    Save-ApiKey -ApiKey $apiKey
    Write-Host (U 'API Key \u5df2\u6309\u53c2\u6570\u8981\u6c42\u52a0\u5bc6\u4fdd\u5b58\u5230\u811a\u672c\u540c\u76ee\u5f55\u7684\u672c\u5730\u914d\u7f6e\u4e2d\u3002') -ForegroundColor Green
}

$env:DEEPSEEK_API_KEY = $apiKey
$env:ANTHROPIC_BASE_URL = "https://api.deepseek.com/anthropic"
$env:ANTHROPIC_AUTH_TOKEN = $apiKey
$env:ANTHROPIC_MODEL = "deepseek-v4-flash"
$env:ANTHROPIC_DEFAULT_OPUS_MODEL = "deepseek-v4-flash"
$env:ANTHROPIC_DEFAULT_SONNET_MODEL = "deepseek-v4-flash"
$env:ANTHROPIC_DEFAULT_HAIKU_MODEL = "deepseek-v4-flash"
$env:CLAUDE_CODE_SUBAGENT_MODEL = "deepseek-v4-flash"
$env:CLAUDE_CODE_EFFORT_LEVEL = "max"
$env:API_TIMEOUT_MS = "600000"

Set-Location -LiteralPath $ProjectPath

Write-Host ""
Write-Host (U 'Claude Code \u5373\u5c06\u542f\u52a8') -ForegroundColor Green
Write-Host ((U '\u9879\u76ee\u76ee\u5f55\uff1a') + $ProjectPath)
Write-Host (U '\u5f53\u524d\u6a21\u578b\uff1adeepseek-v4-flash')
Write-Host (U '\u63a5\u53e3\u5730\u5740\uff1ahttps://api.deepseek.com/anthropic')
Write-Host ""

if ($NoLaunch) {
    Write-Host (U '\u5df2\u5b8c\u6210\u73af\u5883\u51c6\u5907\uff0c\u6309\u8981\u6c42\u4e0d\u542f\u52a8 Claude Code\u3002') -ForegroundColor Yellow
    exit 0
}

if ($script:ClaudeBinaryPath) {
    & $script:ClaudeBinaryPath
} else {
    claude
}
