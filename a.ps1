<#
.SYNOPSIS
  Reliable replacement for "ollama pull" that cannot silently accept a corrupt download.

.DESCRIPTION
  Accepted references:
    hf.co/owner/repo[:TAG]          Hugging Face GGUF (also huggingface.co/... or full URL)
    owner/repo[:TAG]                tried on the Ollama registry first; if not found there, treated as a Hugging Face repo
    name[:tag]                      Ollama registry, native pull with retries

  Hugging Face path:
    1. Lists the repo through the HF API and finds the GGUF for TAG (single file or .partNofM).
       No TAG: Q4_K_M if present, otherwise the first GGUF (same rule Ollama uses).
    2. Reads the publisher's SHA-256 (LFS oid) for each file from HF itself.
    3. Downloads with curl.exe: resumable, retrying, stall-detecting.
    4. Hashes the finished file; mismatch -> delete and re-download from zero. Never imported unverified.
    5. Imports it with "ollama create" under the name "ollama pull" would have used.

.NOTES
  Windows PowerShell 5.1 compatible. Private/gated repos: set $env:HF_TOKEN first.
  Run with:  powershell -ExecutionPolicy Bypass -File .\Pull-Model.ps1 <model>
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Model,
    [string]$OllamaDir = 'F:\backup\ollama',
    [string]$ModelsDir,
    [string]$WorkDir,
    [string]$Name,
    [string]$File,
    [string]$ModelfileExtra,
    [int]$MaxAttempts = 4,
    [switch]$KeepGguf,
    [switch]$SkipHash,
    [switch]$NoStartServer
)

$ErrorActionPreference = 'Stop'
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
} catch { }

if (-not $ModelsDir) { $ModelsDir = Join-Path $OllamaDir 'models' }
if (-not $WorkDir)   { $WorkDir   = Join-Path $OllamaDir 'import' }

function Say([string]$Text) {
    Write-Host ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $Text)
}

function Get-OllamaExe {
    $p = Join-Path $OllamaDir 'ollama.exe'
    if (Test-Path -LiteralPath $p) { return $p }
    $c = Get-Command ollama.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    throw "ollama.exe not found in '$OllamaDir' or on PATH."
}

function Get-ApiBase {
    $h = $env:OLLAMA_HOST
    if (-not $h) { return 'http://127.0.0.1:11434' }
    $h = $h.Trim()
    if ($h -notmatch '^https?://') { $h = 'http://' + $h }
    $h = $h -replace '//0\.0\.0\.0', '//127.0.0.1'
    if ($h -notmatch ':\d+/?$') { $h = $h.TrimEnd('/') + ':11434' }
    return $h.TrimEnd('/')
}

function Test-Server([string]$Base) {
    try {
        Invoke-RestMethod -Uri "$Base/api/version" -TimeoutSec 3 | Out-Null
        return $true
    } catch {
        return $false
    }
}

function Confirm-Server([string]$Exe) {
    $base = Get-ApiBase
    if (Test-Server $base) {
        Say "Ollama server already running at $base (it keeps whatever models folder it was started with)."
        return
    }
    if ($NoStartServer) { throw "Ollama server is not running at $base and -NoStartServer was given." }
    Say "Starting ollama serve with OLLAMA_MODELS=$ModelsDir"
    Start-Process -FilePath $Exe -ArgumentList 'serve' -WindowStyle Hidden | Out-Null
    for ($i = 0; $i -lt 60; $i++) {
        Start-Sleep -Seconds 1
        if (Test-Server $base) { Say 'Server is up.'; return }
    }
    throw "Ollama server did not come up within 60 seconds."
}

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Get-Leaf([string]$Path) {
    return $Path.Split('/')[-1]
}

function Get-FreeBytes([string]$Dir) {
    try {
        $full = [IO.Path]::GetFullPath($Dir)
        $root = [IO.Path]::GetPathRoot($full)
        $di = New-Object IO.DriveInfo($root)
        return [pscustomobject]@{ Root = $root.ToUpperInvariant(); Free = [long]$di.AvailableFreeSpace }
    } catch {
        return $null
    }
}

function Test-CurlRetryAll {
    try {
        $t = (& curl.exe --help all 2>&1 | Out-String)
        return ($t -match '--retry-all-errors')
    } catch {
        return $false
    }
}

function Invoke-CurlResume([string]$Url, [string]$Dest, [long]$Size) {
    $stalls = 0
    while ($true) {
        $have = 0L
        if (Test-Path -LiteralPath $Dest) { $have = (Get-Item -LiteralPath $Dest).Length }
        if ($have -eq $Size) { return }
        if ($have -gt $Size) {
            Remove-Item -LiteralPath $Dest -Force
            $have = 0L
        }
        $curlArgs = @('-L', '--fail', '--progress-bar', '--connect-timeout', '30',
                      '--speed-limit', '10240', '--speed-time', '60',
                      '--retry', '10', '--retry-delay', '5', '-C', '-', '-o', $Dest)
        if ($script:CurlRetryAll) { $curlArgs += '--retry-all-errors' }
        if ($env:HF_TOKEN) { $curlArgs += @('-H', "Authorization: Bearer $($env:HF_TOKEN)") }
        $curlArgs += $Url
        & curl.exe @curlArgs | Out-Host
        $code = $LASTEXITCODE
        $now = 0L
        if (Test-Path -LiteralPath $Dest) { $now = (Get-Item -LiteralPath $Dest).Length }
        if ($now -eq $Size) { return }
        if ($now -gt $have) { $stalls = 0 } else { $stalls++ }
        if ($stalls -ge 8) { throw "No download progress after 8 consecutive resume attempts (last curl exit code $code)." }
        Say ("Transfer interrupted at {0:N0} of {1:N0} bytes (curl exit {2}); resuming in 5 s..." -f $now, $Size, $code)
        Start-Sleep -Seconds 5
    }
}

function Get-VerifiedFile([string]$Repo, $Entry) {
    $leaf = Get-Leaf $Entry.path
    $dest = Join-Path $WorkDir $leaf
    $size = [long]$Entry.size
    $sha = $null
    if ($Entry.lfs) {
        $size = [long]$Entry.lfs.size
        $sha = ([string]$Entry.lfs.oid).ToLowerInvariant()
    }
    $escaped = ($Entry.path.Split('/') | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
    $url = "https://huggingface.co/$Repo/resolve/main/$escaped"
    $lastBad = $null
    for ($a = 1; $a -le $MaxAttempts; $a++) {
        Say "Downloading $leaf (attempt $a of $MaxAttempts)"
        Invoke-CurlResume $url $dest $size
        if ($SkipHash -or -not $sha) {
            Write-Warning "SHA-256 NOT verified for $leaf (size matches only)."
            return $dest
        }
        Say "Hashing $leaf (can take a minute or two for large files)..."
        $got = Get-Sha256 $dest
        if ($got -eq $sha) {
            Say "SHA-256 verified: $got"
            return $dest
        }
        Write-Warning "SHA-256 mismatch for $leaf`n  published: $sha`n  got:       $got"
        Remove-Item -LiteralPath $dest -Force
        if ($lastBad -and $lastBad -eq $got) {
            throw ("Two clean re-downloads produced the SAME wrong hash ($got). The bytes being served to this machine differ from what Hugging Face publishes " +
                   "(proxy/antivirus/VPN rewriting traffic, or the repo changed mid-run). Fix that, or re-run with -SkipHash if you accept the risk.")
        }
        $lastBad = $got
    }
    throw "Could not obtain a hash-verified copy of $leaf after $MaxAttempts attempts."
}

function Merge-Parts([string[]]$Paths, [string]$Out) {
    $buf = New-Object 'byte[]' 16777216
    $o = [IO.File]::Open($Out, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        foreach ($p in $Paths) {
            $i = [IO.File]::OpenRead($p)
            try {
                while (($n = $i.Read($buf, 0, $buf.Length)) -gt 0) { $o.Write($buf, 0, $n) }
            } finally { $i.Dispose() }
        }
    } finally { $o.Dispose() }
}

function Get-HfTree([string]$Repo) {
    $hdr = @{}
    if ($env:HF_TOKEN) { $hdr['Authorization'] = "Bearer $($env:HF_TOKEN)" }
    $url = "https://huggingface.co/api/models/$Repo/tree/main?recursive=true"
    for ($i = 1; $i -le 5; $i++) {
        try {
            return @(Invoke-RestMethod -Uri $url -Headers $hdr -TimeoutSec 60)
        } catch {
            $msg = $_.Exception.Message
            if ($i -eq 5 -or $msg -match '40[0-9]') {
                throw "Could not list '$Repo' on Hugging Face: $msg"
            }
            Start-Sleep -Seconds 3
        }
    }
}

function Show-GgufAlternatives([string]$Repo) {
    try {
        $leaf = $Repo.Split('/')[-1]
        $q = [Uri]::EscapeDataString(($leaf -replace '(?i)-?gguf$', ''))
        $res = @(Invoke-RestMethod -Uri "https://huggingface.co/api/models?search=$q&filter=gguf&limit=10" -TimeoutSec 30)
        if ($res.Count -gt 0) {
            Write-Host 'GGUF repositories with a similar name (use one of these instead):'
            $res | ForEach-Object { Write-Host ("  hf.co/{0}" -f $_.id) }
        }
    } catch { }
}

function Select-ModelFiles($Tree, [string]$Tag, [string]$Explicit, [string]$Repo) {
    $all = @($Tree | Where-Object {
        $_.type -eq 'file' -and $_.path -match '(?i)\.gguf(\.part\d+of\d+)?$' -and (Get-Leaf $_.path) -notmatch '(?i)mmproj'
    })
    if ($all.Count -eq 0) {
        Show-GgufAlternatives $Repo
        throw "Repository '$Repo' contains no .gguf files (it is probably safetensors/PyTorch weights, which Ollama cannot pull directly)."
    }

    if ($Explicit) {
        $m = @($all | Where-Object { (Get-Leaf $_.path) -ieq $Explicit -or $_.path -ieq $Explicit })
        if ($m.Count -eq 0) { throw "-File '$Explicit' not found. Files: " + (($all | ForEach-Object { $_.path }) -join ', ') }
    } else {
        $useTag = $Tag
        if (-not $useTag) { $useTag = 'Q4_K_M' }
        $rx = '(?i)(^|[._\-/])' + [regex]::Escape($useTag) + '(-\d{5}-of-\d{5})?\.gguf(\.part\d+of\d+)?$'
        $m = @($all | Where-Object { $_.path -match $rx })
        if ($m.Count -eq 0 -and -not $Tag) {
            $f = $all | Sort-Object path | Select-Object -First 1
            Say "No Q4_K_M in this repo; using first GGUF like Ollama does: $($f.path)"
            if ($f.path -match '^(.*)\.part\d+of\d+$') {
                $pre = $Matches[1] + '.part'
                $m = @($all | Where-Object { $_.path.StartsWith($pre) })
            } else {
                $m = @($f)
            }
        }
        if ($m.Count -eq 0) {
            throw "No GGUF matching tag '$useTag'. Available: " + (($all | ForEach-Object { Get-Leaf $_.path }) -join ', ') + "  (use -File <name> to pick one)"
        }
    }

    if (@($m | Where-Object { (Get-Leaf $_.path) -match '-\d{5}-of-\d{5}\.gguf$' }).Count -gt 0) {
        throw 'This quant is a split GGUF (-0000N-of-0000M.gguf shards). Ollama cannot import shards directly; merge them with llama.cpp llama-gguf-split --merge first.'
    }

    $parts = @($m | Where-Object { $_.path -match '\.part\d+of\d+$' })
    $whole = @($m | Where-Object { $_.path -notmatch '\.part\d+of\d+$' })
    if ($whole.Count -eq 1 -and $parts.Count -eq 0) { return ,@($whole[0]) }
    if ($whole.Count -eq 0 -and $parts.Count -gt 0) {
        $sorted = @($parts | Sort-Object { [int]([regex]::Match($_.path, '\.part(\d+)of\d+$').Groups[1].Value) })
        $expect = [int]([regex]::Match($sorted[0].path, '\.part\d+of(\d+)$').Groups[1].Value)
        if ($sorted.Count -ne $expect) { throw "Expected $expect parts but found $($sorted.Count)." }
        return $sorted
    }
    throw ('Ambiguous match, use -File to choose: ' + (($m | ForEach-Object { $_.path }) -join ', '))
}

function Invoke-NativePull([string]$Exe, [string]$Ref, [int]$Attempts) {
    for ($i = 1; $i -le $Attempts; $i++) {
        & $Exe pull $Ref
        if ($LASTEXITCODE -eq 0) { return $true }
        Write-Warning "Pull failed (attempt $i of $Attempts)."
        if ($i -lt $Attempts) { Start-Sleep -Seconds 5 }
    }
    return $false
}

# ---------------------------------------------------------------- main

$exe = Get-OllamaExe
New-Item -ItemType Directory -Force -Path $ModelsDir | Out-Null
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null
$env:OLLAMA_MODELS = $ModelsDir

# Classify the reference
$isHf = $false
$bare = $false
$repo = $null
$tag = $null
$rest = $null
$r = $Model.Trim() -replace '^https?://', ''
if ($r -match '^(hf\.co|huggingface\.co)/(.+)$') {
    $isHf = $true
    $rest = $Matches[2]
} elseif ($r -match '^[^/:]+/[^/:]+(:[^/]+)?$') {
    $bare = $true
    $rest = $r
}
if ($rest) {
    if ($rest -match '^([^/]+/[^/:]+)(?::([^/]+))?(/.*)?$') {
        $repo = $Matches[1]
        $tag = $Matches[2]
    } else {
        throw "Cannot parse reference '$Model'."
    }
}

Confirm-Server $exe

if (-not $isHf -and -not $bare) {
    Say "Ollama registry model; native pull with retries: $Model"
    if (Invoke-NativePull $exe $Model $MaxAttempts) { Say 'Pull completed.'; exit 0 }
    throw "ollama pull failed $MaxAttempts times for $Model."
}

if ($bare) {
    Say "Trying the Ollama registry first: $Model"
    if (Invoke-NativePull $exe $Model 1) { Say 'Pull completed.'; exit 0 }
    Say "Not usable from the Ollama registry; trying Hugging Face repo '$repo'..."
}

if (-not (Get-Command curl.exe -ErrorAction SilentlyContinue)) {
    throw 'curl.exe not found (it ships with Windows 10 1803+ / Windows 11).'
}
$script:CurlRetryAll = Test-CurlRetryAll

$tagLabel = 'default Q4_K_M'
if ($tag) { $tagLabel = $tag }
Say "Resolving $repo ($tagLabel) on Hugging Face..."
try {
    $tree = Get-HfTree $repo
} catch {
    if ($bare -and $MaxAttempts -gt 1) {
        Say 'Not on Hugging Face either; retrying the Ollama registry.'
        if (Invoke-NativePull $exe $Model ($MaxAttempts - 1)) { Say 'Pull completed.'; exit 0 }
    }
    throw
}
$entries = Select-ModelFiles $tree $tag $File $repo
$total = 0L
foreach ($e in $entries) {
    if ($e.lfs) { $total += [long]$e.lfs.size } else { $total += [long]$e.size }
}
Say ("Selected: {0}  ({1:N2} GB)" -f (($entries | ForEach-Object { Get-Leaf $_.path }) -join ' + '), ($total / 1GB))

# Disk space
$isParts = ($entries.Count -gt 1)
$work = Get-FreeBytes $WorkDir
$mdl = Get-FreeBytes $ModelsDir
if ($work -and $mdl) {
    $mult = 1L
    if ($isParts) { $mult = 2L }
    if ($work.Root -eq $mdl.Root) {
        $need = ([Math]::Max($mult, 2L) * $total) + 2GB
        if ($work.Free -lt $need) { throw ("Not enough free space on {0}: need about {1:N1} GB, have {2:N1} GB." -f $work.Root, ($need / 1GB), ($work.Free / 1GB)) }
    } else {
        $needW = ($mult * $total) + 2GB
        $needM = $total + 2GB
        if ($work.Free -lt $needW) { throw ("Not enough free space on {0}: need about {1:N1} GB." -f $work.Root, ($needW / 1GB)) }
        if ($mdl.Free -lt $needM) { throw ("Not enough free space on {0}: need about {1:N1} GB." -f $mdl.Root, ($needM / 1GB)) }
    }
}

# Download + verify
$files = @()
foreach ($e in $entries) { $files += (Get-VerifiedFile $repo $e) }

if ($files.Count -gt 1) {
    $first = Get-Leaf $entries[0].path
    $final = Join-Path $WorkDir ($first -replace '\.part\d+of\d+$', '')
    Say "Joining $($files.Count) parts into $final"
    Merge-Parts $files $final
    if ((Get-Item -LiteralPath $final).Length -ne $total) { throw 'Merged file size does not match the sum of the parts.' }
    $files | ForEach-Object { Remove-Item -LiteralPath $_ -Force }
} else {
    $final = $files[0]
}

# Import into Ollama
if (-not $Name) {
    $label = $tag
    if (-not $label) { $label = 'latest' }
    $Name = "hf.co/${repo}:$label"
}
$mf = Join-Path $WorkDir ('Modelfile.' + ($Name -replace '[^A-Za-z0-9._-]', '_'))
$ggufPath = $final -replace '\\', '/'
$text = 'FROM "' + $ggufPath + '"' + "`n"
if ($ModelfileExtra) { $text += $ModelfileExtra + "`n" }
[IO.File]::WriteAllText($mf, $text, (New-Object Text.UTF8Encoding($false)))

Say "Importing into Ollama as '$Name' (hashes and copies the file into $ModelsDir)..."
& $exe create $Name -f $mf
if ($LASTEXITCODE -ne 0) {
    throw "ollama create failed (exit $LASTEXITCODE). The verified GGUF is kept at: $final  (re-run the same command to retry the import without re-downloading)."
}

Remove-Item -LiteralPath $mf -Force -ErrorAction SilentlyContinue
if (-not $KeepGguf) {
    Remove-Item -LiteralPath $final -Force -ErrorAction SilentlyContinue
    Say 'Removed the temporary GGUF.'
}
Say "Done. Run it with:  ollama run $Name"
