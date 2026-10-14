# Loader. Keep this file ASCII-only: the one-liner downloads it with the system code page.
# Fetches run.ps1 as UTF-8 (GitHub first, then Gitee) and runs it. Must work on PowerShell 2.0.
# TLS 1.2 = 3072; PS2 rejects [Net.SecurityProtocolType]3072 because the name does not exist in .NET 3.5.
[Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
$w = New-Object Net.WebClient
$w.Encoding = [Text.Encoding]::UTF8
$s = $null
foreach ($u in 'https://hejun-coderlife.github.io/store-setup/run.ps1', 'https://gitee.com/hejun202609/store-setup/raw/master/run.ps1') {
    try { $s = $w.DownloadString($u + '?t=' + [DateTime]::Now.Ticks); break } catch { Write-Host "Download failed: $u" }
}
if ($s) { Invoke-Expression $s.TrimStart([char]0xFEFF) }
else { Write-Host 'Cannot reach GitHub or Gitee. Check the network and try again.' -ForegroundColor Red; Read-Host 'Press Enter to exit' }
