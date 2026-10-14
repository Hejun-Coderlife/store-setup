# 门店装机 + 体检：按清单检查运行库、芯烨驱动、业务软件，缺什么补什么，最后把报告发到「门店电脑体检」飞书群。
# 门店用法见 README.md：一行命令下载 go.ps1，go.ps1 再下载本文件用 Invoke-Expression 运行（所以本文件不能靠 $PSScriptRoot）。
# 必须兼容 PowerShell 2.0（Win7）：别用 -in/-notin、ConvertFrom-Json、Get-CimInstance、Get-Content -Raw、[ordered]、
# 成员枚举（$list.Name）、Start-Process -Wait（会等子进程，芯烨驱动的向导不关就永远等下去）。
# 原则：只补装，不卸载、不删除、不改已有的打印机。
# 本机测试：$env:STORESETUP_TEST='1' 时不问问题、不等回车；$env:STORESETUP_HOOK / STORESETUP_STORE / STORESETUP_PAPER（58 或 80）可直接给群地址、店名、小票纸宽。

function Invoke-StoreSetup {
    $ErrorActionPreference = 'Stop'
    $Version = '2026.10.14'
    $Dir = 'C:\ProgramData\StoreSetup'
    $Cache = Join-Path $env:TEMP 'StoreSetup'
    $Test = $env:STORESETUP_TEST -eq '1'
    $UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36'
    $State = @{ NeedReboot = $false }

    # 要管理员权限；32 位 PowerShell 跑在 64 位系统上会被重定向注册表和 System32，也要换成 64 位的重开
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    $isWow = ([IntPtr]::Size -eq 4) -and (Test-Path "$env:SystemRoot\SysNative")
    if (-not $isAdmin -or $isWow) {
        # GitHub 打不开的门店用的是 Gitee 那行命令，这里也要能退到 Gitee
        $boot = "[Net.ServicePointManager]::SecurityProtocol=[Enum]::ToObject([Net.SecurityProtocolType],3072);`$w=New-Object Net.WebClient;try{`$s=`$w.DownloadString('https://hejun-coderlife.github.io/store-setup/go.ps1')}catch{`$s=`$w.DownloadString('https://gitee.com/hejun202609/store-setup/raw/master/go.ps1')};iex `$s"
        $exe = 'powershell.exe'
        if ($isWow) { $exe = "$env:SystemRoot\SysNative\WindowsPowerShell\v1.0\powershell.exe" }
        $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($boot))
        Write-Host '需要管理员权限，正在重新打开（弹出的窗口请点「是」）...' -ForegroundColor Yellow
        Start-Process $exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -EncodedCommand $enc" -Verb RunAs
        return 'relaunched'
    }

    foreach ($d in $Dir, $Cache, (Join-Path $Dir 'logs')) { if (-not (Test-Path $d)) { New-Item -ItemType Directory $d | Out-Null } }
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $LogFile = Join-Path $Dir "logs\$stamp.txt"

    function Write-Log($text, $color) {
        if ($color) { Write-Host $text -ForegroundColor $color } else { Write-Host $text }
        [IO.File]::AppendAllText($LogFile, (Get-Date -Format 'HH:mm:ss') + ' ' + $text + "`r`n", [Text.Encoding]::UTF8)
    }

    function Read-Text($path) {
        if (-not (Test-Path $path)) { return '' }
        return [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8).Trim()
    }

    function Get-RegValue($path, $name) {
        try { return (Get-ItemProperty $path -ErrorAction Stop).$name } catch { return $null }
    }

    function New-WebClient($referer) {
        $wc = New-Object Net.WebClient
        $wc.Headers.Add('User-Agent', $UA)
        if ($referer) { $wc.Headers.Add('Referer', $referer) }
        return $wc
    }

    function Get-Download($url, $file, $referer) {
        $path = Join-Path $Cache $file
        Write-Log "  下载 $file ..."
        (New-WebClient $referer).DownloadFile($url, $path)
        return $path
    }

    # 报告里的失败原因要短、要中文，飞书群里看得懂
    function Get-ShortError($err) {
        $e = $err.Exception
        while ($e.InnerException) { $e = $e.InnerException }
        $m = $e.Message
        if ($m -match 'actively refused|积极拒绝') { return '连不上服务器（可能是代理设置残留）' }
        if ($m -match 'Unable to connect|无法连接|could not be resolved|未能解析') { return '连不上服务器' }
        if ($m -match '\(403\)') { return '下载被拒绝(403)' }
        if ($m -match '\(404\)') { return '下载地址失效(404)' }
        if ($m -match 'timed out|超时') { return '下载超时' }
        if ($m.Length -gt 80) { $m = $m.Substring(0, 80) + '…' }
        return $m
    }

    # 启动后先取一次 Handle，不然 PS2/PS5 都可能拿不到 ExitCode
    function Invoke-Wait($path, $arguments, $timeoutSec) {
        if ($arguments) { $p = Start-Process -FilePath $path -ArgumentList $arguments -PassThru }
        else { $p = Start-Process -FilePath $path -PassThru }
        $null = $p.Handle
        if (-not $p.WaitForExit($timeoutSec * 1000)) {
            # 多半是弹了错误框在等人点；杀掉，后面的项目接着装
            try { $p.Kill() } catch { }
            foreach ($c in @(Get-WmiObject Win32_Process -Filter "ParentProcessId=$($p.Id)" -ErrorAction SilentlyContinue)) { try { Stop-Process -Id $c.ProcessId -Force } catch { } }
            throw ('等了 {0} 分钟还没装完（可能弹了错误框），已强制结束' -f [int]($timeoutSec / 60))
        }
        return $p.ExitCode
    }

    # 0 成功；3010/1641 成功但要重启；1638 已经装了同版或更新的
    function Install-Exe($url, $file, $arguments) {
        $path = Get-Download $url $file
        Write-Log "  安装 $file ..."
        $code = Invoke-Wait $path $arguments 600
        if ($code -eq 3010 -or $code -eq 1641) { $State.NeedReboot = $true; return }
        if ($code -eq 0 -or $code -eq 1638) { return }
        if ($code -eq 5100) { throw '系统不满足安装要求（Win7 需要 SP1）' }
        throw "安装程序返回 $code"
    }

    function Get-InstalledNames {
        $names = @()
        foreach ($k in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                       'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                       'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall') {
            if (-not (Test-Path $k)) { continue }
            foreach ($i in @(Get-ChildItem $k -ErrorAction SilentlyContinue)) {
                $n = $i.GetValue('DisplayName')
                if ($n) { $names += [string]$n }
            }
        }
        return $names
    }

    function Test-Installed($pattern) {
        foreach ($n in @(Get-InstalledNames)) { if ($n -match $pattern) { return $true } }
        return $false
    }

    # 业务软件不能只看卸载记录：淘宝闪购的安装包在没有 D 盘的电脑上照样写记录、建快捷方式、返回成功，文件却一个没装。
    # 所以从卸载命令里取出安装目录，确认里面真有程序
    function Test-AppFiles($pattern) {
        foreach ($k in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                       'HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall',
                       'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall') {
            if (-not (Test-Path $k)) { continue }
            foreach ($i in @(Get-ChildItem $k -ErrorAction SilentlyContinue)) {
                if (-not ([string]$i.GetValue('DisplayName') -match $pattern)) { continue }
                $u = [string]$i.GetValue('UninstallString')
                if ($u -match '^"([^"]+)"') { $u = $matches[1] }
                $dir = Split-Path $u
                if (-not $dir -or -not (Test-Path $dir)) { continue }
                foreach ($f in @(Get-ChildItem $dir -Filter *.exe -ErrorAction SilentlyContinue)) {
                    if ($f.Name -notmatch '^(Uninstall|elevate)') { return $true }
                }
            }
        }
        return $false
    }

    # 芯烨官网驱动页上的 exe 下载链接（页面直接写在 HTML 里）；中文文件名要转义才能下载
    function Get-XprinterUrl($page, $pattern) {
        $wc = New-WebClient
        $wc.Encoding = [Text.Encoding]::UTF8
        $html = $wc.DownloadString($page)
        $url = $null
        foreach ($m in [regex]::Matches($html, 'https?://[^"''\s<>]+\.exe')) { if ($m.Value -match $pattern) { $url = $m.Value; break } }
        if (-not $url) { throw '芯烨官网找不到驱动下载链接（可能改版了）' }
        $cut = $url.LastIndexOf('/') + 1
        $name = $url.Substring($cut)
        if ($name -notmatch '%') { $url = $url.Substring(0, $cut) + [Uri]::EscapeDataString($name) }
        Write-Log "  最新驱动：$([Uri]::UnescapeDataString($name))"
        return $url
    }

    # 芯烨的驱动都不是微软认证的，pnputil 会弹「你想安装这个设备软件吗」卡住 → 先把 .cat 的签名证书加进「受信任的发布者」
    function Import-DriverInf($inf) {
        $cat = @(Get-ChildItem (Split-Path $inf) -Filter *.cat)
        if ($cat.Count -eq 0) { throw "驱动目录里没有 .cat 签名文件：$(Split-Path $inf)" }
        $cert = (Get-AuthenticodeSignature $cat[0].FullName).SignerCertificate
        if (-not $cert) { throw '读不出驱动的签名证书' }
        Write-Log "  信任驱动发布者：$($cert.GetNameInfo('SimpleName', $false))"
        $st = New-Object Security.Cryptography.X509Certificates.X509Store('TrustedPublisher', 'LocalMachine')
        $st.Open('ReadWrite')
        try { $st.Add($cert) } finally { $st.Close() }

        Write-Log '  导入驱动 ...'
        $out = & "$env:SystemRoot\System32\pnputil.exe" -i -a $inf 2>&1
        Write-Log ('  pnputil: ' + (($out | Where-Object { "$_".Trim() }) -join ' / '))
    }

    # ---------- 系统信息 ----------
    $os = Get-WmiObject Win32_OperatingSystem
    $build = [int]$os.BuildNumber
    $is64 = Test-Path "$env:SystemRoot\SysWOW64"
    $bits = '32位'; if ($is64) { $bits = '64位' }
    $osText = '{0} {1} {2}，PowerShell {3}' -f $os.Caption.Trim(), $os.Version, $bits, $PSVersionTable.PSVersion

    Write-Log "门店装机体检 v$Version" Cyan
    Write-Log $osText

    # 先试一下网，报告里能一眼看出是不是网的问题。测不通也照样去下载（实测 HEAD 微软首页会超时误判）
    [Net.ServicePointManager]::SecurityProtocol = [Enum]::ToObject([Net.SecurityProtocolType], 3072)
    $State.Online = $false
    foreach ($u in 'https://www.baidu.com/', 'https://www.microsoft.com/') {
        try {
            $q = [Net.WebRequest]::Create($u); $q.Timeout = 10000
            $q.GetResponse().Close(); $State.Online = $true; break
        } catch { }
    }
    if (-not $State.Online) { Write-Log '这台电脑好像连不上网，缺的东西可能装不了' Red }

    # ---------- 店名和飞书群地址（先问完，后面装东西时人可以走开） ----------
    $store = $env:STORESETUP_STORE
    if (-not $store) { $store = Read-Text 'C:\ProgramData\LoginWatch\store.txt' }   # 外卖监控装过就用它的店名
    if (-not $store) { $store = Read-Text (Join-Path $Dir 'store.txt') }
    if (-not $store -and -not $Test) {
        $store = (Read-Host '请输入店名（直接回车用电脑名）').Trim()
        if ($store) { [IO.File]::WriteAllText((Join-Path $Dir 'store.txt'), $store, [Text.Encoding]::UTF8) }
    }
    if (-not $store) { $store = $env:COMPUTERNAME }

    # 「门店电脑体检」群的机器人地址直接写在程序里（2026-10-14 用户定：店员不用粘贴）。仓库公开，谁都能往群里发消息；
    # 被人乱发就删掉机器人重加一个，换掉这里的地址。本机 feishu.txt 有地址就优先用它（测试用）
    $hook = $env:STORESETUP_HOOK
    if (-not $hook) { $hook = Read-Text (Join-Path $Dir 'feishu.txt') }
    if (-not $hook) { $hook = 'https://open.feishu.cn/open-apis/bot/v2/hook/300fe773-a00d-48dd-9688-2c9c0865b11a' }   # 2026-10-14 换到企业「赫眉」下的群

    # ---------- 插着但没添加的 USB 小票机 ----------
    # 2026-10-14 实测：芯烨原装 XP-58IIH 不报型号（USBPRINT\UnknownPrinter），Windows 不会自己添加打印机，要程序来加。
    # 只管「插着、但没有任何打印机用这个 USB 端口」的；已经有打印机在用的不动。纸宽程序判断不了，问一次。
    # 设备 ID 结尾就是端口名（USBPRINT\UNKNOWNPRINTER\6&FA5842&0&USB001）；错误码 45 = 以前插过、现在没插
    $usbDevices = @(Get-WmiObject Win32_PnPEntity -Filter "PNPDeviceID LIKE 'USBPRINT%'" -ErrorAction SilentlyContinue | Where-Object { $_.ConfigManagerErrorCode -ne 45 })
    $usedPorts = @()
    foreach ($pr in @(Get-WmiObject Win32_Printer -ErrorAction SilentlyContinue)) { $usedPorts += [string]$pr.PortName }
    $freePorts = @()
    foreach ($e in $usbDevices) {
        if ($e.PNPDeviceID -match '(USB\d{3})$' -and $usedPorts -notcontains $matches[1] -and $freePorts -notcontains $matches[1]) { $freePorts += $matches[1] }
    }
    $paper = $env:STORESETUP_PAPER
    if ($freePorts.Count -and -not $paper -and -not $Test) {
        Write-Host "发现 $($freePorts.Count) 台插着、但还没添加的 USB 小票机。" -ForegroundColor Yellow
        while ($paper -ne '58' -and $paper -ne '80') {
            $a = (Read-Host '用的是哪种纸？输入 1 = 58 毫米（窄纸）  2 = 80 毫米（宽纸）').Trim()
            if ($a -eq '1') { $paper = '58' } elseif ($a -eq '2') { $paper = '80' }
        }
    }
    if (-not $paper) { $paper = '58' }

    # ---------- 清单 ----------
    # 每项：Name；Check 返回 $true 表示已经有了；Install 补装（失败就 throw）；Skip 返回跳过原因或 $null
    $items = @()

    # VC++ 2005~2013 已停更，链接是各自的最终版；按卸载列表里的名字判断
    $vcOld = @(
        @('2005', 'x86', 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x86.EXE', '/q:a /c:"msiexec /i vcredist.msi /qn /norestart"', 'Visual C\+\+ 2005 Redistributable$'),
        @('2005', 'x64', 'https://download.microsoft.com/download/8/B/4/8B42259F-5D70-43F4-AC2E-4B208FD8D66A/vcredist_x64.EXE', '/q:a /c:"msiexec /i vcredist.msi /qn /norestart"', 'Visual C\+\+ 2005 Redistributable \(x64\)'),
        @('2008', 'x86', 'https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x86.exe', '/q', 'Visual C\+\+ 2008 Redistributable - x86'),
        @('2008', 'x64', 'https://download.microsoft.com/download/5/D/8/5D8C65CB-C849-4025-8E95-C3966CAFD8AE/vcredist_x64.exe', '/q', 'Visual C\+\+ 2008 Redistributable - x64'),
        @('2010', 'x86', 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe', '/q /norestart', 'Visual C\+\+ 2010 +x86 Redistributable'),
        @('2010', 'x64', 'https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x64.exe', '/q /norestart', 'Visual C\+\+ 2010 +x64 Redistributable'),
        @('2012', 'x86', 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x86.exe', '/install /quiet /norestart', 'Visual C\+\+ 2012 Redistributable \(x86\)'),
        @('2012', 'x64', 'https://download.microsoft.com/download/1/6/B/16B06F60-3B20-4FF2-B699-5E9B7962F9AE/VSU_4/vcredist_x64.exe', '/install /quiet /norestart', 'Visual C\+\+ 2012 Redistributable \(x64\)'),
        @('2013', 'x86', 'https://aka.ms/highdpimfc2013x86enu', '/install /quiet /norestart', 'Visual C\+\+ 2013 Redistributable \(x86\)'),
        @('2013', 'x64', 'https://aka.ms/highdpimfc2013x64enu', '/install /quiet /norestart', 'Visual C\+\+ 2013 Redistributable \(x64\)')
    )
    foreach ($v in $vcOld) {
        $items += @{
            Name = "VC++ $($v[0]) $($v[1])"; Url = $v[2]; Args = $v[3]; Pattern = $v[4]; Arch = $v[1]; File = "vcredist_$($v[0])_$($v[1]).exe"
            Skip = { if ($it.Arch -eq 'x64' -and -not $is64) { '32位系统不需要' } }
            Check = { Test-Installed $it.Pattern }
            Install = { Install-Exe $it.Url $it.File $it.Args }
        }
    }

    # VC++ 2015-2022 一直在更新：没装或低于 14.40 就装最新版（原地升级，方案里确认过可以）
    foreach ($a in 'x86', 'x64') {
        $items += @{
            Name = "VC++ 2015-2022 $a"; Arch = $a; Url = "https://aka.ms/vs/17/release/vc_redist.$a.exe"; File = "vc_redist.$a.exe"
            Skip = { if ($it.Arch -eq 'x64' -and -not $is64) { '32位系统不需要' } }
            Check = {
                foreach ($k in "HKLM:\SOFTWARE\Wow6432Node\Microsoft\VisualStudio\14.0\VC\Runtimes\$($it.Arch)", "HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\$($it.Arch)") {
                    if ((Get-RegValue $k 'Installed') -eq 1 -and [int](Get-RegValue $k 'Minor') -ge 40) { return $true }
                }
                return $false
            }
            Install = { Install-Exe $it.Url $it.File '/install /quiet /norestart' }
        }
    }

    $items += @{
        Name = '.NET 3.5'
        Check = { (Get-RegValue 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5' 'Install') -eq 1 }
        Install = {
            # Win8 起要联网从 Windows 更新下载；Win7 自带，只是没打开（Win7 的 DISM 不认 /all）
            $a = '/online /enable-feature /featurename:NetFx3 /norestart'
            if ($build -ge 9200) { $a += ' /all' }
            Write-Log '  用 DISM 打开 .NET 3.5（要联网，可能要几分钟）...'
            $code = Invoke-Wait "$env:SystemRoot\System32\dism.exe" $a 1800
            if ($code -eq 3010) { $State.NeedReboot = $true } elseif ($code -ne 0) { throw "DISM 返回 $code" }
        }
    }

    # Win10 20H2(19042) 起装 4.8.1，更早的系统（含 Win7 SP1）最高 4.8；Release 号是微软文档里的最低值
    $want481 = $build -ge 19042
    $netName = '.NET 4.8'; $netRelease = 528040; $netUrl = 'https://go.microsoft.com/fwlink/?linkid=2088631'
    if ($want481) { $netName = '.NET 4.8.1'; $netRelease = 533320; $netUrl = 'https://go.microsoft.com/fwlink/?linkid=2203305' }
    $items += @{
        Name = $netName
        Check = { [int](Get-RegValue 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' 'Release') -ge $netRelease }
        Install = { Install-Exe $netUrl 'ndp4x.exe' '/q /norestart' }
    }

    $items += @{
        Name = 'DirectX 9.0c (June 2010)'
        Check = {
            if (-not (Test-Path "$env:SystemRoot\System32\d3dx9_43.dll")) { return $false }
            if ($is64 -and -not (Test-Path "$env:SystemRoot\SysWOW64\d3dx9_43.dll")) { return $false }
            return $true
        }
        Install = {
            $exe = Get-Download 'https://download.microsoft.com/download/8/4/A/84A35BF1-DAFE-4AE8-82AF-AD2AE20B6B14/directx_Jun2010_redist.exe' 'directx_Jun2010_redist.exe'
            $dx = Join-Path $Cache 'dx'
            if (-not (Test-Path $dx)) { New-Item -ItemType Directory $dx | Out-Null }
            $code = Invoke-Wait $exe "/Q /T:`"$dx`"" 600
            if ($code -ne 0) { throw "解压返回 $code" }
            Write-Log '  安装 DirectX ...'
            $code = Invoke-Wait (Join-Path $dx 'DXSETUP.exe') '/silent' 1800
            if ($code -ne 0) { throw "DXSETUP 返回 $code" }
        }
    }

    $items += @{
        Name = 'WebView2'
        Skip = { if ($build -lt 10240) { 'Win7/8 不支持新版 WebView2' } }
        Check = {
            foreach ($k in 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
                           'HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}',
                           'HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}') {
                $pv = Get-RegValue $k 'pv'
                if ($pv -and $pv -ne '0.0.0.0') { return $true }
            }
            return $false
        }
        Install = { Install-Exe 'https://go.microsoft.com/fwlink/p/?LinkId=2124703' 'MicrosoftEdgeWebview2Setup.exe' '/silent /install' }
    }

    # 芯烨小票驱动（2026-10-13 沙盒验证过的装法）：
    # 官网下载要带 UA + Referer，不然 403；/silence 只是静默解压到 Program Files，之后 drvinst.exe 仍弹向导等插打印机 → 直接关掉；
    # INF 里带 USBPRINT 硬件 ID，芯烨原装机插上 USB 后 Windows 会自己建打印机。
    $printEnv = 'Windows NT x86'; if ($is64) { $printEnv = 'Windows x64' }
    $items += @{
        Name = '芯烨小票驱动'
        Check = { Test-Path "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Environments\$printEnv\Drivers\Version-3\Xprinter XP-80" }
        Install = {
            $url = Get-XprinterUrl 'https://www.xprinter.net/Driver_Download/2.html' '票据'
            $exe = Get-Download $url 'xprinter_receipt.exe' 'https://www.xprinter.net/'

            Write-Log '  解压驱动 ...'
            $p = Start-Process $exe -ArgumentList '/silence' -PassThru
            $drv = @()
            for ($i = 0; $i -lt 120 -and $drv.Count -eq 0; $i++) {
                Start-Sleep 1
                $drv = @(Get-WmiObject Win32_Process -Filter "Name='drvinst.exe'" | Where-Object { $_.ExecutablePath -match 'Xprinter' })
            }
            if ($drv.Count -eq 0) { throw '解压后没等到芯烨的 drvinst.exe（安装包可能改了）' }
            $dir = Split-Path $drv[0].ExecutablePath
            Start-Sleep 2
            foreach ($d in $drv) { try { Stop-Process -Id $d.ProcessId -Force } catch { } }
            if (-not $p.HasExited) { try { $p.Kill() } catch { } }

            $inf = Join-Path $dir 'xpposv6.inf'
            if (-not (Test-Path $inf)) {
                $inf = @(Get-ChildItem $dir -Filter *.inf)
                if ($inf.Count -eq 0) { throw "解压目录里没有 INF：$dir" }
                $inf = $inf[0].FullName
            }
            Import-DriverInf $inf

            $models = @()
            foreach ($line in [IO.File]::ReadAllLines($inf)) {
                if ($line -match '^\s*"([^"]+)"\s*=\s*PRNTDRV') { if ($models -notcontains $matches[1]) { $models += $matches[1] } }
            }
            $hasCmdlet = [bool](Get-Command Add-PrinterDriver -ErrorAction SilentlyContinue)
            foreach ($m in $models) {
                if ($hasCmdlet) {
                    try { Add-PrinterDriver -Name $m } catch { Write-Log "  $m 注册失败：$($_.Exception.Message)" Yellow }
                } else {
                    # Win7 没有 Add-PrinterDriver，用 printui
                    $null = Invoke-Wait "$env:SystemRoot\System32\rundll32.exe" "printui.dll,PrintUIEntry /ia /m `"$m`" /f `"$inf`" /q" 300
                }
            }
            Write-Log "  型号：$($models -join '、')"
        }
    }

    # 业务软件（2026-10-13 沙盒验证过）：都是 Electron + NSIS，/S 静默装，几秒装完，自动建桌面快捷方式。美得得按方案不装。
    # 美团：商家后台「下载电脑客户端」按钮用的地址，不用登录，302 到最新版安装包
    $items += @{
        Name = '美团外卖商家版'
        Check = { Test-AppFiles '美团外卖商家' }
        Install = { Install-Exe 'https://e.waimai.meituan.com/gw/version/pc/download' 'meituan_setup.exe' '/S /allusers' }
    }
    # 淘宝闪购零售商家版（原饿了么零售商家版）：官方下载页 market.m.taobao.com/app/eleme_nr_bfe_retail/eb_download 读的 download.json。
    # 安装包默认写死装到 D:\淘宝闪购零售商家版，没 D 盘就装空（还报成功）→ 没有 D 盘时用 /D= 指到 C 盘。
    # /D= 必须放最后，路径不能有空格（实测 Program Files 会被截成 C:\Program）
    $items += @{
        Name = '淘宝闪购零售商家版'
        Check = { Test-AppFiles '淘宝闪购|饿了么|eleme' }
        Install = {
            $wc = New-WebClient
            $wc.Encoding = [Text.Encoding]::UTF8
            $cfg = $wc.DownloadString('https://alsc-retail.elemecdn.com/msConfig/eleme_nr_bfe_retail/eb_download/download.json')
            if ($cfg -notmatch '"downloadUrl"\s*:\s*"([^"]+)"') { throw '淘宝闪购下载配置里找不到安装包地址（可能改版了）' }
            $url = $matches[1]
            $dir = 'C:\淘宝闪购零售商家版'
            $d = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='D:'" -ErrorAction SilentlyContinue
            if ($d -and $d.DriveType -eq 3) { $dir = 'D:\淘宝闪购零售商家版' }
            Install-Exe $url 'eleme_retail_setup.exe' "/S /allusers /D=$dir"
        }
    }

    # ---------- 逐项检查、补装 ----------
    $results = @()
    foreach ($it in $items) {
        $r = @{ Name = $it.Name; Status = ''; Detail = '' }
        try {
            $skip = $null
            if ($it.Skip) { $skip = & $it.Skip }
            if ($skip) { $r.Status = '跳过'; $r.Detail = $skip }
            elseif (& $it.Check) { $r.Status = '已有' }
            elseif ($it.NoAuto) { $r.Status = '没装'; $r.Detail = '请手动装' }
            else {
                Write-Log "缺 $($it.Name)，开始补装" Yellow
                & $it.Install
                if (& $it.Check) { $r.Status = '已补装' } else { $r.Status = '补装失败'; $r.Detail = '装完检查还是没有' }
            }
        } catch {
            $r.Status = '补装失败'; $r.Detail = Get-ShortError $_
        }
        $color = 'Green'
        if ($r.Status -eq '补装失败' -or $r.Status -eq '没装') { $color = 'Red' } elseif ($r.Status -eq '跳过') { $color = 'Gray' }
        $line = '{0}：{1}' -f $r.Name, $r.Status
        if ($r.Detail) { $line += "（$($r.Detail)）" }
        Write-Log $line $color
        $results += $r
    }

    # ---------- 给没添加的 USB 小票机添加打印机（要在芯烨小票驱动装好之后） ----------
    $added = @()
    if ($freePorts.Count) {
        $model = "Xprinter XP-$paper"
        $drv = @(Get-WmiObject Win32_PrinterDriver -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "$model,*" })
        foreach ($port in $freePorts) {
            try {
                if ($drv.Count -eq 0) { throw '芯烨小票驱动没装好' }
                $name = '小票打印机'
                if (Get-WmiObject Win32_Printer -Filter "Name='$name'" -ErrorAction SilentlyContinue) { $name = "小票打印机 $port" }
                Write-Log "给 $port 上的小票机添加打印机：$name（$model）" Yellow
                if (Get-Command Add-Printer -ErrorAction SilentlyContinue) { Add-Printer -Name $name -DriverName $model -PortName $port }
                else {
                    # Win7 没有 Add-Printer：/if 按 INF 加打印机，/u 用已装好的驱动，/z 不共享，/q 不弹框
                    $null = Invoke-Wait "$env:SystemRoot\System32\rundll32.exe" "printui.dll,PrintUIEntry /if /b `"$name`" /r `"$port`" /m `"$model`" /f `"$($drv[0].InfName)`" /u /z /q" 120
                }
                $pr = Get-WmiObject Win32_Printer -Filter "Name='$name'"
                if (-not $pr) { throw '加完找不到这台打印机' }
                $tp = $pr.PrintTestPage()
                $t = "$name（$model，$port）"
                if ($tp.ReturnValue -eq 0) { $t += '，已打测试页'; Write-Log '  已打一张测试页，请看小票机有没有出纸' Green }
                else { $t += "，测试页没打出（$($tp.ReturnValue)）" }
                $added += $t
            } catch {
                $added += "$port 没加上（$(Get-ShortError $_)）"
                Write-Log "  $port 没加上：$($_.Exception.Message)" Red
            }
        }
    }

    # ---------- 打印机（用来攒各店型号清单） ----------
    # 小票机多半不报型号，Windows 叫它「No Printer Attached」，看着像没插 → 报告里写「型号不明」，再写这个口添加成了哪台打印机
    $portOwner = @{}
    foreach ($pr in @(Get-WmiObject Win32_Printer -ErrorAction SilentlyContinue)) { $portOwner[[string]$pr.PortName] = [string]$pr.Name }
    $usb = @()
    foreach ($e in $usbDevices) {
        $hw = ($e.PNPDeviceID -split '\\')[1]
        $t = "$($e.Name) [$hw]"
        if ($hw -eq 'UNKNOWNPRINTER') { $t = '型号不明' }
        $port = '端口不明'; if ($e.PNPDeviceID -match '(USB\d{3})$') { $port = $matches[1] }
        if ($portOwner[$port]) { $t += "，$port，已添加为「$($portOwner[$port])」" } else { $t += "，$port，没添加打印机" }
        $usb += $t
    }
    $printers = @()
    foreach ($pr in @(Get-WmiObject Win32_Printer -ErrorAction SilentlyContinue)) {
        # 跳过虚拟打印机（PDF、传真、XPS、OneNote、远程桌面）；真打印机也可能用微软通用驱动，所以不能按 ^Microsoft 过滤
        if ($pr.DriverName -match 'XPS|PDF|Fax|OneNote|Remote Desktop|Easy Print|Virtual' -or $pr.PortName -match '^(PORTPROMPT|SHRFAX|nul):|^TS\d') { continue }
        $printers += "$($pr.Name)（$($pr.DriverName)，$($pr.PortName)）"
    }

    # ---------- 报告 ----------
    $groups = @{}
    foreach ($r in $results) {
        $t = $r.Name; if ($r.Detail) { $t += "（$($r.Detail)）" }
        if (-not $groups[$r.Status]) { $groups[$r.Status] = @() }
        $groups[$r.Status] += $t
    }
    $lines = @()
    $lines += '【{0}】装机体检 {1:yyyy-MM-dd HH:mm} v{2}' -f $store, (Get-Date), $Version
    $lines += "系统：$osText"
    if (-not $State.Online) { $lines += '⚠ 运行时测网不通' }
    foreach ($s in '补装失败', '没装', '已补装', '跳过') {
        if ($groups[$s]) { $lines += "${s}：" + ($groups[$s] -join '；') }
    }
    $okCount = 0; if ($groups['已有']) { $okCount = $groups['已有'].Count }
    $lines += "已有：$okCount 项"
    if ($State.NeedReboot) { $lines += '⚠ 需要重启电脑才能生效' }
    $u = '无'; if ($usb.Count) { $u = "$($usb.Count) 台：" + ($usb -join '；') }
    $lines += "USB 打印机：$u"
    if ($added.Count) { $lines += '新加打印机：' + ($added -join '；') }
    $pp = '无'; if ($printers.Count) { $pp = $printers -join '；' }
    $lines += "已建打印机：$pp"
    $report = $lines -join "`n"
    [IO.File]::WriteAllText((Join-Path $Dir "logs\$stamp-report.txt"), $report, [Text.Encoding]::UTF8)

    Write-Host ''
    Write-Log $report Cyan

    if ($hook) {
        try {
            $sb = New-Object Text.StringBuilder
            foreach ($ch in $report.ToCharArray()) {
                $c = [int]$ch
                if ($ch -eq '"') { [void]$sb.Append('\"') }
                elseif ($ch -eq '\') { [void]$sb.Append('\\') }
                elseif ($c -lt 32) { [void]$sb.Append(('\u{0:x4}' -f $c)) }
                else { [void]$sb.Append($ch) }
            }
            $json = '{"msg_type":"text","content":{"text":"' + $sb.ToString() + '"}}'
            $wc = New-Object Net.WebClient
            $wc.Encoding = [Text.Encoding]::UTF8
            $wc.Headers.Add('Content-Type', 'application/json; charset=utf-8')
            $resp = $wc.UploadString($hook, 'POST', $json)
            if ($resp -match '"(code|StatusCode)"\s*:\s*0\b') { Write-Log '报告已发到飞书群' Green }
            else { Write-Log "飞书返回：$resp" Red }
        } catch { Write-Log "发飞书失败：$($_.Exception.Message)" Red }
    } else {
        Write-Log "没配置飞书群地址，报告只存在本机：$Dir\logs" Yellow
    }

    if ($State.NeedReboot) { Write-Host '有组件需要重启电脑才能生效，请找空闲时间重启。' -ForegroundColor Yellow }
}

# 交给管理员窗口之后这个窗口直接退出，不然店员看到「按回车关闭」会以为已经跑完了
$relaunched = $false
try { $relaunched = (Invoke-StoreSetup) -contains 'relaunched' }
catch { Write-Host "出错了：$($_.Exception.Message)" -ForegroundColor Red }
if (-not $relaunched -and $env:STORESETUP_TEST -ne '1') { Read-Host '按回车关闭' | Out-Null }
