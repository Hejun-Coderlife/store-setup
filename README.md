# 门店装机体检

一键检查门店电脑缺什么、缺了自动补装：VC++ 运行库、.NET、DirectX、WebView2、Edge、芯烨小票驱动、美团外卖商家版、淘宝闪购零售商家版、美得得 ERP（要有 D 盘）、WPS、赫眉外卖监控；插着但没添加的 USB 小票机会自动添加。最后生成一份体检报告。

门店电脑上打开 https://hejun-coderlife.github.io/store-setup/ ，点「复制命令」，按 Win+R，Ctrl+V，回车。命令就是下面这一行：

```
powershell -ep bypass -c "[Net.ServicePointManager]::SecurityProtocol=[Enum]::ToObject([Net.SecurityProtocolType],3072);iex ((New-Object Net.WebClient).DownloadString('https://hejun-coderlife.github.io/store-setup/go.ps1'))"
```

**打不开 GitHub 的电脑**用这一行（从 Gitee 下载，用法一样：Win+R，粘贴，回车）：

```
powershell -ep bypass -c "[Net.ServicePointManager]::SecurityProtocol=[Enum]::ToObject([Net.SecurityProtocolType],3072);iex ((New-Object Net.WebClient).DownloadString('https://gitee.com/hejun202609/store-setup/raw/master/go.ps1'))"
```

支持 Windows 7 SP1 及以上。只补装、不卸载，已经装好的不动。
