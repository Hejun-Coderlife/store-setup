# 门店装机体检

一键检查门店电脑缺什么、缺了自动补装：VC++ 运行库、.NET、DirectX、WebView2、芯烨小票驱动、美团外卖商家版、淘宝闪购零售商家版；插着但没添加的 USB 小票机会自动添加。最后生成一份体检报告。

在门店电脑上按 Win+R（或打开命令提示符），粘贴下面这一行回车：

```
powershell -ep bypass -c "[Net.ServicePointManager]::SecurityProtocol=[Enum]::ToObject([Net.SecurityProtocolType],3072);iex ((New-Object Net.WebClient).DownloadString('https://hejun-coderlife.github.io/store-setup/go.ps1'))"
```

支持 Windows 7 SP1 及以上。只补装、不卸载，已经装好的不动。
