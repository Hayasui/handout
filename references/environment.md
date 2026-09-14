# 本机渲染链

## python 用哪一个

装依赖的隔离环境是 `C:\Users\78572\.workbuddy\binaries\python\envs\default`，
`python-docx` 与 `pypdf` 都在里面，跑脚本一律用它的解释器：

```
C:\Users\78572\.workbuddy\binaries\python\envs\default\Scripts\python.exe
```

用管家的基础版（`...\binaries\python\versions\3.13.12\python.exe`）会报找不到 `docx`。

## PDF 怎么出

Word 没装，LibreOffice 没有，WPS 装了但 COM 被安全策略拦掉——**不能拿 Office 系把 docx 转 PDF**。
pandoc 在 `C:\Program Files\Pandoc\pandoc.exe`，本流程用不上。

所以 HTML → PDF 走 Edge 的无头模式。Edge 在
`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`，脚本会自动找，找不到才需要
`--edge` 指定。调用式样：

```
msedge --headless=new --disable-gpu --no-first-run --no-pdf-header-footer
       --user-data-dir=<纯 ASCII 目录> --print-to-pdf=<纯 ASCII 输出路径>
       file:///<纯 ASCII 的 html 路径>
```

Edge 会往 stderr 吐一堆同步失败、字体回退之类的抱怨，**只要看到 `N bytes written to file`
就算成功**，别被吓到。脚本会把这个尾巴截一段写进报告。

## 为什么 HTML 与 PDF 要先落到 ASCII 路径

`file:///` 对中文路径要做百分号编码，Edge 拿到之后经常打不开。所以 HTML 与 PDF 先写在
纯 ASCII 的中转目录（默认 `C:\Agent\_handout_tmp`），成功了再把 PDF 拷回中文路径。
中转目录里还会生成一份 Edge 的 profile，几十兆，会在下一次运行时复用，不必清。

## 在这台机器上怎么调脚本

bash 起不来（`dirname: command not found`，coreutils 不在 PATH），`ls`、`grep`、`wc` 一律没有。
用 PowerShell 调 python 时又抓不到原生命令的标准输出，`*>` 写出来是 UTF-16，读不出来。

稳的做法是在纯 ASCII 目录放一个壳子脚本：自己往 UTF-8 文件里写报告，再用读文件的方式看。
现成的壳子可以照着 `C:\Agent\_run_handout.py` 抄，它把标准输出与异常分别写成
`_run_out.txt` 与 `_run_err.txt`。

另外，**别在命令行里传中文路径**，PowerShell 偶发 GBK 乱码，脚本会报找不到文件，报错里的路径
也是乱的。中文路径写进 JSON 配置或写进源码，命令行只传配置文件的路径。

## 依赖

`python-docx` 负责 docx，`pypdf` 只在核对分页时用。两个渲染器的分页不保证逐页一致，
要核对就比页数与正文文字，别比字节。
