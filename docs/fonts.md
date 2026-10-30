# 中文字体与本地预览环境

文档源保留原字体名称和 ASCII、高位西文、东亚、复杂文字角色。支持保存名称不代表本机已经安装该字体；SDK 不在打开文档时下载字体，也不把替代字体改名为源字体。

R1-06 回归覆盖宋体、新宋体、`SimSun/NSimSun`，仿宋、仿宋_GB2312、华文仿宋、`FangSong/FangSong_GB2312/STFangsong`，楷体、楷体_GB2312、华文楷体、`KaiTi/KaiTi_GB2312/STKaiti`，以及下列开源字体。文字替换、历史和文档包重开保留这些名称、中西文角色、UTF-16 及变体选择符。

## 已安装的补充字体

本次开发环境按用户要求安装到 `/usr/local/share/fonts/dagent-office`，Fontconfig 缓存已刷新。字体二进制不提交仓库、不随 SDK 分发；版本、下载地址及逐文件 SHA-256 记录在[字体环境](../examples/fonts/environment.json)。许可与 README 保存在 `/usr/local/share/doc/dagent-office-fonts`。

| 项目 | 版本 | 已安装字形与字重 | 许可及来源 |
| --- | --- | --- | --- |
| Fandol | 0.3 | FandolSong、FandolHei 的 Regular/Bold，FandolFang、FandolKai 的 Regular | GPL v3 加字体例外，见发行包 README/COPYING；[CTAN 项目](https://ctan.org/pkg/fandol) |
| 霞鹜文楷 / LXGW WenKai | 1.530 | Light、Regular、Medium | SIL OFL 1.1；[官方发行](https://github.com/lxgw/LxgwWenKai/releases/tag/v1.530)、[许可](https://github.com/lxgw/LxgwWenKai/blob/v1.530/OFL.txt) |

系统原有 Noto Sans CJK SC、Noto Serif CJK SC 和 Liberation Serif 继续作为 DOCX 参考环境，既有[参考记录](../examples/docx/reference.json)没有因安装新字体而刷新。

## 缺失字体的预览

预览先尝试文档请求的字体列表。列表中均不可用时，仿宋类名称依次尝试 FandolFang、Noto Serif CJK SC、serif；楷体类名称依次尝试 FandolKai、LXGW WenKai、Noto Serif CJK SC、serif。宋体类优先使用 Noto Serif CJK SC，再尝试 FandolSong。字体不存在时仍明确报告请求名称与实际选用名称，源 CSS 保持原样。

这些是预览替代，不承诺与微软或华文字体相同的字形、字宽或断行。已有正版字体可由宿主安装后使用；当前未从第三方下载或安装微软宋体、仿宋、楷体等专有字体。

浏览器用固定字体用例验证 FandolSong/FandolFang/FandolKai/LXGW WenKai 实际可选，以及仿宋、楷体名称的显式替代、中西文分工和缺字诊断。测量诊断仍受[DOCX 读取说明](docx-import.md#中文字体)所列边界约束。
