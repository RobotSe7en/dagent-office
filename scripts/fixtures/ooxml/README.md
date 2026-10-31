# 独立原生 XML 校验

`OfficeOpenXML-XMLSchema-Transitional.zip` 是未修改的 ECMA-376 Part 4 第五版（2016）电子附件，仅供开发验收，不进入发布包。来源为 [ECMA 官方 Part 4](https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip)，SHA-256 为 `d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc`；版权与许可见 [LICENSE.txt](LICENSE.txt)。

`check-pptx-output.py` 使用通用 lxml 校验完整标准部件和嵌套常量工作簿；现代批注采用所选 [MS-PPTX](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-pptx/161bc2c9-98fc-46b7-852b-ba7ee77e2e54) 结构检查，其富文本正文使用原规范的 `a:CT_TextBody`。该选定现代检查不等于完整的 Microsoft 扩展 Schema 或 PowerPoint/WPS 实机验收。

本机开发使用 lxml 6.0.2（BSD），CI 使用系统 `python3-lxml`；版本记录进入验收报告。lxml 只运行于开发校验脚本，不作为 SDK 的 Office 转换或运行时依赖。
