# 独立 Word 闭环

R1-14 将已有公开 API 接入 Playground，并以独立安装的五个 tarball 验证普通 JavaScript、React 与 Node 宿主。HTML 文档包仍是持久化源；DOCX 是当前修订生成的原生可编辑 OOXML，PDF 是共享 HTML 画布的输出。

## Playground

运行 `pnpm dev`，打开 DOCX 或 `.zip` 文档包，也可使用“打开报告协作样例”。进入“编辑正文”修改文字，返回预览后统一提交批注，载入固定示例提案，查看差异并整批应用。人工输入与提案共享事务、单调修订、冲突检查及撤销重做。

“导出 DOCX”在宿主拥有的短生命周期 Worker 中读取冻结的文档包。完成、取消、超时、切换文档或修订变化都会结束该 Worker。导出结果必须匹配文档身份与修订才会下载。“导出 PDF”使用开发 Node 子进程和同样的修订检查。打开损坏文件失败时保留当前文档。

DOCX 下载后显示并可下载完整导出报告。精确锚点转为原生批注；已解决状态原生保留。矩形、需要重新定位等不能绑定到 Word 文字的批注明确列为未导出。下载文档包可完整保留这些信息；重新定位需要用户或宿主明确提交事务。详情见 [DOCX 导出](docx-export.md)。

## 独立消费验收

安装 [固定浏览器与字体](development.md)以及 LibreOffice、Poppler 后运行：

```bash
pnpm check
node scripts/playwright.mjs test word-closure.spec.ts
pnpm check:docx-export
pnpm check:word-closure
```

`check:word-closure` 将五个本地包安装到仓库外的临时目录，不使用源代码别名或工作区依赖。临时目录只用于消费验证，验证后清理。普通 JavaScript 宿主没有 React；React 宿主独立安装 React 和 ReactDOM。两者均构建后启动临时本机服务，并在实际浏览器运行。

| 宿主 | 输入 | 验证 |
| --- | --- | --- |
| 普通 JavaScript | HTML 文档包 | 独立只读预览、真实键盘输入、批注请求、固定提案与差异、应用、撤销重做、明确重定位、DOCX／文档包 |
| React | DOCX 导入 | 同一工作台和事务流程，使用可选 `editor/react` 入口 |
| Node | 上述保存产物 | 仅公开包入口读取、DOCX 导入、PDF API、子进程字节协议、PDF 可搜索正文、LibreOffice 重存后重新导入与已解决批注核对 |

浏览器界面用例另核对未重定位批注的报告、损坏文件、切换文档、取消与导出期间编辑。Chromium、Firefox 和 WebKit 运行同一闭环；分页另有 Chromium 109 与真实 SafariDriver 门槛。各测试的覆盖范围以测试代码和 CI 产物为准。

`.cache/word-closure-acceptance/` 保存实际 DOCX、HTML 包、PDF、界面截图、报告、tarball 哈希及运行环境记录；CI 上传这些产物。原生 Word 语义的 21 个 Office 重开样例继续由 `check:docx-export` 验证。

Office 重开基线采用用户确认的 LibreOffice；Word／WPS 尚未验证。该闭环使用公开合成文档与固定提案，没有真实模型执行、正式宿主存储或 npm 发布。达智真实智能体往返属于 R1-15，当前按用户决定延期；下一阶段为[第二轮 PPT](iteration-2.md)，不将第一轮宿主验收标为通过。

## 远程开发访问

Vite 可以监听 `0.0.0.0`。PDF 转换中间件默认仅允许本机连接；远程开发需由宿主明确列出当前同源地址，例如：

```bash
DAGENT_OFFICE_DEV_ALLOWED_ORIGINS=http://192.168.31.221:5174 \
  pnpm --filter @dagent-office/playground exec vite --host 0.0.0.0 --port 5174 --strictPort
```

列表以逗号分隔，仅接受确切的 HTTP(S) origin，不能使用通配符、凭据或路径。请求的 Host、端口、Origin 和开发标记仍需匹配。这个开发选项不替代正式宿主的认证、授权和部署边界。
