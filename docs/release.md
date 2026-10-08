# 发布流程

推送一个以 `v` 开头的 tag，GitHub Actions 就会自动编译 VSIX 并发布到 GitHub Release。

## 概述

- **触发条件**：推送以 `v` 开头（匹配 `v*`）的 tag，例如 `v2.0`。只推送 `main` 的提交不会触发。
- **工作流**：[`.github/workflows/release.yml`](../.github/workflows/release.yml)，在 `windows-2022` 上用 MSBuild 编译 Release，把 `bin/Release/EncodingDisplayExtension.vsix` 上传到该 tag 对应的 Release。
- **注意**：GitHub 使用的是 tag 所指向的那个提交里的工作流文件，所以该提交必须已经包含 `release.yml`。

## 发布步骤

1. **确认 `main` 已推送**

   ```powershell
   git push origin main
   ```

2. **升版本号**
   - 打开 `source.extension.vsixmanifest`，修改 `<Identity ... Version="x.y" ...>` 里的 `Version`。VS 以这个版本号判断是否为新版本。
   - tag 名 = `v` + 清单版本号（清单 `2.0` 对应 tag `v2.0`）。工作流不会校验两者是否一致，需要自己保证。
   - 提交并推送到 `main`。

3. **打 tag 并推送**（要打 tag 的提交必须已经推送到远程）

   ```powershell
   git tag v2.0
   git push origin v2.0
   ```

   TortoiseGit 的做法：右键仓库 → TortoiseGit → Create Tag，填写 `v2.0`，基于 HEAD；然后 Push，选择推送该 tag。

4. **查看运行情况**：打开 [Actions 页面](https://github.com/Alex-hj/EncodingDisplayExtension/actions)，会出现名为 Release 的运行记录，一般几分钟内完成。

5. **取成品**：打开 [Releases 页面](https://github.com/Alex-hj/EncodingDisplayExtension/releases)，对应版本下有 `EncodingDisplayExtension.vsix`，下载后双击即可安装。

同一个 tag 名只对应一次发布。要重新发布请升版本号，或按下文“重来”先删除再重建。

## 首次使用建议先用测试 tag

工作流的真实运行只有在推送 tag 时才会发生。建议先推送一个测试 tag，例如 `v0.0.1-test`，确认流程跑通后，在 Releases 页面删除这个测试 Release，再删除 tag（见下文“重来”）。测试 Release 是公开的，会短暂可见。

## 失败排查

- **日志里出现 403 或 `Resource not accessible by integration`**：到仓库 Settings → Actions → General → Workflow permissions，改成 Read and write permissions。工作流里已声明 `permissions: contents: write`，正常不需要改，只有仓库或组织的策略限制了权限时才会出现。
- **提示找不到 vsix**：说明前面的“编译 VSIX”步骤失败了，点开该步骤查看日志。
- **推送 tag 后没有运行记录**：依次检查：
  - tag 是否以 `v` 开头；
  - tag 指向的提交里是否包含 `.github/workflows/release.yml`；
  - 仓库是否禁用了 Actions（Settings → Actions → General）。

## 重来

先在 Releases 页面删除该 Release，再删除 tag：

```powershell
git push origin --delete v2.0
git tag -d v2.0
```

然后修复问题，重新打 tag 并推送。如果只是偶发失败（比如网络问题），也可以在 Actions 页面对失败的运行点 Re-run jobs，它会复用同一个 tag 和提交。

## 本地复现编译

排查编译问题时，可以在 Visual Studio 2022 的 Developer PowerShell 里运行与工作流相同的命令：

```powershell
msbuild EncodingDisplayExtension.csproj /restore /p:Configuration=Release /p:DeployExtension=false
```

产物在 `bin\Release\EncodingDisplayExtension.vsix`。`DeployExtension=false` 表示不把扩展部署到 VS 实验实例；省略它时，本地命令行编译会像 F5 一样自动部署。

## 说明

- **runner 固定为 `windows-2022`**：该镜像装的是 VS 2022（17.x）和扩展开发工作负载，与清单里的 `[17.0, 18.0)` 一致。截至 2026-09-29，`windows-latest` 指向的镜像只装了 VS 2026。如果 GitHub 日后下线 `windows-2022`，需要修改 `runs-on`。
- **使用的 action**：`actions/checkout@v7`、`microsoft/setup-msbuild@v3`、`softprops/action-gh-release@v3`。

## 可选改进（当前未启用）

- 在 `softprops/action-gh-release` 的 `with` 里加 `generate_release_notes: true`，自动生成 release 说明。
- 发布前校验 tag 与清单版本号一致。
- 增加 `workflow_dispatch` 触发方式，支持手动发布。
- 另加一个在 PR 和 push 时只编译不发布的工作流。
