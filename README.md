# Norio Excel 编辑器 Skill

`norio-office-excel` 是配合 [`@norio-office/office-excel`](https://www.npmjs.com/package/@norio-office/office-excel) 使用的 AI Skill。它约束 AI 通过编辑器公开的能力、目标授权和结构化 Patch 创建或修改工作簿，并可在离线环境校验数据、生成或读取标准 `.xlsx` 文件。

Skill 入口是 [`norio-office-excel/SKILL.md`](norio-office-excel/SKILL.md)。协议说明位于 `references/`，机器可读 Schema 位于 `schemas/`，可执行样例位于 `examples/`。

## 接入说明

宿主应用需要安装兼容版本的 `@norio-office/office-excel`，注册组件公开的 AI 工具，再将本 Skill 提供给 AI 运行环境。**只复制 Skill 不会安装组件、注册工具、建立协作连接或授予工作簿权限。** 离线 CLI 需要 Node.js 20+ 和已安装的兼容 npm 包。

## 维护与发布

修改公共协议后，在仓库根目录执行：

```bash
npm run skill:sync
npm run skill:verify
```

Skill 改动应先随主项目提交并推送到 CNB 的 `main`。之后显式执行 `npm run push -- --skills`，脚本才会通过 Git subtree 将 `packages/skills` 单独同步到 [GitHub Skill 仓库](https://github.com/cgq001/norio-office-office-excel)。普通 `git pull` 和 `git push` 仍只操作主项目的 CNB 远端。
