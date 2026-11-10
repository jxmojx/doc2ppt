# doc2ppt Agent Skill 分发分支（develop）

本分支**只含 Skill 封装**，用于 Agent 技能系统直接拉取；完整项目在 [main](https://github.com/jxmojx/doc2ppt)。

## 安装

```bash
git clone -b develop --depth 1 https://github.com/jxmojx/doc2ppt.git doc2ppt-skill
cd doc2ppt-skill
# 把 skill/doc2ppt 复制进你的 Agent 技能目录，或把 skill/doc2ppt/SKILL.md 内容贴入 Agent 的系统提示
```

SKILL.md：完整的使用流程（环境自检 → 输入整理 → 转换 → 验收交付）与已知边界。
