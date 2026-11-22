# odyssey-income · 五年收入推演

把斯坦福人生设计课里的「奥德赛计划」做成了一个 AI skill。

奥德赛计划能帮你想清楚未来五年的三种活法，但它不告诉你每条路大概能挣多少，也不告诉你这周该做什么。这个 skill 把后半段补上了：它一层一层问你，推演出三条五年路线和每年的收入，帮你找出最想走的那条，再从五年后倒推到今天先做的一件事，最后出一份 PPT。

## 它会带你走五步

1. **一问一答**：10 个问题，从做过什么、擅长什么、接受不了什么，问到愿意为了什么放弃什么。答得越具体，后面越准。
2. **三条五年路线**：按现在的方向走下去；现在这条路走不了了；不考虑钱最想过的生活。每条写清每年走到哪、收入的保守数和乐观数、五年后普通的一天、要付的代价、选错了会怎样。
3. **你自己选**：三个问题（想要、后悔、承受），它不替你选。
4. **倒推到今天**：5 年 → 3 年 → 1 年 → 90 天 → 这个月 → 这周 → 今天。
5. **一份 8 页 PPT**：文字都能直接改。

## 安装

Codex：

```bash
git clone https://github.com/heloraai/odyssey-income ~/.codex/skills/odyssey-income
```

Claude Code：

```bash
git clone https://github.com/heloraai/odyssey-income ~/.claude/skills/odyssey-income
```

其他能读 skill 的 AI 工具（比如豆包工作）：按它导入技能的方式导入这个文件夹，或者把 `SKILL.md` 的全文发给它。

出 PPT 需要 Python 3 和 python-pptx：

```bash
python3 -m pip install python-pptx
```

Windows 上把 `~` 换成 `%USERPROFILE%`，`python3` 换成 `python` 或 `py -3`。

不能运行脚本的工具也能用，它会把同样的内容直接写成文档给你。

## 使用

对它说：

```text
推演我的五年收入
```

然后照实回答。想快一点，可以一次把情况全说了，它只会补问缺的那几项。

## 文件

| 文件 | 作用 |
|---|---|
| `SKILL.md` | 给 AI 看的完整流程 |
| `scripts/make_deck.py` | 把结果画成 PPT；`python3 scripts/make_deck.py --check` 可以自检 |
| `assets/deck-example.json` | 一个虚构人物的示例结果 |

## 说明

- 收入是按你给的信息推出来的，不是保证。拿到结果后，建议找做这行的人问一问。
- 路线里有副业、收费、开公司时，动手前先确认合同、签证或单位规定允不允许。
- 奥德赛计划出自 Bill Burnett 和 Dave Evans 的《Designing Your Life》。这个 skill 是独立实现，和斯坦福大学、原书作者没有关系。

## License

MIT
