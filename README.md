# ForumTouchFish

把论坛伪装成 Microsoft Word 文档界面的油猴脚本：隐藏头像，顶部加上 Word 标题栏和功能区，底部加上状态栏。

| 脚本 | 适用站点 | 安装 |
| --- | --- | --- |
| V2EX Word 摸鱼版 | v2ex.com | [安装](https://raw.githubusercontent.com/RerrentLinden/ForumTouchFish/main/v2ex-moyu.user.js) |
| LINUX DO / IDC Flare Word 摸鱼版 | linux.do、idcflare.com | [安装](https://raw.githubusercontent.com/RerrentLinden/ForumTouchFish/main/linuxdo-idcflare-word-moyu.user.js) |
| NodeSeek / DeepFlood Word 摸鱼版 | nodeseek.com、deepflood.com | [安装](https://raw.githubusercontent.com/RerrentLinden/ForumTouchFish/main/nodeseek-deepflood-moyu.user.js) |

先装好 Tampermonkey 或 Violentmonkey，再点对应的「安装」。

- 点标题栏中间的「工作记录」可以收起或展开功能区和状态栏；往下滚动时也会自动收起。
- 功能区里「样式」后面的按钮是站点的快捷操作，如发帖、点赞或收藏、回复。
- 窗口变窄时，标题栏、功能区和状态栏按 Word 的顺序逐级收起（先折叠段落、字体分组，再隐藏次要按钮），任何宽度下都不会挤出窗口或互相遮挡。
- 看大图、弹窗、全屏编辑器等浮层都排在功能区和状态栏之间，锚点跳转也会停在顶栏下方，Word 界面不会挡住内容。
