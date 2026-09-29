// ==UserScript==
// @name         LINUX DO / IDC Flare Word 摸鱼版
// @namespace    https://codex.local/userscripts
// @version      1.1.1
// @description  隐藏头像，并把 LINUX DO / IDC Flare 伪装成 Microsoft Word 文档界面。
// @author       Codex
// @match        https://linux.do/*
// @match        https://www.linux.do/*
// @match        https://idcflare.com/*
// @match        https://www.idcflare.com/*
// @run-at       document-start
// @grant        none
// @noframes
// ==/UserScript==

(() => {
  "use strict";

  /*
   * ===== Word 界面：三个脚本共用这一段，修改时三处保持一致 =====
   * 标题栏、选项卡、功能区和状态栏照 Word for Mac 画。窗口变窄时按 Word 的顺序逐级收缩：
   * 带 data-fit 的部件按数字从小到大收起（隐藏、只留图标、缩成小按钮，或整组折成一个按钮），
   * 直到这一行放得下，所以任何宽度下都不会溢出或互相遮挡。
   */
  const WORD_CHROME_CSS = `
    :root {
      --word-blue: #185abd;
      --word-ui-height: 148px;
      --word-status-height: 30px;
      --word-chrome-font: -apple-system, BlinkMacSystemFont, "PingFang SC", "Hiragino Sans GB",
        "Segoe UI", "Microsoft YaHei", sans-serif;
    }

    html.codex-word-ribbon-hidden {
      --word-ui-height: 38px;
      --word-status-height: 0px;
    }

    #codex-word-ui,
    #codex-word-status {
      position: fixed;
      z-index: 2147483646;
      right: 0;
      left: 0;
      box-sizing: border-box;
      margin: 0;
      color: #242424;
      font: 13px/1.2 var(--word-chrome-font);
      letter-spacing: normal;
      text-align: left;
      text-transform: none;
      -webkit-font-smoothing: antialiased;
      user-select: none;
    }

    #codex-word-ui *,
    #codex-word-status * {
      box-sizing: border-box;
    }

    #codex-word-ui :is(a, button),
    #codex-word-status :is(a, button) {
      margin: 0;
      color: inherit;
      font: inherit;
      letter-spacing: normal;
      text-decoration: none;
      text-transform: none;
      box-shadow: none;
    }

    #codex-word-ui svg,
    #codex-word-status svg {
      display: block;
      flex: none;
      width: 20px;
      height: 20px;
      overflow: visible;
      fill: none;
      stroke: currentColor;
      stroke-width: 1.25;
      stroke-linecap: round;
      stroke-linejoin: round;
    }

    #codex-word-ui svg text,
    #codex-word-status svg text {
      font-family: var(--word-chrome-font);
      fill: currentColor;
      stroke: none;
    }

    #codex-word-ui svg .is-accent {
      fill: #2f7bd6;
    }

    #codex-word-ui svg .is-card {
      fill: #fff;
    }

    #codex-word-ui {
      top: 0;
      height: var(--word-ui-height);
      overflow: hidden;
      background: #f5f5f5;
      transition: height 180ms ease;
    }

    .codex-word-titlebar {
      position: relative;
      display: flex;
      align-items: center;
      gap: 12px;
      height: 38px;
      padding: 0 10px 0 14px;
      overflow: hidden;
      color: #fff;
      background: var(--word-blue);
    }

    .codex-word-title-left {
      display: flex;
      flex: 0 1 auto;
      align-items: center;
      gap: 1px;
      min-width: 0;
      overflow: hidden;
    }

    .codex-word-title-right {
      display: flex;
      flex: none;
      align-items: center;
      gap: 6px;
      margin-left: auto;
    }

    .codex-word-traffic {
      display: flex;
      flex: none;
      gap: 8px;
      margin-right: 22px;
    }

    .codex-word-traffic i {
      width: 12px;
      height: 12px;
      border-radius: 50%;
      box-shadow: inset 0 0 0 0.5px rgb(0 0 0 / 18%);
    }

    .codex-word-traffic i:nth-child(1) { background: #ff5f57; }
    .codex-word-traffic i:nth-child(2) { background: #febc2e; }
    .codex-word-traffic i:nth-child(3) { background: #28c840; }

    .codex-word-autosave {
      display: flex;
      flex: none;
      align-items: center;
      gap: 6px;
      height: 28px;
      margin: 0 2px 0 -6px;
      padding: 0 6px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 600;
      white-space: nowrap;
      cursor: default;
    }

    .codex-word-switch {
      display: block;
      width: 30px;
      height: 17px;
      padding: 2px;
      border-radius: 999px;
      background: rgb(255 255 255 / 92%);
      transition: background-color 150ms ease;
    }

    .codex-word-switch::before {
      display: block;
      width: 13px;
      height: 13px;
      content: "";
      border-radius: 50%;
      background: #fff;
      box-shadow: 0 0 0 1px rgb(0 0 0 / 22%), 0 1px 1px rgb(0 0 0 / 18%);
      transition: transform 150ms ease, background-color 150ms ease;
    }

    /* 自动保存打开：圆点移到右边；蓝色标题栏上是白底蓝点，深色标题栏上是蓝底白点（见深色模式）。 */
    .codex-word-autosave.is-on .codex-word-switch {
      background: #fff;
    }

    .codex-word-autosave.is-on .codex-word-switch::before {
      background: var(--word-blue);
      box-shadow: 0 1px 1px rgb(0 0 0 / 18%);
      transform: translateX(13px);
    }

    :is(.codex-word-tb-btn, .codex-word-tb-split) {
      display: inline-flex;
      flex: none;
      align-items: center;
      justify-content: center;
      height: 28px;
      border-radius: 6px;
    }

    .codex-word-tb-btn {
      width: 28px;
    }

    .codex-word-tb-split {
      padding: 0 1px 0 4px;
    }

    #codex-word-ui .codex-word-tb-split .codex-word-caret svg {
      width: 11px;
      height: 11px;
    }

    :is(.codex-word-autosave, .codex-word-tb-btn, .codex-word-tb-split, .codex-word-doc-title, .codex-word-pill-edit, .codex-word-round):hover {
      background: rgb(255 255 255 / 14%);
    }

    :is(.codex-word-tb-btn, .codex-word-tb-split).is-disabled {
      opacity: 0.5;
    }

    .codex-word-overflow {
      display: none;
    }

    .codex-word-title-left:has(.is-folded) .codex-word-overflow,
    .codex-word-tabs:has(.is-folded) .codex-word-overflow {
      display: inline-flex;
    }

    .codex-word-doc-title {
      position: absolute;
      top: 50%;
      left: 50%;
      max-width: 36%;
      overflow: hidden;
      padding: 4px 10px;
      border: 0;
      border-radius: 6px;
      background: transparent;
      font-size: 13px;
      font-weight: 600;
      white-space: nowrap;
      text-overflow: ellipsis;
      cursor: default;
      transform: translate(-50%, -50%);
    }

    .codex-word-pill {
      display: inline-flex;
      flex: none;
      align-items: center;
      gap: 6px;
      height: 28px;
      padding: 0 9px 0 11px;
      border-radius: 999px;
      font-size: 13px;
      font-weight: 600;
      white-space: nowrap;
    }

    #codex-word-ui .codex-word-pill svg {
      width: 18px;
      height: 18px;
    }

    #codex-word-ui .codex-word-pill .codex-word-caret svg {
      width: 12px;
      height: 12px;
    }

    .codex-word-pill-edit {
      background: rgb(255 255 255 / 7%);
      box-shadow: inset 0 0 0 1px rgb(255 255 255 / 38%);
    }

    .codex-word-pill-share {
      color: #183a70;
      background: #c9dcfb;
    }

    .codex-word-pill-share:hover {
      background: #d8e6fc;
    }

    .codex-word-round {
      display: inline-flex;
      flex: none;
      align-items: center;
      justify-content: center;
      width: 30px;
      height: 30px;
      border-radius: 50%;
      box-shadow: inset 0 0 0 1px rgb(255 255 255 / 38%);
    }

    .codex-word-tabs {
      display: flex;
      align-items: center;
      gap: 2px;
      height: 30px;
      padding: 0 8px;
      overflow: hidden;
    }

    .codex-word-tab {
      position: relative;
      display: inline-flex;
      flex: none;
      align-items: center;
      height: 26px;
      padding: 0 11px;
      border-radius: 6px;
      font-size: 13.5px;
      white-space: nowrap;
      cursor: default;
    }

    a.codex-word-tab {
      cursor: pointer;
    }

    .codex-word-tab:hover {
      background: #e8e8e8;
    }

    .codex-word-tab.is-active {
      font-weight: 600;
    }

    .codex-word-tab.is-active::after {
      position: absolute;
      right: 11px;
      bottom: 0;
      left: 11px;
      height: 3px;
      content: "";
      border-radius: 2px;
      background: var(--word-blue);
    }

    .codex-word-tabs .codex-word-overflow {
      flex: none;
      align-items: center;
      height: 26px;
      padding: 0 6px;
      color: #8a8a8a;
    }

    #codex-word-ui .codex-word-tabs .codex-word-overflow svg {
      width: 15px;
      height: 15px;
    }

    .codex-word-ribbon {
      height: 80px;
      padding: 1px 8px 3px;
    }

    .codex-word-card {
      display: flex;
      align-items: stretch;
      height: 100%;
      padding: 0 4px;
      overflow: hidden;
      border-radius: 10px;
      background: #fff;
      box-shadow: 0 0 0 0.5px rgb(0 0 0 / 7%), 0 1px 3px rgb(0 0 0 / 10%);
      scrollbar-width: none;
    }

    .codex-word-card.is-overflowing {
      overflow-x: auto;
    }

    .codex-word-group {
      position: relative;
      display: flex;
      flex: none;
      align-items: center;
      gap: 2px;
      padding: 0 7px;
    }

    .codex-word-group:not(.is-first)::before {
      position: absolute;
      top: 10px;
      bottom: 10px;
      left: 0;
      width: 1px;
      content: "";
      background: #e1e1e1;
    }

    .codex-word-more {
      display: none;
      flex: none;
      align-items: center;
      margin-left: auto;
      padding: 0 2px 0 6px;
      color: #9a9a9a;
    }

    .codex-word-card:has(.is-trailing.is-folded) .codex-word-more {
      display: flex;
    }

    #codex-word-ui .codex-word-more svg {
      width: 10px;
      height: 10px;
    }

    .codex-word-actions-group:has(#codex-word-actions:empty) {
      display: none;
    }

    #codex-word-actions {
      display: flex;
      align-items: center;
      gap: 2px;
    }

    .codex-word-stack {
      display: flex;
      flex-direction: column;
      justify-content: center;
      gap: 6px;
    }

    .codex-word-stack.is-mini {
      gap: 0;
    }

    .codex-word-line {
      display: flex;
      align-items: center;
      gap: 1px;
    }

    .codex-word-sep {
      flex: none;
      width: 1px;
      height: 18px;
      margin: 0 5px;
      background: #e1e1e1;
    }

    .codex-word-btn {
      display: inline-flex;
      flex: none;
      align-items: center;
      justify-content: center;
      min-width: 26px;
      height: 26px;
      padding: 0 3px;
      border-radius: 5px;
      color: #424242;
    }

    .codex-word-stack.is-mini .codex-word-btn {
      min-width: 24px;
      height: 22px;
    }

    #codex-word-ui .codex-word-stack.is-mini svg {
      width: 17px;
      height: 17px;
    }

    #codex-word-ui .codex-word-caret svg {
      width: 9px;
      height: 9px;
      stroke-width: 2;
    }

    #codex-word-ui .codex-word-btn .codex-word-caret svg {
      margin-left: 1px;
    }

    .codex-word-btn.is-on {
      background: #e6e6e6;
    }

    .codex-word-btn.is-disabled {
      opacity: 0.38;
    }

    :is(.codex-word-btn, .codex-word-big):hover {
      background: #f0f0f0;
    }

    .codex-word-glyph {
      display: inline-block;
      min-width: 12px;
      font-size: 15px;
      line-height: 1;
      text-align: center;
    }

    .codex-word-glyph sub,
    .codex-word-glyph sup {
      color: #2f7bd6;
      font-size: 9px;
      line-height: 0;
    }

    .codex-word-combo {
      display: inline-flex;
      flex: none;
      align-items: center;
      justify-content: space-between;
      gap: 4px;
      height: 26px;
      margin-right: 3px;
      padding: 0 4px 0 8px;
      border: 1px solid #c6c6c6;
      border-radius: 5px;
      background: #fff;
      font-size: 13px;
      white-space: nowrap;
    }

    #codex-word-ui .codex-word-combo .codex-word-caret svg {
      width: 11px;
      height: 11px;
    }

    .codex-word-big {
      display: inline-flex;
      flex: none;
      flex-direction: column;
      align-items: center;
      gap: 4px;
      min-width: 40px;
      height: 66px;
      padding: 6px 4px 2px;
      border: 0;
      border-radius: 6px;
      background: transparent;
      color: #3b3b3b;
      font-size: 11.5px;
      line-height: 1.2;
      text-align: center;
      white-space: nowrap;
      cursor: default;
    }

    .codex-word-big-top {
      position: relative;
      display: flex;
      align-items: center;
    }

    .codex-word-big-top .codex-word-caret {
      position: absolute;
      left: 100%;
      margin-left: 1px;
    }

    .codex-word-big.has-caret {
      padding-right: 14px;
    }

    #codex-word-ui .codex-word-big svg {
      width: 32px;
      height: 32px;
      stroke-width: 1.35;
    }

    #codex-word-ui .codex-word-big .codex-word-caret svg {
      width: 9px;
      height: 9px;
      stroke-width: 2;
    }

    .codex-word-big .codex-word-label {
      color: #242424;
    }

    .codex-word-cmd {
      cursor: pointer;
    }

    .codex-word-cmd:disabled {
      opacity: 0.4;
      cursor: default;
    }

    .codex-word-fold {
      display: none;
    }

    /* 逐级收缩：普通部件直接隐藏；icon 只留图标；fold 整组折成一个按钮；small 论坛操作改成三行小按钮。 */
    :is(#codex-word-ui, #codex-word-status) [data-fit]:not([data-fit-mode]).is-folded,
    :is(#codex-word-ui, #codex-word-status) .is-fit-hidden {
      display: none !important;
    }

    [data-fit-mode="icon"].is-folded .codex-word-label {
      display: none;
    }

    [data-fit-mode="fold"].is-folded > :not(.codex-word-fold) {
      display: none;
    }

    [data-fit-mode="fold"].is-folded > .codex-word-fold {
      display: inline-flex;
    }

    [data-fit-mode="small"].is-folded #codex-word-actions {
      display: grid;
      grid-template-rows: repeat(3, 22px);
      grid-auto-flow: column;
      gap: 0 2px;
      align-content: center;
    }

    [data-fit-mode="small"].is-folded .codex-word-cmd {
      flex-direction: row;
      justify-content: flex-start;
      gap: 5px;
      min-width: 0;
      height: 22px;
      padding: 0 6px 0 3px;
      border-radius: 4px;
      font-size: 12px;
    }

    #codex-word-ui [data-fit-mode="small"].is-folded .codex-word-cmd svg {
      width: 17px;
      height: 17px;
    }

    #codex-word-status {
      bottom: 0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 18px;
      height: 30px;
      padding: 0 14px 0 22px;
      overflow: hidden;
      color: #616161;
      background: #f7f7f7;
      border-top: 1px solid #d6d6d6;
      font-size: 12px;
      transition: opacity 140ms ease, transform 180ms ease;
    }

    .codex-word-status-side {
      display: flex;
      flex: none;
      align-items: center;
      gap: 18px;
      white-space: nowrap;
    }

    .codex-word-status-side:last-child {
      margin-left: auto;
    }

    .codex-word-status-item {
      display: inline-flex;
      flex: none;
      align-items: center;
      gap: 6px;
    }

    #codex-word-status .codex-word-status-side svg {
      width: 17px;
      height: 17px;
    }

    .codex-word-views {
      display: flex;
      gap: 4px;
    }

    .codex-word-view {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 30px;
      height: 22px;
      border-radius: 4px;
    }

    .codex-word-view.is-active {
      color: #2b2b2b;
      background: #c4c4c4;
    }

    .codex-word-zoom {
      display: flex;
      flex: none;
      align-items: center;
      gap: 8px;
      font-size: 15px;
    }

    .codex-word-slider {
      position: relative;
      width: 96px;
      height: 3px;
      border-radius: 2px;
      background: #b0b0b0;
    }

    .codex-word-slider::after {
      position: absolute;
      top: 50%;
      left: 50%;
      width: 20px;
      height: 13px;
      content: "";
      border-radius: 7px;
      background: #fff;
      box-shadow: 0 0 0 0.5px rgb(0 0 0 / 22%), 0 1px 2px rgb(0 0 0 / 22%);
      transform: translate(-50%, -50%);
    }

    :is(.codex-word-tabs, .codex-word-ribbon) {
      transition: opacity 140ms ease, transform 180ms ease;
    }

    html.codex-word-ribbon-hidden :is(.codex-word-tabs, .codex-word-ribbon) {
      opacity: 0;
      transform: translateY(-6px);
      pointer-events: none;
    }

    html.codex-word-ribbon-hidden #codex-word-status {
      opacity: 0;
      transform: translateY(100%);
      pointer-events: none;
    }

    html.codex-word-dark #codex-word-ui {
      color: #e4e4e4;
      background: #1f1f1f;
    }

    html.codex-word-dark .codex-word-titlebar {
      background: #2b2b2b;
    }

    html.codex-word-dark .codex-word-pill-share {
      color: #0f2a55;
      background: #a8c7f7;
    }

    html.codex-word-dark .codex-word-autosave.is-on .codex-word-switch {
      background: #5b9cf0;
    }

    html.codex-word-dark .codex-word-autosave.is-on .codex-word-switch::before {
      background: #fff;
    }

    html.codex-word-dark .codex-word-tab:hover,
    html.codex-word-dark :is(.codex-word-btn, .codex-word-big):hover {
      background: #3a3a3a;
    }

    html.codex-word-dark .codex-word-tab.is-active::after {
      background: #5b9cf0;
    }

    html.codex-word-dark .codex-word-card {
      background: #2c2c2c;
      box-shadow: 0 0 0 0.5px rgb(255 255 255 / 9%), 0 1px 3px rgb(0 0 0 / 45%);
    }

    html.codex-word-dark :is(.codex-word-btn, .codex-word-big, .codex-word-big .codex-word-label) {
      color: #dcdcdc;
    }

    html.codex-word-dark .codex-word-btn.is-on {
      background: #474747;
    }

    html.codex-word-dark .codex-word-combo {
      color: #e4e4e4;
      background: #1f1f1f;
      border-color: #555;
    }

    html.codex-word-dark :is(.codex-word-group:not(.is-first)::before, .codex-word-sep) {
      background: #454545;
    }

    html.codex-word-dark #codex-word-ui svg .is-card {
      fill: #2c2c2c;
    }

    html.codex-word-dark #codex-word-status {
      color: #b4b4b4;
      background: #1f1f1f;
      border-color: #353535;
    }

    html.codex-word-dark .codex-word-view.is-active {
      color: #fff;
      background: #4a4a4a;
    }

    html.codex-word-dark .codex-word-slider {
      background: #666;
    }

    @media print {
      #codex-word-ui,
      #codex-word-status {
        display: none !important;
      }
    }

    @media (prefers-reduced-motion: reduce) {
      #codex-word-ui,
      #codex-word-status,
      .codex-word-tabs,
      .codex-word-ribbon,
      .codex-word-switch,
      .codex-word-switch::before {
        transition: none !important;
      }
    }
  `;

  const wordSvg = (body, box = 20) =>
    `<svg viewBox="0 0 ${box} ${box}" aria-hidden="true" focusable="false">${body}</svg>`;
  const wordText = (x, y, size, text, extra = "") =>
    `<text x="${x}" y="${y}" font-size="${size}" ${extra}>${text}</text>`;
  const wordLines = (x1, x2, ys) => ys.map((y) => `M${x1} ${y}H${x2}`).join("");

  const WORD_ICONS = {
    caret: wordSvg('<path d="M6 8l4 4 4-4"/>'),
    chevrons: wordSvg('<path d="M5 6l4 4-4 4M10.5 6l4 4-4 4"/>'),
    arrowRight: wordSvg('<path d="M7 4.5l7 5.5-7 5.5z" fill="currentColor" stroke="none"/>'),
    home: wordSvg('<path d="M3.75 9.25 10 3.75l6.25 5.5"/><path d="M5.25 8v8.25h3.5v-4.5h2.5v4.5h3.5V8"/>'),
    save: wordSvg(
      '<path d="M4.25 3.75h9.5l2.5 2.5v10H4.25z"/><path d="M6.75 3.75v3.5h5.5v-3.5"/><path d="M6.75 16.25v-5h6.5v5"/>',
    ),
    undo: wordSvg('<path d="M7.75 4.75 4.5 8l3.25 3.25"/><path d="M5 8h6.5a4.25 4.25 0 0 1 0 8.5H9"/>'),
    redo: wordSvg('<path d="M16 10a6 6 0 1 1-1.76-4.24"/><path d="M14.6 2.9v3.2h-3.2"/>'),
    print: wordSvg(
      '<path d="M6.25 7V3.75h7.5V7"/><rect x="3.75" y="7" width="12.5" height="6.5" rx="1.5"/><path d="M6.25 11.25h7.5v5h-7.5z"/>',
    ),
    more: wordSvg(
      '<g fill="currentColor" stroke="none"><circle cx="4.5" cy="10" r="1.2"/><circle cx="10" cy="10" r="1.2"/><circle cx="15.5" cy="10" r="1.2"/></g>',
    ),
    comment: wordSvg(
      '<path d="M3.75 5.25a1.5 1.5 0 0 1 1.5-1.5h9.5a1.5 1.5 0 0 1 1.5 1.5v6.5a1.5 1.5 0 0 1-1.5 1.5H9.5l-3.75 3v-3h-.5a1.5 1.5 0 0 1-1.5-1.5z"/>',
    ),
    pencil: wordSvg('<path d="M12.75 4.25l3 3-8.5 8.5H4.25v-3z"/><path d="M11 6l3 3"/>'),
    share: wordSvg(
      '<path d="M9 4.25H5.75a1.5 1.5 0 0 0-1.5 1.5v8.5a1.5 1.5 0 0 0 1.5 1.5h8.5a1.5 1.5 0 0 0 1.5-1.5V11"/><path d="M12 3.75h4.25V8"/><path d="M16.25 3.75 10 10"/>',
    ),
    search: wordSvg('<circle cx="8.75" cy="8.75" r="5"/><path d="m12.5 12.5 3.75 3.75"/>'),
    paste: wordSvg(
      '<rect x="4.5" y="6.5" width="16" height="21" rx="2" stroke="#b86f2c" fill="#f9dcb8"/><rect x="9" y="4" width="7" height="4.5" rx="1.2" class="is-card"/><rect x="13" y="12.5" width="13.5" height="16" rx="1.5" class="is-card"/>',
      32,
    ),
    cut: wordSvg('<circle cx="6" cy="14.5" r="2.25"/><circle cx="14" cy="14.5" r="2.25"/><path d="M7.4 12.7 13.5 3.5M12.6 12.7 6.5 3.5"/>'),
    copy: wordSvg('<rect x="7.25" y="6.25" width="9" height="11" rx="1"/><path d="M4.75 13.5V4.25a1 1 0 0 1 1-1h7"/>'),
    painter: wordSvg(
      '<rect x="3.5" y="3.5" width="10.5" height="4.5" rx="1" stroke="#c77a2e" fill="#f3c48f"/><path d="M14 5.75h2.25v4.5H9.75v2.25"/><rect x="8.5" y="12.5" width="2.5" height="5" rx=".8" fill="currentColor" stroke="none"/>',
    ),
    grow: wordSvg(`${wordText(2, 16.5, 15, "A")}<path d="M13.5 6.5 15.5 4.5 17.5 6.5"/>`),
    shrink: wordSvg(`${wordText(3, 16.5, 12, "A")}<path d="M13.5 4.5 15.5 6.5 17.5 4.5"/>`),
    clear: wordSvg(
      `${wordText(2, 14.5, 14, "A")}<path d="M10.5 15.5 15 11l3 3-2.5 2.5h-3z" stroke="#a2479d" fill="#e9b3e6"/>`,
    ),
    phonetic: wordSvg(`${wordText(7, 6.5, 6.5, "abc")}${wordText(4.5, 18, 12.5, "A", 'class="is-accent"')}`),
    charBorder: wordSvg(`<rect x="3" y="3" width="14" height="14" rx=".5"/>${wordText(5.9, 14.8, 11.5, "A")}`),
    effects: wordSvg(wordText(3.5, 16.5, 16, "A", 'font-weight="700" style="fill:#fff;stroke:#2f7bd6;stroke-width:.9"')),
    highlight: wordSvg(
      '<path d="M6.5 12.5 12.75 4.5l2.5 2-6.25 8H6.5z"/><path d="M11.25 6.25l2.5 2"/><rect x="3" y="16" width="14" height="2.5" rx=".5" fill="#fff200" stroke="none"/>',
    ),
    fontColor: wordSvg(`${wordText(4.3, 14.2, 13.5, "A")}<rect x="3" y="16" width="14" height="2.5" rx=".5" fill="#e3222b" stroke="none"/>`),
    shading: wordSvg(`<rect x="3" y="3" width="14" height="14" fill="#c9c9c9" stroke="none"/>${wordText(5.7, 14.8, 11.5, "A")}`),
    enclosed: wordSvg(`<circle cx="10" cy="10" r="7.25"/>${wordText(5.4, 13.6, 9.5, "字")}`),
    bullets: wordSvg(
      `<g class="is-accent" stroke="none"><circle cx="4.5" cy="5" r="1.4"/><circle cx="4.5" cy="10" r="1.4"/><circle cx="4.5" cy="15" r="1.4"/></g><path d="${wordLines(8.5, 17, [5, 10, 15])}"/>`,
    ),
    numbering: wordSvg(
      `${wordText(2.4, 7, 5.5, "1", 'class="is-accent"')}${wordText(2.4, 12, 5.5, "2", 'class="is-accent"')}${wordText(2.4, 17, 5.5, "3", 'class="is-accent"')}<path d="${wordLines(8.5, 17, [5, 10, 15])}"/>`,
    ),
    multilevel: wordSvg(
      `${wordText(2.4, 7, 5.5, "1", 'class="is-accent"')}${wordText(5, 12, 5.5, "a", 'class="is-accent"')}${wordText(7.6, 17, 5.5, "i", 'class="is-accent"')}<path d="M7 5h10M10 10h7M11.5 15H17"/>`,
    ),
    outdent: wordSvg('<path d="M9 5h8M9 10h8M9 15h8M3 5h3M3 15h3"/><path d="M6 7.75 3.5 10 6 12.25" stroke="#2f7bd6"/>'),
    indent: wordSvg('<path d="M9 5h8M9 10h8M9 15h8M3 5h3M3 15h3"/><path d="M3.5 7.75 6 10l-2.5 2.25" stroke="#2f7bd6"/>'),
    cjk: wordSvg(`<path d="M3 4.5h14M5 3 3 4.5 5 6M15 3l2 1.5L15 6" stroke="#2f7bd6"/>${wordText(4.5, 18, 13, "A")}`),
    sort: wordSvg(
      `${wordText(2.5, 8.5, 7.5, "A", 'class="is-accent"')}${wordText(2.5, 17.5, 7.5, "Z")}<path d="M14 3.5v12.75M11.5 13.75 14 16.25l2.5-2.5"/>`,
    ),
    marks: wordSvg('<path d="M16.5 4v6H5"/><path d="M8 7 5 10l3 3"/><path d="M3 15.5h3.5M5.25 14l1.5 1.5-1.5 1.5" stroke="#2f7bd6"/>'),
    alignLeft: wordSvg('<path d="M3 4.5h14M3 8h9M3 11.5h14M3 15h9"/>'),
    alignCenter: wordSvg('<path d="M3 4.5h14M5.5 8h9M3 11.5h14M5.5 15h9"/>'),
    alignRight: wordSvg('<path d="M3 4.5h14M8 8h9M3 11.5h14M8 15h9"/>'),
    justify: wordSvg(`<path d="${wordLines(3, 17, [4.5, 8, 11.5, 15])}"/>`),
    distributed: wordSvg(
      `<path d="M3 3v14M17 3v14"/><path d="${wordLines(5.5, 14.5, [9.5, 12.5, 15.5])}"/><path d="M5.5 5.5h9M7 4 5.5 5.5 7 7M13 4l1.5 1.5L13 7" stroke="#2f7bd6"/>`,
    ),
    spacing: wordSvg(
      '<path d="M9 4.5h8M9 8h8M9 11.5h8M9 15h8"/><path d="M4.5 3.5v13M3 5.25l1.5-1.75L6 5.25M3 14.75l1.5 1.75L6 14.75" stroke="#2f7bd6"/>',
    ),
    fill: wordSvg(
      '<path d="M4 9.5 9.5 4l5.5 5.5-5.5 5.5z"/><path d="M4 9.5h11"/><path d="M16 12.25s1.5 1.8 1.5 2.8a1.5 1.5 0 0 1-3 0c0-1 1.5-2.8 1.5-2.8z" fill="#2f7bd6" stroke="none"/>',
    ),
    borders: wordSvg(
      '<path d="M3.5 3.5h13v13h-13zM3.5 10h13M10 3.5v13" stroke-dasharray="1.2 1.8"/><path d="M3 16.75h14" stroke-width="1.6"/>',
    ),
    styles: wordSvg(
      `${wordText(2, 25, 26, "A", 'font-weight="300"')}<path d="M29 7.5 17.5 21.5" stroke-width="1.6"/><path d="M17.5 21.5c-2.6-.6-4.6 1.1-4.4 3.3.1 1.2-.5 2.2-1.6 2.9 3.3.5 6.9-.8 7.4-3.4z" fill="#2f7bd6" stroke="#2f7bd6"/>`,
      32,
    ),
    stylePane: wordSvg(
      '<rect x="4.5" y="4.5" width="21" height="21" rx="2"/><path d="M28.5 12 18 24.5" stroke-width="1.6"/><path d="M18 24.5c-2.3-.4-4 1.2-3.8 3.1l-.1.9c2.6.5 5.6-.6 5.9-2.9z" fill="#2f7bd6" stroke="#2f7bd6"/>',
      32,
    ),
    dictate: wordSvg(
      '<rect x="12" y="4" width="8" height="15" rx="4" fill="#bcdcf6" stroke="#3b82d4"/><path d="M8.5 15a7.5 7.5 0 0 0 15 0M16 22.5V28"/>',
      32,
    ),
    addins: wordSvg(
      '<g stroke="#d96f55"><rect x="5.5" y="5.5" width="9.5" height="9.5"/><rect x="17" y="5.5" width="9.5" height="9.5"/><rect x="5.5" y="17" width="9.5" height="9.5"/><rect x="17" y="17" width="9.5" height="9.5"/></g>',
      32,
    ),
    editor: wordSvg(
      '<path d="M7.5 25.5 22 8.5l3.5 3L11 28.5H7.5z" stroke="#3b82d4" fill="#d6e7fb"/><path d="M13 27.5h14M17 23h10M21.5 18.5H27" stroke="#3b82d4"/>',
      32,
    ),
    fontGroup: wordSvg(wordText(5, 27, 30, "A", 'font-weight="300"'), 32),
    paragraphGroup: wordSvg(`<path d="${wordLines(4, 28, [8, 13, 18, 23])}"/>`, 32),
    compose: wordSvg(
      '<path d="M8.5 4.5h10l6 6v17h-16z"/><path d="M18.5 4.5v6h6"/><path d="M16.5 15v9M12 19.5h9" stroke="#2f7bd6" stroke-width="1.8"/>',
      32,
    ),
    like: wordSvg(
      '<path d="M16 26.5s-10-5.8-10-12.8a5.4 5.4 0 0 1 10-2.9 5.4 5.4 0 0 1 10 2.9c0 7-10 12.8-10 12.8z" stroke="#d13438" fill="#fbe3e4"/>',
      32,
    ),
    bookmark: wordSvg('<path d="M9.5 5h13v22l-6.5-4.75L9.5 27z" stroke="#2f7bd6" fill="#dcebfb"/>', 32),
    star: wordSvg(
      '<path d="m16 4.75 3.4 6.9 7.6 1.1-5.5 5.4 1.3 7.6L16 22.15l-6.8 3.6 1.3-7.6L5 12.75l7.6-1.1z" stroke="#c28a00" fill="#fdf0c2"/>',
      32,
    ),
    shareTopic: wordSvg(
      '<path d="M13 8.5H8a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-5"/><path d="M18 5.5h8v8M26 5.5 15 16.5" stroke="#2f7bd6" stroke-width="1.8"/>',
      32,
    ),
    reply: wordSvg(
      '<path d="M13 8.5 6 15.5l7 7" stroke="#2f7bd6" stroke-width="1.8"/><path d="M6.5 15.5H19a7 7 0 0 1 7 7V26"/>',
      32,
    ),
    quote: wordSvg(
      '<path d="M7 22v-5.5A7 7 0 0 1 14 9.5M7 22h5.5v-5.5H7M18.5 22v-5.5a7 7 0 0 1 7-7M18.5 22H24v-5.5h-5.5" stroke="#2f7bd6"/>',
      32,
    ),
    dislike: wordSvg(
      '<g transform="rotate(180 16 16)"><path d="M10.5 14v12.5H6V14z"/><path d="M10.5 14l5-8.5a2.2 2.2 0 0 1 3.8 2.1L17.75 12h6.5a2 2 0 0 1 2 2.3l-1.6 9.6a2 2 0 0 1-2 1.6H10.5" stroke="#5c6f8a" fill="#e4e9f0"/></g>',
      32,
    ),
    drumstick: wordSvg(
      '<path d="M19 5a8 8 0 0 1 6.6 12.5c-2.2 3.2-6.7 3.9-9.6 1.6L11.4 23.7l.3 1.7a2.2 2.2 0 1 1-2.7 2.7 2.2 2.2 0 1 1-2.5-3.5l1.7.3 4.6-4.6c-2.3-2.9-1.6-7.4 1.6-9.6A8 8 0 0 1 19 5z" stroke="#b86f2c" fill="#f6d7ae"/>',
      32,
    ),
    proofing: wordSvg(
      '<circle cx="7.5" cy="6.5" r="2.75"/><path d="M2.5 16.5c.4-3.2 2.4-5 5-5 1.1 0 2.1.3 2.9.9"/><path d="M11.5 16.75l.6-2.35 5-5 1.75 1.75-5 5z"/>',
    ),
    a11y: wordSvg(
      '<circle cx="9" cy="3.75" r="1.6"/><path d="M3.5 7 9 8.25 14.5 7M9 8.25v4.25L6.5 17.5M9 12.5l1.5 2.5"/><circle cx="14.75" cy="14.75" r="3.25" fill="currentColor" stroke="none"/><path d="m13.35 14.8.95.95 1.85-1.9" stroke="#f7f7f7"/>',
    ),
    focus: wordSvg('<path d="M11.5 3.5h5v5M16.5 3.5 11 9M8.5 16.5h-5v-5M3.5 16.5 9 11"/>'),
    printLayout: wordSvg(
      '<path d="M5 2.75h7l3.5 3.5v11H5z" fill="currentColor" stroke="none"/><path d="M7.5 9h5M7.5 11.5h5M7.5 14h3.5" stroke="#c4c4c4"/>',
    ),
    webLayout: wordSvg(
      '<circle cx="10" cy="10" r="7"/><path d="M3 10h14M10 3c1.9 2.1 2.8 4.4 2.8 7S11.9 14.9 10 17c-1.9-2.1-2.8-4.4-2.8-7S8.1 5.1 10 3z"/>',
    ),
    outline: wordSvg('<path d="M3.5 5.5h13M6.5 10h10M3.5 14.5h13"/>'),
    draft: wordSvg('<path d="M3.5 4h9M3.5 8h6M3.5 12h3.5"/><path d="M8.5 16.75l.5-2.5 6.75-6.75 2 2L11 16.25z"/>'),
  };

  const wordButton = (icon, title, extra = "") =>
    `<span class="codex-word-btn${extra}" title="${title}">${WORD_ICONS[icon]}</span>`;
  const wordDropdown = (icon, title, extra = "") =>
    `<span class="codex-word-btn${extra}" title="${title}">${WORD_ICONS[icon]}<span class="codex-word-caret">${WORD_ICONS.caret}</span></span>`;
  const wordBig = (icon, label, { caret = false, fold = false, attrs = "" } = {}) =>
    `<span class="codex-word-big${caret ? " has-caret" : ""}${fold ? " codex-word-fold" : ""}" title="${label.replace("<br>", "")}" ${attrs}><span class="codex-word-big-top">${WORD_ICONS[icon]}${caret ? `<span class="codex-word-caret">${WORD_ICONS.caret}</span>` : ""}</span><span class="codex-word-label">${label}</span></span>`;

  /* 功能区里的论坛操作：长成 Word 的大按钮（图标在上、文字在下），窄屏时缩成小按钮。 */
  const wordCommand = ({ icon, label, title = label, href }) => {
    const command = document.createElement(href ? "a" : "button");
    command.className = "codex-word-big codex-word-cmd";
    if (href) {
      command.href = href;
    } else {
      command.type = "button";
    }
    command.title = title;
    command.innerHTML = `${WORD_ICONS[icon] || ""}<span class="codex-word-label"></span>`;
    command.querySelector(".codex-word-label").textContent = label;
    return command;
  };

  /* 按 data-fit 的顺序逐级收缩，直到这一行不再溢出；实在放不下时这一行改为可横向滚动。 */
  const fitWordRow = (row, overflowing = () => row.scrollWidth > row.clientWidth + 1) => {
    const parts = [...row.querySelectorAll("[data-fit], [data-fit-hide]")];
    const levels = [
      ...new Set(parts.flatMap((part) => [part.dataset.fit, part.dataset.fitHide]).filter(Boolean).map(Number)),
    ].sort((a, b) => a - b);

    parts.forEach((part) => part.classList.remove("is-folded", "is-fit-hidden"));
    for (const level of levels) {
      if (!overflowing()) break;
      parts.forEach((part) => {
        if (Number(part.dataset.fit) === level) part.classList.add("is-folded");
        if (Number(part.dataset.fitHide) === level) part.classList.add("is-fit-hidden");
      });
    }

    const groups = [...row.children].filter((child) => child.classList.contains("codex-word-group"));
    const first = groups.find((group) => group.getClientRects().length > 0);
    groups.forEach((group) => group.classList.toggle("is-first", group === first));
    row.classList.toggle("is-overflowing", overflowing());
  };

  /* 标题栏另外要求居中的文档名不压到两侧按钮上。 */
  const titleOverflowing = (row) => () => {
    const left = row.querySelector(".codex-word-title-left");
    const right = row.querySelector(".codex-word-title-right");
    const title = row.querySelector(".codex-word-doc-title");
    if (left.scrollWidth > left.clientWidth + 1 || row.scrollWidth > row.clientWidth + 1) return true;
    if (!title.getClientRects().length) return false;
    const box = title.getBoundingClientRect();
    return box.left < left.getBoundingClientRect().right + 12 || box.right > right.getBoundingClientRect().left - 12;
  };

  const createWordChrome = ({ home = "/", insert = home, onToggle } = {}) => {
    const I = WORD_ICONS;
    const tabs = ["绘图", "设计", "布局", "引用", "邮件", "审阅", "视图"]
      .map((name, index, all) => `<span class="codex-word-tab" data-fit="${all.length - index}">${name}</span>`)
      .join("");

    document.body.insertAdjacentHTML(
      "afterbegin",
      `
        <div id="codex-word-ui" role="banner" aria-label="Word 摸鱼导航">
          <div class="codex-word-titlebar">
            <div class="codex-word-title-left" aria-hidden="true">
              <span class="codex-word-traffic"><i></i><i></i><i></i></span>
              <span class="codex-word-autosave" data-fit="22"><span class="codex-word-label">自动保存</span><i class="codex-word-switch"></i></span>
              <span class="codex-word-tb-btn" title="主页" data-fit="20">${I.home}</span>
              <span class="codex-word-tb-btn" title="保存" data-fit="18">${I.save}</span>
              <span class="codex-word-tb-split is-disabled" title="撤消" data-fit="16">${I.undo}<span class="codex-word-caret">${I.caret}</span></span>
              <span class="codex-word-tb-btn is-disabled" title="无法重复" data-fit="14">${I.redo}</span>
              <span class="codex-word-tb-btn" title="打印" data-fit="12">${I.print}</span>
              <span class="codex-word-tb-btn" title="自定义快速访问工具栏" data-fit="10">${I.more}</span>
              <span class="codex-word-tb-btn codex-word-overflow" title="更多命令">${I.chevrons}</span>
            </div>
            <button class="codex-word-doc-title" type="button" title="收起或展开功能区" data-fit="5">工作记录</button>
            <div class="codex-word-title-right" aria-hidden="true">
              <span class="codex-word-tb-btn" title="批注" data-fit="34">${I.comment}</span>
              <span class="codex-word-pill codex-word-pill-edit" title="编辑" data-fit="30" data-fit-mode="icon" data-fit-hide="36">${I.pencil}<span class="codex-word-label">编辑</span><span class="codex-word-caret">${I.caret}</span></span>
              <span class="codex-word-pill codex-word-pill-share" title="共享" data-fit="32" data-fit-mode="icon">${I.share}<span class="codex-word-label">共享</span><span class="codex-word-caret">${I.caret}</span></span>
              <span class="codex-word-round" title="搜索 (Cmd + Ctrl + U)">${I.search}</span>
            </div>
          </div>
          <nav class="codex-word-tabs" aria-label="功能区选项卡">
            <a class="codex-word-tab is-active" href="${home}">开始</a>
            <a class="codex-word-tab" href="${insert}">插入</a>
            ${tabs}
            <span class="codex-word-overflow" aria-hidden="true">${I.chevrons}</span>
          </nav>
          <div class="codex-word-ribbon">
            <div class="codex-word-card">
              <div class="codex-word-group" data-fit-hide="48" aria-hidden="true">
                ${wordBig("paste", "粘贴", { caret: true })}
                <span class="codex-word-stack is-mini" data-fit="46">
                  ${wordButton("cut", "剪切", " is-disabled")}
                  ${wordButton("copy", "复制", " is-disabled")}
                  ${wordButton("painter", "格式刷")}
                </span>
              </div>
              <div class="codex-word-group" data-fit="14" data-fit-mode="fold" data-fit-hide="44" aria-hidden="true">
                <div class="codex-word-stack">
                  <div class="codex-word-line">
                    <span class="codex-word-combo" style="width: 104px">等线 (正文)<span class="codex-word-caret">${I.caret}</span></span>
                    <span class="codex-word-combo" style="width: 54px">11<span class="codex-word-caret">${I.caret}</span></span>
                    ${wordButton("grow", "增大字体")}${wordButton("shrink", "缩小字体")}
                    <span class="codex-word-sep"></span>
                    <span class="codex-word-btn" title="更改大小写"><span class="codex-word-glyph">Aa</span><span class="codex-word-caret">${I.caret}</span></span>
                    <span class="codex-word-sep"></span>
                    ${wordButton("clear", "清除格式")}${wordButton("phonetic", "拼音指南")}${wordButton("charBorder", "字符边框")}
                  </div>
                  <div class="codex-word-line">
                    <span class="codex-word-btn" title="加粗"><b class="codex-word-glyph">B</b></span>
                    <span class="codex-word-btn" title="倾斜"><i class="codex-word-glyph" style="font-family: Georgia, serif">I</i></span>
                    <span class="codex-word-btn" title="下划线"><u class="codex-word-glyph">U</u><span class="codex-word-caret">${I.caret}</span></span>
                    <span class="codex-word-btn" title="删除线"><s class="codex-word-glyph">ab</s></span>
                    <span class="codex-word-btn" title="下标"><span class="codex-word-glyph">x<sub>2</sub></span></span>
                    <span class="codex-word-btn" title="上标"><span class="codex-word-glyph">x<sup>2</sup></span></span>
                    <span class="codex-word-sep"></span>
                    ${wordDropdown("effects", "文本效果")}${wordDropdown("highlight", "文本突出显示颜色")}${wordDropdown("fontColor", "字体颜色")}
                    ${wordButton("shading", "字符底纹")}${wordButton("enclosed", "带圈字符")}
                  </div>
                </div>
                ${wordBig("fontGroup", "字体", { caret: true, fold: true })}
              </div>
              <div class="codex-word-group" data-fit="12" data-fit-mode="fold" data-fit-hide="42" aria-hidden="true">
                <div class="codex-word-stack">
                  <div class="codex-word-line">
                    ${wordDropdown("bullets", "项目符号")}${wordDropdown("numbering", "编号")}${wordDropdown("multilevel", "多级列表")}
                    <span class="codex-word-sep"></span>
                    ${wordButton("outdent", "减少缩进量")}${wordButton("indent", "增加缩进量")}
                    <span class="codex-word-sep"></span>${wordDropdown("cjk", "中文版式")}
                    <span class="codex-word-sep"></span>${wordButton("sort", "排序")}
                    <span class="codex-word-sep"></span>${wordButton("marks", "显示编辑标记")}
                  </div>
                  <div class="codex-word-line">
                    ${wordButton("alignLeft", "左对齐", " is-on")}${wordButton("alignCenter", "居中")}${wordButton("alignRight", "右对齐")}${wordButton("justify", "两端对齐")}${wordButton("distributed", "分散对齐")}
                    <span class="codex-word-sep"></span>
                    ${wordDropdown("spacing", "行和段落间距")}
                    <span class="codex-word-sep"></span>${wordDropdown("fill", "底纹")}${wordDropdown("borders", "边框")}
                  </div>
                </div>
                ${wordBig("paragraphGroup", "段落", { caret: true, fold: true })}
              </div>
              <div class="codex-word-group" data-fit-hide="40" aria-hidden="true">
                ${wordBig("styles", "样式", { caret: true })}
                ${wordBig("stylePane", "样式<br>窗格", { attrs: 'data-fit="10"' })}
              </div>
              <div class="codex-word-group codex-word-actions-group" data-fit="30" data-fit-mode="small">
                <div id="codex-word-actions" data-fit="50" data-fit-mode="icon" role="toolbar" aria-label="论坛快捷操作"></div>
              </div>
              <div class="codex-word-group is-trailing" data-fit="24" aria-hidden="true">${wordBig("dictate", "听写")}</div>
              <div class="codex-word-group is-trailing" data-fit="22" aria-hidden="true">${wordBig("addins", "加载项")}</div>
              <div class="codex-word-group is-trailing" data-fit="20" aria-hidden="true">${wordBig("editor", "编辑器")}</div>
              <span class="codex-word-more" aria-hidden="true">${I.arrowRight}</span>
            </div>
          </div>
        </div>
        <div id="codex-word-status" aria-hidden="true">
          <span class="codex-word-status-side">
            <span class="codex-word-status-item" data-fit="18">第 1 页，共 1 页</span>
            <span class="codex-word-status-item" id="codex-word-count" data-fit="16">0 个字</span>
            <span class="codex-word-status-item" title="校对" data-fit="14">${I.proofing}</span>
            <span class="codex-word-status-item" data-fit="12">简体中文(中国大陆)</span>
            <span class="codex-word-status-item" data-fit="10">${I.a11y}辅助功能: 一切就绪</span>
          </span>
          <span class="codex-word-status-side">
            <span class="codex-word-status-item" data-fit="20" data-fit-mode="icon" data-fit-hide="26">${I.focus}<span class="codex-word-label">专注</span></span>
            <span class="codex-word-views" data-fit="22">
              <span class="codex-word-view is-active" title="页面视图">${I.printLayout}</span>
              <span class="codex-word-view" title="Web 版式">${I.webLayout}</span>
              <span class="codex-word-view" title="大纲">${I.outline}</span>
              <span class="codex-word-view" title="草稿">${I.draft}</span>
            </span>
            <span class="codex-word-zoom" data-fit="24"><span>−</span><span class="codex-word-slider"></span><span>+</span></span>
            <span class="codex-word-status-item">100%</span>
          </span>
        </div>
      `,
    );

    const ui = document.getElementById("codex-word-ui");
    const status = document.getElementById("codex-word-status");
    const actions = document.getElementById("codex-word-actions");
    const wordCount = document.getElementById("codex-word-count");
    const titlebar = ui.querySelector(".codex-word-titlebar");
    const rows = [
      [titlebar, titleOverflowing(titlebar)],
      [ui.querySelector(".codex-word-tabs")],
      [ui.querySelector(".codex-word-card")],
      [status],
    ];

    let fitQueued = false;
    const fit = () => {
      if (fitQueued) return;
      fitQueued = true;
      requestAnimationFrame(() => {
        fitQueued = false;
        rows.forEach(([row, overflowing]) => fitWordRow(row, overflowing));
      });
    };

    let width = 0;
    new ResizeObserver(() => {
      if (ui.clientWidth === width) return;
      width = ui.clientWidth;
      fit();
    }).observe(ui);
    const contentObserver = new MutationObserver(fit);
    const watch = { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ["disabled", "hidden"] };
    contentObserver.observe(actions, watch);
    contentObserver.observe(wordCount, watch);
    document.fonts?.ready.then(fit);

    ui.querySelector(".codex-word-doc-title").addEventListener("click", () => {
      document.documentElement.classList.toggle("codex-word-ribbon-hidden");
      onToggle?.();
    });

    /* 自动保存开关只是样子，没有实际功能；开关状态按站点记在 localStorage，换页后保持。存储不可用时只在当前页切换。 */
    const autosave = ui.querySelector(".codex-word-autosave");
    try {
      autosave.classList.toggle("is-on", localStorage.getItem("codex-word-autosave") === "on");
    } catch {}
    autosave.addEventListener("click", () => {
      const on = autosave.classList.toggle("is-on");
      try {
        localStorage.setItem("codex-word-autosave", on ? "on" : "off");
      } catch {}
    });

    return { actions, wordCount };
  };
  /* ===== Word 界面结束 ===== */

  const root = document.documentElement;
  let manualToggleUntil = 0;
  root.classList.add("codex-word-theme", "codex-word-discourse");

  const style = document.createElement("style");
  style.id = "codex-discourse-word-style";
  style.textContent = `${WORD_CHROME_CSS}
    :root {
      --word-ink: #242424;
      --word-line: #d7d7d7;
      --word-paper: #fff;
      --word-workspace: #f5f5f5;
      --codex-discourse-header-height: 52px;
      --header-offset: calc(var(--word-ui-height) + var(--codex-discourse-header-height));
    }

    html.codex-word-theme,
    html.codex-word-theme body {
      background: var(--word-workspace) !important;
      background-image: none !important;
    }

    html.codex-word-theme body {
      box-sizing: border-box !important;
      min-height: 100vh !important;
      padding-top: var(--word-ui-height) !important;
      padding-bottom: var(--word-status-height) !important;
      color: var(--word-ink) !important;
      font-family: Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
      transition: padding-top 180ms ease;
    }

    /* 只隐藏头像与对应占位；正文图片、Logo、表情和 Boost 文案不受影响。 */
    img.avatar,
    .topic-avatar,
    .post-avatar,
    .topic-list .posters,
    .topic-list-header .posters,
    .topic-map__users-list,
    #user-card .user-card-avatar,
    .user-main .user-profile-avatar {
      display: none !important;
    }

    .discourse-boosts__bubble > a:has(> img.avatar:only-child) {
      display: none !important;
    }

    .post__row {
      gap: 0 !important;
    }

    /* 当前页面用不上的操作（如列表页的点赞、回复）直接隐藏。 */
    html.codex-word-discourse .codex-word-cmd:disabled {
      display: none;
    }

    html.codex-word-theme .codex-word-picker-hosted .discourse-reactions-picker.is-expanded {
      position: fixed !important;
      z-index: 2147483647 !important;
      top: var(--codex-reaction-picker-y) !important;
      left: var(--codex-reaction-picker-x) !important;
      transform: translateX(-50%) !important;
    }

    html.codex-word-theme .d-header-wrap {
      top: var(--word-ui-height) !important;
      transition: top 160ms ease;
    }

    html.codex-word-theme #main-outlet {
      box-sizing: border-box !important;
      width: min(calc(100% - 64px), 1180px) !important;
      max-width: 1180px !important;
      min-height: calc(100vh - var(--word-ui-height) - 92px);
      margin: 28px auto 48px !important;
      padding: 28px 40px 44px !important;
      background: var(--word-paper) !important;
      border: 1px solid #d0d0d0;
      box-shadow: 0 1px 5px rgb(0 0 0 / 18%);
    }

    html.codex-word-theme #main-outlet > .welcome-banner {
      padding-top: 0 !important;
    }

    html.codex-word-theme :is(.topic-list, .topic-list-body, .topic-post, .topic-body) {
      max-width: 100% !important;
    }

    html.codex-word-theme #reply-control.open:not(.fullscreen) {
      box-sizing: border-box !important;
      right: auto !important;
      left: 50% !important;
      width: min(calc(100vw - 64px), 1180px) !important;
      max-height: min(48vh, 430px) !important;
      transform: translateX(-50%);
      border: 1px solid #bdbdbd;
      box-shadow: 0 0 12px rgb(0 0 0 / 22%);
    }

    html.codex-word-theme #reply-control.open:not(.fullscreen) .reply-area {
      width: 100% !important;
    }

    /*
     * Word 界面相当于窗口边框：原本铺满视口或贴着视口边缘的浮层，
     * 都排进功能区和状态栏之间，界面本身不盖住任何内容。
     */
    html.codex-word-theme #reply-control:not(.fullscreen) {
      bottom: var(--word-status-height) !important;
    }

    html.codex-word-theme.fullscreen-composer #reply-control.fullscreen {
      bottom: var(--word-status-height);
      height: calc(
        var(--composer-vh, 1vh) * 100 - var(--word-ui-height) - var(--word-status-height)
      ) !important;
    }

    html.codex-word-theme .d-modal {
      top: var(--word-ui-height);
      height: calc(100% - var(--word-ui-height) - var(--word-status-height));
    }

    html.codex-word-theme .d-modal__container {
      max-height: min(
        var(--modal-max-height, 80vh),
        calc(100vh - var(--word-ui-height) - var(--word-status-height) - 24px)
      ) !important;
    }

    /* 灯箱遮罩仍铺满窗口；图片由 fitLightbox() 让开两栏，按钮和图片说明在这里挪进来。 */
    html.codex-word-theme .pswp__top-bar {
      top: var(--word-ui-height);
    }

    html.codex-word-theme .pswp__button--arrow {
      margin-top: calc((var(--word-ui-height) - var(--word-status-height)) / 2 - 50px);
    }

    html.codex-word-theme .pswp__caption {
      bottom: calc(var(--word-status-height) + var(--safe-area-inset-bottom, 0px));
    }

    html.codex-word-theme .chat-drawer-outlet-container {
      bottom: var(--word-status-height);
      max-height: calc(100% - var(--header-offset) - 15px - var(--word-status-height));
    }

    html.codex-word-theme .with-topic-progress {
      bottom: calc(
        env(safe-area-inset-bottom) + var(--composer-height, 0px) + var(--word-status-height)
      );
    }

    html.codex-word-theme .menu-panel.drop-down {
      max-height: calc(100dvh - var(--header-offset) - 1em - var(--word-status-height));
    }

    @media (width >= 48rem) {
      html.codex-word-theme .sidebar-wrapper {
        height: calc(
          var(--composer-vh, 1dvh) * 100 - var(--main-outlet-offset, 0px) - var(--word-status-height)
        );
      }
    }

    /* LINUX DO 站点组件的操作提示，原本离视口顶部 10%。 */
    html.codex-word-theme #messageToast {
      top: calc(var(--word-ui-height) + 10%);
    }

    html.codex-word-dark {
      --word-ink: #f3f3f3;
      --word-line: #454545;
      --word-paper: #252526;
      --word-workspace: #1f1f1f;
    }

    html.codex-word-dark :is(#main-outlet, #reply-control.open:not(.fullscreen)) {
      color: var(--word-ink) !important;
      border-color: var(--word-line);
    }

    @media (max-width: 900px) {
      html.codex-word-theme #main-outlet {
        width: calc(100% - 24px) !important;
        margin: 12px auto 36px !important;
        padding: 20px 18px 34px !important;
      }

      html.codex-word-theme #reply-control.open:not(.fullscreen) {
        width: calc(100vw - 24px) !important;
      }
    }

    @media print {
      html.codex-word-theme body {
        padding: 0 !important;
        background: #fff !important;
      }

      html.codex-word-theme .d-header-wrap {
        top: 0 !important;
      }

      html.codex-word-theme #main-outlet {
        width: 100% !important;
        max-width: none !important;
        margin: 0 !important;
        border: 0;
        box-shadow: none;
      }
    }

    @media (prefers-reduced-motion: reduce) {
      html.codex-word-theme body,
      html.codex-word-theme .d-header-wrap {
        transition: none !important;
      }
    }
  `;
  (document.head || root).append(style);

  /*
   * 图片灯箱（PhotoSwipe）按整个视口给图片排版。它在 init() 之前把实例写到 window.pswp，
   * 这里接住实例，把功能区和状态栏的高度加进内边距，图片的适配、缩放和拖动范围就都落在两栏之间。
   */
  const fitLightbox = (pswp) => {
    const { paddingFn, padding } = pswp.options;
    pswp.options.paddingFn = (viewportSize, itemData, index) => {
      const base = paddingFn?.(viewportSize, itemData, index) ?? padding ?? {};
      const chrome = getComputedStyle(root);
      return {
        ...base,
        top: (base.top || 0) + parseFloat(chrome.getPropertyValue("--word-ui-height")),
        bottom: (base.bottom || 0) + parseFloat(chrome.getPropertyValue("--word-status-height")),
      };
    };
  };

  const watchLightbox = () => {
    let current = window.pswp;
    Object.defineProperty(window, "pswp", {
      configurable: true,
      get: () => current,
      set: (pswp) => {
        current = pswp;
        if (typeof pswp?.on !== "function") return;
        fitLightbox(pswp);
        // 灯箱关闭时 PhotoSwipe 会 delete window.pswp，连同这个访问器一起删掉，关闭后重新装上。
        pswp.on("destroy", () => queueMicrotask(watchLightbox));
      },
    });
  };
  watchLightbox();

  const isTopicPage = () => location.pathname.startsWith("/t/");
  const mainPost = () => document.querySelector('#post_1, .topic-post[data-post-number="1"]');
  const officialBookmark = () =>
    mainPost()?.querySelector(".post-action-menu__bookmark") ||
    document.querySelector(".topic-footer-main-buttons .bookmark-menu-trigger");
  const reactionWrapper = () => mainPost()?.querySelector(".discourse-reactions-reaction-button");
  const dispatchReactionPointer = (type) => {
    const wrapper = reactionWrapper();
    if (!wrapper) return;
    const event = new Event(type, { bubbles: true, cancelable: true });
    Object.defineProperty(event, "pointerType", { value: "mouse" });
    wrapper.dispatchEvent(event);
  };
  const openOfficialBookmark = () => {
    const bookmark = officialBookmark();
    if (bookmark) return bookmark.click();

    const post = mainPost();
    const showMore = post?.querySelector(".post-action-menu__show-more");
    if (!post || !showMore) return;

    const observer = new MutationObserver(() => {
      const revealedBookmark = officialBookmark();
      if (!revealedBookmark) return;
      observer.disconnect();
      revealedBookmark.click();
    });
    observer.observe(post, { childList: true, subtree: true });
    showMore.click();
    setTimeout(() => observer.disconnect(), 1000);
  };

  const actions = [
    {
      icon: "compose",
      label: "发帖",
      find: () => document.body,
      run: () => {
        const event = new KeyboardEvent("keypress", { key: "c", code: "KeyC", bubbles: true, cancelable: true });
        Object.defineProperty(event, "which", { value: 99 });
        document.dispatchEvent(event);
      },
    },
    {
      icon: "like",
      label: "点赞",
      find: () => isTopicPage() && mainPost()?.querySelector(".reaction-button"),
      enter: (button) => {
        const wrapper = reactionWrapper();
        if (!wrapper) return;
        const rect = button.getBoundingClientRect();
        root.style.setProperty("--codex-reaction-picker-x", `${rect.left + rect.width / 2}px`);
        root.style.setProperty("--codex-reaction-picker-y", `${rect.bottom - 1}px`);
        wrapper.closest(".discourse-reactions-actions")?.classList.add("codex-word-picker-hosted");
        dispatchReactionPointer("pointerover");
      },
      leave: () => dispatchReactionPointer("pointerout"),
    },
    {
      icon: "bookmark",
      label: "书签",
      find: () =>
        isTopicPage() && (officialBookmark() || mainPost()?.querySelector(".post-action-menu__show-more")),
      run: openOfficialBookmark,
    },
    {
      icon: "shareTopic",
      label: "分享",
      find: () => isTopicPage() && document.querySelector(".topic-footer-main-buttons .share-and-invite"),
    },
    {
      icon: "reply",
      label: "回复",
      find: () =>
        isTopicPage() &&
        (mainPost()?.querySelector(".post-action-menu__reply") ||
          document.querySelector(".topic-footer-main-buttons .create")),
    },
  ];

  const sync = () => {
    root.classList.toggle("codex-word-dark", getComputedStyle(root).colorScheme.includes("dark"));

    const header = document.querySelector(".d-header-wrap");
    if (header) {
      root.style.setProperty("--codex-discourse-header-height", `${header.offsetHeight}px`);
    }

    const wordCount = document.getElementById("codex-word-count");
    const mainPostContent = mainPost()?.querySelector(".cooked");
    const count = mainPostContent ? Array.from(mainPostContent.innerText.replace(/\s/g, "")).length : 0;
    if (wordCount) wordCount.textContent = `${count} 个字`;

    document.querySelectorAll("[data-codex-action]").forEach((button) => {
      const action = actions[Number(button.dataset.codexAction)];
      button.disabled = !action?.find();
    });
  };

  const setupAutoHide = () => {
    let touchY;

    const toggle = (hidden) => {
      // 灯箱打开时滚轮和方向键用来平移、切换图片，功能区保持不动。
      if (window.pswp) return;
      root.classList.toggle("codex-word-ribbon-hidden", hidden && window.scrollY > 48);
    };

    window.addEventListener("wheel", (event) => {
      if (Math.abs(event.deltaY) >= 4) toggle(event.deltaY > 0);
    }, { passive: true });
    window.addEventListener("touchstart", (event) => {
      touchY = event.touches[0]?.clientY;
    }, { passive: true });
    window.addEventListener("touchmove", (event) => {
      const nextY = event.touches[0]?.clientY;
      if (touchY !== undefined && nextY !== undefined && Math.abs(touchY - nextY) >= 4) {
        toggle(touchY > nextY);
      }
      touchY = nextY;
    }, { passive: true });
    window.addEventListener("keydown", (event) => {
      if (event.target.closest?.("input, textarea, [contenteditable]")) return;
      const down = ["ArrowDown", "PageDown", "End"].includes(event.key) || (event.key === " " && !event.shiftKey);
      const up = ["ArrowUp", "PageUp", "Home"].includes(event.key) || (event.key === " " && event.shiftKey);
      if (down || up) toggle(down);
    });
    window.addEventListener("scroll", () => {
      if (performance.now() < manualToggleUntil) return;
      if (window.scrollY <= 48) toggle(false);
    }, { passive: true });
  };

  const mount = () => {
    if (!document.body || document.getElementById("codex-word-ui")) return;

    const chrome = createWordChrome({
      home: "/",
      insert: "/new-topic",
      onToggle: () => {
        manualToggleUntil = performance.now() + 250;
        // 灯箱开着时按新的两栏高度重新排版图片。
        window.pswp?.updateSize(true);
      },
    });

    actions.forEach((action, index) => {
      const button = wordCommand({ icon: action.icon, label: action.label });
      button.dataset.codexAction = index;
      button.addEventListener("click", () => {
        action.run ? action.run() : action.find()?.click();
      });
      button.addEventListener("pointerenter", () => action.enter?.(button));
      button.addEventListener("pointerleave", () => action.leave?.());
      chrome.actions.append(button);
    });

    sync();
    matchMedia("(prefers-color-scheme: dark)").addEventListener("change", sync);
    setupAutoHide();

    let queued = false;
    new MutationObserver(() => {
      if (queued) return;
      queued = true;
      requestAnimationFrame(() => {
        queued = false;
        sync();
      });
    }).observe(document.body, { childList: true, subtree: true, characterData: true });
  };

  if (document.body) {
    mount();
  } else {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  }
})();
