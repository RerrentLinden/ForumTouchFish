// ==UserScript==
// @name         V2EX Word 摸鱼版
// @namespace    https://codex.local/userscripts
// @version      2.1.6
// @description  隐藏头像，并把 V2EX 伪装成 Microsoft Word 文档界面。
// @author       Codex
// @match        https://v2ex.com/*
// @match        https://www.v2ex.com/*
// @run-at       document-start
// @grant        none
// @noframes
// ==/UserScript==

(() => {
  "use strict";

  let manualToggleUntil = 0;
  document.documentElement.classList.add("codex-word-theme", "codex-word-v2ex");

  const style = document.createElement("style");
  style.id = "codex-v2ex-moyu-style";
  style.textContent = `
    /* 首页、主题、回复、账号卡片及右侧栏头像。 */
    img.avatar,
    a:has(> img.avatar:only-child),
    td:has(> a > img.avatar),
    td:has(> img.avatar),
    .header > .fr:has(img.avatar),
    td:has(> a > img.avatar) + td:empty,
    td:has(> img.avatar) + td:empty,
    #Rightbar img,
    #Rightbar a:has(> img:only-child),
    #Rightbar td:has(> a > img),
    #Rightbar td:has(> img),
    #Rightbar td:has(> a > img) + td:empty,
    #Rightbar td:has(> img) + td:empty {
      display: none !important;
    }

    :root {
      --word-blue: #185abd;
      --word-link: #0563c1;
      --word-ink: #242424;
      --word-line: #dedede;
      --word-workspace: #e7e7e7;
      --word-paper: #fff;
      --word-ui-height: 132px;
      --word-status-height: 29px;
    }

    html.codex-word-theme {
      background: var(--word-workspace) !important;
    }

    html.codex-word-theme body {
      box-sizing: border-box !important;
      margin: 0 !important;
      padding-top: var(--word-ui-height) !important;
      padding-bottom: var(--word-status-height) !important;
      color: var(--word-ink) !important;
      background: var(--word-workspace) !important;
      background-image: none !important;
      font-family: Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
      transition: padding-top 180ms ease;
    }

    #codex-word-ui {
      position: fixed;
      z-index: 2147483646;
      inset: 0 0 auto;
      height: var(--word-ui-height);
      color: #202020;
      background: #f8f8f8;
      border-bottom: 1px solid #c9c9c9;
      box-shadow: 0 1px 4px rgb(0 0 0 / 18%);
      font: 13px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
      overflow: hidden;
      transition: height 180ms ease, box-shadow 180ms ease;
    }

    #codex-word-ui * {
      box-sizing: border-box;
    }

    #codex-word-ui :is(a, button) {
      margin: 0;
      box-shadow: none;
      font: inherit;
      letter-spacing: normal;
      text-transform: none;
    }

    .codex-word-titlebar {
      display: grid;
      grid-template-columns: 1fr auto 1fr;
      align-items: center;
      height: 34px;
      padding: 0 14px;
      color: #fff;
      background: var(--word-blue);
    }

    .codex-word-left,
    .codex-word-right {
      display: flex;
      align-items: center;
      gap: 12px;
      min-width: 0;
    }

    .codex-word-right {
      justify-content: flex-end;
      color: rgb(255 255 255 / 82%);
    }

    .codex-word-dots,
    .codex-word-quick {
      display: flex;
      gap: 7px;
    }

    .codex-word-dots i {
      width: 12px;
      height: 12px;
      border-radius: 50%;
    }

    .codex-word-dots i:nth-child(1) { background: #ff5f57; }
    .codex-word-dots i:nth-child(2) { background: #febc2e; }
    .codex-word-dots i:nth-child(3) { background: #28c840; }

    .codex-word-autosave {
      display: flex;
      align-items: center;
      gap: 6px;
      white-space: nowrap;
    }

    .codex-word-switch {
      width: 28px;
      height: 14px;
      padding: 2px;
      border-radius: 999px;
      background: rgb(255 255 255 / 34%);
    }

    .codex-word-switch::before {
      display: block;
      width: 10px;
      height: 10px;
      content: "";
      border-radius: 50%;
      background: #fff;
    }

    .codex-word-quick {
      gap: 10px;
      color: rgb(255 255 255 / 88%);
      font-size: 16px;
    }

    .codex-word-title {
      overflow: hidden;
      padding: 0 8px;
      color: inherit;
      background: transparent;
      border: 0;
      cursor: pointer;
      font-weight: 600;
      text-align: center;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .codex-word-tabs {
      display: flex;
      align-items: end;
      height: 31px;
      padding: 0 15px;
      gap: 22px;
      overflow: hidden;
      background: #fff;
      border-bottom: 1px solid #ddd;
      transition: opacity 140ms ease, transform 180ms ease;
    }

    .codex-word-tabs :is(a, span) {
      height: 31px;
      padding-top: 9px;
      color: #202020 !important;
      text-decoration: none !important;
      white-space: nowrap;
    }

    .codex-word-tabs .is-active {
      color: var(--word-blue) !important;
      border-bottom: 2px solid var(--word-blue);
      font-weight: 600;
    }

    .codex-word-tabs a:hover {
      background: #f0f4fa;
    }

    .codex-word-ribbon {
      display: flex;
      align-items: stretch;
      height: 67px;
      padding: 7px 16px 5px;
      gap: 12px;
      overflow: hidden;
      background: #fafafa;
      transition: opacity 140ms ease, transform 180ms ease;
    }

    html.codex-word-theme.codex-word-ribbon-hidden {
      --word-ui-height: 34px;
      --word-status-height: 0px;
    }

    html.codex-word-ribbon-hidden #codex-word-ui {
      box-shadow: 0 1px 3px rgb(0 0 0 / 14%);
    }

    html.codex-word-ribbon-hidden :is(.codex-word-tabs, .codex-word-ribbon) {
      opacity: 0;
      transform: translateY(-8px);
      pointer-events: none;
    }

    .codex-word-group {
      display: flex;
      align-items: center;
      gap: 7px;
      min-width: 0;
      padding-right: 12px;
      border-right: 1px solid #ddd;
      white-space: nowrap;
    }

    .codex-word-group:last-child {
      border-right: 0;
    }

    .codex-word-select,
    .codex-word-tool,
    .codex-word-style {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      min-width: 29px;
      height: 29px;
      padding: 0 8px;
      border: 1px solid #d2d2d2;
      border-radius: 2px;
      background: #fff;
    }

    .codex-word-select {
      justify-content: space-between;
      min-width: 114px;
    }

    .codex-word-size {
      min-width: 46px;
    }

    .codex-word-tool {
      border-color: transparent;
      background: transparent;
      font-size: 15px;
    }

    .codex-word-style {
      flex: 0 0 122px;
      width: 122px;
      height: 49px;
      padding: 4px 8px 3px;
      overflow: hidden;
      flex-direction: column;
      font: 16px/1 Georgia, "Times New Roman", serif;
      white-space: nowrap;
    }

    .codex-word-style small {
      display: block;
      margin-top: 4px;
      font: 11px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
    }

    #codex-word-status {
      position: fixed;
      z-index: 2147483646;
      inset: auto 0 0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      height: var(--word-status-height);
      padding: 0 18px;
      color: #555;
      background: #f4f4f4;
      border-top: 1px solid #cfcfcf;
      font: 12px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
      overflow: hidden;
      transition: height 180ms ease, opacity 140ms ease, transform 180ms ease;
    }

    html.codex-word-ribbon-hidden #codex-word-status {
      opacity: 0;
      transform: translateY(100%);
      pointer-events: none;
      border-top-width: 0;
    }

    .codex-word-status-left,
    .codex-word-status-right {
      display: flex;
      align-items: center;
      gap: 18px;
      white-space: nowrap;
    }

    html.codex-word-theme img {
      max-width: 100%;
      height: auto;
    }

    html.codex-word-v2ex #Top {
      position: sticky !important;
      z-index: 2 !important;
      top: var(--word-ui-height) !important;
      display: block !important;
      height: 44px !important;
      background: #fff !important;
      border-bottom: 1px solid #d2d2d2 !important;
      box-shadow: none !important;
    }

    html.codex-word-v2ex #Top > .content {
      box-sizing: border-box !important;
      width: min(1220px, calc(100vw - 32px)) !important;
      height: 44px !important;
      margin: 0 auto !important;
    }

    html.codex-word-v2ex #Wrapper {
      box-sizing: border-box !important;
      width: 100% !important;
      min-height: calc(100vh - var(--word-ui-height) - var(--word-status-height)) !important;
      padding: 28px 16px 72px !important;
      background: var(--word-workspace) !important;
    }

    html.codex-word-v2ex #Wrapper > .content {
      display: grid !important;
      grid-template: "main right" auto / minmax(0, 1fr) 260px !important;
      align-items: start !important;
      gap: 16px !important;
      width: min(1220px, calc(100vw - 32px)) !important;
      margin: 0 auto !important;
    }

    html.codex-word-v2ex #Leftbar {
      display: none !important;
    }

    html.codex-word-v2ex #Main {
      grid-area: main;
      box-sizing: border-box !important;
      float: none !important;
      width: 100% !important;
      margin: 0 !important;
      min-height: 1060px;
      padding: 44px 42px 64px !important;
      background: var(--word-paper) !important;
      border: 1px solid #d0d0d0;
      box-shadow: 0 2px 8px rgb(0 0 0 / 18%);
    }

    html.codex-word-v2ex #Rightbar {
      position: sticky !important;
      grid-area: right;
      top: calc(var(--word-ui-height) + 58px) !important;
      display: block !important;
      float: none !important;
      width: 100% !important;
      max-height: calc(100vh - var(--word-ui-height) - var(--word-status-height) - 28px);
      margin: 0 !important;
      overflow: auto;
      scrollbar-width: thin;
    }

    html.codex-word-v2ex #Rightbar .box {
      width: auto !important;
      margin-bottom: 12px !important;
      border: 1px solid #d5d5d5 !important;
      border-radius: 0 !important;
      box-shadow: 0 1px 3px rgb(0 0 0 / 10%) !important;
    }

    html.codex-word-v2ex #Bottom {
      display: block !important;
      color: #555 !important;
      background: #f7f7f7 !important;
      border-top: 1px solid #d5d5d5 !important;
    }

    html.codex-word-v2ex #Main :is(.box, .cell, .header, .inner) {
      background-color: transparent !important;
      border-radius: 0 !important;
      box-shadow: none !important;
    }

    html.codex-word-v2ex #Main .box {
      border: 0 !important;
    }

    html.codex-word-v2ex #Main .cell {
      padding: 14px 8px !important;
      border-bottom: 1px solid var(--word-line) !important;
    }

    html.codex-word-v2ex #Main .header {
      padding: 4px 8px 18px !important;
      border-bottom: 1px solid var(--word-line) !important;
    }

    html.codex-word-v2ex #Main :is(h1, h2, h3) {
      color: var(--word-ink) !important;
      font-family: Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
    }

    html.codex-word-v2ex #Main h1 {
      font-size: 24px !important;
      line-height: 1.35 !important;
    }

    html.codex-word-v2ex #Main :is(.topic_content, .markdown_body) {
      color: var(--word-ink) !important;
      font: 16px/1.78 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
    }

    html.codex-word-v2ex #Main :is(input, textarea, select, button) {
      border-color: #b8b8b8 !important;
      border-radius: 2px !important;
      font-family: inherit !important;
    }

    html.codex-word-v2ex #Main textarea {
      box-sizing: border-box !important;
      width: 100% !important;
      background: #fff !important;
    }

    #codex-word-action-group {
      margin: 0;
    }

    .codex-word-native-actions {
      display: flex;
      align-items: center;
      height: 100%;
      gap: 2px;
    }

    :is(.codex-word-native-action, .codex-word-compose-link) {
      display: inline-flex !important;
      align-items: center !important;
      justify-content: center !important;
      min-width: 60px;
      height: 34px !important;
      margin: 0 !important;
      padding: 0 8px !important;
      color: #222 !important;
      background: transparent !important;
      border: 1px solid transparent !important;
      border-radius: 2px !important;
      box-shadow: none !important;
      font: 13px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
      font-weight: 400 !important;
      letter-spacing: normal !important;
      text-transform: none !important;
      cursor: pointer;
      text-decoration: none !important;
      white-space: nowrap;
    }

    :is(.codex-word-native-action, .codex-word-compose-link):hover {
      color: var(--word-blue) !important;
      background: #edf3fb !important;
      border-color: #c9d9f0 !important;
    }

    .codex-word-native-action:disabled {
      color: #999 !important;
      cursor: default;
    }

    /*
     * 原站在点楼层回复或输入框获得焦点时给回复框加 .reply-box-sticky，粘在视口底部。
     * 上面把 #Main 内的 .box 设为透明，停靠时须补回底色，否则会和楼层文字重叠。
     */
    html.codex-word-v2ex #reply-box.reply-box-sticky {
      bottom: var(--word-status-height) !important;
      max-height: min(46vh, 420px) !important;
      overflow: auto !important;
      background: var(--word-paper) !important;
      border-top: 1px solid #b7b7b7 !important;
      box-shadow: 0 -8px 24px rgb(0 0 0 / 16%) !important;
    }

    #reply-box:not(.reply-box-sticky) .codex-word-close-editor {
      display: none !important;
    }

    .codex-word-close-editor {
      display: inline-flex !important;
      align-items: center;
      justify-content: center;
      width: 24px !important;
      height: 24px !important;
      margin-left: 8px !important;
      padding: 0 6px !important;
      color: #444 !important;
      background: transparent !important;
      border: 1px solid #b8b8b8 !important;
      border-radius: 2px !important;
      font: 20px/1 Aptos, Calibri, sans-serif !important;
      cursor: pointer;
    }

    .codex-word-close-editor:hover {
      color: var(--word-blue) !important;
      background: #edf3fb !important;
      border-color: #c9d9f0 !important;
    }

    @media (max-width: 1050px) {
      .codex-word-ribbon .codex-word-group:nth-child(2),
      .codex-word-ribbon .codex-word-group:nth-child(3) {
        display: none;
      }
    }

    @media (max-width: 980px) {
      html.codex-word-v2ex #Wrapper > .content {
        grid-template: "main" auto "right" auto / minmax(0, 1fr) !important;
      }

      html.codex-word-v2ex #Rightbar {
        position: static !important;
        max-height: none;
      }
    }

    @media (max-width: 760px) {
      :root { --word-ui-height: 99px; }
      .codex-word-titlebar { grid-template-columns: 1fr auto; }
      .codex-word-title { display: none; }
      .codex-word-tabs { gap: 13px; }
      .codex-word-ribbon { height: 34px; padding-top: 2px; }
      .codex-word-ribbon .codex-word-group:not(:first-child):not(#codex-word-action-group),
      .codex-word-style { display: none; }
      html.codex-word-v2ex #Main {
        min-height: 800px;
        padding: 30px 22px 48px !important;
      }
    }

    @media print {
      #codex-word-ui,
      #codex-word-status { display: none !important; }
      html.codex-word-theme body {
        padding: 0 !important;
        background: #fff !important;
      }
      html.codex-word-v2ex #Main {
        border: 0 !important;
        box-shadow: none !important;
      }
    }

    @media (prefers-reduced-motion: reduce) {
      html.codex-word-theme body,
      #codex-word-ui,
      .codex-word-tabs,
      .codex-word-ribbon,
      #codex-word-status,
      html.codex-word-v2ex #Top {
        transition: none !important;
      }
    }
  `;
  (document.head || document.documentElement).append(style);

  /*
   * 收藏节点使用 closed Shadow DOM；V2EX 同时把 Shadow Root 暴露在
   * 组件的 _shadowRoot 属性上，因此在组件定义后处理图标。
   */
  const hideNodeSidebarIcon = (node) => {
    const avatarLink = node?._shadowRoot?.getElementById("linkAvatar");
    avatarLink?.style.setProperty("display", "none", "important");
  };

  const scanNodeSidebars = (root) => {
    if (root?.matches?.("node-sidebar")) hideNodeSidebarIcon(root);
    root?.querySelectorAll?.("node-sidebar").forEach(hideNodeSidebarIcon);
  };

  const observer = new MutationObserver((records) => {
    records.forEach((record) => {
      record.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) scanNodeSidebars(node);
      });
    });
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });
  customElements.whenDefined("node-sidebar").then(() => scanNodeSidebars(document));

  const findTopicAction = (name) =>
    [...document.querySelectorAll("#Main .topic_buttons a")].find((item) => item.textContent.trim().includes(name));

  /*
   * 停靠状态完全交给原站的 .reply-box-sticky：点楼层回复、输入框获得焦点、
   * 「取消回复框停靠」都走原站逻辑，这里只补关闭按钮和功能区入口。
   */
  const setupReplyDock = () => {
    const box = document.getElementById("reply-box");
    const undock = document.getElementById("undock-button");
    if (!box || !undock) return null;

    const button = document.createElement("button");
    button.type = "button";
    button.className = "codex-word-close-editor";
    button.title = "关闭回复窗口";
    button.setAttribute("aria-label", "关闭回复窗口");
    button.textContent = "×";
    button.addEventListener("click", () => undock.click());
    undock.parentElement.append(button);

    // 与原站 setReplyBoxSticky() 相同：停靠回复框并显示「取消回复框停靠」。
    return () => {
      box.classList.add("reply-box-sticky");
      undock.style.display = "inline-block";
    };
  };

  const setupAutoHideRibbon = () => {
    let anchorY = window.scrollY;
    let ignoreUntil = 0;
    let queued = false;

    const update = () => {
      queued = false;
      const y = Math.max(0, window.scrollY);
      const delta = y - anchorY;

      if (performance.now() < manualToggleUntil) {
        anchorY = y;
        return;
      }

      if (y <= 48) {
        document.documentElement.classList.remove("codex-word-ribbon-hidden");
        anchorY = y;
        return;
      }

      if (performance.now() < ignoreUntil || Math.abs(delta) < 12) return;

      const hidden = delta > 0;
      const changed =
        document.documentElement.classList.contains("codex-word-ribbon-hidden") !== hidden;
      document.documentElement.classList.toggle("codex-word-ribbon-hidden", hidden);
      anchorY = y;
      if (changed) ignoreUntil = performance.now() + 220;
    };

    window.addEventListener(
      "scroll",
      () => {
        if (queued) return;
        queued = true;
        requestAnimationFrame(update);
      },
      { passive: true },
    );
  };

  const mountRibbonActions = (openReply) => {
    const group = document.getElementById("codex-word-action-group");
    if (!group) return;

    const toolbar = document.createElement("div");
    toolbar.className = "codex-word-native-actions";
    toolbar.setAttribute("role", "toolbar");
    toolbar.setAttribute("aria-label", "V2EX 快捷操作");

    const compose = document.createElement("a");
    compose.className = "codex-word-compose-link";
    compose.href = "/write";
    compose.textContent = "＋ 发帖";
    toolbar.append(compose);

    if (/^\/t\/\d+/.test(location.pathname)) {
      [
        { icon: "☆", label: "收藏", source: () => findTopicAction("收藏") },
        { icon: "♡", label: "感谢", source: () => findTopicAction("感谢") },
      ].forEach((action) => {
        const button = document.createElement("button");
        const source = action.source();
        button.type = "button";
        button.className = "codex-word-native-action";
        button.disabled = !source;
        button.textContent = `${action.icon} ${source?.textContent.trim() || action.label}`;
        button.addEventListener("click", () => action.source()?.click());
        toolbar.append(button);
      });

      if (openReply) {
        const reply = document.createElement("button");
        reply.type = "button";
        reply.className = "codex-word-native-action";
        reply.textContent = "↩ 回复";
        reply.addEventListener("click", openReply);
        toolbar.append(reply);
      }
    }

    group.append(toolbar);
  };

  const mount = () => {
    if (!document.body || document.getElementById("codex-word-ui")) return;

    document.body.insertAdjacentHTML(
      "afterbegin",
      `
        <nav id="codex-word-ui" aria-label="Word 摸鱼导航">
          <div class="codex-word-titlebar">
            <div class="codex-word-left">
              <span class="codex-word-dots" aria-hidden="true"><i></i><i></i><i></i></span>
              <span class="codex-word-autosave">自动保存 <i class="codex-word-switch"></i></span>
              <span class="codex-word-quick" aria-hidden="true">⌂ ▣ ↶ ↷</span>
            </div>
            <button class="codex-word-title" type="button" title="切换功能区和状态栏">工作记录 - Word</button>
            <div class="codex-word-right">⌕ 搜索 (Cmd + Ctrl + U)</div>
          </div>
          <div class="codex-word-tabs">
            <a href="/">文件</a>
            <a class="is-active" href="/">开始</a>
            <a href="/write">插入</a>
            <span>绘图</span><span>设计</span><span>布局</span><span>引用</span><span>邮件</span><span>审阅</span><span>视图</span>
          </div>
          <div class="codex-word-ribbon">
            <div class="codex-word-group">
              <span class="codex-word-select">等线 <b>⌄</b></span>
              <span class="codex-word-select codex-word-size">11 <b>⌄</b></span>
              <span class="codex-word-tool"><b>B</b></span>
              <span class="codex-word-tool"><i>I</i></span>
              <span class="codex-word-tool"><u>U</u></span>
            </div>
            <div class="codex-word-group">
              <span class="codex-word-tool">☷</span><span class="codex-word-tool">≡</span>
              <span class="codex-word-tool">☰</span><span class="codex-word-tool">↕</span>
              <span class="codex-word-tool">A⌄</span>
            </div>
            <div class="codex-word-group">
              <span class="codex-word-style">AaBbCcDdEe<small>正文</small></span>
              <span class="codex-word-style">AaBbCcDdEe<small>无间隔</small></span>
            </div>
            <div class="codex-word-group" id="codex-word-action-group"></div>
          </div>
        </nav>
        <div id="codex-word-status" aria-hidden="true">
          <span class="codex-word-status-left"><span>第 1 页，共 1 页</span><span id="codex-word-word-count">0 个字</span><span>简体中文（中国大陆）</span></span>
          <span class="codex-word-status-right"><span>辅助功能：一切就绪</span><span>▤　◎　☰　−　100%　＋</span></span>
        </div>
      `,
    );

    document.querySelector(".codex-word-title").addEventListener("click", () => {
      manualToggleUntil = performance.now() + 250;
      document.documentElement.classList.toggle("codex-word-ribbon-hidden");
    });

    const mainPost = document.querySelector("#Main .topic_content");
    const wordCount = mainPost ? [...mainPost.innerText.replace(/\s+/g, "")].length : 0;
    document.getElementById("codex-word-word-count").textContent = `${wordCount} 个字`;
    mountRibbonActions(setupReplyDock());
    setupAutoHideRibbon();
  };

  if (document.body) {
    mount();
  } else {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  }
})();
