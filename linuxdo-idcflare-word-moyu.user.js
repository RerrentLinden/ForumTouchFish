// ==UserScript==
// @name         LINUX DO / IDC Flare Word 摸鱼版
// @namespace    https://codex.local/userscripts
// @version      1.0.13
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

  const root = document.documentElement;
  let manualToggleUntil = 0;
  root.classList.add("codex-word-theme", "codex-word-discourse");

  const style = document.createElement("style");
  style.id = "codex-discourse-word-style";
  style.textContent = `
    :root {
      --codex-word-blue: #185abd;
      --codex-word-ink: #242424;
      --codex-word-line: #d7d7d7;
      --codex-word-paper: #fff;
      --codex-word-workspace: #e7e7e7;
      --codex-word-ui-height: 132px;
      --codex-word-status-height: 29px;
      --codex-discourse-header-height: 52px;
      --header-offset: calc(var(--codex-word-ui-height) + var(--codex-discourse-header-height));
    }

    html.codex-word-theme,
    html.codex-word-theme body {
      background: var(--codex-word-workspace) !important;
      background-image: none !important;
    }

    html.codex-word-theme body {
      box-sizing: border-box !important;
      min-height: 100vh !important;
      padding-top: var(--codex-word-ui-height) !important;
      padding-bottom: var(--codex-word-status-height) !important;
      color: var(--codex-word-ink) !important;
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

    #codex-word-ui {
      position: fixed;
      z-index: 2147483646;
      inset: 0 0 auto;
      height: var(--codex-word-ui-height);
      overflow: hidden;
      color: #202020;
      background: #f8f8f8;
      border-bottom: 1px solid #c9c9c9;
      box-shadow: 0 1px 4px rgb(0 0 0 / 18%);
      font: 13px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
      transition: height 180ms ease, box-shadow 180ms ease;
    }

    #codex-word-ui *,
    #codex-word-status * {
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
      background: var(--codex-word-blue);
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

    .codex-word-dots {
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
      display: flex;
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

    .codex-word-tabs a,
    .codex-word-tabs span {
      height: 31px;
      padding-top: 9px;
      color: #202020 !important;
      text-decoration: none !important;
      white-space: nowrap;
    }

    .codex-word-tabs .is-active {
      color: var(--codex-word-blue) !important;
      border-bottom: 2px solid var(--codex-word-blue);
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
      flex-direction: column;
      width: 122px;
      height: 49px;
      padding: 4px 8px 3px;
      overflow: hidden;
      font: 16px/1 Georgia, "Times New Roman", serif;
      white-space: nowrap;
    }

    .codex-word-style small {
      display: block;
      margin-top: 4px;
      font: 11px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
    }

    #codex-word-actions {
      flex: 1 1 auto;
      height: 100%;
      align-items: center;
      justify-content: flex-start;
      overflow: hidden;
    }

    .codex-word-action {
      display: inline-flex !important;
      flex: 0 0 auto;
      align-items: center !important;
      justify-content: center !important;
      min-width: 60px;
      height: 34px !important;
      margin: 0 !important;
      padding: 0 8px !important;
      color: #202020 !important;
      background: transparent !important;
      border: 1px solid transparent !important;
      border-radius: 2px !important;
      box-shadow: none !important;
      font: 13px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
      font-weight: 400 !important;
      letter-spacing: normal !important;
      text-decoration: none !important;
      text-transform: none !important;
      white-space: nowrap;
      cursor: pointer;
    }

    .codex-word-action:hover {
      color: var(--codex-word-blue) !important;
      background: #edf3fb !important;
      border-color: #c9d9f0 !important;
    }

    .codex-word-action:disabled {
      display: none !important;
    }

    html.codex-word-theme .codex-word-picker-hosted .discourse-reactions-picker.is-expanded {
      position: fixed !important;
      z-index: 2147483647 !important;
      top: var(--codex-reaction-picker-y) !important;
      left: var(--codex-reaction-picker-x) !important;
      transform: translateX(-50%) !important;
    }

    html.codex-word-theme .d-header-wrap {
      top: var(--codex-word-ui-height) !important;
      transition: top 160ms ease;
    }

    html.codex-word-theme #main-outlet {
      box-sizing: border-box !important;
      width: min(calc(100% - 64px), 1180px) !important;
      max-width: 1180px !important;
      min-height: calc(100vh - var(--codex-word-ui-height) - 92px);
      margin: 28px auto 48px !important;
      padding: 28px 40px 44px !important;
      background: var(--codex-word-paper) !important;
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
      bottom: var(--codex-word-status-height) !important;
    }

    html.codex-word-theme.fullscreen-composer #reply-control.fullscreen {
      bottom: var(--codex-word-status-height);
      height: calc(
        var(--composer-vh, 1vh) * 100 - var(--codex-word-ui-height) - var(--codex-word-status-height)
      ) !important;
    }

    html.codex-word-theme .d-modal {
      top: var(--codex-word-ui-height);
      height: calc(100% - var(--codex-word-ui-height) - var(--codex-word-status-height));
    }

    html.codex-word-theme .d-modal__container {
      max-height: min(
        var(--modal-max-height, 80vh),
        calc(100vh - var(--codex-word-ui-height) - var(--codex-word-status-height) - 24px)
      ) !important;
    }

    /* 灯箱遮罩仍铺满窗口；图片由 fitLightbox() 让开两栏，按钮和图片说明在这里挪进来。 */
    html.codex-word-theme .pswp__top-bar {
      top: var(--codex-word-ui-height);
    }

    html.codex-word-theme .pswp__button--arrow {
      margin-top: calc((var(--codex-word-ui-height) - var(--codex-word-status-height)) / 2 - 50px);
    }

    html.codex-word-theme .pswp__caption {
      bottom: calc(var(--codex-word-status-height) + var(--safe-area-inset-bottom, 0px));
    }

    html.codex-word-theme .chat-drawer-outlet-container {
      bottom: var(--codex-word-status-height);
      max-height: calc(100% - var(--header-offset) - 15px - var(--codex-word-status-height));
    }

    html.codex-word-theme .with-topic-progress {
      bottom: calc(
        env(safe-area-inset-bottom) + var(--composer-height, 0px) + var(--codex-word-status-height)
      );
    }

    html.codex-word-theme .menu-panel.drop-down {
      max-height: calc(100dvh - var(--header-offset) - 1em - var(--codex-word-status-height));
    }

    @media (width >= 48rem) {
      html.codex-word-theme .sidebar-wrapper {
        height: calc(
          var(--composer-vh, 1dvh) * 100 - var(--main-outlet-offset, 0px) - var(--codex-word-status-height)
        );
      }
    }

    /* LINUX DO 站点组件的操作提示，原本离视口顶部 10%。 */
    html.codex-word-theme #messageToast {
      top: calc(var(--codex-word-ui-height) + 10%);
    }

    #codex-word-status {
      position: fixed;
      z-index: 2147483646;
      inset: auto 0 0;
      display: flex;
      align-items: center;
      justify-content: space-between;
      height: var(--codex-word-status-height);
      padding: 0 18px;
      color: #555;
      background: #f4f4f4;
      border-top: 1px solid #cfcfcf;
      font: 12px/1.2 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif;
      overflow: hidden;
      transition: height 180ms ease, opacity 140ms ease, transform 180ms ease;
    }

    .codex-word-status-left,
    .codex-word-status-right {
      display: flex;
      align-items: center;
      gap: 18px;
      white-space: nowrap;
    }

    html.codex-word-ribbon-hidden {
      --codex-word-ui-height: 34px;
      --codex-word-status-height: 0px;
    }

    html.codex-word-ribbon-hidden :is(.codex-word-tabs, .codex-word-ribbon) {
      opacity: 0;
      transform: translateY(-8px);
      pointer-events: none;
    }

    html.codex-word-ribbon-hidden #codex-word-status {
      opacity: 0;
      transform: translateY(100%);
      pointer-events: none;
      border-top-width: 0;
    }

    html.codex-word-dark {
      --codex-word-ink: #f3f3f3;
      --codex-word-line: #454545;
      --codex-word-paper: #252526;
      --codex-word-workspace: #1b1b1f;
    }

    html.codex-word-dark #codex-word-ui {
      color: var(--codex-word-ink);
      background: #2b2b2b;
      border-bottom-color: var(--codex-word-line);
    }

    html.codex-word-dark :is(.codex-word-tabs, .codex-word-ribbon) {
      color: var(--codex-word-ink);
      background: var(--codex-word-paper);
      border-color: var(--codex-word-line);
    }

    html.codex-word-dark .codex-word-tabs :is(a, span) {
      color: var(--codex-word-ink) !important;
    }

    html.codex-word-dark .codex-word-tabs :is(.is-active, a:hover),
    html.codex-word-dark .codex-word-action:hover {
      color: #75b7ff !important;
      background: #333a44 !important;
      border-color: #4f6b8a !important;
    }

    html.codex-word-dark .codex-word-group {
      border-color: var(--codex-word-line);
    }

    html.codex-word-dark :is(.codex-word-select, .codex-word-style) {
      color: var(--codex-word-ink);
      background: #333;
      border-color: #555;
    }

    html.codex-word-dark .codex-word-action {
      color: var(--codex-word-ink) !important;
    }

    html.codex-word-dark :is(#main-outlet, #reply-control.open:not(.fullscreen)) {
      color: var(--codex-word-ink) !important;
      border-color: var(--codex-word-line);
    }

    html.codex-word-dark #codex-word-status {
      color: #bdbdbd;
      background: var(--codex-word-paper);
      border-color: var(--codex-word-line);
    }

    @media (max-width: 1050px) {
      .codex-word-ribbon .codex-word-group:nth-child(2),
      .codex-word-ribbon .codex-word-group:nth-child(3) {
        display: none;
      }
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

    @media (max-width: 760px) {
      :root { --codex-word-ui-height: 99px; }
      .codex-word-titlebar { grid-template-columns: 1fr auto; }
      .codex-word-title { display: none; }
      .codex-word-tabs { gap: 13px; }
      .codex-word-ribbon { height: 34px; padding-top: 2px; }
      .codex-word-ribbon .codex-word-group:not(:first-child):not(#codex-word-actions),
      .codex-word-style { display: none; }
    }

    @media print {
      #codex-word-ui,
      #codex-word-status {
        display: none !important;
      }

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
      html.codex-word-theme .d-header-wrap,
      #codex-word-ui,
      .codex-word-tabs,
      .codex-word-ribbon,
      #codex-word-status {
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
        top: (base.top || 0) + parseFloat(chrome.getPropertyValue("--codex-word-ui-height")),
        bottom: (base.bottom || 0) + parseFloat(chrome.getPropertyValue("--codex-word-status-height")),
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
      label: "＋ 发帖",
      find: () => document.body,
      run: () => {
        const event = new KeyboardEvent("keypress", { key: "c", code: "KeyC", bubbles: true, cancelable: true });
        Object.defineProperty(event, "which", { value: 99 });
        document.dispatchEvent(event);
      },
    },
    {
      label: "♡ 点赞",
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
      label: "☆ 书签",
      find: () =>
        isTopicPage() && (officialBookmark() || mainPost()?.querySelector(".post-action-menu__show-more")),
      run: openOfficialBookmark,
    },
    {
      label: "↗ 分享",
      find: () => isTopicPage() && document.querySelector(".topic-footer-main-buttons .share-and-invite"),
    },
    {
      label: "↩ 回复",
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
            <a href="/new-topic">插入</a>
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
            <div class="codex-word-group" aria-hidden="true">
              <span class="codex-word-tool">☷</span>
              <span class="codex-word-tool">≡</span>
              <span class="codex-word-tool">☰</span>
              <span class="codex-word-tool">↕</span>
              <span class="codex-word-tool">A⌄</span>
            </div>
            <div class="codex-word-group" aria-hidden="true">
              <span class="codex-word-style">AaBbCcDdEe<small>正文</small></span>
              <span class="codex-word-style">AaBbCcDdEe<small>无间隔</small></span>
            </div>
            <div class="codex-word-group" id="codex-word-actions" role="toolbar" aria-label="论坛快捷操作">
              ${actions
                .map(
                  (action, index) =>
                    `<button class="codex-word-action" type="button" data-codex-action="${index}">${action.label}</button>`,
                )
                .join("")}
            </div>
          </div>
        </nav>
        <div id="codex-word-status" aria-hidden="true">
          <span class="codex-word-status-left">
            <span>第 1 页，共 1 页</span>
            <span id="codex-word-count">0 个字</span>
            <span>简体中文（中国大陆）</span>
          </span>
          <span class="codex-word-status-right">
            <span>辅助功能：一切就绪</span>
            <span>▤　◎　☰　−　100%　＋</span>
          </span>
        </div>
      `,
    );

    document.querySelector(".codex-word-title").addEventListener("click", () => {
      manualToggleUntil = performance.now() + 250;
      root.classList.toggle("codex-word-ribbon-hidden");
      // 灯箱开着时按新的两栏高度重新排版图片。
      window.pswp?.updateSize(true);
    });

    document.querySelectorAll("[data-codex-action]").forEach((button) => {
      const action = actions[Number(button.dataset.codexAction)];
      button.addEventListener("click", () => {
        action?.run ? action.run() : action?.find()?.click();
      });
      button.addEventListener("pointerenter", () => action?.enter?.(button));
      button.addEventListener("pointerleave", () => action?.leave?.());
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
