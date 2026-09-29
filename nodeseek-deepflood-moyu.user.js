// ==UserScript==
// @name         NodeSeek / DeepFlood Word 摸鱼版
// @namespace    https://codex.local/userscripts
// @version      2.4.6
// @description  隐藏头像，并把 NodeSeek / DeepFlood 伪装成 Microsoft Word 文档界面。
// @author       Codex
// @match        https://nodeseek.com/*
// @match        https://www.nodeseek.com/*
// @match        https://deepflood.com/*
// @match        https://www.deepflood.com/*
// @run-at       document-start
// @grant        none
// @noframes
// ==/UserScript==

(() => {
  "use strict";

  const home = "/";
  const create = "/new-discussion";
  let manualToggleUntil = 0;

  document.documentElement.classList.add("codex-word-theme", "codex-word-nodeseek");

  const style = document.createElement("style");
  style.id = "codex-nsk-moyu-style";
  style.textContent = `
    /*
     * 两站头像地址统一包含 /avatar/；类名选择器兼容首页、帖子和侧栏。
     * 只包含头像的外层一并收起，正文图片与站点图标不受影响。
     */
    :where(
      img.avatar-normal,
      img.avatar-small,
      img.avatar-large,
      img[src*="/avatar/"]
    ),
    :where(
      .post-list-item > a:has(> img[src*="/avatar/"]),
      .post-list-item > a:has(> img[class*="avatar"]),
      .user-head > a:has(> img[src*="/avatar/"]),
      .user-head > a:has(> img[class*="avatar"]),
      .avatar-wrapper:has(img[src*="/avatar/"]),
      .avatar-wrapper:has(img[class*="avatar"])
    ) {
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
      background: var(--word-workspace) !important;
      background-image: none !important;
      color: var(--word-ink) !important;
      font-family: Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
      transition: padding-top 180ms ease;
    }

    #codex-word-ui {
      position: fixed;
      z-index: 2147483646;
      top: 0;
      right: 0;
      left: 0;
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
      right: 0;
      bottom: 0;
      left: 0;
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

    /* NodeSeek / DeepFlood：保留原生入口，只把它们排进 Word 工作区。 */
    html.codex-word-nodeseek body > footer,
    html.codex-word-nodeseek #fast-nav-button-group {
      display: none !important;
    }

    html.codex-word-nodeseek body > header {
      position: sticky !important;
      z-index: 2 !important;
      top: var(--word-ui-height) !important;
      right: auto !important;
      left: auto !important;
      display: block !important;
      width: 100% !important;
      height: 44px !important;
      background: #fff !important;
      border-bottom: 1px solid #d2d2d2 !important;
      box-shadow: none !important;
      transition: top 180ms ease;
    }

    html.codex-word-nodeseek body > header #nsk-head {
      display: flex !important;
      align-items: center !important;
      width: min(1320px, calc(100vw - 32px)) !important;
      max-width: none !important;
      height: 44px !important;
      margin: 0 auto !important;
      gap: 12px;
    }

    html.codex-word-nodeseek body > header .site-title {
      position: static !important;
      display: flex !important;
      align-items: center !important;
      flex: 0 0 auto;
      margin: 0 !important;
      white-space: nowrap;
    }

    html.codex-word-nodeseek body > header .search-box {
      position: static !important;
      display: flex !important;
      align-items: center !important;
      width: min(360px, 34vw) !important;
      margin: 0 0 0 auto !important;
    }

    html.codex-word-nodeseek body > header .search-box input {
      width: 100% !important;
      height: 30px !important;
      border: 1px solid #b8b8b8 !important;
      border-radius: 2px !important;
      background: #fff !important;
    }

    html.codex-word-nodeseek body > header :is(#nsx-icon-group, .right-button-group) {
      position: static !important;
      display: flex !important;
      align-items: center !important;
      flex: 0 0 auto;
      margin: 0 !important;
      gap: 4px;
    }

    html.codex-word-nodeseek body > header .color-theme-switcher {
      position: static !important;
      display: flex !important;
      align-items: center !important;
      flex: 0 0 auto;
      margin: 0 0 0 4px !important;
    }

    html.codex-word-nodeseek #nsk-frame {
      box-sizing: border-box !important;
      width: 100% !important;
      margin-top: 0 !important;
      min-height: calc(100vh - var(--word-ui-height) - var(--word-status-height)) !important;
      padding: 20px 16px 72px !important;
      background: var(--word-workspace) !important;
    }

    html.codex-word-nodeseek #nsk-body {
      display: grid !important;
      grid-template-columns: 150px minmax(0, 1fr) 240px !important;
      align-items: start !important;
      gap: 16px !important;
      width: min(1320px, calc(100vw - 32px)) !important;
      max-width: none !important;
      margin: 0 auto !important;
      background: transparent !important;
      border-radius: 0 !important;
      box-shadow: none !important;
    }

    html.codex-word-nodeseek #nsk-body:has(> #nsk-body-left:only-child) {
      display: block !important;
      width: min(1080px, calc(100vw - 48px)) !important;
    }

    html.codex-word-nodeseek :is(#nsk-left-panel-container, #nsk-right-panel-container) {
      position: sticky !important;
      top: calc(var(--word-ui-height) + 58px) !important;
      right: auto !important;
      bottom: auto !important;
      left: auto !important;
      display: block !important;
      box-sizing: border-box !important;
      width: auto !important;
      max-height: calc(100vh - var(--word-ui-height) - var(--word-status-height) - 28px);
      margin: 0 !important;
      overflow: auto;
      transform: none !important;
      scrollbar-width: thin;
      transition: top 180ms ease;
    }

    html.codex-word-nodeseek :is(#nsk-left-panel-container, #nsk-right-panel-container) :is(.nsk-panel, .user-card) {
      width: auto !important;
      margin: 0 0 12px !important;
      padding: 10px !important;
      background: #fff !important;
      border: 1px solid #d5d5d5 !important;
      border-radius: 0 !important;
      box-shadow: 0 1px 3px rgb(0 0 0 / 10%) !important;
    }

    html.codex-word-nodeseek #nsk-left-panel-container .category-list {
      display: block !important;
    }

    html.codex-word-nodeseek #nsk-left-panel-container .category-list a {
      display: flex !important;
      align-items: center !important;
      min-height: 28px;
      padding: 4px 8px !important;
      color: #444 !important;
      text-decoration: none !important;
      border-left: 2px solid transparent;
    }

    html.codex-word-nodeseek #nsk-left-panel-container .category-list a:hover,
    html.codex-word-nodeseek #nsk-left-panel-container .category-list a.active {
      color: var(--word-blue) !important;
      background: #edf3fb !important;
      border-left-color: var(--word-blue);
    }

    html.codex-word-nodeseek #nsk-right-panel-container :is(.btn.new-discussion, a.new-discussion) {
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      box-sizing: border-box !important;
      width: 100% !important;
      min-height: 34px;
      margin: 0 0 12px !important;
      color: #fff !important;
      background: var(--word-blue) !important;
      border: 1px solid var(--word-blue) !important;
      border-radius: 2px !important;
      text-decoration: none !important;
    }

    html.codex-word-nodeseek #nsk-right-panel-container :is(input, select, button) {
      box-sizing: border-box !important;
      max-width: 100% !important;
      border-color: #b8b8b8 !important;
      border-radius: 2px !important;
      font-family: inherit !important;
    }

    html.codex-word-nodeseek #nsk-body-left {
      box-sizing: border-box !important;
      width: 100% !important;
      min-width: 0 !important;
      min-height: 1060px;
      margin: 0 !important;
      padding: 44px 42px 64px !important;
      background: var(--word-paper) !important;
      border: 1px solid #d0d0d0;
      box-shadow: 0 2px 8px rgb(0 0 0 / 18%);
    }

    html.codex-word-nodeseek #nsk-body-left :is(
      .post-list,
      .nsk-post-wrapper,
      .nsk-post,
      .comment-container,
      .comments,
      .content-item,
      .nsk-panel
    ) {
      background: transparent !important;
      border-radius: 0 !important;
      box-shadow: none !important;
    }

    html.codex-word-nodeseek #nsk-body-left :is(.post-list > li, .content-item) {
      border-right: 0 !important;
      border-bottom: 1px solid var(--word-line) !important;
      border-left: 0 !important;
    }

    html.codex-word-nodeseek #nsk-body-left .post-list > li {
      padding: 13px 8px !important;
    }

    html.codex-word-nodeseek #nsk-body-left .content-item {
      padding-right: 8px !important;
      padding-left: 8px !important;
    }

    html.codex-word-nodeseek #nsk-body-left .post-content {
      color: var(--word-ink) !important;
      font: 16px/1.78 Aptos, Calibri, "Segoe UI", "Microsoft YaHei", sans-serif !important;
    }

    html.codex-word-nodeseek #nsk-body-left :is(input, textarea, select, button, .CodeMirror) {
      border-color: #b8b8b8 !important;
      border-radius: 2px !important;
      font-family: inherit !important;
    }

    html.codex-word-nodeseek #nsk-body-left .post-list-controler {
      display: flex !important;
      align-items: center !important;
      justify-content: space-between !important;
      gap: 12px;
      margin-bottom: 12px !important;
    }

    html.codex-word-nodeseek #nsk-body-left [role="navigation"][aria-label="pagination"] {
      display: flex !important;
      align-items: center !important;
      justify-content: center !important;
      gap: 3px;
    }

    html.codex-word-nodeseek #nsk-body-left [role="navigation"][aria-label="pagination"] > * {
      min-width: 28px;
      min-height: 28px;
      padding: 5px 7px !important;
      color: #444 !important;
      background: #fff !important;
      border: 1px solid #d5d5d5 !important;
      border-radius: 2px !important;
      text-align: center;
      text-decoration: none !important;
    }

    html.codex-word-nodeseek #nsk-body-left [role="navigation"][aria-label="pagination"] > :is(.pager-cur, [aria-current="page"]) {
      color: #fff !important;
      background: var(--word-blue) !important;
      border-color: var(--word-blue) !important;
    }

    html.codex-word-nodeseek #nsk-frame-block:not(:empty) {
      display: block !important;
      box-sizing: border-box !important;
      width: min(1080px, calc(100vw - 48px)) !important;
      max-width: none !important;
      margin: 0 auto !important;
    }

    html.codex-word-nodeseek #user-setting-panel {
      box-sizing: border-box !important;
      width: 100% !important;
      max-width: none !important;
      min-height: 900px;
      margin: 0 !important;
      padding: 34px 48px 56px !important;
      background: var(--word-paper) !important;
      border: 1px solid #d0d0d0 !important;
      border-radius: 0 !important;
      box-shadow: 0 2px 8px rgb(0 0 0 / 18%) !important;
    }

    html.codex-word-nodeseek #user-setting-panel .selector {
      display: grid !important;
      grid-template-columns: 190px minmax(0, 1fr) !important;
      align-items: start !important;
      gap: 28px;
    }

    html.codex-word-nodeseek #user-setting-panel .selector-left-side {
      width: auto !important;
      background: #f7f7f7 !important;
      border: 1px solid #dedede !important;
    }

    html.codex-word-nodeseek #user-setting-panel .selector-right-side {
      width: auto !important;
      min-width: 0 !important;
    }

    #codex-word-action-group:not(.has-native-actions):not(.has-compose-actions) {
      display: none;
    }

    :is(.codex-word-native-actions, .codex-word-compose-actions) {
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

    html.codex-word-nodeseek .md-editor.codex-word-floating-reply {
      position: fixed !important;
      z-index: 2147483645 !important;
      right: auto !important;
      bottom: var(--word-status-height) !important;
      left: var(--codex-reply-left, 8px) !important;
      box-sizing: border-box !important;
      width: var(--codex-reply-width, calc(100vw - 16px)) !important;
      max-width: none !important;
      max-height: calc(100vh - var(--word-ui-height) - var(--word-status-height) - 24px) !important;
      margin: 0 !important;
      overflow: auto !important;
      transform: none !important;
      background: #fff !important;
      border: 1px solid #b7b7b7 !important;
      box-shadow: 0 10px 30px rgb(0 0 0 / 24%) !important;
    }

    html.codex-word-nodeseek .md-editor.codex-word-floating-reply:has(> #editor-body.fullscreen-editor) {
      height: calc(100vh - var(--word-ui-height) - var(--word-status-height) - 24px) !important;
      overflow: hidden !important;
    }

    html.codex-word-nodeseek .md-editor.codex-word-floating-reply > #editor-body.fullscreen-editor {
      position: static !important;
      width: 100% !important;
      height: 100% !important;
      z-index: auto !important;
    }

    .codex-word-close-editor {
      flex: 0 0 auto;
    }

    html.codex-word-nodeseek:has(body.dark-layout) {
      --word-link: #75b7ff;
      --word-ink: #f3f3f3;
      --word-line: #454545;
      --word-workspace: #1b1b1f;
      --word-paper: #252526;
      color-scheme: dark;
    }

    html.codex-word-nodeseek:has(body.dark-layout) #codex-word-ui {
      color: var(--word-ink);
      background: #2b2b2b;
      border-bottom-color: var(--word-line);
    }

    html.codex-word-nodeseek:has(body.dark-layout) :is(.codex-word-tabs, .codex-word-ribbon) {
      color: var(--word-ink);
      background: var(--word-paper);
      border-color: var(--word-line);
    }

    html.codex-word-nodeseek:has(body.dark-layout) .codex-word-tabs :is(a, span) {
      color: var(--word-ink) !important;
    }

    html.codex-word-nodeseek:has(body.dark-layout) .codex-word-tabs :is(.is-active, a:hover),
    html.codex-word-nodeseek:has(body.dark-layout) :is(.codex-word-native-action, .codex-word-compose-link):hover {
      color: #75b7ff !important;
      background: #333a44 !important;
      border-color: #4f6b8a !important;
    }

    html.codex-word-nodeseek:has(body.dark-layout) .codex-word-group {
      border-color: var(--word-line);
    }

    html.codex-word-nodeseek:has(body.dark-layout) :is(.codex-word-select, .codex-word-style) {
      color: var(--word-ink);
      background: #333;
      border-color: #555;
    }

    html.codex-word-nodeseek:has(body.dark-layout) #codex-word-status {
      color: #bdbdbd;
      background: var(--word-paper);
      border-color: var(--word-line);
    }

    html.codex-word-nodeseek:has(body.dark-layout) body > header,
    html.codex-word-nodeseek:has(body.dark-layout) :is(
      #nsk-left-panel-container,
      #nsk-right-panel-container
    ) :is(.nsk-panel, .user-card),
    html.codex-word-nodeseek:has(body.dark-layout) .md-editor.codex-word-floating-reply {
      color: var(--word-ink) !important;
      background: var(--word-paper) !important;
      border-color: var(--word-line) !important;
    }

    html.codex-word-nodeseek:has(body.dark-layout) body > header .search-box input,
    html.codex-word-nodeseek:has(body.dark-layout) #nsk-body-left [role="navigation"][aria-label="pagination"] > * {
      color: var(--word-ink) !important;
      background: #333 !important;
      border-color: #555 !important;
    }

    html.codex-word-nodeseek:has(body.dark-layout) #nsk-left-panel-container .category-list a,
    html.codex-word-nodeseek:has(body.dark-layout) :is(.codex-word-native-action, .codex-word-compose-link) {
      color: var(--word-ink) !important;
    }

    html.codex-word-nodeseek:has(body.dark-layout) #nsk-left-panel-container .category-list :is(a:hover, a.active) {
      color: #75b7ff !important;
      background: #333a44 !important;
      border-left-color: #75b7ff;
    }

    html.codex-word-nodeseek:has(body.dark-layout) :is(
      #nsk-body-left,
      #user-setting-panel,
      #user-setting-panel .selector-left-side
    ) {
      color: var(--word-ink) !important;
      background: var(--word-paper) !important;
      border-color: var(--word-line) !important;
    }

    @media (max-width: 1050px) {
      .codex-word-ribbon .codex-word-group:nth-child(2),
      .codex-word-ribbon .codex-word-group:nth-child(3) {
        display: none;
      }

      html.codex-word-nodeseek #nsk-body {
        grid-template-columns: 1fr !important;
      }

      html.codex-word-nodeseek :is(#nsk-left-panel-container, #nsk-right-panel-container) {
        position: static !important;
        max-height: none;
      }

      html.codex-word-nodeseek #nsk-left-panel-container {
        display: grid !important;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 12px;
      }

      html.codex-word-nodeseek #nsk-left-panel-container .category-list {
        margin: 0 !important;
      }

      html.codex-word-nodeseek #nsk-right-panel-container {
        display: grid !important;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 12px;
      }
    }

    @media (max-width: 760px) {
      :root {
        --word-ui-height: 99px;
      }

      .codex-word-titlebar {
        grid-template-columns: 1fr auto;
      }

      .codex-word-title {
        display: none;
      }

      .codex-word-tabs {
        gap: 13px;
        overflow: hidden;
      }

      .codex-word-ribbon {
        height: 34px;
        padding-top: 2px;
      }

      .codex-word-ribbon .codex-word-group:not(:first-child):not(#codex-word-action-group),
      .codex-word-style {
        display: none;
      }

      html.codex-word-nodeseek #nsk-body-left {
        min-height: 800px;
        padding: 30px 22px 48px !important;
      }

      html.codex-word-nodeseek body > header #nsk-head {
        width: calc(100vw - 20px) !important;
        overflow-x: auto;
      }

      html.codex-word-nodeseek body > header .search-box {
        min-width: 190px;
      }

      html.codex-word-nodeseek #nsk-left-panel-container,
      html.codex-word-nodeseek #nsk-right-panel-container,
      html.codex-word-nodeseek #user-setting-panel .selector {
        grid-template-columns: 1fr !important;
      }

      html.codex-word-nodeseek #user-setting-panel {
        padding: 28px 22px 44px !important;
      }
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

      html.codex-word-nodeseek #nsk-body-left {
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
      html.codex-word-nodeseek body > header,
      html.codex-word-nodeseek :is(#nsk-left-panel-container, #nsk-right-panel-container) {
        transition: none !important;
      }
    }
  `;
  (document.head || document.documentElement).append(style);

  const postActions = [
    { label: "点赞", icon: "♡", matches: (item) => item.title === "点赞" },
    { label: "鸡腿", icon: "♧", matches: (item) => item.title === "加鸡腿" },
    { label: "点踩", icon: "♢", matches: (item) => item.title === "反对" },
    { label: "收藏", icon: "☆", matches: (item) => item.title === "收藏" },
    { label: "引用", icon: "❝", matches: (item) => item.textContent.trim() === "引用" },
    { label: "回复", icon: "↩", matches: (item) => item.textContent.trim() === "回复" },
  ];

  const composeSites = [
    {
      key: "ns",
      label: "NS 发帖",
      host: "nodeseek.com",
      href: "https://www.nodeseek.com/new-discussion",
    },
    {
      key: "df",
      label: "DF 发帖",
      host: "deepflood.com",
      href: "https://www.deepflood.com/new-discussion",
    },
  ];

  const syncComposeActions = (group) => {
    const path = location.pathname;
    const isPost = /^\/post-\d+(?:-\d+)?\/?$/.test(path);
    const isForumList =
      path === "/" || /^\/page-\d+\/?$/.test(path) || /^\/categories(?:\/|$)/.test(path);
    const targets = isPost
      ? composeSites.filter((site) => location.hostname.endsWith(site.host))
      : isForumList
        ? composeSites
        : [];
    const signature = targets.map((site) => site.key).join(",");
    let toolbar = group.querySelector(".codex-word-compose-actions");

    group.classList.toggle("has-compose-actions", targets.length > 0);
    if (!targets.length) {
      toolbar?.remove();
      delete group.dataset.composeSites;
      return;
    }
    if (toolbar && group.dataset.composeSites === signature) return;

    toolbar?.remove();
    toolbar = document.createElement("nav");
    toolbar.className = "codex-word-compose-actions";
    toolbar.setAttribute("aria-label", "论坛发帖快捷入口");

    targets.forEach((site) => {
      const link = document.createElement("a");
      link.className = "codex-word-compose-link";
      link.href = site.href;
      link.textContent = `＋ ${site.label}`;
      toolbar.append(link);
    });

    group.dataset.composeSites = signature;
    group.append(toolbar);
  };

  const findPostAction = (action) => {
    const menu = document.querySelector("#nsk-body-left .content-item[id='0'] .comment-menu");
    return [...(menu?.querySelectorAll(".menu-item") || [])].find(action.matches);
  };

  const syncPostActions = () => {
    const group = document.getElementById("codex-word-action-group");
    if (!group) return;

    const mainPost = document.querySelector("#nsk-body-left .content-item[id='0'] .post-content");
    const wordCount = mainPost ? [...mainPost.innerText.replace(/\s+/g, "")].length : 0;
    const wordCountLabel = document.getElementById("codex-word-word-count");
    if (wordCountLabel?.textContent !== `${wordCount} 个字`) {
      wordCountLabel.textContent = `${wordCount} 个字`;
    }

    syncComposeActions(group);
    const sources = postActions.map(findPostAction);
    const available = sources.some(Boolean);
    group.classList.toggle("has-native-actions", available);

    let toolbar = group.querySelector(".codex-word-native-actions");
    if (!available) {
      toolbar?.remove();
      return;
    }

    if (!toolbar) {
      toolbar = document.createElement("div");
      toolbar.className = "codex-word-native-actions";
      toolbar.setAttribute("role", "toolbar");
      toolbar.setAttribute("aria-label", "主贴快捷操作");

      postActions.forEach((action, index) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "codex-word-native-action";
        button.dataset.actionIndex = index;
        button.addEventListener("click", () => findPostAction(action)?.click());
        toolbar.append(button);
      });
      group.append(toolbar);
    }

    [...toolbar.children].forEach((button, index) => {
      const action = postActions[index];
      const source = sources[index];
      const value = source?.querySelector("span")?.textContent.trim();
      const text = `${action.icon} ${action.label}${value && value !== action.label ? ` ${value}` : ""}`;
      button.disabled = !source;
      if (button.textContent !== text) button.textContent = text;
    });
  };

  const setupFloatingReply = () => {
    if (!/^\/post-\d+(?:-\d+)?\/?$/.test(location.pathname)) return;

    const align = () => {
      const editor = document.querySelector(".md-editor.codex-word-floating-reply");
      const paper = document.querySelector("#nsk-body-left");
      if (!editor || !paper) return;
      const rect = paper.getBoundingClientRect();
      editor.style.setProperty("--codex-reply-left", `${rect.left}px`);
      editor.style.setProperty("--codex-reply-width", `${rect.width}px`);
    };

    const close = (editor) => {
      editor.classList.remove("codex-word-floating-reply");
      editor.querySelector(".fullscreen-editor")?.classList.remove("fullscreen-editor");
      editor.querySelector(".split-view-container")?.classList.remove("split-view-container");
      editor.querySelector(".codex-word-close-editor")?.remove();
    };

    const open = () => {
      const editor = document.querySelector(".md-editor");
      if (!editor) return;
      const source = editor.querySelector("#editor-body .window_header > :last-child");
      if (!source) return;

      editor.classList.add("codex-word-floating-reply");
      if (!editor.querySelector(".codex-word-close-editor")) {
        const button = source.cloneNode(true);
        button.removeAttribute("id");
        button.classList.add("codex-word-close-editor");
        button.title = "关闭回复窗口";
        button.setAttribute("aria-label", "关闭回复窗口");
        const icon = button.querySelector("span");
        if (icon) {
          icon.className = "i-icon i-icon-close";
          icon.innerHTML =
            '<svg width="16" height="16" viewBox="0 0 48 48" fill="none"><path d="M8 8L40 40M8 40L40 8" stroke="currentColor" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/></svg>';
        }
        button.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          close(editor);
        });
        source.after(button);
      }

      requestAnimationFrame(() => {
        align();
        editor.querySelector(".CodeMirror")?.CodeMirror?.refresh();
      });
    };

    window.addEventListener("resize", align);
    document.addEventListener(
      "click",
      (event) => {
        const item = event.target.closest(".menu-item");
        if (!item) return;
        const action = item.title || item.textContent.trim();
        if (["引用", "回复", "编辑"].includes(action)) open();
      },
      true,
    );
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
            <a href="${home}">文件</a>
            <a class="is-active" href="${home}">开始</a>
            <a href="${create}">插入</a>
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
              <span class="codex-word-tool">☷</span>
              <span class="codex-word-tool">≡</span>
              <span class="codex-word-tool">☰</span>
              <span class="codex-word-tool">↕</span>
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

    syncPostActions();
    setupFloatingReply();
    setupAutoHideRibbon();
    let syncQueued = false;
    new MutationObserver(() => {
      if (syncQueued) return;
      syncQueued = true;
      requestAnimationFrame(() => {
        syncQueued = false;
        syncPostActions();
      });
    }).observe(document.body, { childList: true, subtree: true, characterData: true });
  };

  if (document.body) {
    mount();
  } else {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  }
})();
