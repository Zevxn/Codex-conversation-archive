(function () {
  "use strict";

  const ROOT_SELECTOR = ".markdown-body";
  const TARGET_SELECTORS = [
    "p", "h1", "h2", "h3", "h4", "h5", "h6",
    "ul", "ol", "li", "blockquote", "table", "hr",
    ".code-block", ".mermaid-block", ".katex-display", "pre"
  ];
  const TARGET_SELECTOR = TARGET_SELECTORS.join(",");
  const IGNORED_SELECTOR = "button, input, textarea, select, summary, .md-context-menu, .code-block-toolbar";

  function notify(message) {
    window.ArchiveApp?.showToast?.(message);
  }

  function eventElement(event) {
    const target = event?.target;
    return target && typeof target.closest === "function" ? target : null;
  }

  function pointerPosition(event) {
    const clientX = Number.isFinite(event?.clientX)
      ? event.clientX
      : (Number.isFinite(event?.pageX) ? event.pageX - window.scrollX : 0);
    const clientY = Number.isFinite(event?.clientY)
      ? event.clientY
      : (Number.isFinite(event?.pageY) ? event.pageY - window.scrollY : 0);
    const pageX = Number.isFinite(event?.pageX) ? event.pageX : clientX + window.scrollX;
    const pageY = Number.isFinite(event?.pageY) ? event.pageY : clientY + window.scrollY;
    return { clientX, clientY, pageX, pageY };
  }

  function normalizeNewlines(value) {
    return String(value || "").replace(/\r\n?/g, "\n");
  }

  function makeFence(language, value) {
    const code = normalizeNewlines(value);
    const longestBacktickRun = Math.max(0, ...(code.match(/`+/g) || []).map((run) => run.length));
    const fence = "`".repeat(Math.max(3, longestBacktickRun + 1));
    const content = code.endsWith("\n") ? code : `${code}\n`;
    return `\n\n${fence}${language || ""}\n${content}${fence}\n\n`;
  }

  function renderMathForWord(markdown) {
    if (typeof katex === "undefined") return String(markdown || "");

    const codeMap = new Map();
    const mathMap = new Map();
    let serial = 0;
    let protectedSource = String(markdown || "");

    const nextToken = (kind) => `CODEXSELECTION${kind}${serial++}TOKEN`;
    const storeCode = (value) => {
      const token = nextToken("CODE");
      codeMap.set(token, value);
      return token;
    };
    const storeMath = (value) => {
      const token = nextToken("MATH");
      mathMap.set(token, value);
      return token;
    };
    const render = (expression, displayMode) => {
      try {
        return katex.renderToString(expression.trim(), {
          output: "mathml",
          displayMode,
          throwOnError: false,
          trust: false,
          strict: "ignore"
        });
      } catch {
        return displayMode ? `$$${expression}$$` : `$${expression}$`;
      }
    };

    const fencedCode = /(^|\n)([ \t]*)([`~]{3,})([^\n]*\n[\s\S]*?)(?:\n\2\3[ \t]*)(?=\n|$)/g;
    protectedSource = protectedSource.replace(
      fencedCode,
      (match, prefix) => `${prefix}${storeCode(match.slice(prefix.length))}`
    );
    protectedSource = protectedSource.replace(/(`+)([^`\n]+?)\1/g, (match) => storeCode(match));

    protectedSource = protectedSource.replace(/\$\$[\s\S]+?\$\$/g, (match) => {
      return storeMath(render(match.slice(2, -2), true));
    });
    protectedSource = protectedSource.replace(/\\\[[\s\S]+?\\\]/g, (match) => {
      return storeMath(render(match.slice(2, -2), true));
    });
    protectedSource = protectedSource.replace(/\\\([^\n]+?\\\)/g, (match) => {
      return storeMath(render(match.slice(2, -2), false));
    });
    protectedSource = protectedSource.replace(
      /(^|[^\\$])\$(?!\s)([^$\n]*?\S)\$/g,
      (match, prefix, expression) => `${prefix}${storeMath(render(expression, false))}`
    );

    for (const [token, math] of mathMap) {
      protectedSource = protectedSource.split(token).join(math);
    }
    for (const [token, code] of codeMap) {
      protectedSource = protectedSource.split(token).join(code);
    }
    return protectedSource;
  }

  async function copyMarkdownAsWordToClipboard(markdown) {
    if (typeof marked === "undefined") throw new Error("Marked 未加载");

    if (typeof ClipboardItem === "undefined" || !navigator.clipboard?.write) {
      await window.ArchiveMarkdown.copyText(markdown);
      return;
    }

    const bodyHtml = marked.parse(renderMathForWord(markdown));
    const fullHtml = `
      <!doctype html>
      <html
        xmlns:o="urn:schemas-microsoft-com:office:office"
        xmlns:w="urn:schemas-microsoft-com:office:word"
        xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
      >
        <head>
          <meta charset="utf-8">
          <title>Markdown to Word</title>
          <style>
            body { font-family: Calibri, sans-serif; }
            math { font-family: "Cambria Math", serif; }
            table { border-collapse: collapse; width: 100%; margin: 10px 0; }
            td, th { border: 1px solid #000; padding: 5px; }
            img { max-width: 100%; }
          </style>
        </head>
        <body>${bodyHtml}</body>
      </html>
    `;
    const clipboardData = new ClipboardItem({
      "text/html": new Blob([fullHtml], { type: "text/html" }),
      "text/plain": new Blob([markdown], { type: "text/plain" })
    });
    await navigator.clipboard.write([clipboardData]);
  }

  class MarkdownBlockSelector {
    constructor() {
      this.hoverClass = "md-draggable-hover";
      this.selectedClass = "md-selected";
      this.isEnabled = false;
      this.isBoxSelecting = false;
      this.isDragging = false;
      this.ignoreNextClick = false;
      this.pendingBoxSelection = null;
      this.isMultiSelect = false;
      this.initialSelection = new Set();
      this.selectedElements = new Set();
      this.startPos = { x: 0, y: 0 };
      this.currentMousePos = { x: 0, y: 0 };
      this.startClientPos = { x: 0, y: 0 };
      this.currentClientPos = { x: 0, y: 0 };
      this.scrollParent = null;
      this.startScrollPos = { top: 0, left: 0 };
      this.previousBodyCursor = "";
      this.lastAltKeyTime = 0;
      this.tempActionData = null;

      this.turndownService = typeof TurndownService === "function"
        ? new TurndownService({
          headingStyle: "atx",
          codeBlockStyle: "fenced",
          bulletListMarker: "-",
          emDelimiter: "*"
        })
        : null;
      if (this.turndownService) {
        this.turndownService.keep(["sub", "sup"]);
        this.turndownService.addRule("selectionRaw", {
          filter: (node) => node.nodeType === 1 && node.classList.contains("selection-raw"),
          replacement: (content, node) => node.textContent || ""
        });
        if (typeof turndownPluginGfm !== "undefined") {
          this.turndownService.use(turndownPluginGfm.gfm);
        }
      }

      this.createUI();
      this.bindEvents();
      window.ArchiveSelection = this;
    }

    createUI() {
      this.marquee = document.createElement("div");
      this.marquee.className = "selection-marquee";
      this.marquee.style.display = "none";
      document.body.appendChild(this.marquee);

      this.badge = document.createElement("div");
      this.badge.className = "selection-drag-badge";
      document.body.appendChild(this.badge);

      this.contextMenu = document.createElement("div");
      this.contextMenu.className = "md-context-menu";
      this.contextMenu.hidden = true;
      this.contextMenu.setAttribute("role", "menu");
      this.contextMenu.addEventListener("contextmenu", (event) => event.preventDefault());
      this.contextMenu.append(
        this.createContextMenuItem("复制 Markdown", "markdown"),
        this.createContextMenuItem("复制 Word 格式", "word")
      );
      document.body.appendChild(this.contextMenu);
    }

    createContextMenuItem(label, type) {
      const item = document.createElement("button");
      item.type = "button";
      item.className = "md-context-menu-item";
      item.setAttribute("role", "menuitem");
      item.textContent = label;
      item.addEventListener("click", async (event) => {
        event.preventDefault();
        event.stopPropagation();
        const action = this.tempActionData;
        this.hideContextMenu();
        if (action) await this.copyToClipboard(action.markdown, action.count, type);
      });
      return item;
    }

    bindEvents() {
      this.bound = {
        mouseover: (event) => this.handleMouseOver(event),
        mouseout: (event) => this.handleMouseOut(event),
        mousedown: (event) => this.handleMouseDown(event),
        mousemove: (event) => this.handleMouseMove(event),
        mouseup: (event) => this.handleMouseUp(event),
        click: (event) => this.handleClick(event),
        documentClick: (event) => this.handleDocumentClick(event),
        dragstart: (event) => this.handleDragStart(event),
        dragend: () => this.handleDragEnd(),
        keydown: (event) => this.handleKeyDown(event),
        keyup: (event) => this.handleKeyUp(event),
        contextmenu: (event) => this.handleContextMenu(event),
        scroll: (event) => this.handleScroll(event)
      };

      document.addEventListener("mouseover", this.bound.mouseover);
      document.addEventListener("mouseout", this.bound.mouseout);
      document.addEventListener("mousedown", this.bound.mousedown, true);
      document.addEventListener("mousemove", this.bound.mousemove, true);
      document.addEventListener("mouseup", this.bound.mouseup, true);
      document.addEventListener("click", this.bound.click, true);
      document.addEventListener("click", this.bound.documentClick, true);
      document.addEventListener("dragstart", this.bound.dragstart, true);
      document.addEventListener("dragend", this.bound.dragend);
      document.addEventListener("keydown", this.bound.keydown, true);
      document.addEventListener("keyup", this.bound.keyup, true);
      document.addEventListener("contextmenu", this.bound.contextmenu, true);
      document.addEventListener("scroll", this.bound.scroll, true);
    }

    handleKeyUp(event) {
      if (event.key !== "Alt") return;
      event.preventDefault();
      const now = Date.now();
      if (now - this.lastAltKeyTime < 400) {
        this.lastAltKeyTime = 0;
        this.toggleDraggerMode();
      } else {
        this.lastAltKeyTime = now;
      }
    }

    toggleDraggerMode() {
      this.isEnabled = !this.isEnabled;
      document.body.classList.toggle("block-selection-mode", this.isEnabled);
      if (this.isEnabled) {
        notify("选块模式已开启");
        return;
      }

      this.pendingBoxSelection = null;
      this.cancelBoxSelection();
      this.hideContextMenu();
      this.clearSelection();
      document.querySelectorAll(`.${this.hoverClass}`).forEach((element) => {
        element.classList.remove(this.hoverClass);
      });
      document.querySelectorAll('[data-selection-draggable="true"]').forEach((element) => {
        element.removeAttribute("draggable");
        delete element.dataset.selectionDraggable;
      });
      notify("选块模式已关闭");
    }

    isIgnoredInteraction(element) {
      return Boolean(element?.closest(IGNORED_SELECTOR));
    }

    getMarkdownRoot(element) {
      const markdownRoot = element?.closest(ROOT_SELECTOR);
      if (markdownRoot) return markdownRoot;

      // 消息卡片的内边距不属于 .markdown-body，但应允许从这里开始框选。
      const message = element?.closest(".message");
      return message?.querySelector(ROOT_SELECTOR) || null;
    }

    getTarget(event) {
      const source = eventElement(event);
      if (!source || this.isIgnoredInteraction(source)) return null;

      const root = source.closest(ROOT_SELECTOR);
      if (!root) return null;
      const element = source.closest(TARGET_SELECTOR);
      if (!element || !root.contains(element)) return null;
      return this.normalizeTarget(element, root);
    }

    normalizeTarget(element, root) {
      const codeBlock = element.closest(".code-block");
      if (codeBlock && root.contains(codeBlock)) return codeBlock;

      const mermaidBlock = element.closest(".mermaid-block");
      if (mermaidBlock && root.contains(mermaidBlock)) return mermaidBlock;

      const mathBlock = element.closest(".katex-display");
      if (mathBlock && root.contains(mathBlock)) return mathBlock;

      const table = element.closest("table");
      if (table && root.contains(table)) return table;

      if (element.tagName === "P" && element.parentElement?.tagName === "LI") {
        return element.parentElement;
      }
      return element;
    }

    getSelectionCandidates() {
      const candidates = new Set();
      document.querySelectorAll(ROOT_SELECTOR).forEach((root) => {
        root.querySelectorAll(TARGET_SELECTOR).forEach((element) => {
          if (this.isIgnoredInteraction(element)) return;
          const normalized = this.normalizeTarget(element, root);
          if (normalized && root.contains(normalized)) candidates.add(normalized);
        });
      });
      return Array.from(candidates);
    }

    markDraggable(element) {
      element.setAttribute("draggable", "true");
      element.dataset.selectionDraggable = "true";
    }

    unmarkDraggable(element) {
      if (element.dataset.selectionDraggable !== "true") return;
      element.removeAttribute("draggable");
      delete element.dataset.selectionDraggable;
    }

    handleMouseOver(event) {
      if (!this.isEnabled || this.isBoxSelecting) return;
      const target = this.getTarget(event);
      if (!target) return;

      document.querySelectorAll(`.${this.hoverClass}`).forEach((element) => {
        if (element !== target) {
          element.classList.remove(this.hoverClass);
          if (!this.selectedElements.has(element)) this.unmarkDraggable(element);
        }
      });
      if (this.selectedElements.has(target)) return;

      target.classList.add(this.hoverClass);
      this.markDraggable(target);
    }

    handleMouseOut(event) {
      const target = this.getTarget(event);
      if (!target || (event.relatedTarget && target.contains(event.relatedTarget))) return;
      target.classList.remove(this.hoverClass);
      if (!this.selectedElements.has(target)) this.unmarkDraggable(target);
    }

    getScrollParent(node) {
      let parent = node?.parentElement;
      while (parent) {
        const style = window.getComputedStyle(parent);
        if (style.overflowY === "auto" || style.overflowY === "scroll") return parent;
        parent = parent.parentElement;
      }
      return document.scrollingElement || document.documentElement;
    }

    getScrollPosition(parent) {
      if (this.isDocumentScrollParent(parent)) {
        return { top: window.scrollY, left: window.scrollX };
      }
      return { top: parent?.scrollTop || 0, left: parent?.scrollLeft || 0 };
    }

    isDocumentScrollParent(parent) {
      return parent === document.documentElement
        || parent === document.body
        || parent === document.scrollingElement;
    }

    beginBoxSelection(event, target, root) {
      const position = pointerPosition(event);
      const additive = event.ctrlKey || event.metaKey;
      if (additive) {
        this.isMultiSelect = true;
        this.initialSelection = new Set(this.selectedElements);
      } else {
        this.clearSelection();
        this.isMultiSelect = false;
        this.initialSelection = new Set();
      }

      this.ignoreNextClick = false;
      this.isBoxSelecting = true;
      this.previousBodyCursor = document.body.style.cursor;
      document.body.style.cursor = "crosshair";
      this.startPos = { x: position.pageX, y: position.pageY };
      this.currentMousePos = { x: position.pageX, y: position.pageY };
      this.startClientPos = { x: position.clientX, y: position.clientY };
      this.currentClientPos = { x: position.clientX, y: position.clientY };
      this.scrollParent = this.getScrollParent(target || root);
      this.startScrollPos = this.getScrollPosition(this.scrollParent);
      this.scrollParent?.addEventListener("scroll", this.bound.scroll, { passive: true });
      this.updateMarquee(position.pageX, position.pageY, 0, 0);
      this.marquee.style.display = "block";
    }

    handleMouseDown(event) {
      if (!this.isEnabled || event.button !== 0) return;
      const source = eventElement(event);
      if (!source || this.isIgnoredInteraction(source)) return;
      const root = this.getMarkdownRoot(source);
      if (!root) return;

      const target = this.getTarget(event);
      if ((event.ctrlKey || event.metaKey) || !target) {
        this.pendingBoxSelection = null;
        event.preventDefault();
        event.stopPropagation();
        this.beginBoxSelection(event, target, root);
        return;
      }

      // 未选中的块先进入“待判定”状态：短点击仍然执行单选，真正移动后才切换为框选。
      // 已选中的块保留原生拖拽，用于把选中内容拖到其他应用中复制。
      if (!this.selectedElements.has(target)) {
        this.pendingBoxSelection = { event, target, root, startPosition: pointerPosition(event) };
        event.preventDefault();
        event.stopPropagation();
      }
    }

    handleMouseMove(event) {
      if (!this.isEnabled) return;
      const position = pointerPosition(event);
      this.currentMousePos = { x: position.pageX, y: position.pageY };
      this.currentClientPos = { x: position.clientX, y: position.clientY };

      if (this.pendingBoxSelection && !this.isBoxSelecting) {
        const { event: startEvent, target, root, startPosition } = this.pendingBoxSelection;
        const moved = Math.hypot(
          position.clientX - startPosition.clientX,
          position.clientY - startPosition.clientY
        );
        if (moved >= 4) {
          this.pendingBoxSelection = null;
          this.beginBoxSelection(startEvent, target, root);
          this.currentMousePos = { x: position.pageX, y: position.pageY };
          this.currentClientPos = { x: position.clientX, y: position.clientY };
        }
      }

      if (this.isBoxSelecting) {
        event.preventDefault();
        event.stopPropagation();
        this.updateBoxSelection();
        return;
      }

      if (!this.getTarget(event)) {
        document.querySelectorAll(`.${this.hoverClass}`).forEach((element) => {
          element.classList.remove(this.hoverClass);
          if (!this.selectedElements.has(element)) this.unmarkDraggable(element);
        });
      }
    }

    handleMouseUp() {
      this.pendingBoxSelection = null;
      this.finishBoxSelection();
    }

    finishBoxSelection() {
      if (!this.isBoxSelecting) return;
      this.isBoxSelecting = false;
      this.marquee.style.display = "none";
      document.body.style.cursor = this.previousBodyCursor;
      this.scrollParent?.removeEventListener("scroll", this.bound.scroll);
      this.scrollParent = null;
    }

    cancelBoxSelection() {
      this.ignoreNextClick = false;
      this.finishBoxSelection();
    }

    getScrollBounds(parent) {
      if (this.isDocumentScrollParent(parent)) {
        return {
          left: window.scrollX,
          top: window.scrollY,
          right: window.scrollX + window.innerWidth,
          bottom: window.scrollY + window.innerHeight
        };
      }
      const rect = parent.getBoundingClientRect();
      return {
        left: rect.left + window.scrollX,
        top: rect.top + window.scrollY,
        right: rect.right + window.scrollX,
        bottom: rect.bottom + window.scrollY
      };
    }

    updateBoxSelection() {
      if (!this.isBoxSelecting || !this.scrollParent) return;

      const scrollPosition = this.getScrollPosition(this.scrollParent);
      const scrollDeltaX = this.isDocumentScrollParent(this.scrollParent)
        ? 0
        : scrollPosition.left - this.startScrollPos.left;
      const scrollDeltaY = this.isDocumentScrollParent(this.scrollParent)
        ? 0
        : scrollPosition.top - this.startScrollPos.top;
      const rawStartX = this.startPos.x - scrollDeltaX;
      const rawStartY = this.startPos.y - scrollDeltaY;
      const rawCurrentX = this.currentClientPos.x + window.scrollX;
      const rawCurrentY = this.currentClientPos.y + window.scrollY;

      if (Math.hypot(rawCurrentX - rawStartX, rawCurrentY - rawStartY) < 3) return;
      this.ignoreNextClick = true;

      const boxLeft = Math.min(rawStartX, rawCurrentX);
      const boxRight = Math.max(rawStartX, rawCurrentX);
      const boxTop = Math.min(rawStartY, rawCurrentY);
      const boxBottom = Math.max(rawStartY, rawCurrentY);
      const bounds = this.getScrollBounds(this.scrollParent);
      const visibleLeft = Math.max(boxLeft, bounds.left);
      const visibleTop = Math.max(boxTop, bounds.top);
      const visibleRight = Math.min(boxRight, bounds.right);
      const visibleBottom = Math.min(boxBottom, bounds.bottom);

      if (visibleRight <= visibleLeft || visibleBottom <= visibleTop) {
        this.marquee.style.display = "none";
      } else {
        this.marquee.style.display = "block";
        this.updateMarquee(visibleLeft, visibleTop, visibleRight - visibleLeft, visibleBottom - visibleTop);
      }

      this.detectIntersection({
        left: boxLeft,
        top: boxTop,
        width: boxRight - boxLeft,
        height: boxBottom - boxTop
      });
    }

    handleScroll() {
      this.hideContextMenu();
      if (this.isBoxSelecting) requestAnimationFrame(() => this.updateBoxSelection());
    }

    updateMarquee(left, top, width, height) {
      this.marquee.style.left = `${left}px`;
      this.marquee.style.top = `${top}px`;
      this.marquee.style.width = `${width}px`;
      this.marquee.style.height = `${height}px`;
    }

    detectIntersection(rect) {
      const right = rect.left + rect.width;
      const bottom = rect.top + rect.height;
      this.getSelectionCandidates().forEach((element) => {
        const box = element.getBoundingClientRect();
        if (box.width <= 0 || box.height <= 0) return;

        const left = box.left + window.scrollX;
        const top = box.top + window.scrollY;
        const elementRight = left + box.width;
        const elementBottom = top + box.height;
        const overlaps = !(elementRight < rect.left || left > right || elementBottom < rect.top || top > bottom);

        if (overlaps) {
          this.addToSelection(element);
        } else if (this.isMultiSelect) {
          if (!this.initialSelection.has(element)) this.removeFromSelection(element);
        } else if (this.selectedElements.has(element)) {
          this.removeFromSelection(element);
        }
      });
    }

    addToSelection(element) {
      if (this.selectedElements.has(element)) return;
      this.selectedElements.add(element);
      element.classList.add(this.selectedClass);
      element.classList.remove(this.hoverClass);
      this.markDraggable(element);
    }

    removeFromSelection(element) {
      if (!this.selectedElements.has(element)) return;
      this.selectedElements.delete(element);
      element.classList.remove(this.selectedClass, this.hoverClass);
      this.unmarkDraggable(element);
    }

    clearSelection() {
      this.selectedElements.forEach((element) => {
        element.classList.remove(this.selectedClass);
        this.unmarkDraggable(element);
      });
      this.selectedElements.clear();
    }

    pruneSelection() {
      this.selectedElements.forEach((element) => {
        if (!element.isConnected || !element.closest(ROOT_SELECTOR)) this.removeFromSelection(element);
      });
    }

    getUniqueTopLevelElements() {
      this.pruneSelection();
      const elements = Array.from(this.selectedElements);
      return elements
        .filter((element) => !elements.some((parent) => parent !== element && parent.contains(element)))
        .sort((left, right) => {
          const position = left.compareDocumentPosition(right);
          return position & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
        });
    }

    isElementInSelection(target) {
      this.pruneSelection();
      return Array.from(this.selectedElements).some((element) => element === target || element.contains(target));
    }

    generateMarkdown(target = null) {
      this.pruneSelection();
      let items = [];
      if (target && !this.selectedElements.has(target)) {
        items = [target];
      } else if (this.selectedElements.size > 0) {
        items = this.getUniqueTopLevelElements();
      } else if (target) {
        items = [target];
      } else {
        return null;
      }
      if (!items.length) return null;

      const container = document.createElement("div");
      items.forEach((element) => {
        container.append(element.cloneNode(true), document.createTextNode("\n\n"));
      });
      this.sanitizeDomForMarkdown(container);

      let markdown = this.turndownService
        ? this.turndownService.turndown(container.innerHTML)
        : container.textContent || "";
      markdown = normalizeNewlines(markdown)
        .replace(/\n{3,}/g, "\n\n")
        .replace(/^[ \t]+```/gm, "```")
        .replace(/\\(\$)/g, "$")
        .trim();
      return { markdown, count: items.length };
    }

    sanitizeDomForMarkdown(root) {
      root.querySelectorAll(".selection-marquee, .selection-drag-badge, .md-context-menu, .code-block-toolbar").forEach((element) => {
        element.remove();
      });

      Array.from(root.querySelectorAll(".katex-display")).forEach((element) => {
        const annotation = element.querySelector('annotation[encoding="application/x-tex"]');
        const latex = annotation?.textContent?.trim();
        if (!latex) return;
        const raw = document.createElement("span");
        raw.className = "selection-raw";
        raw.textContent = `\n\n$$${latex}$$\n\n`;
        element.replaceWith(raw);
      });

      Array.from(root.querySelectorAll(".katex")).forEach((element) => {
        if (element.closest(".katex-display")) return;
        const annotation = element.querySelector('annotation[encoding="application/x-tex"]');
        const latex = annotation?.textContent?.trim();
        if (!latex) return;
        const raw = document.createElement("span");
        raw.className = "selection-raw";
        raw.textContent = `$${latex}$`;
        element.replaceWith(raw);
      });

      Array.from(root.querySelectorAll(".code-block")).forEach((block) => {
        const code = block.querySelector("pre > code");
        if (!code) return;
        const language = block.dataset.mdLanguage || Array.from(code.classList)
          .find((name) => name.startsWith("language-"))
          ?.slice("language-".length) || "";
        const raw = document.createElement("pre");
        raw.className = "selection-raw";
        raw.textContent = makeFence(language === "text" ? "" : language, block.dataset.mdSource ?? code.textContent ?? "");
        block.replaceWith(raw);
      });

      Array.from(root.querySelectorAll(".mermaid-block")).forEach((block) => {
        const source = block.dataset.mdSource;
        if (source == null) return;
        const raw = document.createElement("pre");
        raw.className = "selection-raw";
        raw.textContent = makeFence("mermaid", source);
        block.replaceWith(raw);
      });

      Array.from(root.querySelectorAll("table")).forEach((table) => {
        const markdown = this.convertTableToMarkdown(table);
        const raw = document.createElement("pre");
        raw.className = "selection-raw";
        raw.textContent = `\n\n${markdown}\n\n`;
        table.replaceWith(raw);
      });
    }

    convertTableToMarkdown(table) {
      const rows = Array.from(table.querySelectorAll("tr"));
      if (!rows.length) return "";

      return rows.map((row, rowIndex) => {
        const cells = Array.from(row.children).filter((cell) => ["TH", "TD"].includes(cell.tagName));
        const values = cells.map((cell) => {
          let value = this.turndownService
            ? this.turndownService.turndown(cell.innerHTML)
            : cell.textContent || "";
          value = value.replace(/\r?\n/g, "<br>").replace(/\|/g, "\\|").trim();
          return value;
        });
        const rowMarkdown = `| ${values.join(" | ")} |`;
        return rowIndex === 0
          ? `${rowMarkdown}\n| ${values.map(() => "---").join(" | ")} |`
          : rowMarkdown;
      }).join("\n");
    }

    async handleKeyDown(event) {
      if (!this.isEnabled || !(event.ctrlKey || event.metaKey) || !["c", "C"].includes(event.key)) return;
      if (/INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName || "")) return;
      if (document.activeElement?.isContentEditable) return;

      this.pruneSelection();
      if (!this.selectedElements.size) return;
      const result = this.generateMarkdown();
      if (!result?.markdown) return;

      event.preventDefault();
      event.stopPropagation();
      await this.copyToClipboard(result.markdown, result.count);
    }

    async handleContextMenu(event) {
      if (!this.isEnabled) return;
      const source = eventElement(event);
      if (!source || this.isIgnoredInteraction(source)) {
        this.hideContextMenu();
        return;
      }

      const target = this.getTarget(event);
      if (!target) {
        this.hideContextMenu();
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      let result;
      if (this.isElementInSelection(target)) {
        result = this.generateMarkdown();
      } else {
        this.clearSelection();
        this.addToSelection(target);
        result = this.generateMarkdown(target);
      }

      if (result?.markdown) {
        this.tempActionData = result;
        this.showContextMenu(event.clientX, event.clientY);
      }
    }

    showContextMenu(x, y) {
      this.contextMenu.hidden = false;
      const menuWidth = this.contextMenu.offsetWidth || 136;
      const menuHeight = this.contextMenu.offsetHeight || 76;
      const left = x + menuWidth > window.innerWidth ? Math.max(8, x - menuWidth) : x;
      const top = y + menuHeight > window.innerHeight ? Math.max(8, y - menuHeight) : y;
      this.contextMenu.style.left = `${left}px`;
      this.contextMenu.style.top = `${top}px`;
    }

    hideContextMenu() {
      this.contextMenu.hidden = true;
      this.tempActionData = null;
    }

    handleDocumentClick(event) {
      if (eventElement(event)?.closest(".md-context-menu")) return;
      this.hideContextMenu();
    }

    handleClick(event) {
      if (!this.isEnabled || this.isDragging) return;
      if (eventElement(event)?.closest(".md-context-menu")) return;
      if (this.ignoreNextClick) {
        this.ignoreNextClick = false;
        event.preventDefault();
        event.stopPropagation();
        return;
      }
      if (this.isBoxSelecting) return;

      const source = eventElement(event);
      if (!source || this.isIgnoredInteraction(source)) return;
      const root = source.closest(ROOT_SELECTOR);
      if (!root) return;
      const target = this.getTarget(event);
      event.preventDefault();
      event.stopPropagation();

      if (event.ctrlKey || event.metaKey) {
        if (!target) return;
        if (this.selectedElements.has(target)) this.removeFromSelection(target);
        else this.addToSelection(target);
        return;
      }

      if (target) {
        const isOnlySelf = this.selectedElements.size === 1 && this.selectedElements.has(target);
        if (!isOnlySelf) {
          this.clearSelection();
          this.addToSelection(target);
        }
      } else {
        this.clearSelection();
      }
    }

    handleDragStart(event) {
      if (!this.isEnabled) return;
      const source = eventElement(event);
      if (!source || this.isIgnoredInteraction(source)) return;

      // 某些浏览器会在可拖拽块上先触发 dragstart；此时将它转为框选，避免原生拖拽抢走手势。
      if (this.pendingBoxSelection) {
        const { event: startEvent, target, root } = this.pendingBoxSelection;
        this.pendingBoxSelection = null;
        event.preventDefault();
        this.beginBoxSelection(startEvent, target, root);
        this.updateBoxSelection();
        return;
      }

      const target = this.getTarget(event);
      if (!target || !event.dataTransfer) return;

      this.isDragging = true;
      if (!this.selectedElements.has(target)) {
        this.clearSelection();
        this.addToSelection(target);
      }

      const result = this.generateMarkdown();
      if (!result?.markdown) return;
      event.dataTransfer.setData("text/plain", result.markdown);
      event.dataTransfer.effectAllowed = "copy";
      this.badge.textContent = `📝 ${result.count} 个块`;
      event.dataTransfer.setDragImage(this.badge, 0, 10);
    }

    handleDragEnd() {
      window.setTimeout(() => { this.isDragging = false; }, 50);
    }

    async copyToClipboard(markdown, count, type = "markdown") {
      try {
        if (type === "word") await copyMarkdownAsWordToClipboard(markdown);
        else await window.ArchiveMarkdown.copyText(markdown);
        notify(type === "word"
          ? `已复制 ${count} 个块的 Word 格式`
          : `已复制 ${count} 个块的 Markdown`);
      } catch (error) {
        console.error("选中内容复制失败", error);
        notify("选中内容复制失败");
      }
    }
  }

  const initialize = () => new MarkdownBlockSelector();
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initialize);
  else initialize();
})();
