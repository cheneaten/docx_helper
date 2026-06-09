/* ============================================================
   DOCX 编辑器 - 主应用逻辑
   ============================================================ */

(function() {
    'use strict';

    if (typeof mammoth === 'undefined') {
        document.getElementById('editor').innerHTML =
            '<div style="color:#dc2626;padding:40px;text-align:center">' +
            '<h2>依赖库加载失败</h2>' +
            '<p>mammoth.js 库未加载。请检查网络连接后刷新页面。</p>' +
            '<button onclick="location.reload()" style="padding:8px 20px;margin-top:12px;cursor:pointer">重新加载</button>' +
            '</div>';
        console.error('mammoth.js 未加载');
    }
    if (typeof JSZip === 'undefined') console.error('JSZip 未加载');
    if (typeof saveAs === 'undefined') console.error('FileSaver.js 未加载');

    var $ = function(id) { return document.getElementById(id); };

    // 离开确认标志：设为 true 后跳过 beforeunload 弹窗（防止双重弹窗）
    var _allowNavigation = false;
    // 离开操作类型：'reload' | 'hardReload' | 'close'
    var _leaveAction = null;
    var editor = $('editor');
    var fileInput = $('fileInput');
    var importBtn = $('importBtn');
    var exportBtn = $('exportBtn');
    var printBtn = $('printBtn');
    var loadingOverlay = $('loadingOverlay');
    var loadingText = $('loadingText');
    var docStatus = $('docStatus');
    var tocContainer = $('tocContainer');
    var tocRefresh = $('tocRefresh');
    var tocCollapseAll = $('tocCollapseAll');
    var tocExpandAll = $('tocExpandAll');
    var searchInput = $('searchInput');
    var replaceInput = $('replaceInput');
    var searchBtn = $('searchBtn');
    var replaceBtn = $('replaceBtn');
    var replaceAllBtn = $('replaceAllBtn');
    var prevMatch = $('prevMatch');
    var nextMatch = $('nextMatch');
    var matchInfo = $('matchInfo');
    var imgCenterToggle = $('imgCenterToggle');
    var bodyFont = $('bodyFont');
    var bodySize = $('bodySize');
    var bodyLineHeight = $('bodyLineHeight');
    var applyBodyFormat = $('applyBodyFormat');
    var editorContainer = $('editorContainer');

    // ============================================================
    // 📑 多文档标签管理器（会话切换模式）
    // ============================================================
    var MAX_TABS = 3;
    var TAB_DB_VERSION = 2; // IndexedDB 版本号

    function createDocumentSession(docId, title) {
        return {
            id: docId,                    // 'doc_1', 'doc_2', 'doc_3'
            title: title || '文档 1',      // 标签显示名
            customTitle: false,           // 用户是否手动编辑过标签名
            html: '',                     // editor.innerHTML 快照
            undoHistory: [],              // 撤销栈
            undoIndex: -1,
            imageDataMap: new Map(),      // 图片数据（内存中为 Map）
            imageCounter: 0,
            scrollTop: 0,
            currentMatches: [],
            currentMatchIndex: -1,
            searchText: '',
            isImporting: false,
            anchors: [],                  // 锚点列表
            foldPoints: [],              // 折叠标记点
            foldRegions: [],             // 折叠区域
            sourceFileName: null,         // 导入的原始文件名
            sourceImportTime: null,       // 导入时间 ISO 字符串
            savedAt: null,
            _dirty: false,               // 未保存标记（用于标签页显示 *）
            _contentLoaded: true,         // 内容是否已加载（新建文档默认已加载）
            _lastSavedHtml: '',           // 上次保存的 HTML，用于增量保存比对
            _imagesChanged: false         // 图片数据是否有变更
        };
    }

    var tabManager = {
        sessions: [],
        activeIndex: 0,
        tabIdCounter: 0,

        init: function() {
            var self = this;
            var tabList = document.getElementById('tabList');
            var addBtn = document.getElementById('addTabBtn');
            if (!tabList || !addBtn) {
                console.error('标签栏 DOM 未找到');
                return;
            }
            // 从 IndexedDB 恢复或创建默认文档
            return this.loadAllFromDB().then(function(sessions) {
                if (sessions && sessions.length > 0) {
                    self.sessions = sessions;
                    self.activeIndex = 0;
                    // 恢复 _meta_ 中的 activeIndex
                    // activeIndex 在 loadAllFromDB 中已设置
                } else {
                    // 首次启动：创建默认文档
                    var s = createDocumentSession('doc_1', '文档 1');
                    self.sessions = [s];
                    self.activeIndex = 0;
                }
                self.tabIdCounter = self.sessions.length;
                // 🔄 懒加载：仅加载活跃标签页的完整内容
                var active = self.getActive();
                if (active && !active._contentLoaded) {
                    return self.loadSessionContentFromDB(active.id).then(function() {
                        return self._initActiveSession(active, tabList, addBtn);
                    });
                } else {
                    return self._initActiveSession(active, tabList, addBtn);
                }
            });
        },

        // 初始化活跃会话到编辑器（从 init 中提取）
        _initActiveSession: function(active, tabList, addBtn) {
            var self = this;
            // 将活跃会话状态同步到全局变量
            syncStateFromSession(active);
            // 恢复编辑器内容（抑制 input 事件避免误标记为未保存）
            suppressCaptureUntil = Date.now() + 500;
            if (active.html && !isPlaceholderContent(active.html)) {
                editor.innerHTML = active.html;
            }
            // 恢复滚动位置
            if (active.scrollTop) {
                setTimeout(function() { editorContainer.scrollTop = active.scrollTop; }, 50);
            }
            self.renderTabs();
            // 初始更新统计信息
            setTimeout(function() { updateStats(); }, 100);
            // 绑定事件
            addBtn.addEventListener('click', function() { self.addTab(); });
            // 绑定全局 tooltip 隐藏
            document.addEventListener('click', function(e) {
                if (!e.target.closest('.tab-item')) {
                    var tip = document.getElementById('tabTooltip');
                    if (tip) tip.classList.remove('visible');
                }
            });
            return self;
        },

        getActive: function() {
            return this.sessions[this.activeIndex] || null;
        },

        switchTo: function(index) {
            if (index === this.activeIndex) return;
            if (index < 0 || index >= this.sessions.length) return;
            var self = this;
            // 1. 完整保存当前会话到 DB（异步后台执行）
            syncStateToSession();
            saveDocumentFull();
            // 2. 切换活跃索引
            this.activeIndex = index;
            var session = this.getActive();
            if (!session) return;
            // 🔄 懒加载：如果目标会话内容未加载，先从 DB 加载
            if (!session._contentLoaded) {
                var overlay = document.getElementById('loadingOverlay');
                var overlayText = document.getElementById('loadingText');
                if (overlay) overlay.classList.remove('hidden');
                if (overlayText) overlayText.textContent = '正在加载文档...';
                return this.loadSessionContentFromDB(session.id).then(function(loaded) {
                    if (overlay) overlay.classList.add('hidden');
                    if (!loaded) {
                        showToast('文档加载失败', 'error', 2000);
                        return;
                    }
                    // 3. 加载目标会话到编辑器
                    self.loadSessionToEditor(session);
                    // 4. 刷新 UI
                    self.renderTabs();
                    // 5. 保存活跃索引到 DB
                    self.saveMetaToDB();
                });
            }
            // 3. 加载目标会话（已加载的情况）
            this.loadSessionToEditor(session);
            // 4. 刷新 UI
            this.renderTabs();
            // 5. 保存活跃索引到 DB
            this.saveMetaToDB();
        },

        addTab: function(title) {
            if (this.sessions.length >= MAX_TABS) {
                showToast('最多只能创建 ' + MAX_TABS + ' 个文档标签', 'warning', 2000);
                return;
            }
            // 保存当前会话
            this.saveActiveSession();
            // 创建新会话
            this.tabIdCounter++;
            var docId = 'doc_' + this.tabIdCounter;
            var defaultTitle = title || ('文档 ' + this.tabIdCounter);
            var session = createDocumentSession(docId, defaultTitle);
            this.sessions.push(session);
            this.activeIndex = this.sessions.length - 1;
            // 加载新会话（空白编辑器）
            this.loadSessionToEditor(session);
            this.renderTabs();
            this.saveAllToDB();
            showToast('已创建: ' + defaultTitle, 'info', 1500);
        },

        closeTab: function(index) {
            if (this.sessions.length <= 1) {
                showToast('至少保留一个文档标签', 'warning', 1500);
                return;
            }
            var session = this.sessions[index];
            var docId = session.id;
            // 从数组中移除
            this.sessions.splice(index, 1);
            // 从 IndexedDB 删除
            this.deleteFromDB(docId);
            // 调整活跃索引
            if (index <= this.activeIndex) {
                this.activeIndex = Math.max(0, this.activeIndex - 1);
            }
            if (this.activeIndex >= this.sessions.length) {
                this.activeIndex = this.sessions.length - 1;
            }
            // 加载新活跃会话
            var newActive = this.getActive();
            if (newActive) {
                this.loadSessionToEditor(newActive);
            }
            this.renderTabs();
            this.saveMetaToDB();
            showToast('已关闭标签', 'info', 1200);
        },

        saveActiveSession: function() {
            var session = this.getActive();
            if (!session) return;
            // 保存 HTML 内容
            session.html = editor.innerHTML;
            // 保存撤销栈（当前引用）
            session.undoHistory = undoHistory;
            session.undoIndex = undoIndex;
            // 保存图片数据（当前引用）
            session.imageDataMap = imageDataMap;
            session.imageCounter = imageCounter;
            // 保存滚动位置
            session.scrollTop = editorContainer ? editorContainer.scrollTop : 0;
            // 保存搜索状态
            session.currentMatches = currentMatches;
            session.currentMatchIndex = currentMatchIndex;
            session.searchText = searchText;
            session.isImporting = isImporting;
            // 保存锚点
            session.anchors = anchors.slice();
            // 保存折叠数据
            session.foldPoints = foldPoints.slice();
            session.foldRegions = foldRegions.map(function(r) {
                return { id: r.id, startPid: r.startPid, endPid: r.endPid, isFolded: r.isFolded };
            });
        },

        loadSessionToEditor: function(session) {
            if (!session) return;
            // 抑制切换标签时触发的 input 事件（避免误标记为未保存）
            suppressCaptureUntil = Date.now() + 500;
            // 恢复编辑器内容
            editor.innerHTML = session.html || '';
            // 恢复撤销栈（浅拷贝数组）
            undoHistory = (session.undoHistory || []).slice();
            undoIndex = session.undoIndex != null ? session.undoIndex : -1;
            // 恢复图片数据
            imageDataMap = session.imageDataMap || new Map();
            imageCounter = session.imageCounter || 0;
            // 恢复搜索状态
            currentMatches = (session.currentMatches || []).slice();
            currentMatchIndex = session.currentMatchIndex != null ? session.currentMatchIndex : -1;
            searchText = session.searchText || '';
            isImporting = session.isImporting || false;
            anchors = (session.anchors || []).slice();
            foldPoints = (session.foldPoints || []).slice();
            foldRegions = (session.foldRegions || []).slice();
            // 切换标签后刷新装订线和面板
            renderAnchorGutter();
            renderAnchorPanel();
            renderFoldGutter();
            ensureParagraphIds();
            setTimeout(function() { reapplyFoldRegions(); }, 100);
            // 同步搜索输入框
            if (searchInput) searchInput.value = session.searchText || '';
            if (replaceInput) replaceInput.value = '';
            // 重置撤销控制变量（切换后重新计数）
            undoBlocked = false;
            lastBeforeInputTime = 0;
            suppressCaptureUntil = 0;
            justRestored = false;
            // 恢复滚动位置
            if (session.scrollTop) {
                setTimeout(function() { editorContainer.scrollTop = session.scrollTop; }, 30);
            }
            // 刷新 UI
            applyHeadingStylesToEditor();
            if (typeof applyBodyFormatFn === 'function') applyBodyFormatFn();
            generateTOC();
            renumber();
            if (imgCenterToggle && imgCenterToggle.checked) centerAllImages();
            if (typeof applyAllCodeThemes === 'function') applyAllCodeThemes();
            updateStats();
            // 更新搜索面板
            updateSearchUI();
            // 更新状态
            setStatus('已切换到: ' + (session.title || '文档'));
        },

        renameTab: function(index, newTitle) {
            if (index < 0 || index >= this.sessions.length) return;
            var session = this.sessions[index];
            session.title = newTitle || ('文档 ' + (index + 1));
            this.renderTabs();
            if (session._contentLoaded) {
                this.saveAllToDB();
            } else {
                // 🔄 懒加载：直接更新 DB 中的标题，不触碰未加载的内容
                this._updateSessionTitleInDB(session);
            }
        },

        startRename: function(index) {
            var self = this;
            var session = this.sessions[index];
            if (!session) return;
            var tabItems = document.querySelectorAll('.tab-item');
            var tabItem = tabItems[index];
            if (!tabItem) return;
            var nameSpan = tabItem.querySelector('.tab-name');
            if (!nameSpan) return;
            // 替换为 input
            var input = document.createElement('input');
            input.type = 'text';
            input.className = 'tab-name-input';
            input.value = session.title;
            input.setAttribute('data-tab-index', index);
            nameSpan.replaceWith(input);
            input.focus();
            input.select();
            // 确认：Enter 或失焦
            function commit() {
                var newTitle = input.value.trim();
                if (!newTitle) newTitle = '文档 ' + (index + 1);
                self.commitRename(index, newTitle);
            }
            input.addEventListener('keydown', function(e) {
                if (e.key === 'Enter') { e.preventDefault(); commit(); }
                if (e.key === 'Escape') { input.value = session.title; commit(); }
            });
            input.addEventListener('blur', function() { commit(); });
        },

        commitRename: function(index, newTitle) {
            var session = this.sessions[index];
            if (!session) return;
            session.title = newTitle;
            session.customTitle = true;
            this.renderTabs();
            if (session._contentLoaded) {
                this.saveAllToDB();
            } else {
                // 🔄 懒加载：直接更新 DB 中的标题
                this._updateSessionTitleInDB(session);
            }
        },

        // 🔄 懒加载：仅更新 DB 中文档的标题（不修改未加载的内容）
        _updateSessionTitleInDB: function(session) {
            if (!session || !session.id) return;
            openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readwrite');
                var store = tx.objectStore('docs');
                var getReq = store.get(session.id);
                getReq.onsuccess = function() {
                    var doc = getReq.result;
                    if (doc) {
                        doc.title = session.title;
                        doc.customTitle = session.customTitle;
                        store.put(doc);
                    }
                };
            }).catch(function() {});
        },

        // ---- Tooltip ----
        showTooltip: function(session, event) {
            var tip = document.getElementById('tabTooltip');
            if (!tip || !session) return;
            var html = '';
            if (session.sourceFileName) {
                html += '<div class="tooltip-row"><span class="tooltip-label">📄 来源文件：</span><span class="tooltip-value">' + escHtml(session.sourceFileName) + '</span></div>';
                if (session.sourceImportTime) {
                    var d = new Date(session.sourceImportTime);
                    var timeStr = d.toLocaleString('zh-CN');
                    html += '<div class="tooltip-row"><span class="tooltip-label">📅 导入时间：</span><span class="tooltip-value">' + timeStr + '</span></div>';
                }
            } else {
                html = '<div class="tooltip-row"><span class="tooltip-label">📝 手动创建文档</span></div>';
            }
            tip.innerHTML = html;
            tip.classList.add('visible');
            // 定位 tooltip
            var rect = event.target.closest('.tab-item').getBoundingClientRect();
            var tipW = tip.offsetWidth || 280;
            var left = rect.left + (rect.width - tipW) / 2;
            if (left < 8) left = 8;
            if (left + tipW > window.innerWidth - 8) left = window.innerWidth - tipW - 8;
            tip.style.left = left + 'px';
            tip.style.top = (rect.bottom + 6) + 'px';
        },

        hideTooltip: function() {
            var tip = document.getElementById('tabTooltip');
            if (tip) tip.classList.remove('visible');
        },

        // ---- 持久化 ----
        saveAllToDB: function() {
            var self = this;
            this.saveActiveSession();
            return openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readwrite');
                var store = tx.objectStore('docs');
                // 保存每个已加载的文档（未加载的文档内容未变，跳过以避免覆盖）
                self.sessions.forEach(function(s) {
                    if (!s._contentLoaded) return; // 🔄 懒加载：跳过未加载的会话
                    s._dirty = false; // 清除未保存标记
                    store.put({
                        id: s.id,
                        title: s.title,
                        customTitle: s.customTitle,
                        content: s.html,
                        imageData: s.imageDataMap ? Array.from(s.imageDataMap.entries()) : [],
                        undoHistory: (s.undoHistory || []).slice(),
                        undoIndex: s.undoIndex,
                        anchors: (s.anchors || []).slice(),
                        foldPoints: (s.foldPoints || []).slice(),
                        foldRegions: (s.foldRegions || []).slice(),
                        scrollTop: s.scrollTop,
                        sourceFileName: s.sourceFileName,
                        sourceImportTime: s.sourceImportTime,
                        savedAt: new Date().toISOString()
                    });
                });
                // 保存元信息
                store.put({
                    id: '_meta_',
                    activeIndex: self.activeIndex,
                    tabCount: self.sessions.length,
                    tabIdCounter: self.tabIdCounter
                });
                return new Promise(function(resolve) { tx.oncomplete = resolve; tx.onerror = resolve; });
            }).then(function() {
                self.renderTabs(); // 刷新标签（去除 * 标记）
            }).catch(function(err) {
                console.warn('批量保存失败:', err);
            });
        },

        loadAllFromDB: function() {
            var self = this;
            return openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readonly');
                var store = tx.objectStore('docs');
                var getAllReq = store.getAll();
                return new Promise(function(resolve) {
                    getAllReq.onsuccess = function() {
                        var allDocs = getAllReq.result || [];
                        var sessions = [];
                        var meta = null;
                        allDocs.forEach(function(doc) {
                            if (doc.id === '_meta_') {
                                meta = doc;
                            } else if (doc.id && doc.id.indexOf('doc_') === 0) {
                                // 🔄 懒加载：仅恢复元数据，不加载完整内容
                                var s = createDocumentSession(doc.id, doc.title || '文档');
                                s.customTitle = doc.customTitle || false;
                                s.sourceFileName = doc.sourceFileName || null;
                                s.sourceImportTime = doc.sourceImportTime || null;
                                s.savedAt = doc.savedAt || null;
                                // 内容字段延迟加载（标记为未加载）
                                s.html = '';
                                s.imageDataMap = new Map();
                                s.imageCounter = 0;
                                s.undoHistory = [];
                                s.undoIndex = -1;
                                s.anchors = [];
                                s.foldPoints = [];
                                s.foldRegions = [];
                                s.scrollTop = 0;
                                s._contentLoaded = false;
                                sessions.push(s);
                            }
                            // 兼容旧版单文档数据（key='editor'）
                            else if (doc.id === 'editor') {
                                var s = createDocumentSession('doc_1', '文档 1');
                                s.html = doc.content || '';
                                s.imageDataMap = new Map(doc.imageData || []);
                                s.scrollTop = doc.scrollTop || 0;
                                s.savedAt = doc.savedAt || null;
                                s.undoHistory = [];
                                s.undoIndex = -1;
                                sessions.push(s);
                                // 异步删除旧数据
                                setTimeout(function() {
                                    openDocDB().then(function(delDb) {
                                        var delTx = delDb.transaction('docs', 'readwrite');
                                        delTx.objectStore('docs').delete('editor');
                                    }).catch(function() {});
                                }, 1000);
                            }
                        });
                        if (meta) {
                            self.activeIndex = meta.activeIndex || 0;
                            self.tabIdCounter = meta.tabIdCounter || sessions.length;
                            if (self.activeIndex >= sessions.length) self.activeIndex = 0;
                        }
                        resolve(sessions);
                    };
                    getAllReq.onerror = function() { resolve([]); };
                });
            }).catch(function() { return []; });
        },

        // 🔄 懒加载：从 DB 加载单个文档的完整内容
        loadSessionContentFromDB: function(docId) {
            var self = this;
            // 找到对应会话
            var session = null;
            for (var i = 0; i < self.sessions.length; i++) {
                if (self.sessions[i].id === docId) { session = self.sessions[i]; break; }
            }
            if (!session) return Promise.resolve(null);
            // 如果已加载，直接返回
            if (session._contentLoaded) return Promise.resolve(session);
            return openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readonly');
                var store = tx.objectStore('docs');
                var getReq = store.get(docId);
                return new Promise(function(resolve) {
                    getReq.onsuccess = function() {
                        var doc = getReq.result;
                        if (doc) {
                            session.html = doc.content || '';
                            session.imageDataMap = new Map(doc.imageData || []);
                            session.imageCounter = 0;
                            session.undoHistory = (doc.undoHistory || []).slice();
                            session.undoIndex = doc.undoIndex != null ? doc.undoIndex : -1;
                            session.anchors = (doc.anchors || []).slice();
                            session.foldPoints = (doc.foldPoints || []).slice();
                            session.foldRegions = (doc.foldRegions || []).slice();
                            session.scrollTop = doc.scrollTop || 0;
                            session.sourceFileName = doc.sourceFileName || null;
                            session.sourceImportTime = doc.sourceImportTime || null;
                            session._contentLoaded = true;
                            session._lastSavedHtml = session.html;
                        }
                        resolve(session);
                    };
                    getReq.onerror = function() { resolve(null); };
                });
            }).catch(function() { return null; });
        },

        deleteFromDB: function(docId) {
            openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readwrite');
                tx.objectStore('docs').delete(docId);
            }).catch(function() {});
        },

        saveMetaToDB: function() {
            openDocDB().then(function(db) {
                var tx = db.transaction('docs', 'readwrite');
                tx.objectStore('docs').put({
                    id: '_meta_',
                    activeIndex: tabManager.activeIndex,
                    tabCount: tabManager.sessions.length,
                    tabIdCounter: tabManager.tabIdCounter
                });
            }).catch(function() {});
        },

        renderTabs: function() {
            var self = this;
            var tabList = document.getElementById('tabList');
            var addBtn = document.getElementById('addTabBtn');
            if (!tabList) return;
            // 清空
            tabList.innerHTML = '';
            // 渲染每个标签
            this.sessions.forEach(function(session, index) {
                var tabItem = document.createElement('div');
                tabItem.className = 'tab-item' + (index === self.activeIndex ? ' active' : '');
                tabItem.setAttribute('data-tab-index', index);

                var nameSpan = document.createElement('span');
                nameSpan.className = 'tab-name';
                nameSpan.textContent = (session.title || ('文档 ' + (index + 1))) + (session._dirty ? ' *' : '');
                nameSpan.title = ''; // 使用自定义 tooltip

                var closeBtn = document.createElement('button');
                closeBtn.className = 'tab-close';
                closeBtn.innerHTML = '&#10005;';
                closeBtn.title = '关闭标签';
                closeBtn.addEventListener('click', function(e) {
                    e.stopPropagation();
                    e.preventDefault();
                    self.closeTab(index);
                });

                tabItem.appendChild(nameSpan);
                if (self.sessions.length > 1) {
                    tabItem.appendChild(closeBtn);
                }

                // 点击切换
                tabItem.addEventListener('click', function(e) {
                    if (e.target === closeBtn) return;
                    self.switchTo(index);
                });

                // 双击编辑标签名
                tabItem.addEventListener('dblclick', function(e) {
                    if (e.target === closeBtn) return;
                    self.startRename(index);
                });

                // Hover tooltip
                tabItem.addEventListener('mouseenter', function(e) {
                    self.showTooltip(session, e);
                });
                tabItem.addEventListener('mouseleave', function() {
                    self.hideTooltip();
                });
                tabItem.addEventListener('mousemove', function(e) {
                    // 更新 tooltip 位置（但不重建内容）
                    var tip = document.getElementById('tabTooltip');
                    if (tip && tip.classList.contains('visible')) {
                        var rect = tabItem.getBoundingClientRect();
                        var tipW = tip.offsetWidth || 280;
                        var left = rect.left + (rect.width - tipW) / 2;
                        if (left < 8) left = 8;
                        if (left + tipW > window.innerWidth - 8) left = window.innerWidth - tipW - 8;
                        tip.style.left = left + 'px';
                        tip.style.top = (rect.bottom + 6) + 'px';
                    }
                });

                tabList.appendChild(tabItem);
            });
            // 更新新建按钮状态
            if (addBtn) {
                addBtn.disabled = this.sessions.length >= MAX_TABS;
            }
        },

        // 更新活跃标签名（从 H1 自动提取，仅当用户未手动编辑时）
        updateActiveTabTitleFromH1: function() {
            var session = this.getActive();
            if (!session) return;
            // 如果用户手动编辑过标签名，不覆盖
            if (session.customTitle) return;
            var h1 = editor.querySelector('h1');
            if (h1) {
                var text = h1.textContent.trim();
                if (text && text.length > 0) {
                    var newTitle = text.substring(0, 30);
                    if (newTitle !== session.title) {
                        session.title = newTitle;
                        this.renderTabs();
                    }
                }
            }
        }
    };

    // 同步全局状态变量到活跃会话（切换前调用）
    function syncStateToSession() {
        var session = tabManager.getActive();
        if (!session) return;
        session.html = editor ? editor.innerHTML : '';
        session.undoHistory = undoHistory;
        session.undoIndex = undoIndex;
        session.imageDataMap = imageDataMap;
        session.imageCounter = imageCounter;
        session.scrollTop = editorContainer ? editorContainer.scrollTop : 0;
        session.currentMatches = currentMatches;
        session.currentMatchIndex = currentMatchIndex;
        session.searchText = searchText;
        session.isImporting = isImporting;
        session.anchors = anchors.slice();
        session.foldPoints = foldPoints.slice();
        session.foldRegions = foldRegions.map(function(r) {
            return { id: r.id, startPid: r.startPid, endPid: r.endPid, isFolded: r.isFolded };
        });
    }

    // 从会话恢复全局状态变量（切换后调用）
    function syncStateFromSession(session) {
        if (!session) return;
        undoHistory = (session.undoHistory || []).slice();
        undoIndex = session.undoIndex != null ? session.undoIndex : -1;
        imageDataMap = session.imageDataMap || new Map();
        imageCounter = session.imageCounter || 0;
        currentMatches = (session.currentMatches || []).slice();
        currentMatchIndex = session.currentMatchIndex != null ? session.currentMatchIndex : -1;
        searchText = session.searchText || '';
        isImporting = session.isImporting || false;
        anchors = (session.anchors || []).slice();
        foldPoints = (session.foldPoints || []).slice();
        foldRegions = (session.foldRegions || []).slice();
    }

    // ===== 会话级状态变量（由 TabManager 管理） =====
    var currentMatches = [];
    var currentMatchIndex = -1;
    var searchText = '';
    var isImporting = false;
    var headingConfig = {};
    var imageDataMap = new Map();
    var imageCounter = 0;
    var anchors = [];  // 锚点列表 [{id, pid, text, createdAt}]
    var foldPoints = [];  // 折叠标记点 [{id, pid, text, createdAt}]
    var foldRegions = []; // 折叠区域 [{id, startPid, endPid, isFolded}]

    // ===== 增强版撤销/重做管理器（捕获所有编辑操作） =====
    var undoHistory = [];
    var undoIndex = -1;
    var MAX_UNDO = 80;
    var undoBlocked = false;
    var lastBeforeInputTime = 0;  // beforeinput 冷却计时
    var BEFORE_INPUT_COOLDOWN = 300; // ms，防止逐键快照过多
    var suppressCaptureUntil = 0; // 程序化操作期间抑制事件捕获（时间戳，ms）
    var justRestored = false;     // 撤销/重做后标记，强制下一次 saveUndoState 跳过去重

    // ---- 光标位置序列化/恢复 ----
    function saveCursorPath() {
        try {
            var sel = window.getSelection();
            if (!sel.rangeCount || !editor.contains(sel.anchorNode)) return null;
            var range = sel.getRangeAt(0);
            // 仅保存折叠光标（不保存选区范围，撤销后选区通常不需要保留）
            var node = range.startContainer;
            var offset = range.startOffset;
            // 构建从 startContainer 向上的路径（子节点索引列表）
            var path = [];
            var cur = node;
            while (cur && cur !== editor) {
                var parent = cur.parentNode;
                if (!parent) { path = null; break; }
                var idx = Array.prototype.indexOf.call(parent.childNodes, cur);
                path.unshift({ idx: idx, tag: parent.tagName || '' });
                cur = parent;
            }
            if (!path) return null;
            return { path: path, offset: offset, startNodeType: node.nodeType };
        } catch(e) { return null; }
    }

    function restoreCursorPath(cursorPath) {
        if (!cursorPath || !cursorPath.path) return;
        try {
            var cur = editor;
            for (var i = 0; i < cursorPath.path.length; i++) {
                var childNodes = cur.childNodes;
                var idx = cursorPath.path[i].idx;
                // 索引可能因为 DOM 变化而失效，做边界检查
                if (idx >= childNodes.length) idx = childNodes.length - 1;
                if (idx < 0) return;
                cur = childNodes[idx];
            }
            // 确保 offset 不越界
            var maxOffset = cur.nodeType === Node.TEXT_NODE ? cur.textContent.length : cur.childNodes.length;
            var off = Math.min(cursorPath.offset, maxOffset);
            var range = document.createRange();
            range.setStart(cur, off);
            range.collapse(true);
            var sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        } catch(e) { /* 光标恢复失败不影响撤销功能 */ }
    }

    // ---- 统一快照保存 ----
    function saveUndoState(description) {
        if (undoBlocked) return;
        // 截断未来分支（新编辑后不能再重做）
        if (undoIndex < undoHistory.length - 1) {
            undoHistory = undoHistory.slice(0, undoIndex + 1);
        }
        var html = editor.innerHTML;
        // 跳过占位符和空内容
        if (!html || html.indexOf('class="placeholder"') >= 0) return;
        // 注意：不进行 HTML 去重检查。
        // 原因：如果去重跳过了保存，但后续程序化 DOM 操作不触发 beforeinput/input，
        // 撤销栈中将永远缺少该操作前的快照，导致撤销跳回错误状态。
        // 通过 suppressCaptureUntil 机制防止"污染快照"已足够。
        justRestored = false;
        var snap = {
            html: html,
            scrollTop: editorContainer.scrollTop,
            cursorPath: saveCursorPath(),
            description: description || '编辑',
            timestamp: Date.now()
        };
        undoHistory.push(snap);
        if (undoHistory.length > MAX_UNDO) undoHistory.shift();
        undoIndex = undoHistory.length - 1;
        // 抑制后续事件捕获 200ms，防止程序化 DOM 操作触发 beforeinput/input 产生污染快照
        suppressCaptureUntil = Date.now() + 200;
        lastBeforeInputTime = Date.now();
    }

    // 兼容旧接口
    function undoPushSnapshot() { saveUndoState('操作'); }

    // ---- 快照恢复 ----
    function restoreSnapshot(snap, directionLabel) {
        undoBlocked = true;
        editor.innerHTML = snap.html;
        // 抑制后续 renumber/generateTOC 引发的 beforeinput/input 事件
        suppressCaptureUntil = Date.now() + 300;
        undoBlocked = false;
        if (snap.scrollTop !== undefined) editorContainer.scrollTop = snap.scrollTop;
        if (snap.cursorPath) restoreCursorPath(snap.cursorPath);
        generateTOC();
        updateStats();
        renumber();
        applyHeadingStylesToEditor();
        setStatus(directionLabel + ': ' + (snap.description || '编辑'));
        var desc = snap.description || '编辑';
        showToast(directionLabel + ': ' + desc, 'info', 1500);
        // 标记刚恢复状态，强制下一次 saveUndoState 跳过去重（即使 HTML 与栈顶相同）
        justRestored = true;
    }

    function undoPerform() {
        if (undoIndex <= 0) { showToast('无法继续撤销', 'warning', 1200); return; }
        undoIndex--;
        restoreSnapshot(undoHistory[undoIndex], '撤销');
    }

    function redoPerform() {
        if (undoIndex >= undoHistory.length - 1) { showToast('无法继续重做', 'warning', 1200); return; }
        undoIndex++;
        restoreSnapshot(undoHistory[undoIndex], '重做');
    }

    // ---- beforeinput: 捕获所有浏览器原生编辑操作 ----
    editor.addEventListener('beforeinput', function(e) {
        if (undoBlocked) return;
        // 拦截浏览器原生撤销/重做，改用自定义实现（始终生效，不受抑制影响）
        if (e.inputType === 'historyUndo') {
            e.preventDefault();
            undoPerform();
            return;
        }
        if (e.inputType === 'historyRedo') {
            e.preventDefault();
            redoPerform();
            return;
        }
        // 程序化操作抑制期：saveUndoState 调用后 200ms 内跳过事件捕获，防止污染快照
        if (Date.now() < suppressCaptureUntil) return;
        // 冷却期内不重复保存（避免逐键快照）
        var now = Date.now();
        if (now - lastBeforeInputTime < BEFORE_INPUT_COOLDOWN) return;
        lastBeforeInputTime = now;
        // 根据 inputType 生成描述
        var desc = '编辑';
        var it = e.inputType || '';
        if (it.indexOf('insertText') >= 0) desc = '输入文本';
        else if (it.indexOf('insertFromPaste') >= 0) desc = '粘贴';
        else if (it.indexOf('insertFromDrop') >= 0) desc = '拖放';
        else if (it.indexOf('delete') >= 0) desc = '删除';
        else if (it.indexOf('formatBold') >= 0) desc = '加粗';
        else if (it.indexOf('formatItalic') >= 0) desc = '斜体';
        else if (it.indexOf('formatUnderline') >= 0) desc = '下划线';
        else if (it.indexOf('formatStrike') >= 0) desc = '删除线';
        else if (it.indexOf('formatFontName') >= 0) desc = '字体';
        else if (it.indexOf('formatFontSize') >= 0) desc = '字号';
        else if (it.indexOf('formatFontColor') >= 0) desc = '文字颜色';
        else if (it.indexOf('formatBackColor') >= 0) desc = '背景色';
        else if (it.indexOf('formatJustify') >= 0 || it.indexOf('formatAlign') >= 0) desc = '对齐';
        else if (it.indexOf('formatIndent') >= 0 || it.indexOf('formatOutdent') >= 0) desc = '缩进';
        else if (it.indexOf('insertOrderedList') >= 0) desc = '有序列表';
        else if (it.indexOf('insertUnorderedList') >= 0) desc = '无序列表';
        else if (it.indexOf('insertLink') >= 0) desc = '链接';
        else if (it.indexOf('insertHorizontalRule') >= 0) desc = '分割线';
        else if (it.indexOf('insertFromComposition') >= 0) desc = '输入文本';
        else desc = '编辑';
        saveUndoState(desc);
    });
    // 确保 beforeinput 后 input 事件不再重复触发快照
    editor.addEventListener('input', function() {
        if (!undoBlocked && Date.now() >= suppressCaptureUntil) {
            // 如果 beforeinput 已经保存过快照（最近 500ms 内），不再重复保存
            var now = Date.now();
            if (now - lastBeforeInputTime > 500) {
                saveUndoState('编辑');
            }
            lastBeforeInputTime = 0; // 重置冷却
            // 标记当前文档为未保存状态，更新标签页显示 *
            var activeSession = tabManager.getActive();
            if (activeSession) {
                activeSession._dirty = true;
                // 轻量更新：仅修改活跃标签的文字，不重建全部 DOM
                var activeTabName = document.querySelector('.tab-item.active .tab-name');
                if (activeTabName) {
                    activeTabName.textContent = (activeSession.title || '文档 1') + ' *';
                }
            }
            // 用户编辑后重置离开标志（取消离开后重新编辑时重新启用拦截）
            resetLeaveFlag();
        }
        handleEditorChange();
        triggerAutoSave();
        // 统计信息防抖更新（避免逐键重算阻塞渲染）
        debouncedUpdateStats();
    });

    function showToast(message, type, duration) {
        type = type || 'info';
        duration = duration || 3000;
        var container = document.querySelector('.toast-container');
        if (!container) {
            container = document.createElement('div');
            container.className = 'toast-container';
            document.body.appendChild(container);
        }
        var toast = document.createElement('div');
        toast.className = 'toast ' + type;
        toast.textContent = message;
        container.appendChild(toast);
        setTimeout(function() {
            toast.classList.add('out');
            setTimeout(function() { toast.remove(); }, 300);
        }, duration);
    }

    var currentTheme = localStorage.getItem('docx-editor-theme') || 'light';

    function setTheme(theme) {
        currentTheme = theme;
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('docx-editor-theme', theme);
        var btn = document.getElementById('themeToggle');
        if (btn) {
            btn.querySelector('.theme-icon').textContent = theme === 'dark' ? '☀️' : '\u{1f319}';
            btn.title = theme === 'dark' ? '切换到亮色主题' : '切换到暗色主题';
        }
        // 代码主题为 auto 时跟随页面
        if (getCodeThemeSetting && getCodeThemeSetting() === 'auto') {
            if (typeof applyAllCodeThemes === 'function') applyAllCodeThemes();
        }
    }

    function toggleTheme() {
        setTheme(currentTheme === 'dark' ? 'light' : 'dark');
    }
    setTheme(currentTheme);

    function showLoading(text) {
        loadingText.textContent = text || '正在处理...';
        loadingOverlay.classList.remove('hidden');
    }

    function hideLoading() {
        loadingOverlay.classList.add('hidden');
    }

    function setStatus(text) {
        docStatus.textContent = text;
    }

    function debounce(fn, delay) {
        delay = delay || 300;
        var timer;
        return function() {
            var ctx = this;
            var args = arguments;
            clearTimeout(timer);
            timer = setTimeout(function() { fn.apply(ctx, args); }, delay);
        };
    }

    function throttle(fn, limit) {
        limit = limit || 100;
        var inThrottle = false;
        return function() {
            if (!inThrottle) {
                fn.apply(this, arguments);
                inThrottle = true;
                setTimeout(function() { inThrottle = false; }, limit);
            }
        };
    }

    var numberingEnabled = true;
    var NUMBER_CLASS = 'heading-number';
    var hasExistingNumbering = false;

    function detectExistingNumbering() {
        var hs = editor.querySelectorAll('h1, h2, h3, h4, h5, h6');
        for (var i = 0; i < hs.length; i++) {
            var t = hs[i].textContent.trim();
            if (/^[\d]+[\.．、]/.test(t) || /^[IVXLCDM]+[\.．、]/.test(t) || /^[A-Z][\.．、]/.test(t) || /^[一二三四五六七八九十百]+[、]/.test(t) || /^（[一二三四五六七八九十百]+）/.test(t)) {
                hasExistingNumbering = true; return true;
            }
        }
        hasExistingNumbering = false; return false;
    }

    function stripTextPrefixes() {
        var hs = editor.querySelectorAll('h1, h2, h3, h4, h5, h6');
        hs.forEach(function(h) {
            var w = document.createTreeWalker(h, NodeFilter.SHOW_TEXT, null, false);
            var ft = null;
            while (w.nextNode()) { if (w.currentNode.textContent.trim()) { ft = w.currentNode; break; } }
            if (!ft) return;
            var m = ft.textContent.match(/^(\s*[\d]+[\.．、](?:[\d]+[\.．、])*\s*|\s*[\d]+[\.．、]\s*|\s*[IVXLCDM]+[\.．、]\s*|\s*[A-Z][\.．、]\s*|\s*[一二三四五六七八九十百]+[、]\s*|\s*（[一二三四五六七八九十百]+）\s*)/);
            if (m) ft.textContent = ft.textContent.substring(m[0].length);
        });
        hasExistingNumbering = false;
    }

    function renumber() {
        editor.querySelectorAll('.' + NUMBER_CLASS).forEach(function(el) { el.remove(); });
        stripTextPrefixes();
        applyNumberingSpans();
    }

    function applyNumberingSpans() {
        var hs = editor.querySelectorAll('h1, h2, h3, h4, h5, h6');
        if (hs.length === 0) return;
        var c = [0,0,0,0,0,0];
        hs.forEach(function(h) {
            var lv = parseInt(h.tagName.substring(1)) - 1;
            for (var i = lv + 1; i < c.length; i++) c[i] = 0;
            c[lv]++;
            var ns = document.createElement('span');
            ns.className = NUMBER_CLASS;
            ns.textContent = getChineseNumberText(c.slice(0, lv + 1), lv + 1);
            h.insertBefore(ns, h.firstChild);
        });
        generateTOC();
    }

    function applyNumbering() { renumber(); }

    // ===== AI 标题助手 =====
    function applyHintFromJSON(jsonStr) {
        var items;
        try { items = JSON.parse(jsonStr); } catch(e) { showToast('JSON 解析失败', 'error'); return; }
        if (!Array.isArray(items) || items.length === 0) { showToast('无效的 hint 格式', 'error'); return; }
        for (var i = 0; i < items.length; i++) {
            if (!items[i].level || !items[i].text || !items[i].anchor) {
                showToast('第 ' + (i + 1) + ' 项缺少字段', 'error'); return;
            }
            if (items[i].level < 1 || items[i].level > 6) { showToast('第 ' + (i + 1) + ' 项 level 无效', 'error'); return; }
            items[i].position = items[i].position || 'before';
        }
        var ops = [];
        for (var i = 0; i < items.length; i++) {
            var target = findAnchorNode(items[i].anchor);
            if (!target) { showToast('未找到锚点 "' + items[i].anchor.substring(0, 15) + '..."', 'warning'); continue; }
            ops.push({ item: items[i], target: target });
        }
        if (ops.length === 0) { showToast('未找到任何锚点', 'error'); return; }
        ops.sort(function(a, b) { return getNodePosition(b.target) - getNodePosition(a.target); });
        saveUndoState('AI 插入标题'); // 记录 AI 插入标题前状态
        for (var i = 0; i < ops.length; i++) {
            var h = document.createElement('H' + ops[i].item.level);
            h.textContent = ops[i].item.text;
            var cfg = headingConfig[ops[i].item.level];
            if (cfg) { h.style.fontFamily = cfg.family; h.style.fontSize = cfg.size; h.style.fontWeight = cfg.bold ? 'bold' : 'normal'; if (cfg.color && cfg.color !== '#000000') h.style.color = cfg.color; }
            if (ops[i].item.position === 'before') {
                ops[i].target.parentNode.insertBefore(h, ops[i].target);
            } else if (ops[i].target.nextSibling) {
                ops[i].target.parentNode.insertBefore(h, ops[i].target.nextSibling);
            } else {
                ops[i].target.parentNode.appendChild(h);
            }
        }
        renumber(); generateTOC();
        showToast('已插入 ' + ops.length + ' 个标题', 'success');
        setStatus('Hint 标题已应用');
    }

    function findAnchorNode(anchor) {
        var w = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null, false);
        var n;
        while (n = w.nextNode()) {
            if (n.textContent.indexOf(anchor) >= 0) {
                var b = n.parentNode;
                while (b && b !== editor && !/^P$|^DIV$|^H[1-6]$|^LI$|^TD$|^TH$/i.test(b.tagName)) b = b.parentNode;
                return b || n.parentNode;
            }
        }
        return null;
    }

    function getNodePosition(node) {
        var p = 0, w = document.createTreeWalker(editor, NodeFilter.SHOW_ALL, null, false), n;
        while (n = w.nextNode()) { if (n === node) return p; p++; }
        return -1;
    }

    function clearHint() {
        document.getElementById('hintInput').value = '';
        errorLineNum = null;
        updateLineNumbers();
        var st = document.getElementById('jsonStatus');
        if (st) st.className = 'json-status hidden';
        lastErrorPos = null;
    }

    function downloadRuleFile() {
        fetch('./format-rule.md').then(function(r) { return r.text(); }).then(function(text) {
            var blob = new Blob([text], { type: 'text/markdown' });
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a'); a.href = url; a.download = 'format-rule.md';
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }).catch(function() { showToast('规则文件下载失败', 'error'); });
    }


    function getChineseNumberText(prefix, level) {
        if (level === 1) return toChineseNum(prefix[0]) + '、';
        if (level === 2) return '（' + toChineseNum(prefix[1]) + '）';
        return prefix[level - 1] + '.';
    }

    function toChineseNum(n) {
        var cn = ['零','一','二','三','四','五','六','七','八','九','十'];
        if (n <= 10) return cn[n];
        if (n < 20) return '十' + (n > 10 ? cn[n - 10] : '');
        var s = '';
        if (n >= 100) { s += cn[Math.floor(n / 100)] + '百'; n = n % 100; }
        if (n >= 10) { s += cn[Math.floor(n / 10)] + '十'; n = n % 10; }
        if (n > 0) s += cn[n];
        return s;
    }

    var CONFIG_FILENAME = 'docx-editor-config.json';
    var configFileHandle = null;

    function collectAllSettings() {
        readHeadingConfig();
        var cd = document.getElementById('codeDetectToggle');
        return {
            version: 3,
            headingConfig: JSON.parse(JSON.stringify(headingConfig)),
            bodyFont: bodyFont.value,
            bodySize: bodySize.value,
            bodyLineHeight: bodyLineHeight.value,
            imgCenter: imgCenterToggle.checked,
            codeDetect: cd ? cd.checked : true,
            codeFoldLines: (function(v){var n=parseInt(v);return isNaN(n)||n<0?8:n;})(document.getElementById('codeFoldLines').value),
            codeTheme: getCodeThemeSetting(),
            theme: currentTheme
        };
    }

    function applySettings(settings) {
        if (!settings) return;
        if (settings.theme) setTheme(settings.theme);
        if (settings.imgCenter !== undefined) {
            imgCenterToggle.checked = settings.imgCenter;
            if (settings.imgCenter) centerAllImages();
            else uncenterAllImages();
        }
        if (settings.bodyFont) bodyFont.value = settings.bodyFont;
        if (settings.bodySize) bodySize.value = settings.bodySize;
        if (settings.bodyLineHeight) bodyLineHeight.value = settings.bodyLineHeight;
        if (settings.headingConfig) {
            Object.keys(settings.headingConfig).forEach(function(level) {
                var cfg = settings.headingConfig[level];
                var row = document.querySelector('.format-row[data-level="' + level + '"]');
                if (!row) return;
                if (cfg.family) row.querySelector('.font-family').value = cfg.family;
                if (cfg.size) row.querySelector('.font-size').value = cfg.size;
                if (cfg.bold !== undefined) row.querySelector('.font-bold').checked = cfg.bold;
                if (cfg.color) row.querySelector('.font-color').value = cfg.color;
            });
        }
        applyHeadingStylesToEditor();
        applyBodyFormatFn();
        if (settings.codeDetect !== undefined) {
            var cd = document.getElementById('codeDetectToggle');
            if (cd) cd.checked = settings.codeDetect;
        }
        if (settings.codeFoldLines) {
            var fl = document.getElementById('codeFoldLines');
            if (fl) fl.value = String(settings.codeFoldLines);
        }
        if (settings.codeTheme) {
            var ct = document.getElementById('codeThemeSelect');
            if (ct) { ct.value = settings.codeTheme; applyAllCodeThemes(); }
        }
        setStatus('配置已加载');
    }

    function saveConfigToStorage() {
        try {
            localStorage.setItem('docx-editor-config', JSON.stringify(collectAllSettings()));
        } catch(e) { console.warn('save failed', e); }
    }

    function loadConfigFromStorage() {
        try {
            var data = localStorage.getItem('docx-editor-config');
            if (data) { applySettings(JSON.parse(data)); return true; }
        } catch(e) { /* ignore */ }
        return false;
    }

    function loadConfigFromProjectDir() {
        return fetch('./' + CONFIG_FILENAME)
            .then(function(r) { if (!r.ok) throw Error('no file'); return r.json(); })
            .then(function(s) {
                if (s && s.version) { applySettings(s); try { localStorage.setItem('docx-editor-config', JSON.stringify(s)); } catch(e) {} return true; }
                return false;
            })
            .catch(function() { return false; });
    }

    function saveConfigToFile() {
        var s = collectAllSettings();
        var blob = new Blob([JSON.stringify(s, null, 2)], { type: 'application/json' });
        if (configFileHandle) { doSaveHandle(configFileHandle, blob); return; }
        if (window.showSaveFilePicker) {
            showSaveFilePicker({ suggestedName: CONFIG_FILENAME, types: [{ description: 'Config', accept: { 'application/json': ['.json'] } }] })
            .then(function(h) { configFileHandle = h; return doSaveHandle(h, blob); })
            .catch(function(err) { if (err.name !== 'AbortError') fallbackDownloadConfig(blob); });
        } else { fallbackDownloadConfig(blob); }
    }

    function doSaveHandle(handle, blob) {
        return handle.createWritable().then(function(w) { return w.write(blob).then(function() { return w.close(); }); })
        .then(function() { showToast('Config saved to project directory', 'success'); })
        .catch(function() { configFileHandle = null; fallbackDownloadConfig(blob); });
    }

    function fallbackDownloadConfig(blob) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url; a.download = CONFIG_FILENAME;
        document.body.appendChild(a); a.click(); document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast('Config downloaded, place in project directory', 'info');
    }

    function loadConfigFromFile(file) {
        if (!file) return;
        var reader = new FileReader();
        reader.onload = function(e) {
            try {
                var s = JSON.parse(e.target.result);
                if (!s || !s.version) { showToast('Invalid config file', 'error'); return; }
                applySettings(s);
                saveConfigToStorage();
                showToast('Config imported', 'success');
            } catch(err) { showToast('Parse failed: ' + err.message, 'error'); }
        };
        reader.readAsText(file);
    }

    function readHeadingConfig() {
        headingConfig = {};
        document.querySelectorAll('.format-row[data-level]').forEach(function(row) {
            var level = row.dataset.level;
            headingConfig[level] = {
                family: row.querySelector('.font-family').value,
                size: row.querySelector('.font-size').value,
                bold: row.querySelector('.font-bold').checked,
                color: row.querySelector('.font-color').value
            };
        });
    }

    function applyHeadingStylesToEditor() {
        readHeadingConfig();
        var el = document.getElementById('editor-heading-styles');
        if (!el) { el = document.createElement('style'); el.id = 'editor-heading-styles'; document.head.appendChild(el); }
        var css = '';
        for (var i = 1; i <= 6; i++) {
            var c = headingConfig[i];
            if (c) css += '.editor h' + i + ' { font-family:"' + c.family + '";font-size:' + c.size + ';font-weight:' + (c.bold ? 'bold' : 'normal') + ';color:' + c.color + '; }\n';
        }
        el.textContent = css;
        setStatus('Styles updated');
    }

    function applyBodyFormatFn() {
        var font = bodyFont.value, size = bodySize.value, lh = bodyLineHeight.value;
        var el = document.getElementById('editor-body-styles');
        if (!el) { el = document.createElement('style'); el.id = 'editor-body-styles'; document.head.appendChild(el); }
        el.textContent = '.editor,.editor p,.editor div:not([class*=heading]):not(h1):not(h2):not(h3):not(h4):not(h5):not(h6){font-family:"' + font + '";font-size:' + size + ';line-height:' + lh + '}.editor table,.editor td,.editor th{font-family:"' + font + '";font-size:' + size + '}';
        setStatus('Body format applied: ' + font + ', ' + size);
        showToast('Body format applied', 'success');
    }

    // ===== 从 DOCX document.xml 提取内联格式并应用到 HTML =====
    // mammoth.js 默认只转换粗体/斜体/删除线等语义标签，
    // 文字颜色/背景色/字号/字体/下划线等需从原始 XML 中提取
    async function enrichFormattingFromDocx(container, arrayBuffer) {
        var LOG = console.log.bind(console, '[格式化导入]');
        try {
            // 1. 打开 DOCX ZIP 读取 word/document.xml
            LOG('步骤1: 打开 ZIP...');
            var zip = new JSZip();
            var z = await zip.loadAsync(arrayBuffer);
            var docXmlFile = z.file('word/document.xml');
            if (!docXmlFile) { LOG('失败: 找不到 word/document.xml'); return; }
            LOG('找到 word/document.xml, 大小约 ' + (docXmlFile._data ? docXmlFile._data.uncompressedSize || '?' : '?') + ' 字节');

            var xmlString = await docXmlFile.async('string');
            LOG('步骤2: XML 字符串长度 = ' + xmlString.length);

            // 2. 解析 XML
            var parser = new DOMParser();
            var xmlDoc = parser.parseFromString(xmlString, 'text/xml');
            if (!xmlDoc || !xmlDoc.documentElement) { LOG('失败: XML 解析返回空'); return; }
            LOG('XML documentElement tagName = ' + xmlDoc.documentElement.tagName);

            // 3. 使用 tagName（含前缀）方式遍历 XML
            var bodies = xmlDoc.getElementsByTagName('w:body');
            LOG('步骤3: getElementsByTagName("w:body") 找到 ' + (bodies ? bodies.length : 0) + ' 个');

            // 尝试备用方式
            if (!bodies || bodies.length === 0) {
                bodies = xmlDoc.getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'body');
                LOG('备用 NS 方式: 找到 ' + (bodies ? bodies.length : 0) + ' 个');
            }
            if (!bodies || bodies.length === 0) { LOG('失败: 无法找到 w:body'); return; }

            // 4. 从每个 w:p 中提取 runs
            var allPs = bodies[0].getElementsByTagName('w:p');
            LOG('步骤4: 找到 ' + allPs.length + ' 个 w:p 段落');

            var xmlParagraphs = [];
            var totalRuns = 0;
            var runsWithFormat = 0;
            for (var pi = 0; pi < allPs.length; pi++) {
                var wp = allPs[pi];
                var runs = [];
                var wrs = wp.getElementsByTagName('w:r');
                totalRuns += wrs.length;
                for (var ri = 0; ri < wrs.length; ri++) {
                    var wr = wrs[ri];
                    var text = '';
                    var tEls = wr.getElementsByTagName('w:t');
                    for (var ti = 0; ti < tEls.length; ti++) {
                        text += tEls[ti].textContent || '';
                    }
                    var brEls = wr.getElementsByTagName('w:br');
                    for (var bi = 0; bi < brEls.length; bi++) { text += '\n'; }

                    if (text.length === 0) continue;

                    var fmt = extractRunFormatFromXml(wr);
                    if (fmt) runsWithFormat++;
                    runs.push({ text: text, fmt: fmt });
                }
                if (runs.length > 0) {
                    xmlParagraphs.push({ runs: runs });
                }
            }
            LOG('总计 ' + totalRuns + ' 个 w:r, 其中 ' + runsWithFormat + ' 个有格式, ' + xmlParagraphs.length + ' 个非空段落');

            if (xmlParagraphs.length === 0) { LOG('失败: 没有解析到任何段落'); return; }

            // 打印前 5 个有格式的 run 作为样本
            var sampleCount = 0;
            for (var xi = 0; xi < xmlParagraphs.length && sampleCount < 5; xi++) {
                for (var xri = 0; xri < xmlParagraphs[xi].runs.length && sampleCount < 5; xri++) {
                    var rf = xmlParagraphs[xi].runs[xri];
                    if (rf.fmt) {
                        LOG('  样本' + (sampleCount+1) + ': text="' + rf.text.substring(0,30) + '" fmt=' + JSON.stringify(rf.fmt));
                        sampleCount++;
                    }
                }
            }
            if (sampleCount === 0) { LOG('没有找到任何带格式的 run，跳过'); return; }

            // 5. 收集 HTML 块级元素
            var htmlBlocks = [];
            var blockWalker = document.createTreeWalker(
                container,
                NodeFilter.SHOW_ELEMENT,
                {
                    acceptNode: function(node) {
                        var t = node.tagName.toLowerCase();
                        if (t === 'p' || (t.length === 2 && t[0] === 'h' && t[1] >= '1' && t[1] <= '6')) {
                            return NodeFilter.FILTER_ACCEPT;
                        }
                        return NodeFilter.FILTER_SKIP;
                    }
                },
                false
            );
            var block;
            while ((block = blockWalker.nextNode())) {
                htmlBlocks.push(block);
            }
            LOG('步骤5: 找到 ' + htmlBlocks.length + ' 个 HTML 块元素 (p, h1-h6)');

            if (htmlBlocks.length === 0) { LOG('失败: HTML 中没有块元素'); return; }

            // 6. 文本内容匹配：对每个有格式的 XML 段落，在 HTML 块中找文本匹配
            var appliedCount = 0;
            var missedCount = 0;
            for (var xi2 = 0; xi2 < xmlParagraphs.length; xi2++) {
                var hasFmt = false;
                for (var ri2 = 0; ri2 < xmlParagraphs[xi2].runs.length; ri2++) {
                    if (xmlParagraphs[xi2].runs[ri2].fmt) { hasFmt = true; break; }
                }
                if (!hasFmt) continue;

                // 计算此 XML 段落的完整文本
                var xmlParaText = '';
                for (var ri3 = 0; ri3 < xmlParagraphs[xi2].runs.length; ri3++) {
                    xmlParaText += xmlParagraphs[xi2].runs[ri3].text;
                }
                if (!xmlParaText.trim()) continue;

                // 在 HTML 块中搜索包含此文本的最佳匹配
                var bestBlock = null;
                var bestScore = 0;
                for (var hi2 = 0; hi2 < htmlBlocks.length; hi2++) {
                    var blockText = htmlBlocks[hi2].textContent;
                    if (blockText.indexOf(xmlParaText) >= 0) {
                        bestBlock = htmlBlocks[hi2];
                        bestScore = xmlParaText.length;
                        break;
                    }
                    var matchLen = 0;
                    var maxCheck = Math.min(xmlParaText.length, blockText.length);
                    while (matchLen < maxCheck && xmlParaText[matchLen] === blockText[matchLen]) {
                        matchLen++;
                    }
                    if (matchLen > bestScore) {
                        bestScore = matchLen;
                        bestBlock = htmlBlocks[hi2];
                    }
                }

                if (bestBlock && bestScore >= 4) {
                    applyFormattingToBlock(bestBlock, xmlParagraphs[xi2].runs);
                    appliedCount++;
                } else if (missedCount < 5) {
                    missedCount++;
                    LOG('未匹配 #' + missedCount +
                        ': XML全文前50字="' + xmlParaText.substring(0, 50) + '"' +
                        ' 最佳得分=' + bestScore +
                        (bestBlock ? ' HTML前50字="' + bestBlock.textContent.substring(0, 50) + '"' : ''));
                }
            }
            LOG('完成: 对 ' + appliedCount + ' 个有格式段落应用了格式 (共扫描 ' + xmlParagraphs.length + ' 个XML段落, ' + htmlBlocks.length + ' 个HTML块)');
        } catch(e) {
            console.warn('[格式化导入] 异常:', e.message || e, e.stack);
        }
    }

    // 从单个 w:r 元素中提取格式信息
    function extractRunFormatFromXml(wr) {
        var rPrs = wr.getElementsByTagName('w:rPr');
        if (!rPrs || rPrs.length === 0) {
            // 尝试无前缀
            rPrs = wr.getElementsByTagName('rPr');
        }
        if (!rPrs || rPrs.length === 0) return null;
        var rPr = rPrs[0];

        var fmt = {};

        // 辅助：获取子元素（同时尝试 w: 前缀和无前缀）
        function getChildEls(tagName) {
            var els = rPr.getElementsByTagName('w:' + tagName);
            if (!els || els.length === 0) els = rPr.getElementsByTagName(tagName);
            return els || [];
        }

        // 辅助：获取属性值（多策略）
        function getAttr(el, attrName) {
            // 1. 带 w: 前缀
            var v = el.getAttribute('w:' + attrName);
            if (v !== null && v !== undefined && v !== '') return v;
            // 2. 无前缀
            v = el.getAttribute(attrName);
            if (v !== null && v !== undefined && v !== '') return v;
            // 3. 命名空间方式
            var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
            if (el.getAttributeNS) {
                v = el.getAttributeNS(W, attrName);
                if (v) return v;
            }
            // 4. 遍历所有属性按 localName 查找
            if (el.attributes) {
                for (var ai = 0; ai < el.attributes.length; ai++) {
                    var a = el.attributes[ai];
                    if (a.localName === attrName || a.name === 'w:' + attrName || a.name === attrName) {
                        return a.value;
                    }
                }
            }
            return '';
        }

        // 文字颜色 — 取最后一个（OOXML 中同属性后者覆盖前者）
        var colorEls = getChildEls('color');
        if (colorEls.length > 0) {
            var c = getAttr(colorEls[colorEls.length - 1], 'val');
            if (c && c !== 'auto' && c !== '000000') fmt.color = '#' + c.toUpperCase();
        }

        // 高亮（背景色）— 取最后一个
        var hlEls = getChildEls('highlight');
        if (hlEls.length > 0) {
            var hlVal = getAttr(hlEls[hlEls.length - 1], 'val');
            if (hlVal && hlVal !== 'none') {
                var hlMap = { yellow:'#FFFF00', green:'#00FF00', cyan:'#00FFFF',
                    magenta:'#FF00FF', blue:'#0000FF', red:'#FF0000', darkBlue:'#00008B',
                    darkCyan:'#008B8B', darkGreen:'#006400', darkMagenta:'#8B008B',
                    darkRed:'#8B0000', darkYellow:'#9B870C', darkGray:'#A9A9A9',
                    lightGray:'#D3D3D3', black:'#000000', white:'#FFFFFF' };
                fmt.backgroundColor = hlMap[hlVal] || '#FFFF00';
            }
        }

        // 底纹填充（背景色）— 取最后一个
        if (!fmt.backgroundColor) {
            var shdEls = getChildEls('shd');
            if (shdEls.length > 0) {
                var fill = getAttr(shdEls[shdEls.length - 1], 'fill');
                if (fill && fill !== 'auto' && fill !== 'FFFFFF') {
                    fmt.backgroundColor = '#' + fill.toUpperCase();
                }
            }
        }

        // 字号 — 取最后一个
        var szEls = getChildEls('sz');
        if (szEls.length > 0) {
            var sz = parseInt(getAttr(szEls[szEls.length - 1], 'val'));
            if (sz > 0) fmt.fontSize = (sz / 2) + 'pt';
        }

        // 字体 — 取最后一个
        var rfEls = getChildEls('rFonts');
        if (rfEls.length > 0) {
            var rf = rfEls[rfEls.length - 1];
            var font = getAttr(rf, 'eastAsia') || getAttr(rf, 'ascii') || getAttr(rf, 'hAnsi') || '';
            if (font) fmt.fontFamily = font;
        }

        // 下划线 — 取最后一个
        var uEls = getChildEls('u');
        if (uEls.length > 0) {
            var uVal = getAttr(uEls[uEls.length - 1], 'val');
            if (!uVal || (uVal !== 'none' && uVal !== 'false')) fmt.textDecoration = 'underline';
        }

        // 粗体 — 取最后一个
        var bEls = getChildEls('b');
        if (bEls.length > 0) {
            var bVal = getAttr(bEls[bEls.length - 1], 'val');
            if (!bVal || (bVal !== 'false' && bVal !== '0')) fmt.fontWeight = 'bold';
        }

        // 斜体 — 取最后一个
        var iEls = getChildEls('i');
        if (iEls.length > 0) {
            var iVal = getAttr(iEls[iEls.length - 1], 'val');
            if (!iVal || (iVal !== 'false' && iVal !== '0')) fmt.fontStyle = 'italic';
        }

        // 删除线 — 取最后一个
        var strikeEls = getChildEls('strike');
        if (strikeEls.length > 0) {
            var sVal = getAttr(strikeEls[strikeEls.length - 1], 'val');
            if (!sVal || (sVal !== 'false' && sVal !== '0')) {
                fmt.textDecoration = fmt.textDecoration ? fmt.textDecoration + ' line-through' : 'line-through';
            }
        }

        // 过滤"正文默认格式"：只有字号和字体且正好是默认宋体正文 → 不视为特殊格式
        var keys = Object.keys(fmt);
        var isOnlyBodyFormat = (keys.length <= 2) && !fmt.color && !fmt.backgroundColor &&
            !fmt.fontWeight && !fmt.fontStyle && !fmt.textDecoration &&
            (fmt.fontSize === '10.5pt' || fmt.fontSize === '12pt' || fmt.fontSize === '14pt') &&
            fmt.fontFamily === '宋体';
        if (isOnlyBodyFormat) return null;

        return Object.keys(fmt).length > 0 ? fmt : null;
    }

    // 将一个 XML 段落的 runs 格式应用到对应的 HTML 块元素
    var _applyDebugCount = 0;
    function applyFormattingToBlock(block, runs) {
        // 收集块内所有文本节点
        var textNodes = [];
        var walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, null, false);
        var node;
        while ((node = walker.nextNode())) {
            if (node.textContent.length > 0) {
                textNodes.push(node);
            }
        }
        if (textNodes.length === 0) return;

        // 构建 HTML 全文和偏移
        var htmlText = '';
        var htmlOffsets = [];
        for (var i = 0; i < textNodes.length; i++) {
            htmlOffsets.push(htmlText.length);
            htmlText += textNodes[i].textContent;
        }

        // 从后往前处理每个有格式的 run
        for (var ri = runs.length - 1; ri >= 0; ri--) {
            var fmt = runs[ri].fmt;
            if (!fmt) continue;
            var runText = runs[ri].text;
            if (!runText) continue;

            // 直接在 HTML 文本中搜索 run 的文本
            var foundAt = htmlText.indexOf(runText);

            // 如果精确匹配失败，尝试去掉首尾空白后匹配
            if (foundAt === -1) {
                var trimmed = runText.replace(/^\s+|\s+$/g, '');
                if (trimmed.length > 0) foundAt = htmlText.indexOf(trimmed);
                if (foundAt >= 0) runText = trimmed;
            }

            if (foundAt >= 0) {
                var applied = wrapRangeInSpan(textNodes, htmlOffsets, foundAt, foundAt + runText.length, fmt);
                if (applied) {
                    // 刷新文本节点和偏移 —— DOM 已被修改，旧引用失效
                    textNodes = [];
                    htmlOffsets = [];
                    htmlText = '';
                    var walker2 = document.createTreeWalker(block, NodeFilter.SHOW_TEXT, null, false);
                    var n2;
                    while ((n2 = walker2.nextNode())) {
                        if (n2.textContent.length > 0) {
                            htmlOffsets.push(htmlText.length);
                            htmlText += n2.textContent;
                            textNodes.push(n2);
                        }
                    }
                }
            } else if (_applyDebugCount < 5) {
                _applyDebugCount++;
                console.log('[格式应用] 匹配失败 #' + _applyDebugCount +
                    ': runText="' + runText.substring(0, 40) + '"' +
                    ' htmlBlock前40字="' + htmlText.substring(0, 40) + '"' +
                    ' tag=' + block.tagName);
            }
        }
    }

    // 在文本节点数组的指定范围内包装 span 元素，返回是否成功
    function wrapRangeInSpan(textNodes, nodeOffsets, rangeStart, rangeEnd, fmt) {
        var anyApplied = false;
        // 从后往前遍历：后面的节点先处理，DOM 修改不影响前面节点的偏移
        for (var i = textNodes.length - 1; i >= 0; i--) {
            var nStart = nodeOffsets[i];
            var nEnd = nStart + textNodes[i].textContent.length;

            // 反向遍历的 continue/break 条件与前向遍历相反
            if (nStart >= rangeEnd) continue; // 节点在范围后面 → 跳过（前面可能还有）
            if (nEnd <= rangeStart) break;     // 节点在范围前面 → 终止（更前面的也在范围外）

            var localStart = Math.max(nStart, rangeStart) - nStart;
            var localEnd = Math.min(nEnd, rangeEnd) - nStart;

            if (localStart >= localEnd) continue;

            var text = textNodes[i].textContent;
            var before = text.substring(0, localStart);
            var match = text.substring(localStart, localEnd);
            var after = text.substring(localEnd);

            if (!match) continue;

            // 创建带格式的 span
            var span = document.createElement('span');
            for (var key in fmt) {
                if (fmt.hasOwnProperty(key)) {
                    span.style[key] = fmt[key];
                }
            }
            span.textContent = match;

            var parent = textNodes[i].parentNode;
            if (!parent) continue;

            // 替换原文本节点
            var frag = document.createDocumentFragment();
            if (before) frag.appendChild(document.createTextNode(before));
            frag.appendChild(span);
            if (after) frag.appendChild(document.createTextNode(after));
            parent.replaceChild(frag, textNodes[i]);
            anyApplied = true;
        }
        return anyApplied;
    }

    async function importDocx(file) {
        if (!file || isImporting) return;
        saveUndoState('导入文档'); // 记录导入前状态（可以撤销回到当前文档）
        isImporting = true;
        showLoading('Reading file...');
        try {
            var buf = await file.arrayBuffer();
            if (typeof mammoth === 'undefined') throw new Error('mammoth.js not loaded');
            var result = null, lastError = null;
            var strategies = [
                function() { return mammoth.convertToHtml({ arrayBuffer: buf }, {
                    convertImage: mammoth.images.dataUri,
                    styleMap: [
                        "p[style-name='Heading 1'] => h1:fresh", "p[style-name='Heading 2'] => h2:fresh",
                        "p[style-name='Heading 3'] => h3:fresh", "p[style-name='Heading 4'] => h4:fresh",
                        "p[style-name='Heading 5'] => h5:fresh", "p[style-name='Heading 6'] => h6:fresh",
                        "p[style-name='heading 1'] => h1:fresh", "p[style-name='heading 2'] => h2:fresh",
                        "p[style-name='heading 3'] => h3:fresh", "p[style-name='heading 4'] => h4:fresh",
                        "p[style-name='heading 5'] => h5:fresh", "p[style-name='heading 6'] => h6:fresh",
                        "p[style-name='标题 1'] => h1:fresh", "p[style-name='标题 2'] => h2:fresh",
                        "p[style-name='标题 3'] => h3:fresh", "p[style-name='标题 4'] => h4:fresh",
                        "p[style-name='标题 5'] => h5:fresh", "p[style-name='标题 6'] => h6:fresh",
                        "p[style-name='heading1'] => h1:fresh", "p[style-name='heading2'] => h2:fresh",
                        "p[style-name='heading3'] => h3:fresh", "p[style-name='heading4'] => h4:fresh",
                        "p[style-name='heading5'] => h5:fresh", "p[style-name='heading6'] => h6:fresh"
                    ]
                }); },
                function() { return mammoth.convertToHtml({ arrayBuffer: buf }, { convertImage: mammoth.images.dataUri }); },
                function() { return mammoth.convertToHtml({ arrayBuffer: buf }); }
            ];
            for (var s = 0; s < strategies.length; s++) {
                try { result = await strategies[s](); if (result && typeof result.value === 'string') break; }
                catch(e) { lastError = e; result = null; }
            }
            if (!result) throw lastError || new Error('All strategies failed');

            showLoading('Rendering...');
            var html = result.value;
            if (typeof html !== 'string') html = String(html || '');

            var tempDiv = document.createElement('div');
            tempDiv.innerHTML = html;

            // Post-process: heading class -> real headings
            tempDiv.querySelectorAll('p[class*="Heading"],p[class*="heading"],p[class*="Titre"]').forEach(function(p) {
                var m = (p.className || '').match(/(?:Heading|heading|Titre)[_\s]*(\d)/i);
                var lv = m ? parseInt(m[1]) : 0;
                if (lv >= 1 && lv <= 6) {
                    var h = document.createElement('H' + lv);
                    h.innerHTML = p.innerHTML;
                    if (p.id) h.id = p.id;
                    p.parentNode.replaceChild(h, p);
                }
            });

            // Bold heuristic detection (改进：优先按中文编号模式判断级别)
            tempDiv.querySelectorAll('p').forEach(function(p) {
                if (/^H[1-6]$/i.test(p.tagName) || (p.closest && p.closest('table')) || p.textContent.trim().length < 2) return;
                if (p.childNodes.length === 1 && p.childNodes[0].nodeType === Node.ELEMENT_NODE && (p.childNodes[0].tagName === 'STRONG' || p.childNodes[0].tagName === 'B')) {
                    var text = p.textContent.trim();
                    var lv = null;

                    // 中文编号模式检测
                    if (/^[一二三四五六七八九十百千]+[、]/.test(text)) lv = 1;           // 一、 → H1
                    else if (/^[（\(][一二三四五六七八九十百千]+[）\)]/.test(text)) lv = 2; // （一）→ H2
                    else if (/^\d+[\.、]/.test(text)) lv = 3;                            // 1. → H3
                    else if (/^[（\(]\d+[）\)]/.test(text)) lv = 4;                       // （1）→ H4

                    // 无编号，按长度推算（中文阈值适当降低）
                    if (lv === null) {
                        var len = text.length;
                        lv = len <= 6 ? 1 : len <= 15 ? 2 : len <= 25 ? 3 : len <= 40 ? 4 : 5;
                    }

                    var h = document.createElement('H' + lv);
                    h.innerHTML = p.innerHTML;
                    if (p.id) h.id = p.id;
                    p.parentNode.replaceChild(h, p);
                }
            });

            // ===== 从原始 DOCX 提取内联格式（颜色、字号、字体、下划线、背景色） =====
            // mammoth.js 默认只转换粗体/斜体，其他内联格式需从 document.xml 提取
            await enrichFormattingFromDocx(tempDiv, buf);

            // 1×1 表格 → 美观代码块
            var codeDetectToggle = document.getElementById('codeDetectToggle');
            if (!codeDetectToggle || codeDetectToggle.checked) {
                tempDiv.querySelectorAll('table').forEach(function(tbl) {
                    var rows = tbl.querySelectorAll('tr');
                    if (rows.length !== 1) return;
                    var firstRowCells = rows[0].querySelectorAll('td, th');
                    if (firstRowCells.length !== 1) return;
                    var cell = firstRowCells[0];

                    var parts = [];
                    (function collectText(node) {
                        if (node.nodeType === Node.TEXT_NODE) {
                            parts.push(node.textContent);
                        } else if (node.nodeType === Node.ELEMENT_NODE) {
                            var tag = node.tagName.toLowerCase();
                            if (tag === 'br') { parts.push('\n'); }
                            else if (tag === 'p' || tag === 'div') {
                                for (var i = 0; i < node.childNodes.length; i++) collectText(node.childNodes[i]);
                                parts.push('\n');
                            } else {
                                for (var i = 0; i < node.childNodes.length; i++) collectText(node.childNodes[i]);
                            }
                        }
                    })(cell);
                    var raw = parts.join('');
                    raw = raw.replace(/^\n+/, '').replace(/\n+$/, '');
                    if (!raw.trim()) return;

                    // 保存原始表格 HTML (base64) 供导出使用
                    var originalHtml = tbl.outerHTML;
                    var encoded = btoa(unescape(encodeURIComponent(originalHtml)));

                    // 折叠配置
                    var foldLines = (function(v){var n=parseInt(v);return isNaN(n)||n<0?8:n;})(document.getElementById('codeFoldLines').value);
                    var lineCount = raw.split('\n').length;
                    var shouldFold = foldLines > 0 && lineCount > foldLines;

                    var codeBlock = buildCodeBlock(raw, encoded, lineCount, foldLines, shouldFold, 'Code');
                    tbl.parentNode.replaceChild(codeBlock, tbl);
                });
            }

            // Extract images
            tempDiv.querySelectorAll('img[src^="data:"]').forEach(function(img) {
                var m = (img.getAttribute('src') || '').match(/^data:([^;]+);base64,(.+)$/);
                if (m) {
                    var id = 'img_' + (++imageCounter);
                    img.setAttribute('data-img-id', id);
                    imageDataMap.set(id, { contentType: m[1], base64: m[2], altText: img.getAttribute('alt') || '' });
                }
            });

            // 清理标题内嵌套的内联字号/字体（导入 DOCX 中 run 级格式会覆盖标题统一格式）
            tempDiv.querySelectorAll('h1, h2, h3, h4, h5, h6').forEach(function(h) {
                h.querySelectorAll('[style]').forEach(function(el) {
                    var style = el.getAttribute('style') || '';
                    var changed = false;
                    // 移除 font-size 和 font-family（标题格式由配置统一控制）
                    if (/font-size\s*:/i.test(style)) {
                        style = style.replace(/font-size\s*:\s*[^;"]+;?\s*/gi, '');
                        changed = true;
                    }
                    if (/font-family\s*:/i.test(style)) {
                        style = style.replace(/font-family\s*:\s*[^;"]+;?\s*/gi, '');
                        changed = true;
                    }
                    if (changed) {
                        if (style.trim()) {
                            el.setAttribute('style', style);
                        } else {
                            el.removeAttribute('style');
                        }
                    }
                });
            });

            var frag = document.createDocumentFragment();
            while (tempDiv.childNodes.length) {
                var child = tempDiv.removeChild(tempDiv.firstChild);
                if (child.nodeType === Node.TEXT_NODE && child.textContent.trim()) {
                    var p = document.createElement('p'); p.appendChild(child); frag.appendChild(p);
                } else { frag.appendChild(child); }
            }

            editor.innerHTML = '';
            editor.appendChild(frag);
            editor.querySelectorAll('.placeholder').forEach(function(el) { el.remove(); });
            saveUndoState('文档导入完成'); // 记录导入完成后的快照

            applyHeadingStylesToEditor();
            applyBodyFormatFn();
            if (imgCenterToggle.checked) centerAllImages();

            var fs = (file.size / 1024 / 1024).toFixed(1);
            setStatus('Loaded: ' + file.name + ' (' + (file.size > 1048576 ? fs + ' MB' : (file.size / 1024).toFixed(1) + ' KB') + ')');
            showToast('Document loaded successfully', 'success');

            // 记录文件来源信息（用于标签 tooltip）
            var session = tabManager.getActive();
            if (session) {
                session.sourceFileName = file.name;
                session.sourceImportTime = new Date().toISOString();
                session._imagesChanged = true; // 导入新文档，图片数据已更新
                // 如果用户未手动编辑过标签名，自动更新为文件名
                if (!session.customTitle) {
                    session.title = file.name.length > 30 ? file.name.substring(0, 27) + '...' : file.name;
                }
                tabManager.renderTabs();
                tabManager.saveAllToDB();
            }

            generateTOC();
            updateStats();
            if (!detectExistingNumbering()) {
                renumber();
            }
            // 清理旧锚点和折叠数据（新文档不含）
            anchors = [];
            foldPoints = [];
            foldRegions = [];
            ensureParagraphIds();
            renderAnchorGutter();
            renderFoldGutter();

            if (result.messages && result.messages.length) {
                var ws = result.messages.filter(function(m) { return m.type === 'warning'; });
                if (ws.length) showToast('Conversion completed with ' + ws.length + ' warnings', 'warning', 4000);
            }
        } catch (err) {
            console.error('Import failed:', err);
            var msg = err.message || String(err) || 'Unknown error';
            showToast('Import failed: ' + msg, 'error');
            setStatus('Import failed: ' + msg.substring(0, 60));
        } finally {
            hideLoading();
            isImporting = false;
        }
    }

    // ===== 代码块函数（供导入和折叠配置使用） =====
    function showCopyFeedback(block) {
        var tip = block.querySelector('.code-block-copied');
        if (tip) { tip.classList.add('show'); setTimeout(function() { tip.classList.remove('show'); }, 1500); }
        showToast('代码已复制', 'success', 1500);
    }
    function fallbackCopy(text, block) {
        var ta = document.createElement('textarea');
        ta.value = text; ta.style.position = 'fixed'; ta.style.left = '-9999px';
        document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); showCopyFeedback(block); } catch(e) { showToast('复制失败，请手动选择复制', 'error'); }
        document.body.removeChild(ta);
    }

    function buildCodeBlock(raw, encoded, lineCount, foldLines, shouldFold, lang) {
        lang = lang || 'Code';
        var LH = 19;
        var codeBlock = document.createElement('div');
        codeBlock.className = 'code-block';
        codeBlock.setAttribute('data-otable', encoded || '');
        codeBlock.setAttribute('data-lines', lineCount);
        codeBlock.setAttribute('data-fold', foldLines);

        // 生成行号
        var lineNumHtml = '';
        for (var i = 1; i <= lineCount; i++) {
            lineNumHtml += '<div>' + i + '</div>';
        }

        var header = document.createElement('div');
        header.className = 'code-block-header';
        header.innerHTML =
            '<div class="code-block-header-left">' +
                '<div class="code-block-dots">' +
                    '<span class="dot-red"></span><span class="dot-yellow"></span><span class="dot-green"></span>' +
                '</div>' +
                '<span class="code-block-lang">' + escHtml(lang) + '</span>' +
            '</div>' +
            '<div class="code-block-header-right"></div>';
        var headerRight = header.querySelector('.code-block-header-right');

        // 代码主题切换按钮（独立于页面主题）
        var themeBtn = document.createElement('button');
        themeBtn.className = 'code-block-btn';
        themeBtn.textContent = '🌓';
        themeBtn.title = '切换代码主题';
        themeBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            cycleCodeTheme(codeBlock);
        });
        headerRight.appendChild(themeBtn);

        var foldBtn = document.createElement('button');
        foldBtn.className = 'code-block-btn';
        foldBtn.textContent = shouldFold ? '📂 展开 (' + lineCount + ' 行)' : '📁 收起 (' + lineCount + ' 行)';
        foldBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            var wrap = codeBlock.querySelector('.code-block-fold');
            wrap.classList.toggle('collapsed');
            var nowCollapsed = wrap.classList.contains('collapsed');
            this.textContent = nowCollapsed ? '📂 展开 (' + lineCount + ' 行)' : '📁 收起 (' + lineCount + ' 行)';
            var bd = codeBlock.querySelector('.code-block-body');
            if (nowCollapsed) {
                var fl = parseInt(codeBlock.getAttribute('data-fold')) || 8;
                wrap.style.maxHeight = (fl * LH + 24) + 'px';
                if (bd) bd.style.maxHeight = '';
            } else {
                wrap.style.maxHeight = '';
                if (bd) bd.style.maxHeight = '60vh';
            }
        });
        headerRight.appendChild(foldBtn);

        var copyBtn = document.createElement('button');
        copyBtn.className = 'code-block-btn';
        copyBtn.textContent = '📋 复制';
        copyBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            if (navigator.clipboard) {
                navigator.clipboard.writeText(raw).then(function() { showCopyFeedback(codeBlock); })
                    .catch(function() { fallbackCopy(raw, codeBlock); });
            } else { fallbackCopy(raw, codeBlock); }
        });
        headerRight.appendChild(copyBtn);
        codeBlock.appendChild(header);

        var bodyWrap = document.createElement('div');
        bodyWrap.className = 'code-block-fold' + (shouldFold ? ' collapsed' : '');
        if (shouldFold) bodyWrap.style.maxHeight = (foldLines * LH + 24) + 'px';

        var body = document.createElement('div');
        body.className = 'code-block-body';
        if (!shouldFold) body.style.maxHeight = '60vh';

        // 行号 + 代码 Flex 布局
        var wrapDiv = document.createElement('div');
        wrapDiv.className = 'code-content-wrap';

        var lineNumDiv = document.createElement('div');
        lineNumDiv.className = 'code-line-nums';
        lineNumDiv.innerHTML = lineNumHtml;
        wrapDiv.appendChild(lineNumDiv);

        var pre = document.createElement('pre');
        var code = document.createElement('code');
        code.textContent = raw;
        pre.appendChild(code);
        wrapDiv.appendChild(pre);

        body.appendChild(wrapDiv);

        if (shouldFold) {
            var fade = document.createElement('div');
            fade.className = 'code-block-fade';
            body.appendChild(fade);
        }

        bodyWrap.appendChild(body);
        codeBlock.appendChild(bodyWrap);

        var copiedTip = document.createElement('div');
        copiedTip.className = 'code-block-copied';
        copiedTip.textContent = '✓ 已复制';
        codeBlock.appendChild(copiedTip);

        // 应用代码主题
        applyCodeTheme(codeBlock);

        return codeBlock;
    }

    function getCodeThemeSetting() {
        var sel = document.getElementById('codeThemeSelect');
        return sel ? sel.value : 'auto';
    }

    function applyCodeTheme(codeBlock) {
        var setting = getCodeThemeSetting();
        codeBlock.classList.remove('code-theme-dark', 'code-theme-light');
        if (setting === 'dark') codeBlock.classList.add('code-theme-dark');
        else if (setting === 'light') codeBlock.classList.add('code-theme-light');
        else {
            var pageTheme = document.documentElement.getAttribute('data-theme') || 'light';
            codeBlock.classList.add(pageTheme === 'dark' ? 'code-theme-dark' : 'code-theme-light');
        }
    }

    function cycleCodeTheme(codeBlock) {
        var current = codeBlock.classList.contains('code-theme-dark') ? 'dark'
            : codeBlock.classList.contains('code-theme-light') ? 'light' : 'auto';
        codeBlock.classList.remove('code-theme-dark', 'code-theme-light');
        if (current === 'dark') codeBlock.classList.add('code-theme-light');
        else if (current === 'light') codeBlock.classList.add('code-theme-dark');
        else codeBlock.classList.add('code-theme-dark');
    }

    function applyAllCodeThemes() {
        editor.querySelectorAll('.code-block').forEach(function(cb) { applyCodeTheme(cb); });
    }

    function updateAllCodeBlocks() {
        var newFoldLines = (function(v){var n=parseInt(v);return isNaN(n)||n<0?8:n;})(document.getElementById('codeFoldLines').value);
        var LH = 19;
        editor.querySelectorAll('.code-block').forEach(function(cb) {
            var lineCount = parseInt(cb.getAttribute('data-lines')) || 0;
            var shouldFold = newFoldLines > 0 && lineCount > newFoldLines;
            var foldEl = cb.querySelector('.code-block-fold');
            var foldBtn = cb.querySelector('.code-block-header-right .code-block-btn:nth-child(2)');
            var fade = cb.querySelector('.code-block-fade');
            var body = cb.querySelector('.code-block-body');
            if (!foldEl) return;

            if (shouldFold) {
                foldEl.classList.add('collapsed');
                foldEl.style.maxHeight = (newFoldLines * LH + 24) + 'px';
                if (body) body.style.maxHeight = '';
                if (foldBtn) foldBtn.textContent = '📂 展开 (' + lineCount + ' 行)';
                if (!fade && body) {
                    var newFade = document.createElement('div');
                    newFade.className = 'code-block-fade';
                    body.appendChild(newFade);
                }
            } else {
                foldEl.classList.remove('collapsed');
                foldEl.style.maxHeight = '';
                if (body) body.style.maxHeight = '60vh';
                if (foldBtn) foldBtn.textContent = '📁 收起 (' + lineCount + ' 行)';
                if (fade) fade.remove();
            }
            cb.setAttribute('data-fold', newFoldLines);
        });
    }

    function centerAllImages() {
        saveUndoState('图片居中');
        editor.querySelectorAll('img:not([data-not-center])').forEach(function(img) {
            img.classList.add('img-center');
            img.style.display = 'block';
            img.style.marginLeft = 'auto';
            img.style.marginRight = 'auto';
        });
        setStatus('Images centered');
    }

    function uncenterAllImages() {
        saveUndoState('取消图片居中');
        editor.querySelectorAll('img.img-center').forEach(function(img) {
            img.classList.remove('img-center');
            img.style.display = '';
            img.style.marginLeft = '';
            img.style.marginRight = '';
        });
        setStatus('Images uncentered');
    }

    var collapsedHeadings = new Set();

    function generateTOC() {
        var headings = editor.querySelectorAll('h1, h2, h3, h4, h5, h6');
        if (!headings.length) {
            tocContainer.innerHTML = '<div class="toc-empty">No headings found<br>Use Ctrl+1~6 to add headings</div>';
            return;
        }
        headings.forEach(function(h, i) { if (!h.id) h.id = 'heading-' + i + '-' + Date.now(); });
        var tree = buildTocTree(headings);
        var list = document.createElement('ul');
        list.className = 'toc-list';
        renderTocTree(tree, list);
        tocContainer.innerHTML = '';
        tocContainer.appendChild(list);
        highlightVisibleHeading();
        if (collapsedHeadings.size) {
            tocContainer.querySelectorAll('.toc-item').forEach(function(li) {
                var link = li.querySelector('.toc-link');
                if (link && collapsedHeadings.has(link.dataset.target)) {
                    var cl = li.querySelector('.toc-children');
                    var t = li.querySelector('.toc-toggle');
                    if (cl && t) { cl.classList.add('collapsed'); t.textContent = '▶'; }
                }
            });
        }
    }

    function buildTocTree(headings) {
        var items = [];
        headings.forEach(function(h) {
            items.push({ level: parseInt(h.tagName.substring(1)), text: h.textContent.trim() || '(empty)', id: h.id, element: h, children: [] });
        });
        var root = { level: 0, children: [] }, stack = [root];
        items.forEach(function(item) {
            while (stack.length && stack[stack.length - 1].level >= item.level) stack.pop();
            var p = stack[stack.length - 1];
            p.children = p.children || [];
            p.children.push(item);
            stack.push(item);
        });
        return root.children;
    }

    function escHtml(str) {
        return typeof str === 'string' ? str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : '';
    }

    function renderTocTree(items, parentEl) {
        items.forEach(function(item) {
            var li = document.createElement('li');
            li.className = 'toc-item level-' + item.level;
            var row = document.createElement('div');
            row.className = 'toc-row';
            var link = document.createElement('a');
            link.className = 'toc-link';
            link.href = '#' + item.id;
            link.innerHTML = '<span class="toc-level-badge">H' + item.level + '</span> ' + escHtml(item.text);
            link.dataset.target = item.id;
            link.addEventListener('click', function(e) {
                e.preventDefault();
                var t = document.getElementById(this.dataset.target);
                if (t) {
                    scrollEditorTo(t, 'start');
                    t.style.background = '#fef08a';
                    setTimeout(function() { t.style.background = ''; }, 1500);
                    document.querySelectorAll('.toc-link.active').forEach(function(a) { a.classList.remove('active'); });
                    this.classList.add('active');
                }
            });
            if (item.children && item.children.length) {
                var toggle = document.createElement('span');
                toggle.className = 'toc-toggle';
                toggle.textContent = '▼';
                toggle.addEventListener('click', function(e) {
                    e.stopPropagation();
                    var cl = this.closest('.toc-item').querySelector('.toc-children');
                    var lk = this.closest('.toc-item').querySelector('.toc-link');
                    if (cl) {
                        var c = cl.classList.toggle('collapsed');
                        this.textContent = c ? '▶' : '▼';
                        var id = lk ? lk.dataset.target : null;
                        if (c && id) collapsedHeadings.add(id);
                        else if (id) collapsedHeadings.delete(id);
                    }
                });
                row.appendChild(toggle);
            }
            row.appendChild(link);
            li.appendChild(row);
            if (item.children && item.children.length) {
                var ul = document.createElement('ul');
                ul.className = 'toc-children';
                renderTocTree(item.children, ul);
                li.appendChild(ul);
            }
            parentEl.appendChild(li);
        });
    }

    function highlightVisibleHeading() {
        var headings = editor.querySelectorAll('h1, h2, h3, h4, h5, h6');
        if (!headings.length) return;
        var vc = editorContainer.scrollTop + editorContainer.clientHeight / 3;
        var closest = null, best = Infinity;
        headings.forEach(function(h) {
            var d = Math.abs(h.offsetTop - vc);
            if (d < best) { best = d; closest = h; }
        });
        document.querySelectorAll('.toc-link.active').forEach(function(a) { a.classList.remove('active'); });
        if (closest && closest.id) {
            var l = tocContainer.querySelector('.toc-link[data-target="' + closest.id + '"]');
            if (l) l.classList.add('active');
        }
    }

    function toggleHeading(level) {
        saveUndoState('设置标题');
        var sel = window.getSelection();
        if (!sel.rangeCount) return;
        var node = sel.getRangeAt(0).commonAncestorContainer;
        while (node && node !== editor && !/^H[1-6]$|^P$|^DIV$|^LI$|^TD$/.test(node.tagName || '')) node = node.parentNode;
        if (!node || node === editor) return;
        if (node.tagName === 'H' + level) {
            var p = document.createElement('p');
            p.innerHTML = node.innerHTML;
            if (node.id) p.id = node.id;
            node.parentNode.replaceChild(p, node);
            sel.removeAllRanges();
            var r = document.createRange(); r.selectNodeContents(p); sel.addRange(r);
            renumber(); generateTOC(); setStatus('Unset heading ' + level); return;
        }
        var h = document.createElement('H' + level);
        h.innerHTML = (/^H[1-6]$|^P$|^DIV$/i.test(node.tagName)) ? node.innerHTML : node.textContent;
        if (node.id) h.id = node.id;
        node.parentNode.replaceChild(h, node);
        var cfg = headingConfig[level];
        if (cfg) {
            h.style.fontFamily = cfg.family;
            h.style.fontSize = cfg.size;
            h.style.fontWeight = cfg.bold ? 'bold' : 'normal';
            if (cfg.color && cfg.color !== '#000000') h.style.color = cfg.color;
        }
        sel.removeAllRanges();
        try { var r = document.createRange(); r.setStart(h, 0); r.collapse(true); sel.addRange(r); } catch(e) {}
        renumber(); generateTOC(); setStatus('Set to H' + level);
    }

    function clearHeading() {
        saveUndoState('清除标题');
        var sel = window.getSelection();
        if (!sel.rangeCount) return;
        var node = sel.getRangeAt(0).commonAncestorContainer;
        while (node && node !== editor && !/^H[1-6]$|^P$|^DIV$/.test(node.tagName || '')) node = node.parentNode;
        if (!node || node === editor || !/^H[1-6]$/i.test(node.tagName)) return;
        var p = document.createElement('p');
        p.innerHTML = node.innerHTML;
        if (node.id) p.id = node.id;
        node.parentNode.replaceChild(p, node);
        sel.removeAllRanges();
        var r = document.createRange(); r.selectNodeContents(p); sel.addRange(r);
        renumber(); generateTOC(); setStatus('Heading cleared');
    }

    function getContextAround(node, start, end, maxChars) {
        maxChars = maxChars || 40;
        var text = node.textContent;
        var ctxStart = Math.max(0, start - maxChars);
        var ctxEnd = Math.min(text.length, end + maxChars);
        var prefix = ctxStart > 0 ? '…' : '';
        var suffix = ctxEnd < text.length ? '…' : '';
        return {
            before: prefix + text.substring(ctxStart, start),
            match: text.substring(start, end),
            after: text.substring(end, ctxEnd) + suffix
        };
    }

    function performSearch() {
        var text = searchInput.value.trim();
        if (!text) {
            clearSearchHighlights();
            document.getElementById('searchMatchInfo').textContent = '0 个匹配';
            currentMatches = []; currentMatchIndex = -1;
            renderSearchResults();
            return;
        }

        clearSearchHighlights();
        currentMatches = [];
        var walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null, false), node;
        var lowerText = text.toLowerCase();

        while (node = walker.nextNode()) {
            var content = node.textContent.toLowerCase();
            var idx = content.indexOf(lowerText);
            while (idx !== -1) {
                var ctx = getContextAround(node, idx, idx + text.length);
                currentMatches.push({
                    node: node,
                    start: idx,
                    end: idx + text.length,
                    text: text,
                    context: ctx
                });
                idx = content.indexOf(lowerText, idx + 1);
            }
        }

        highlightAllMatches();
        currentMatchIndex = currentMatches.length > 0 ? 0 : -1;
        updateMatchHighlight();
        updateSearchUI();
        renderSearchResults();
        if (!currentMatches.length) showToast('未找到匹配', 'info', 1500);
    }

    function highlightAllMatches() {
        var groups = new Map();
        currentMatches.forEach(function(m, i) {
            if (!groups.has(m.node)) groups.set(m.node, []);
            groups.get(m.node).push({ start: m.start, end: m.end, idx: i });
        });
        groups.forEach(function(matches, node) {
            var text = node.textContent, parent = node.parentNode, frag = document.createDocumentFragment();
            var sorted = matches.slice().sort(function(a,b){return a.start-b.start;}), last = 0;
            sorted.forEach(function(m) {
                if (m.start > last) frag.appendChild(document.createTextNode(text.substring(last, m.start)));
                var span = document.createElement('span'); span.className = 'search-highlight'; span.dataset.matchIdx = m.idx; span.textContent = text.substring(m.start, m.end);
                frag.appendChild(span); last = m.end;
            });
            if (last < text.length) frag.appendChild(document.createTextNode(text.substring(last)));
            parent.replaceChild(frag, node);
        });
        editor.querySelectorAll('.search-highlight').forEach(function(m) {
            var idx = parseInt(m.dataset.matchIdx);
            if (currentMatches[idx]) {
                currentMatches[idx].element = m;
            }
        });
    }

    function clearSearchHighlights() {
        editor.querySelectorAll('.search-highlight').forEach(function(el) {
            el.parentNode.replaceChild(document.createTextNode(el.textContent), el);
            el.parentNode.normalize();
        });
        currentMatches = []; currentMatchIndex = -1;
    }

    function updateMatchHighlight() {
        editor.querySelectorAll('.search-highlight.active').forEach(function(el) { el.classList.remove('active'); });
        if (currentMatchIndex >= 0 && currentMatchIndex < currentMatches.length && currentMatches[currentMatchIndex].element) {
            currentMatches[currentMatchIndex].element.classList.add('active');
            scrollEditorTo(currentMatches[currentMatchIndex].element, 'center');
        }
    }

    function updateSearchUI() {
        var info = document.getElementById('searchMatchInfo');
        var total = currentMatches.length;
        info.textContent = total + ' 个匹配' + (currentMatchIndex >= 0 ? ' (第 ' + (currentMatchIndex + 1) + ' 个)' : '');
        document.getElementById('prevMatch').disabled = total === 0;
        document.getElementById('nextMatch').disabled = total === 0;
        document.getElementById('replaceBtn').disabled = total === 0;
    }

    function goToNextMatch() {
        if (currentMatches.length) {
            currentMatchIndex = (currentMatchIndex + 1) % currentMatches.length;
            updateMatchHighlight();
            updateSearchUI();
            highlightResultItem(currentMatchIndex);
        }
    }
    function goToPrevMatch() {
        if (currentMatches.length) {
            currentMatchIndex = (currentMatchIndex - 1 + currentMatches.length) % currentMatches.length;
            updateMatchHighlight();
            updateSearchUI();
            highlightResultItem(currentMatchIndex);
        }
    }

    function renderSearchResults() {
        var container = document.getElementById('searchResultList');
        if (!container) return;
        if (!currentMatches.length) {
            var emptyText = searchInput && searchInput.value.trim() ? '未找到匹配' : '输入关键词后点击"查找"';
            container.innerHTML = '<div class="search-result-empty">' + emptyText + '</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < currentMatches.length; i++) {
            var m = currentMatches[i];
            var ctx = m.context || { before: '', match: m.text, after: '' };
            var active = i === currentMatchIndex ? ' active' : '';
            var matchText = escHtml(ctx.match);
            html += '<div class="search-result-item' + active + '" data-idx="' + i + '">';
            html += '<div class="search-result-context">';
            html += '<span class="search-result-index">' + (i + 1) + '</span>';
            html += escHtml(ctx.before);
            html += '<span class="match-highlight' + (active ? ' active' : '') + '">' + matchText + '</span>';
            html += escHtml(ctx.after);
            html += '</div></div>';
        }
        container.innerHTML = html;

        // 点击跳转
        container.querySelectorAll('.search-result-item').forEach(function(item) {
            item.addEventListener('click', function() {
                var idx = parseInt(this.dataset.idx);
                if (idx >= 0 && idx < currentMatches.length) {
                    currentMatchIndex = idx;
                    updateMatchHighlight();
                    updateSearchUI();
                    highlightResultItem(idx);
                }
            });
        });
    }

    function highlightResultItem(idx) {
        var container = document.getElementById('searchResultList');
        if (!container) return;
        container.querySelectorAll('.search-result-item.active').forEach(function(el) { el.classList.remove('active'); });
        var items = container.querySelectorAll('.search-result-item');
        if (items[idx]) {
            items[idx].classList.add('active');
            items[idx].scrollIntoView({ block: 'nearest' });
        }
        // 同步列表中的 match-highlight 高亮
        container.querySelectorAll('.match-highlight.active').forEach(function(el) { el.classList.remove('active'); });
        var activeMatch = items[idx] && items[idx].querySelector('.match-highlight');
        if (activeMatch) activeMatch.classList.add('active');
    }

    function replaceCurrent() {
        saveUndoState('替换文本');
        var rt = replaceInput.value;
        if (!currentMatches.length || currentMatchIndex < 0 || !currentMatches[currentMatchIndex].element) return;
        var m = currentMatches[currentMatchIndex];
        m.element.textContent = rt;
        m.element.classList.remove('active');
        var tn = document.createTextNode(rt);
        m.element.parentNode.replaceChild(tn, m.element);
        tn.parentNode.normalize();
        performSearch();
    }

    function replaceAll() {
        saveUndoState('全部替换');
        var st = searchInput.value.trim(), rt = replaceInput.value;
        if (!st) return;
        clearSearchHighlights();
        var walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT, null, false), node, nodes = [];
        var lowerSt = st.toLowerCase();
        while (node = walker.nextNode()) {
            if (node.textContent.toLowerCase().indexOf(lowerSt) >= 0) nodes.push(node);
        }
        var count = 0;
        nodes.forEach(function(node) {
            var txt = node.textContent;
            var low = txt.toLowerCase(), res = '', last = 0, idx = low.indexOf(lowerSt);
            while (idx !== -1) { res += txt.substring(last, idx) + rt; last = idx + st.length; idx = low.indexOf(lowerSt, last); }
            res += txt.substring(last);
            if (res !== txt) { node.textContent = res; count++; }
        });
        showToast('已替换 ' + count + ' 处', 'success');
        setStatus('Replaced ' + count + ' occurrences');
        performSearch();
    }

    function handleKeyboard(e) {
        var isCtrl = e.ctrlKey || e.metaKey;
        if (isCtrl) {
            // 拦截导航快捷键：Ctrl+R / Ctrl+Shift+R（刷新）
            if ((e.key === 'r' || e.key === 'R') && !e.altKey) {
                e.preventDefault();
                showLeaveConfirmModal(e.shiftKey ? 'hardReload' : 'reload');
                return;
            }
            // 拦截关闭标签页快捷键：Ctrl+W
            if (e.key === 'w' || e.key === 'W') {
                e.preventDefault();
                showLeaveConfirmModal('close');
                return;
            }
            // 拦截 Ctrl+F4（Windows 关闭标签页）
            if (e.key === 'F4') {
                e.preventDefault();
                showLeaveConfirmModal('close');
                return;
            }
            // 撤销/重做（Ctrl+Z、Ctrl+Y、Ctrl+Shift+Z）
            if (e.key === 'z' || e.key === 'Z') {
                e.preventDefault();
                if (e.shiftKey) { redoPerform(); } else { undoPerform(); }
                return;
            }
            if (e.key === 'y' || e.key === 'Y') {
                e.preventDefault();
                redoPerform();
                return;
            }
            switch (e.key) {
                case 'o': case 'O': e.preventDefault(); fileInput.click(); return;
                case 's': case 'S': e.preventDefault(); saveDocumentFull().then(function() { showToast('✅ 已完整保存到本地存储', 'success', 2000); }); return;
                case 'b': case 'B': e.preventDefault(); saveUndoState('加粗'); document.execCommand('bold'); return;
                case 'i': case 'I': e.preventDefault(); saveUndoState('斜体'); document.execCommand('italic'); return;
                case 'u': case 'U': e.preventDefault(); saveUndoState('下划线'); document.execCommand('underline'); return;
                case 'f': case 'F': e.preventDefault(); searchInput.focus(); searchInput.select(); return;
                case 'h': case 'H': e.preventDefault(); replaceInput.focus(); replaceInput.select(); return;
                case '0': e.preventDefault(); clearHeading(); return;
            }
            if (e.key >= '1' && e.key <= '6') { e.preventDefault(); toggleHeading(parseInt(e.key)); return; }
        }
        // 拦截 F5 刷新键（无论是否按下 Ctrl）
        if (e.key === 'F5') {
            e.preventDefault();
            showLeaveConfirmModal('reload');
            return;
        }
        if (e.key === 'F4' && e.altKey) {
            // Alt+F4 关闭窗口 — 无法完全拦截，但尝试弹出确认
            e.preventDefault();
            showLeaveConfirmModal('close');
            return;
        }
        if (e.key === 'F3') { e.preventDefault(); if (e.shiftKey) goToPrevMatch(); else goToNextMatch(); }
        if (e.key === 'F11') { e.preventDefault(); var b = document.getElementById('fullscreenToggle'); if (b) b.click(); }
        // 锚点快捷键
        if (e.key === 'F9') {
            e.preventDefault();
            toggleAnchorAtCursor();
        }
        if (isCtrl && e.shiftKey && (e.key === 'ArrowUp' || e.key === 'Up')) {
            e.preventDefault();
            navigateAnchor(-1);
        }
        if (isCtrl && e.shiftKey && (e.key === 'ArrowDown' || e.key === 'Down')) {
            e.preventDefault();
            navigateAnchor(1);
        }
    }

    var handleEditorChange = debounce(function() {
        generateTOC();
        renumber();
        ensureParagraphIds();
        renderAnchorGutter();
        renderFoldGutter();
        setStatus('Updated');
        // 自动更新标签标题（从 H1 提取，仅当用户未手动编辑时）
        if (tabManager && typeof tabManager.updateActiveTabTitleFromH1 === 'function') {
            tabManager.updateActiveTabTitleFromH1();
        }
    }, 500);

    // ============================================================
    // 📊 文档统计信息更新
    // ============================================================
    function updateStats() {
        var elWord = document.getElementById('statWordCount');
        var elLine = document.getElementById('statLineCount');
        var elImage = document.getElementById('statImageCount');
        var elCode = document.getElementById('statCodeCount');
        var elTable = document.getElementById('statTableCount');
        var elSize = document.getElementById('statFileSize');
        var elFreq = document.getElementById('statFreqWords');
        // 如果统计栏 DOM 不存在（旧版 HTML），静默跳过
        if (!elWord && !elLine) return;

        var html = editor.innerHTML;
        var text = editor.textContent || '';
        // 去除首尾空白
        text = text.trim();

        // ---- 字数统计（中文字符 + 英文单词） ----
        var chineseChars = 0;
        var englishWords = 0;
        if (text.length > 0) {
            // 提取中文字符（含中文标点）
            var cjkMatch = text.match(/[一-鿿㐀-䶿　-〿＀-￯]/g);
            chineseChars = cjkMatch ? cjkMatch.length : 0;
            // 提取英文单词（去除中文后，按空白分词）
            var nonCjk = text.replace(/[一-鿿㐀-䶿　-〿＀-￯]/g, ' ');
            var words = nonCjk.match(/[a-zA-Z0-9]+/g);
            englishWords = words ? words.length : 0;
        }
        var totalWords = chineseChars + englishWords;

        // ---- 行数统计（近似值：按块级元素 + <br> 计算） ----
        var lineCount = 0;
        // 创建临时容器解析 HTML
        var tempDiv = document.createElement('div');
        tempDiv.innerHTML = html;
        // 统计块级元素数量（每个块级元素至少占一行）
        var blocks = tempDiv.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, li, pre, blockquote, hr, table, tr');
        lineCount += blocks.length;
        // 统计 <br> 标签
        var brs = tempDiv.querySelectorAll('br');
        lineCount += brs.length;
        // 如果没有任何块级元素但有文本内容，至少算 1 行
        if (lineCount === 0 && text.length > 0) {
            lineCount = text.split(/\n/).length;
        }
        if (lineCount === 0) lineCount = 0;

        // ---- 图片数 ----
        var images = editor.querySelectorAll('img');
        var imageCount = images.length;

        // ---- 代码块数 ----
        var codeBlocks = editor.querySelectorAll('.code-block, pre[class*="language-"], pre');
        var codeCount = codeBlocks.length;

        // ---- 表格数 ----
        var tables = editor.querySelectorAll('table');
        var tableCount = tables.length;

        // ---- 预计 DOCX 文件大小（基于 HTML 长度估算） ----
        var htmlBytes = new Blob([html]).size;
        // DOCX 是 ZIP 压缩的 XML，大致是 HTML 大小的 40%~80%，取 60%
        var estimatedDocxBytes = Math.round(htmlBytes * 0.6);
        var sizeStr = '';
        if (estimatedDocxBytes < 1024) {
            sizeStr = estimatedDocxBytes + ' B';
        } else if (estimatedDocxBytes < 1024 * 1024) {
            sizeStr = (estimatedDocxBytes / 1024).toFixed(1) + ' KB';
        } else {
            sizeStr = (estimatedDocxBytes / (1024 * 1024)).toFixed(2) + ' MB';
        }

        // ---- 前5高频词（中文2-gram + 英文单词） ----
        var topWordsStr = '';
        if (text.length > 0) {
            var freqMap = {};
            // 英文停用词过滤
            var stopWords = {the:1,a:1,an:1,is:1,are:1,was:1,were:1,be:1,been:1,being:1,
                have:1,has:1,had:1,do:1,does:1,did:1,will:1,would:1,could:1,should:1,
                may:1,might:1,can:1,shall:1,to:1,of:1,in:1,for:1,on:1,with:1,at:1,by:1,
                from:1,as:1,and:1,or:1,but:1,not:1,so:1,if:1,than:1,it:1,its:1,
                this:1,that:1,these:1,those:1,no:1,all:1,some:1,any:1,each:1};
            // 提取英文单词（≥2字母，过滤停用词）
            var enWords = text.match(/[a-zA-Z]{2,}/g);
            if (enWords) {
                enWords.forEach(function(w) {
                    w = w.toLowerCase();
                    if (!stopWords[w]) freqMap[w] = (freqMap[w] || 0) + 1;
                });
            }
            // 提取中文2-gram（两字词组）
            var cn = text.replace(/[^一-鿿]/g, '');
            for (var i = 0; i < cn.length - 1; i++) {
                var bg = cn.substring(i, i + 2);
                freqMap[bg] = (freqMap[bg] || 0) + 1;
            }
            // 排序取前5
            var sorted = [];
            for (var k in freqMap) {
                if (freqMap.hasOwnProperty(k)) sorted.push({w:k, c:freqMap[k]});
            }
            sorted.sort(function(a, b) { return b.c - a.c; });
            // 去重：跳过已被更长词组覆盖的子串
            var result = [];
            var skip = {};
            for (var m = 0; m < sorted.length && result.length < 5; m++) {
                if (skip[sorted[m].w]) continue;
                result.push(sorted[m]);
                // 标记子串（3字词标记其内部2字词）
                if (sorted[m].w.length >= 3) {
                    for (var si = 0; si <= sorted[m].w.length - 2; si++) {
                        skip[sorted[m].w.substring(si, si + 2)] = true;
                    }
                }
            }
            topWordsStr = result.map(function(r) { return r.w + '(' + r.c + ')'; }).join(' ');
        }

        // ---- 更新 DOM ----
        if (elWord) elWord.textContent = totalWords;
        if (elLine) elLine.textContent = lineCount;
        if (elImage) elImage.textContent = imageCount;
        if (elCode) elCode.textContent = codeCount;
        if (elTable) elTable.textContent = tableCount;
        if (elSize) elSize.textContent = sizeStr;
        if (elFreq) {
            elFreq.textContent = topWordsStr || '-';
            elFreq.title = topWordsStr || '暂无高频词数据';
        }
    }

    // 防抖版统计更新（输入时使用，避免逐键重算阻塞渲染）
    var debouncedUpdateStats = debounce(updateStats, 400);

    importBtn.addEventListener('click', function() { fileInput.click(); });
    fileInput.addEventListener('change', function(e) { if (this.files && this.files[0]) importDocx(this.files[0]); this.value = ''; });

    exportBtn.addEventListener('click', async function() {
        // 立即显示加载提示，确保 UI 快速响应
        showLoading('正在准备导出...');
        // 清除防抖定时器并立即保存
        if (autoSaveTimer) { clearTimeout(autoSaveTimer); autoSaveTimer = null; }
        try {
            await saveDocumentState();
            var fn = window.exportDocx || window.exportDocumentAsDocx;
            if (typeof fn === 'function') {
                showLoading('正在生成 DOCX...');
                // 用 setTimeout 让 loading 渲染后再执行导出
                await new Promise(function(resolve) { setTimeout(resolve, 50); });
                await fn();
                showToast('导出成功', 'success');
            } else {
                showToast('导出模块未就绪', 'error');
            }
        } catch(err) {
            showToast('导出失败: ' + (err.message || '未知错误'), 'error');
        } finally {
            hideLoading();
        }
    });

    printBtn.addEventListener('click', function() {
        // 克隆编辑器内容，处理折叠和代码块
        var clone = editor.cloneNode(true);

        // === 移除折叠占位符 ===
        clone.querySelectorAll('.fold-placeholder').forEach(function(el) { el.remove(); });

        // === 展开所有被折叠隐藏的元素 ===
        clone.querySelectorAll('.fold-hidden').forEach(function(el) {
            el.classList.remove('fold-hidden');
        });

        // === 移除锚点装订线圆点（如果编辑器内有的话） ===
        clone.querySelectorAll('.anchor-gutter-dot').forEach(function(el) { el.remove(); });

        // === 将代码块转换为 1×1 表格（与导出格式一致） ===
        clone.querySelectorAll('.code-block').forEach(function(cb) {
            var tableHtml = '';

            // 优先使用原始表格（data-otable）
            var encoded = cb.getAttribute('data-otable');
            if (encoded) {
                try {
                    var decoded = decodeURIComponent(escape(atob(encoded)));
                    var tempDiv = document.createElement('div');
                    tempDiv.innerHTML = decoded;
                    var origTable = tempDiv.querySelector('table');
                    if (origTable) {
                        tableHtml = origTable.outerHTML;
                    }
                } catch(e) {
                    console.warn('打印：解码代码块表格失败', e);
                }
            }

            // 没有原始表格：从代码内容构建 1×1 表格
            if (!tableHtml) {
                var codeEl = cb.querySelector('code') || cb.querySelector('pre');
                var rawText = codeEl ? codeEl.textContent : cb.textContent;
                // 清理首尾换行 + 去除代码块 header/footer 的杂讯
                rawText = rawText.replace(/^\n+/, '').replace(/\n+$/, '');
                // 构建表格：每行一个 <tr><td>，保留缩进
                var lines = rawText.split('\n');
                var tbody = '';
                for (var li = 0; li < lines.length; li++) {
                    var lineText = lines[li].replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                    // 用 &nbsp; 保留前导空格
                    lineText = lineText.replace(/^ /, '&nbsp;').replace(/  /g, ' &nbsp;');
                    tbody += '<tr><td style="font-family:Consolas,monospace;font-size:10pt;padding:2px 8px;border:none;background:#f5f5f5;">' + (lineText || '&nbsp;') + '</td></tr>';
                }
                tableHtml = '<table style="border-collapse:collapse;width:100%;"><tbody>' + tbody + '</tbody></table>';
            }

            // 用表格替换代码块 div
            var tableWrapper = document.createElement('div');
            tableWrapper.innerHTML = tableHtml;
            var newTable = tableWrapper.firstChild;
            cb.parentNode.replaceChild(newTable, cb);
        });

        // === 生成打印样式 ===
        var printStyle = '<style>' +
            'body { font-family: "SimSun", "宋体", serif; padding: 40px; color: #333; line-height: 1.8; }' +
            'table { border-collapse: collapse; width: 100%; margin: 8px 0; }' +
            'td, th { border: 1px solid #333; padding: 6px 8px; }' +
            'img { max-width: 100%; height: auto; }' +
            'h1, h2, h3, h4, h5, h6 { margin: 0.6em 0 0.3em; }' +
            'p { margin: 0.3em 0; }' +
            // 隐藏编辑器的 UI 残留
            '.code-block-header, .code-block-btn, .code-block-dots, .code-block-copied, .code-block-fade,' +
            '.code-line-nums, .code-block-fold, .code-content-wrap { display: none !important; }' +
            // 确保代码块表格样式
            'td[style*="background:#f5f5f5"] { background: #f5f5f5 !important; }' +
            '</style>';

        var w = window.open('', '', 'width=800,height=600');
        w.document.write('<html><head><meta charset="utf-8"><title>打印文档</title>' + printStyle + '</head><body>' + clone.innerHTML + '</body></html>');
        w.document.close();
        // 等待样式和表格渲染完成后弹出打印对话框
        setTimeout(function() { w.focus(); w.print(); }, 300);
    });

    editorContainer.addEventListener('dragover', function(e) { e.preventDefault(); });
    editorContainer.addEventListener('drop', function(e) { e.preventDefault(); if (e.dataTransfer.files.length && e.dataTransfer.files[0].name.endsWith('.docx')) importDocx(e.dataTransfer.files[0]); });

    imgCenterToggle.addEventListener('change', function() { if (this.checked) centerAllImages(); else uncenterAllImages(); });
    var codeDetectToggle = document.getElementById('codeDetectToggle');
    if (codeDetectToggle) {
        codeDetectToggle.addEventListener('change', function() { showToast('代码检测已' + (this.checked ? '开启' : '关闭'), 'info'); saveConfigToStorage(); });
    }
    var codeFoldLines = document.getElementById('codeFoldLines');
    if (codeFoldLines) {
        codeFoldLines.addEventListener('change', function() {
            saveConfigToStorage();
            updateAllCodeBlocks();
        });
    }

    // 代码主题切换
    var codeThemeSelect = document.getElementById('codeThemeSelect');
    if (codeThemeSelect) {
        codeThemeSelect.addEventListener('change', function() {
            saveConfigToStorage();
            applyAllCodeThemes();
        });
    }

    // ===== 插入代码弹窗 =====
    var codeModalOverlay = document.getElementById('codeModalOverlay');
    var codeModalClose = document.getElementById('codeModalClose');
    var codeModalCancel = document.getElementById('codeModalCancel');
    var codeModalInsert = document.getElementById('codeModalInsert');
    var codeInput = document.getElementById('codeInput');
    var codeLangSelect = document.getElementById('codeLangSelect');
    var insertCodeBtn = document.getElementById('insertCodeBtn');

    function showCodeModal() {
        createInsertMarker(); // 在光标位置插入 DOM 标记
        codeModalOverlay.classList.remove('hidden');
        if (codeInput) { codeInput.value = ''; setTimeout(function() { codeInput.focus(); }, 100); }
    }
    function hideCodeModal() { codeModalOverlay.classList.add('hidden'); removeInsertMarker(); }

    if (insertCodeBtn) insertCodeBtn.addEventListener('click', showCodeModal);
    if (codeModalClose) codeModalClose.addEventListener('click', hideCodeModal);
    if (codeModalCancel) codeModalCancel.addEventListener('click', hideCodeModal);
    if (codeModalOverlay) {
        codeModalOverlay.addEventListener('click', function(e) {
            if (e.target === this) hideCodeModal();
        });
    }
    if (codeModalInsert) {
        codeModalInsert.addEventListener('click', function() {
            saveUndoState('插入代码块'); // 记录插入代码前状态
            var raw = codeInput ? codeInput.value : '';
            raw = raw.replace(/^\n+/, '').replace(/\n+$/, '');
            if (!raw.trim()) { showToast('请粘贴代码', 'warning'); return; }
            var lang = codeLangSelect ? codeLangSelect.value : 'Code';
            var lineCount = raw.split('\n').length;
            var foldLines = (function(v){var n=parseInt(v);return isNaN(n)||n<0?8:n;})(document.getElementById('codeFoldLines').value);
            var shouldFold = foldLines > 0 && lineCount > foldLines;
            var marker = getInsertMarker();
            if (marker) {
                var cb = buildCodeBlock(raw, '', lineCount, foldLines, shouldFold, lang);
                marker.parentNode.insertBefore(cb, marker);
                marker.parentNode.removeChild(marker);
                insertMarkerId = null;
                var r2 = document.createRange();
                r2.setStartAfter(cb); r2.collapse(true);
                var sel = window.getSelection();
                sel.removeAllRanges(); sel.addRange(r2);
            } else {
                var cb = buildCodeBlock(raw, '', lineCount, foldLines, shouldFold, lang);
                editor.focus();
                try { document.execCommand('insertHTML', false, cb.outerHTML); } catch(e) { editor.appendChild(cb); }
            }
            hideCodeModal();
            showToast('已插入 ' + lineCount + ' 行代码', 'success');
        });
    }

    // 重新挂接代码块事件（innerHTML/outerHTML 丢失了 DOM 事件）
    function rewireCodeBlock(cb, raw) {
        var foldBtn = cb.querySelector('.code-block-header-right .code-block-btn:nth-child(2)');
        var copyBtn = cb.querySelector('.code-block-header-right .code-block-btn:nth-child(3)');
        var themeBtn = cb.querySelector('.code-block-header-right .code-block-btn:first-child');
        var lineCount = parseInt(cb.getAttribute('data-lines')) || 0;
        var LH = 19;
        if (foldBtn) {
            foldBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                var wrap = cb.querySelector('.code-block-fold');
                wrap.classList.toggle('collapsed');
                var nowCollapsed = wrap.classList.contains('collapsed');
                this.textContent = nowCollapsed ? '📂 展开 (' + lineCount + ' 行)' : '📁 收起 (' + lineCount + ' 行)';
                var bd = cb.querySelector('.code-block-body');
                if (nowCollapsed) {
                    var fl = parseInt(cb.getAttribute('data-fold')) || 8;
                    wrap.style.maxHeight = (fl * LH + 24) + 'px';
                    if (bd) bd.style.maxHeight = '';
                } else { wrap.style.maxHeight = ''; if (bd) bd.style.maxHeight = '60vh'; }
            });
        }
        if (copyBtn) {
            copyBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                if (navigator.clipboard) {
                    navigator.clipboard.writeText(raw).then(function() { showCopyFeedback(cb); })
                        .catch(function() { fallbackCopy(raw, cb); });
                } else { fallbackCopy(raw, cb); }
            });
        }
        if (themeBtn) {
            themeBtn.addEventListener('click', function(e) {
                e.stopPropagation(); cycleCodeTheme(cb);
            });
        }
    }
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && !codeModalOverlay.classList.contains('hidden')) hideCodeModal();
    });

    // ===== 对象浏览（代码块/表格列表） =====
    var objPanel = document.getElementById('objectBrowserPanel');
    var objHeader = document.getElementById('objectBrowserHeader');
    var objClose = document.getElementById('objectBrowserClose');
    var objBrowseCode = document.getElementById('objBrowseCode');
    var objBrowseTable = document.getElementById('objBrowseTable');
    var objBrowseList = document.getElementById('objBrowseList');
    var objBrowseInfo = document.getElementById('objBrowseInfo');
    var objPanelOpen = false;

    if (objHeader) {
        objHeader.addEventListener('mousedown', function(e) {
            if (e.button !== 0) return;
            var r = objPanel.getBoundingClientRect();
            var od = { sx: e.clientX, sy: e.clientY, l: r.left, t: r.top };
            var m = function(ev) { objPanel.style.left = (od.l + ev.clientX - od.sx) + 'px'; objPanel.style.top = (od.t + ev.clientY - od.sy) + 'px'; objPanel.style.right = 'auto'; };
            var u = function() { document.removeEventListener('mousemove', m); document.removeEventListener('mouseup', u); };
            document.addEventListener('mousemove', m);
            document.addEventListener('mouseup', u);
            e.preventDefault();
        });
    }

    function showObjPanel() { if (objPanelOpen) return; objPanelOpen = true; objPanel.classList.remove('hidden'); objPanel.style.left = 'auto'; objPanel.style.right = '50px'; objPanel.style.top = '130px'; }
    function hideObjPanel() { objPanelOpen = false; objPanel.classList.add('hidden'); }

    var openObjBtn = document.getElementById('openObjectBrowser');
    if (openObjBtn) openObjBtn.addEventListener('click', showObjPanel);
    if (objClose) objClose.addEventListener('click', hideObjPanel);
    document.addEventListener('keydown', function(e) { if (e.key === 'Escape' && objPanelOpen) hideObjPanel(); });

    function renderObjList(type) {
        if (!objBrowseList || !objBrowseInfo) return;
        var items = [];
        if (type === 'code') {
            editor.querySelectorAll('.code-block').forEach(function(cb, i) {
                var lang = cb.querySelector('.code-block-lang');
                var lines = parseInt(cb.getAttribute('data-lines')) || 0;
                var code = cb.querySelector('code');
                var text = code ? code.textContent.substring(0, 50).replace(/\n/g, '↵ ') : '';
                if (code && code.textContent.length > 50) text += '…';
                items.push({ i: i, el: cb, label: '#' + (i+1) + ' ' + (lang?lang.textContent:'Code') + ' (' + lines + ' 行)', prev: text });
            });
        } else {
            editor.querySelectorAll('table').forEach(function(tbl, i) {
                var rows = tbl.querySelectorAll('tr').length;
                var cols = rows ? Math.round(tbl.querySelectorAll('td, th').length / rows) : 0;
                var text = (tbl.textContent||'').substring(0, 50).replace(/\n/g, '↵ ');
                if ((tbl.textContent||'').length > 50) text += '…';
                items.push({ i: i, el: tbl, label: '#' + (i+1) + ' 表格 (' + rows + '×' + cols + ')', prev: text });
            });
        }
        objBrowseInfo.textContent = items.length + ' 个';
        if (!items.length) { objBrowseList.innerHTML = '<div class="search-result-empty">未找到' + (type==='code'?'代码块':'表格') + '</div>'; return; }
        var html = '';
        for (var k = 0; k < items.length; k++) {
            var it = items[k];
            html += '<div class="search-result-item" data-idx="' + k + '">';
            html += '<div style="display:flex;justify-content:space-between;align-items:center;">';
            html += '<div style="flex:1;min-width:0;">';
            html += '<div style="font-size:12px;font-weight:500;">' + escHtml(it.label) + '</div>';
            if (it.prev) html += '<div style="font-size:11px;color:var(--text-secondary);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-family:monospace;">' + escHtml(it.prev) + '</div>';
            html += '</div>';
            html += '<button class="obj-del-btn" data-idx="' + k + '" title="删除" style="flex-shrink:0;background:none;border:1px solid var(--border);border-radius:4px;cursor:pointer;padding:2px 8px;font-size:12px;color:#dc2626;">🗑</button>';
            html += '</div></div>';
        }
        objBrowseList.innerHTML = html;
        objBrowseList.querySelectorAll('.search-result-item').forEach(function(row) {
            row.addEventListener('click', function(e) {
                if (e.target.closest('.obj-del-btn')) return;
                var it = items[parseInt(this.dataset.idx)];
                if (!it || !it.el) return;
                scrollEditorTo(it.el, 'center');
                it.el.style.outline = '3px solid var(--primary)'; it.el.style.outlineOffset = '2px';
                setTimeout(function() { it.el.style.outline = ''; }, 2000);
                objBrowseList.querySelectorAll('.search-result-item.active').forEach(function(a) { a.classList.remove('active'); });
                row.classList.add('active');
            });
        });
        objBrowseList.querySelectorAll('.obj-del-btn').forEach(function(btn) {
            btn.addEventListener('click', function(e) {
                e.stopPropagation();
                var it = items[parseInt(this.dataset.idx)];
                if (!it || !it.el || !confirm('确定删除此' + (type==='code'?'代码块':'表格') + '？')) return;
                it.el.parentNode.removeChild(it.el);
                showToast('已删除', 'info');
                renderObjList(type);
            });
        });
    }

    if (objBrowseCode) objBrowseCode.addEventListener('click', function() { renderObjList('code'); });
    if (objBrowseTable) objBrowseTable.addEventListener('click', function() { renderObjList('table'); });
    document.querySelectorAll('#headingControls select, #headingControls input[type="checkbox"]').forEach(function(el) { el.addEventListener('change', applyHeadingStylesToEditor); });
    applyBodyFormat.addEventListener('click', applyBodyFormatFn);

    // ===== 查找与替换浮动面板（可拖动、不遮挡） =====
    var searchFloatingPanel = document.getElementById('searchFloatingPanel');
    var searchPanelHeader = document.getElementById('searchPanelHeader');
    var searchPanelClose = document.getElementById('searchPanelClose');
    var openSearchPanelBtn = document.getElementById('openSearchModal');
    var searchPanelOpen = false;
    var dragData = null;

    function startDrag(e) {
        if (e.button !== 0) return;
        var rect = searchFloatingPanel.getBoundingClientRect();
        dragData = { startX: e.clientX, startY: e.clientY, panelLeft: rect.left, panelTop: rect.top };
        document.addEventListener('mousemove', onDrag);
        document.addEventListener('mouseup', endDrag);
        e.preventDefault();
    }

    function onDrag(e) {
        if (!dragData) return;
        searchFloatingPanel.style.left = (dragData.panelLeft + e.clientX - dragData.startX) + 'px';
        searchFloatingPanel.style.top = (dragData.panelTop + e.clientY - dragData.startY) + 'px';
        searchFloatingPanel.style.right = 'auto';
    }

    function endDrag() {
        dragData = null;
        document.removeEventListener('mousemove', onDrag);
        document.removeEventListener('mouseup', endDrag);
    }

    if (searchPanelHeader) searchPanelHeader.addEventListener('mousedown', startDrag);

    function showSearchPanel(focusReplace) {
        if (searchPanelOpen) return;
        searchPanelOpen = true;
        searchFloatingPanel.classList.remove('hidden');
        searchFloatingPanel.style.left = 'auto';
        searchFloatingPanel.style.right = '40px';
        searchFloatingPanel.style.top = '80px';
        searchInput.value = getEditorSelectedText();
        setTimeout(function() {
            if (focusReplace) { replaceInput.focus(); replaceInput.select(); }
            else { searchInput.focus(); searchInput.select(); }
        }, 100);
    }

    function hideSearchPanel() {
        searchPanelOpen = false;
        searchFloatingPanel.classList.add('hidden');
        clearSearchHighlights();
        var rl = document.getElementById('searchResultList');
        if (rl) rl.innerHTML = '<div class="search-result-empty">输入关键词后点击"查找"</div>';
        document.getElementById('searchMatchInfo').textContent = '0 个匹配';
        document.getElementById('prevMatch').disabled = true;
        document.getElementById('nextMatch').disabled = true;
        document.getElementById('replaceBtn').disabled = true;
        currentMatches = []; currentMatchIndex = -1;
    }

    function getEditorSelectedText() {
        var sel = window.getSelection();
        if (sel.rangeCount) {
            var txt = sel.toString().trim();
            if (txt && editor.contains(sel.getRangeAt(0).commonAncestorContainer)) return txt;
        }
        return '';
    }

    if (openSearchPanelBtn) openSearchPanelBtn.addEventListener('click', function() { showSearchPanel(false); });
    if (searchPanelClose) searchPanelClose.addEventListener('click', hideSearchPanel);

    // ===== 帮助弹窗（可拖动、多标签、非模态） =====
    var helpWindow = document.getElementById('helpWindow');
    var helpHeader = document.getElementById('helpHeader');
    var helpBtn = document.getElementById('helpBtn');
    var helpClose = document.getElementById('helpClose');
    var helpOpen = false;
    var helpDragData = null;

    function helpStartDrag(e) {
        if (e.button !== 0) return;
        if (e.target.closest && e.target.closest('button')) return;
        var rect = helpWindow.getBoundingClientRect();
        helpDragData = { startX: e.clientX, startY: e.clientY, panelLeft: rect.left, panelTop: rect.top };
        document.addEventListener('mousemove', helpOnDrag);
        document.addEventListener('mouseup', helpEndDrag);
        e.preventDefault();
    }

    function helpOnDrag(e) {
        if (!helpDragData) return;
        helpWindow.style.left = (helpDragData.panelLeft + e.clientX - helpDragData.startX) + 'px';
        helpWindow.style.top = (helpDragData.panelTop + e.clientY - helpDragData.startY) + 'px';
        helpWindow.style.transform = 'none';
    }

    function helpEndDrag() {
        helpDragData = null;
        document.removeEventListener('mousemove', helpOnDrag);
        document.removeEventListener('mouseup', helpEndDrag);
    }

    function showHelp() {
        if (helpOpen) { hideHelp(); return; }
        helpOpen = true;
        helpWindow.classList.remove('hidden');
        helpWindow.style.left = '50%';
        helpWindow.style.top = '100px';
        helpWindow.style.transform = 'translateX(-50%)';
        var firstTab = helpWindow.querySelector('.help-tab');
        if (firstTab) firstTab.click();
    }

    function hideHelp() {
        helpOpen = false;
        helpWindow.classList.add('hidden');
    }

    if (helpHeader) helpHeader.addEventListener('mousedown', helpStartDrag);
    if (helpBtn) helpBtn.addEventListener('click', showHelp);
    if (helpClose) helpClose.addEventListener('click', hideHelp);

    // 帮助标签切换
    (function() {
        var tabs = helpWindow.querySelectorAll('.help-tab');
        var contents = helpWindow.querySelectorAll('.help-tab-content');
        tabs.forEach(function(tab) {
            tab.addEventListener('click', function() {
                var target = this.getAttribute('data-tab');
                tabs.forEach(function(t) { t.classList.remove('active'); });
                this.classList.add('active');
                contents.forEach(function(c) {
                    c.classList.remove('active');
                    if (c.id === 'helpTab' + target.charAt(0).toUpperCase() + target.slice(1)) {
                        c.classList.add('active');
                    }
                });
            });
        });
    })();

    // F1 快捷键切换帮助
    document.addEventListener('keydown', function(e) {
        if (e.key === 'F1') {
            e.preventDefault();
            if (helpOpen) { hideHelp(); } else { showHelp(); }
        }
    });

    (function() {
        var origHandler = handleKeyboard;
        handleKeyboard = function(e) {
            var isCtrl = e.ctrlKey || e.metaKey;
            if (isCtrl) {
                if (e.key === 'z' || e.key === 'Z') {
                    e.preventDefault();
                    if (e.shiftKey) { redoPerform(); }
                    else { undoPerform(); }
                    return;
                }
                if (e.key === 'y' || e.key === 'Y') {
                    e.preventDefault();
                    redoPerform();
                    return;
                }
                if (e.key === 'f' || e.key === 'F') { e.preventDefault(); showSearchPanel(false); return; }
                if (e.key === 'h' || e.key === 'H') { e.preventDefault(); showSearchPanel(true); return; }
            }
            if (e.key === 'Escape' && searchPanelOpen) { hideSearchPanel(); return; }
            // Escape 关闭帮助弹窗
            if (e.key === 'Escape' && helpOpen) { hideHelp(); return; }
            origHandler.call(this, e);
        };
    })();

    searchBtn.addEventListener('click', performSearch);
    searchInput.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (e.shiftKey) goToPrevMatch();
            else { performSearch(); if (currentMatches.length) goToNextMatch(); }
        }
    });
    replaceInput.addEventListener('keydown', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            if (e.shiftKey) replaceAll(); else replaceCurrent();
        }
    });
    document.getElementById('nextMatch').addEventListener('click', goToNextMatch);
    document.getElementById('prevMatch').addEventListener('click', goToPrevMatch);
    document.getElementById('replaceBtn').addEventListener('click', replaceCurrent);
    document.getElementById('replaceAllBtn').addEventListener('click', replaceAll);


    tocRefresh.addEventListener('click', generateTOC);
    tocCollapseAll.addEventListener('click', function() { tocContainer.querySelectorAll('.toc-children').forEach(function(el) { el.classList.add('collapsed'); var t = el.parentElement.querySelector('.toc-toggle'); if (t) t.textContent = '▶'; }); });
    tocExpandAll.addEventListener('click', function() { tocContainer.querySelectorAll('.toc-children').forEach(function(el) { el.classList.remove('collapsed'); var t = el.parentElement.querySelector('.toc-toggle'); if (t) t.textContent = '▼'; }); });

    var themeBtn = document.getElementById('themeToggle');
    if (themeBtn) themeBtn.addEventListener('click', toggleTheme);

    var fullscreenBtn = document.getElementById('fullscreenToggle');
    if (fullscreenBtn) {
        fullscreenBtn.addEventListener('click', function() {
            if (!document.fullscreenElement) document.documentElement.requestFullscreen().catch(function(){});
            else document.exitFullscreen().catch(function(){});
        });
    }

    var saveConfigBtn = document.getElementById('saveConfigBtn');
    var loadConfigBtn = document.getElementById('loadConfigBtn');
    var configFileInput = document.getElementById('configFileInput');
    if (saveConfigBtn) saveConfigBtn.addEventListener('click', saveConfigToFile);
    if (loadConfigBtn) loadConfigBtn.addEventListener('click', function() { configFileInput.click(); });
    if (configFileInput) configFileInput.addEventListener('change', function(e) { if (this.files && this.files[0]) loadConfigFromFile(this.files[0]); this.value = ''; });

    document.querySelectorAll('.font-color').forEach(function(input) { input.addEventListener('input', applyHeadingStylesToEditor); });

    // AI 标题助手
    var applyHintBtn = document.getElementById('applyHintBtn');
    var clearHintBtn = document.getElementById('clearHintBtn');
    var downloadRuleLink = document.getElementById('downloadRuleLink');
    if (clearHintBtn) clearHintBtn.addEventListener('click', clearHint);
    if (downloadRuleLink) downloadRuleLink.addEventListener('click', function(e) { e.preventDefault(); downloadRuleFile(); });

    // ===== JSON 检测与格式化（弹窗） =====
    var jsonModalOverlay = document.getElementById('jsonModalOverlay');
    var jsonModalClose = document.getElementById('jsonModalClose');
    var openJsonModalBtn = document.getElementById('openJsonModal');
    var jsonStatus = document.getElementById('jsonStatus');
    var validateJsonBtn = document.getElementById('validateJsonBtn');
    var formatJsonBtn = document.getElementById('formatJsonBtn');
    var hintInputEl = document.getElementById('hintInput');
    var lastErrorPos = null;
    var jsonModalOpen = false;

    function showJsonModal() {
        if (jsonModalOpen) return;
        jsonModalOpen = true;
        jsonModalOverlay.classList.remove('hidden');
        updateLineNumbers();
        setTimeout(function() { if (hintInputEl) hintInputEl.focus(); }, 100);
    }

    function hideJsonModal() {
        jsonModalOpen = false;
        jsonModalOverlay.classList.add('hidden');
        if (jsonStatus) { jsonStatus.className = 'json-status hidden'; }
    }

    // 行号渲染
    var errorLineNum = null;

    function syncTextareaHeight() {
        if (!hintInputEl) return;
        hintInputEl.style.height = 'auto';
        hintInputEl.style.height = hintInputEl.scrollHeight + 'px';
    }

    function updateLineNumbers() {
        var gutter = document.getElementById('hintLineNumbers');
        if (!gutter || !hintInputEl) return;
        var text = hintInputEl.value;
        var lineCount = text.split('\n').length;
        var html = '';
        for (var i = 1; i <= lineCount; i++) {
            var cls = 'ln' + (i === errorLineNum ? ' error' : '');
            html += '<span class="' + cls + '">' + i + '</span>';
        }
        gutter.innerHTML = html;
        // 行号和文本框在同一个滚动容器中，同步高度即可
        syncTextareaHeight();
    }

    if (hintInputEl) {
        hintInputEl.addEventListener('input', function() {
            errorLineNum = null;
            updateLineNumbers();
        });
        // 不监听 scroll — 外层容器统一滚动
    }

    if (openJsonModalBtn) openJsonModalBtn.addEventListener('click', showJsonModal);
    if (jsonModalClose) jsonModalClose.addEventListener('click', hideJsonModal);
    if (jsonModalOverlay) {
        jsonModalOverlay.addEventListener('click', function(e) {
            if (e.target === this) hideJsonModal();
        });
    }
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && jsonModalOpen) hideJsonModal();
    });

    function showJsonStatus(msg, type) {
        if (!jsonStatus) return;
        jsonStatus.className = 'json-status';
        jsonStatus.innerHTML = '';
        if (type) {
            jsonStatus.classList.add(type);
            jsonStatus.classList.remove('hidden');
        } else {
            jsonStatus.classList.add('hidden');
            return;
        }
        if (lastErrorPos && type === 'error') {
            var parts = msg.split('\n');
            for (var i = 0; i < parts.length; i++) {
                var p = parts[i];
                if (p.indexOf('📍') >= 0) {
                    var span = document.createElement('span');
                    span.className = 'err-location';
                    span.textContent = '📍 第 ' + lastErrorPos.line + ' 行，第 ' + lastErrorPos.col + ' 列';
                    span.title = '点击跳转到错误位置';
                    (function(pos) {
                        span.addEventListener('click', function() {
                            jumpToErrorPosition(pos.line, pos.col, pos.msg);
                        });
                    })(lastErrorPos);
                    jsonStatus.appendChild(span);
                    jsonStatus.appendChild(document.createElement('br'));
                } else if (p.indexOf('→') === 0 || p.indexOf('↑') === 0 || p.indexOf('^') === 0) {
                    var pre = document.createElement('span');
                    pre.className = 'err-line-preview';
                    pre.textContent = p;
                    jsonStatus.appendChild(pre);
                } else {
                    jsonStatus.appendChild(document.createTextNode(p));
                    if (i < parts.length - 1) jsonStatus.appendChild(document.createElement('br'));
                }
            }
        } else {
            jsonStatus.textContent = msg;
        }
    }

    function markErrorLine(line, col) {
        errorLineNum = line;
        updateLineNumbers();
        if (!hintInputEl || !line) return;
        var lineHeight = 18;
        hintInputEl.scrollTop = Math.max(0, (line - 3) * lineHeight);
        if (col) {
            hintInputEl.focus();
            var text = hintInputEl.value;
            var lines = text.split('\n');
            var pos = 0;
            for (var i = 0; i < line - 1 && i < lines.length; i++) {
                pos += lines[i].length + 1;
            }
            pos += Math.min(col - 1, lines[line - 1] ? lines[line - 1].length : 0);
            pos = Math.min(pos, text.length);
            if (hintInputEl.setSelectionRange) {
                hintInputEl.setSelectionRange(pos, pos);
            }
        }
    }

    function jumpToErrorPosition(line, col) {
        if (!hintInputEl || !line || !col) return;
        showJsonModal();
        setTimeout(function() {
            markErrorLine(line, col);
            showToast('已定位到第 ' + line + ' 行', 'info', 1500);
        }, 200);
    }

    function validateJson(str) {
        str = str.trim();
        if (!str) {
            showJsonStatus('请输入 JSON 内容', 'warning');
            return null;
        }
        try {
            var parsed = JSON.parse(str);
            var formatted = JSON.stringify(parsed, null, 2);
            lastErrorPos = null;
            errorLineNum = null;
            updateLineNumbers();
            if (Array.isArray(parsed)) {
                var issues = [];
                for (var i = 0; i < parsed.length; i++) {
                    var item = parsed[i];
                    if (!item.level) issues.push('第 ' + (i + 1) + ' 项缺少 level');
                    if (!item.text) issues.push('第 ' + (i + 1) + ' 项缺少 text');
                    if (!item.anchor) issues.push('第 ' + (i + 1) + ' 项缺少 anchor');
                    if (item.level && (item.level < 1 || item.level > 6)) issues.push('第 ' + (i + 1) + ' 项 level 超出范围 (1-6)');
                }
                if (issues.length) {
                    showJsonStatus('⚠ 格式有效，但数据有问题：\n' + issues.join('\n'), 'warning');
                } else {
                    showJsonStatus('✅ JSON 格式正确！共 ' + parsed.length + ' 条标题', 'success');
                }
            } else {
                showJsonStatus('⚠ JSON 格式正确，但需要的是数组 [...]', 'warning');
            }
            return { valid: true, data: parsed, formatted: formatted };
        } catch(e) {
            var msg = e.message || 'JSON 解析失败';
            var line = null, col = null;
            var m = msg.match(/position\s+(\d+)/i);
            if (m) {
                var pos = parseInt(m[1]);
                var before = str.substring(0, pos);
                line = (before.match(/\n/g) || []).length + 1;
                var lastNewline = before.lastIndexOf('\n');
                col = pos - lastNewline;
            }
            var m2 = msg.match(/line\s+(\d+)\s+column\s+(\d+)/i);
            if (m2) { line = parseInt(m2[1]); col = parseInt(m2[2]); }

            lastErrorPos = (line !== null && col !== null) ? { line: line, col: col, msg: msg } : null;

            var errMsg = '❌ ' + msg;
            if (lastErrorPos) {
                errMsg += '\n📍 第 ' + line + ' 行，第 ' + col + ' 列（点击跳转）';
                var lns = str.split('\n');
                if (line >= 1 && line <= lns.length) {
                    errMsg += '\n→ ' + lns[line - 1].substring(0, 60) + (lns[line - 1].length > 60 ? '...' : '');
                    errMsg += '\n' + ' '.repeat(Math.min(col - 1, 60)) + '↑';
                }
            }
            showJsonStatus(errMsg, 'error');
            // 自动高亮错误位置，带上错误消息
            if (line && col) {
                setTimeout(function() { jumpToErrorPosition(line, col, msg); }, 100);
            }
            return { valid: false, error: msg, line: line, col: col };
        }
    }

    if (validateJsonBtn) {
        validateJsonBtn.addEventListener('click', function() {
            var val = hintInputEl ? hintInputEl.value.trim() : '';
            if (!val) {
                showJsonStatus('请输入 JSON 内容', 'warning');
                hintInputEl.focus();
                return;
            }
            validateJson(val);
        });
    }

    if (formatJsonBtn) {
        formatJsonBtn.addEventListener('click', function() {
            var val = hintInputEl ? hintInputEl.value.trim() : '';
            if (!val) {
                showJsonStatus('请输入 JSON 内容', 'warning');
                hintInputEl.focus();
                return;
            }
            var result = validateJson(val);
            if (result && result.valid && hintInputEl) {
                hintInputEl.value = result.formatted;
                showJsonStatus('✅ 已格式化完成（' + result.data.length + ' 条）', 'success');
                // 重新渲染行号并同步高度
                errorLineNum = null;
                updateLineNumbers();
            }
        });
    }

    // 应用标题 — 从弹窗中执行后关闭
    if (applyHintBtn) {
        applyHintBtn.addEventListener('click', function() {
            var input = document.getElementById('hintInput');
            if (input && input.value.trim()) {
                applyHintFromJSON(input.value.trim());
                hideJsonModal();
            } else showToast('请先粘贴 hint JSON', 'warning');
        });
    }

    // ===== IndexedDB 自动保存/恢复 =====
    var DB_NAME = 'DocxEditorDB', DB_VERSION = 2;
    function openDocDB() {
        return new Promise(function(resolve, reject) {
            var req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = function(e) {
                var db = e.target.result;
                if (!db.objectStoreNames.contains('docs')) db.createObjectStore('docs', { keyPath: 'id' });
            };
            req.onsuccess = function() { resolve(req.result); };
            req.onerror = function() { reject(req.error); };
        });
    }
    function isPlaceholderContent(html) {
        return !html || html.indexOf('class="placeholder"') >= 0 || html.trim() === '';
    }

    // 🔄 增量保存：仅保存变更内容，跳过撤销历史以提升效率
    function saveDocumentState() {
        var html = editor.innerHTML;
        if (isPlaceholderContent(html)) return Promise.resolve();
        // 同步当前状态到活跃会话
        syncStateToSession();
        var session = tabManager.getActive();
        if (!session) return Promise.resolve();
        // 增量检查：如果 HTML 与上次保存相同，跳过
        if (session._lastSavedHtml === session.html) return Promise.resolve();
        var si = document.getElementById('saveIndicator');
        if (si) { si.className = 'save-indicator saving'; si.style.display = 'inline-block'; }
        setStatus('保存中...');
        var docId = session.id;
        return openDocDB().then(function(db) {
            var tx = db.transaction('docs', 'readwrite');
            var store = tx.objectStore('docs');
            // 增量保存：只保存核心内容，跳过庞大的 undoHistory
            var saveData = {
                id: docId,
                title: session.title,
                customTitle: session.customTitle,
                content: session.html,
                scrollTop: session.scrollTop,
                sourceFileName: session.sourceFileName,
                sourceImportTime: session.sourceImportTime,
                savedAt: new Date().toISOString()
            };
            // 仅在图片有变更时保存图片数据
            if (session._imagesChanged) {
                saveData.imageData = session.imageDataMap ? Array.from(session.imageDataMap.entries()) : [];
                session._imagesChanged = false;
            }
            // 始终保存锚点和折叠数据（数据量小）
            saveData.anchors = (session.anchors || []).slice();
            saveData.foldPoints = (session.foldPoints || []).slice();
            saveData.foldRegions = (session.foldRegions || []).slice();
            store.put(saveData);
            // 同时保存元信息
            store.put({
                id: '_meta_',
                activeIndex: tabManager.activeIndex,
                tabCount: tabManager.sessions.length,
                tabIdCounter: tabManager.tabIdCounter
            });
            return new Promise(function(resolve) { tx.oncomplete = resolve; tx.onerror = resolve; });
        }).then(function() {
            // 记录已保存的 HTML，用于下次增量比对
            session._lastSavedHtml = session.html;
            // 清除未保存标记（仅当仍是活跃标签时才更新 DOM 文字）
            session._dirty = false;
            if (tabManager.getActive() === session) {
                tabManager.renderTabs();
            }
            if (si) si.className = 'save-indicator saved';
            setStatus('已自动保存');
            setTimeout(function() {
                if (docStatus && docStatus.textContent === '已自动保存') setStatus('就绪');
                if (si) setTimeout(function() { si.style.display = 'none'; }, 500);
            }, 2000);
        }).catch(function(err) {
            console.warn('Auto-save failed:', err);
            if (si) { si.className = 'save-indicator hidden'; }
            setStatus('保存失败');
        });
    }

    // 🔄 完整保存：保存所有数据（包括撤销历史），用于 Ctrl+S / 标签切换 / 离开页面
    function saveDocumentFull() {
        var html = editor.innerHTML;
        if (isPlaceholderContent(html)) return Promise.resolve();
        syncStateToSession();
        var session = tabManager.getActive();
        if (!session) return Promise.resolve();
        var si = document.getElementById('saveIndicator');
        if (si) { si.className = 'save-indicator saving'; si.style.display = 'inline-block'; }
        setStatus('保存中...');
        var docId = session.id;
        return openDocDB().then(function(db) {
            var tx = db.transaction('docs', 'readwrite');
            var store = tx.objectStore('docs');
            store.put({
                id: docId,
                title: session.title,
                customTitle: session.customTitle,
                content: session.html,
                imageData: session.imageDataMap ? Array.from(session.imageDataMap.entries()) : [],
                undoHistory: (session.undoHistory || []).slice(),
                undoIndex: session.undoIndex,
                anchors: (session.anchors || []).slice(),
                foldPoints: (session.foldPoints || []).slice(),
                foldRegions: (session.foldRegions || []).slice(),
                scrollTop: session.scrollTop,
                sourceFileName: session.sourceFileName,
                sourceImportTime: session.sourceImportTime,
                savedAt: new Date().toISOString()
            });
            store.put({
                id: '_meta_',
                activeIndex: tabManager.activeIndex,
                tabCount: tabManager.sessions.length,
                tabIdCounter: tabManager.tabIdCounter
            });
            return new Promise(function(resolve) { tx.oncomplete = resolve; tx.onerror = resolve; });
        }).then(function() {
            session._lastSavedHtml = session.html;
            session._imagesChanged = false;
            session._dirty = false;
            // 仅当仍是活跃标签时才刷新标签 UI（异步保存可能在切换标签后完成）
            if (tabManager.getActive() === session) {
                tabManager.renderTabs();
            }
            if (si) si.className = 'save-indicator saved';
            setStatus('已保存');
            setTimeout(function() {
                if (docStatus && docStatus.textContent === '已保存') setStatus('就绪');
                if (si) setTimeout(function() { si.style.display = 'none'; }, 500);
            }, 2000);
        }).catch(function(err) {
            console.warn('Full save failed:', err);
            if (si) { si.className = 'save-indicator hidden'; }
            setStatus('保存失败');
        });
    }
    function loadDocumentState() {
        // 已由 TabManager.init() 处理，此函数保留兼容但直接返回 null
        return Promise.resolve(null);
    }
    var autoSaveTimer = null;
    var periodicSaveTimer = null;
    function triggerAutoSave() {
        if (autoSaveTimer) clearTimeout(autoSaveTimer);
        autoSaveTimer = setTimeout(saveDocumentState, 1500);
        // 每次有输入时重置定时保存计时器（让定时器从最后一次输入重新计时）
        startPeriodicSave();
    }
    function startPeriodicSave() {
        if (periodicSaveTimer) clearInterval(periodicSaveTimer);
        periodicSaveTimer = setInterval(function() {
            if (!isPlaceholderContent(editor.innerHTML)) {
                saveDocumentFull(); // 🔄 定期完整保存，确保撤销历史不丢失
            }
        }, 30000); // 每 30 秒完整保存一次
    }

    // ===== 📌 锚点装订线系统 =====
    var anchorGutter = document.getElementById('anchorGutter');
    var anchorPanel = document.getElementById('anchorPanel');
    var anchorPanelBtn = $('anchorPanelBtn');
    var anchorPanelClose = $('anchorPanelClose');
    var anchorList = $('anchorList');
    var anchorPrev = $('anchorPrev');
    var anchorNext = $('anchorNext');
    var anchorClearAll = $('anchorClearAll');
    var anchorPanelOpen = false;

    // 确保每个段落有唯一 data-pid
    function ensureParagraphIds() {
        var counter = 0;
        var now = Date.now();
        editor.querySelectorAll('p, h1, h2, h3, h4, h5, h6').forEach(function(el) {
            if (!el.dataset.pid) {
                el.dataset.pid = 'p-' + now + '-' + (counter++);
            }
        });
    }

    // 根据 Y 坐标找到最近的段落元素（Y 相对于给定参考元素的顶部）
    function findParagraphAtY(y, refEl) {
        var paras = editor.querySelectorAll('p, h1, h2, h3, h4, h5, h6');
        var closest = null;
        var minDist = Infinity;
        var refRect = (refEl || editor).getBoundingClientRect();
        paras.forEach(function(p) {
            var rect = p.getBoundingClientRect();
            var mid = rect.top - refRect.top + rect.height / 2;
            var dist = Math.abs(mid - y);
            if (dist < minDist) { minDist = dist; closest = p; }
        });
        return closest;
    }

    // 渲染装订线中的锚点圆点
    function renderAnchorGutter() {
        if (!anchorGutter) return;
        anchorGutter.querySelectorAll('.anchor-gutter-dot').forEach(function(d) { d.remove(); });
        var gutterRect = anchorGutter.getBoundingClientRect();
        anchors.forEach(function(anchor) {
            var el = editor.querySelector('[data-pid="' + anchor.pid + '"]');
            if (!el) return;
            if (el.classList.contains('fold-hidden')) return;
            var elRect = el.getBoundingClientRect();
            var topOffset = elRect.top - gutterRect.top + elRect.height / 2 - 5;
            var dot = document.createElement('div');
            dot.className = 'anchor-gutter-dot';
            dot.style.top = topOffset + 'px';
            dot.title = '锚点: ' + anchor.text.substring(0, 30) + '\n点击移除';
            dot.addEventListener('click', function(e) {
                e.stopPropagation();
                removeAnchor(anchor.id);
            });
            anchorGutter.appendChild(dot);
        });
    }

    // 同步装订线高度（flex 布局自动撑高，此函数仅更新圆点位置）
    function syncGutterHeight() {
        // 装订线高度由 flex 布局自动与编辑器对齐，无需手动设置
        renderAnchorGutter();
    }

    // 切换锚点（在指定段落上）
    function toggleAnchor(el) {
        ensureParagraphIds();
        syncGutterHeight();
        var pid = el.dataset.pid;
        if (!pid) return;
        var existingIdx = -1;
        for (var i = 0; i < anchors.length; i++) {
            if (anchors[i].pid === pid) { existingIdx = i; break; }
        }
        if (existingIdx >= 0) {
            anchors.splice(existingIdx, 1);
            showToast('锚点已移除', 'info', 1200);
        } else {
            var text = el.textContent.trim().substring(0, 50);
            if (!text) text = '(空段落)';
            anchors.push({
                id: 'a-' + Date.now(),
                pid: pid,
                text: text,
                createdAt: new Date().toISOString()
            });
            showToast('锚点已添加', 'success', 1200);
        }
        renderAnchorGutter();
        renderAnchorPanel();
        // 标记为未保存
        var session = tabManager.getActive();
        if (session) session._dirty = true;
    }

    // 移除指定锚点
    function removeAnchor(anchorId) {
        for (var i = 0; i < anchors.length; i++) {
            if (anchors[i].id === anchorId) {
                anchors.splice(i, 1);
                break;
            }
        }
        renderAnchorGutter();
        renderAnchorPanel();
        showToast('锚点已移除', 'info', 1200);
    }

    // 将元素滚动到 editorContainer 可视区域内（不影响页面整体滚动）
    function scrollEditorTo(el, block) {
        if (!editorContainer || !el) return;
        block = block || 'center';
        var containerRect = editorContainer.getBoundingClientRect();
        var elRect = el.getBoundingClientRect();
        var offset;
        if (block === 'start') {
            offset = elRect.top - containerRect.top - 20;
        } else if (block === 'nearest') {
            if (elRect.top < containerRect.top + 40) {
                offset = elRect.top - containerRect.top - 40;
            } else if (elRect.bottom > containerRect.bottom - 40) {
                offset = elRect.bottom - containerRect.bottom + 40;
            } else { return; }
        } else {
            offset = elRect.top - containerRect.top - containerRect.height / 2 + elRect.height / 2;
        }
        var targetScroll = editorContainer.scrollTop + offset;
        targetScroll = Math.max(0, Math.min(targetScroll, editorContainer.scrollHeight - editorContainer.clientHeight));
        editorContainer.scrollTo({ top: targetScroll, behavior: 'smooth' });
    }

    // 跳转到指定锚点（仅操作 editorContainer 滚动，避免页面整体滚动）
    function jumpToAnchor(anchorId) {
        var anchor = null;
        for (var i = 0; i < anchors.length; i++) {
            if (anchors[i].id === anchorId) { anchor = anchors[i]; break; }
        }
        if (!anchor) return;
        var el = editor.querySelector('[data-pid="' + anchor.pid + '"]');
        if (!el) {
            showToast('锚点目标段落已不存在', 'warning', 2000);
            removeAnchor(anchorId);
            return;
        }
        // 检查锚点是否在折叠区域内
        if (el.classList.contains('fold-hidden')) {
            var foldInfo = findFoldRegionContaining(anchor.pid);
            if (foldInfo) {
                // 找到该区域在 foldRegions 中的序号
                var regionIdx = -1;
                for (var ri = 0; ri < foldRegions.length; ri++) {
                    if (foldRegions[ri].id === foldInfo.region.id) { regionIdx = ri; break; }
                }
                showToast('📁 锚点在折叠 #' + (regionIdx + 1) + ' 中，请先展开该折叠区域', 'warning', 3500);
            } else {
                showToast('📁 锚点所在段落已被折叠，请先展开', 'warning', 3500);
            }
            return;
        }
        scrollEditorTo(el, 'center');
        // 高亮动画
        el.classList.remove('anchor-flash');
        void el.offsetWidth;
        el.classList.add('anchor-flash');
        // 移动端优化
        if (window.innerWidth < 900 && anchorPanelOpen) {
            hideAnchorPanel();
        }
    }

    // 上一个/下一个锚点导航
    function navigateAnchor(direction) {
        if (!anchors.length) {
            showToast('暂无锚点', 'info', 1500);
            return;
        }
        // 找到当前光标或滚动位置附近的锚点
        var currentIdx = -1;
        var scrollTop = editorContainer ? editorContainer.scrollTop : 0;
        var viewCenter = scrollTop + (editorContainer ? editorContainer.clientHeight / 2 : 300);
        var minDist = Infinity;
        for (var i = 0; i < anchors.length; i++) {
            var el = editor.querySelector('[data-pid="' + anchors[i].pid + '"]');
            if (!el) continue;
            var elTop = el.offsetTop;
            if (direction < 0 && elTop < viewCenter - 20) {
                // 上一个：找视口上方最近的
                var dist = viewCenter - elTop;
                if (dist < minDist) { minDist = dist; currentIdx = i; }
            } else if (direction > 0 && elTop > viewCenter + 20) {
                // 下一个：找视口下方最近的
                var distUp = elTop - viewCenter;
                if (distUp < minDist) { minDist = distUp; currentIdx = i; }
            }
        }
        if (currentIdx < 0) {
            // 如果没找到上方/下方的，循环到最后一个/第一个
            if (direction < 0) currentIdx = anchors.length - 1;
            else currentIdx = 0;
        }
        if (currentIdx >= 0 && currentIdx < anchors.length) {
            jumpToAnchor(anchors[currentIdx].id);
        }
    }

    // 渲染锚点导航面板
    function renderAnchorPanel() {
        if (!anchorList) return;
        if (!anchors.length) {
            anchorList.innerHTML = '<div class="anchor-empty">暂无锚点<br>点击编辑器左侧灰色区域添加锚点</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < anchors.length; i++) {
            var a = anchors[i];
            html += '<div class="anchor-item" data-anchor-id="' + a.id + '">' +
                '<span class="anchor-dot">●</span>' +
                '<span class="anchor-index">#' + (i + 1) + '</span>' +
                '<span class="anchor-text" title="' + escHtmlAttr(a.text) + '">' + escHtml(a.text) + '</span>' +
                '<button class="btn btn-sm btn-secondary anchor-nav-btn" data-action="jump" data-id="' + a.id + '">跳转</button>' +
                '<button class="btn btn-sm btn-secondary anchor-nav-btn" data-action="delete" data-id="' + a.id + '" style="color:#ef4444;">✕</button>' +
                '</div>';
        }
        anchorList.innerHTML = html;
    }

    // HTML 属性值转义
    function escHtmlAttr(str) {
        return typeof str === 'string' ? str.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : '';
    }

    // 锚点面板点击委托
    if (anchorList) {
        anchorList.addEventListener('click', function(e) {
            var target = e.target;
            if (target.classList.contains('anchor-nav-btn')) {
                var action = target.dataset.action;
                var id = target.dataset.id;
                if (action === 'jump') {
                    jumpToAnchor(id);
                } else if (action === 'delete') {
                    removeAnchor(id);
                }
                return;
            }
            // 点击锚点条目本身也跳转
            var item = target.closest('.anchor-item');
            if (item && item.dataset.anchorId) {
                jumpToAnchor(item.dataset.anchorId);
            }
        });
    }

    // 在光标所在段落切换锚点（快捷键 F9）
    function toggleAnchorAtCursor() {
        ensureParagraphIds();
        syncGutterHeight();
        var sel = window.getSelection();
        if (!sel.rangeCount) { showToast('请先将光标放在段落中', 'info', 1500); return; }
        var node = sel.anchorNode;
        if (!node) { showToast('请先将光标放在段落中', 'info', 1500); return; }
        // 向上查找最近的段落元素
        var el = node.nodeType === 3 ? node.parentElement : node;
        while (el && el !== editor) {
            var tag = el.tagName ? el.tagName.toLowerCase() : '';
            if (tag === 'p' || (tag.length === 2 && tag[0] === 'h' && tag[1] >= '1' && tag[1] <= '6')) {
                break;
            }
            el = el.parentElement;
        }
        if (!el || el === editor) {
            showToast('未找到可锚定的段落，请将光标放在正文段落中', 'warning', 2000);
            return;
        }
        toggleAnchor(el);
    }

    // 显示/隐藏锚点面板
    function toggleAnchorPanel() {
        if (anchorPanelOpen) {
            hideAnchorPanel();
        } else {
            showAnchorPanel();
        }
    }

    function showAnchorPanel() {
        if (!anchorPanel) return;
        renderAnchorPanel();
        anchorPanel.classList.remove('hidden');
        anchorPanelOpen = true;
        if (anchorPanelBtn) anchorPanelBtn.classList.add('active');
    }

    function hideAnchorPanel() {
        if (!anchorPanel) return;
        anchorPanel.classList.add('hidden');
        anchorPanelOpen = false;
        if (anchorPanelBtn) anchorPanelBtn.classList.remove('active');
    }

    // 装订线点击事件
    if (anchorGutter) {
        anchorGutter.addEventListener('click', function(e) {
            if (e.target.classList.contains('anchor-gutter-dot')) return;
            ensureParagraphIds();
            var y = e.clientY - anchorGutter.getBoundingClientRect().top;
            var el = findParagraphAtY(y, anchorGutter);
            if (el) toggleAnchor(el);
        });
    }

    // 工具栏按钮
    if (anchorPanelBtn) {
        anchorPanelBtn.addEventListener('click', function() {
            toggleAnchorPanel();
        });
    }

    // 面板关闭按钮
    if (anchorPanelClose) {
        anchorPanelClose.addEventListener('click', function() {
            hideAnchorPanel();
        });
    }

    // 上一个锚点按钮
    if (anchorPrev) {
        anchorPrev.addEventListener('click', function() {
            navigateAnchor(-1);
        });
    }

    // 下一个锚点按钮
    if (anchorNext) {
        anchorNext.addEventListener('click', function() {
            navigateAnchor(1);
        });
    }

    // 清除所有锚点按钮
    if (anchorClearAll) {
        anchorClearAll.addEventListener('click', function() {
            if (!anchors.length) { showToast('没有锚点可清除', 'info', 1200); return; }
            if (confirm('确定要清除全部 ' + anchors.length + ' 个锚点吗？此操作不可撤销。')) {
                anchors = [];
                renderAnchorGutter();
                renderAnchorPanel();
                showToast('已清除全部锚点', 'success', 1500);
            }
        });
    }

    // 全局点击：点击锚点面板外部时关闭面板
    document.addEventListener('click', function(e) {
        if (!anchorPanelOpen) return;
        if (!anchorPanel) return;
        var target = e.target;
        // 检查是否点击在面板内或按钮上
        if (anchorPanel.contains(target)) return;
        if (anchorPanelBtn && anchorPanelBtn.contains(target)) return;
        hideAnchorPanel();
    });

    // ===== 📁 折叠装订线系统（右侧） =====
    var foldGutter = document.getElementById('foldGutter');

    // 渲染折叠装订线
    function renderFoldGutter() {
        if (!foldGutter) return;
        foldGutter.querySelectorAll('.fold-gutter-dot, .fold-bracket').forEach(function(d) { d.remove(); });

        var gutterRect = foldGutter.getBoundingClientRect();

        // 渲染折叠标记点
        foldPoints.forEach(function(fp, idx) {
            var el = editor.querySelector('[data-pid="' + fp.pid + '"]');
            if (!el) return;
            if (el.classList.contains('fold-hidden')) return;
            var elRect = el.getBoundingClientRect();
            var topOffset = elRect.top - gutterRect.top + elRect.height / 2 - 4;
            var dot = document.createElement('div');
            dot.className = 'fold-gutter-dot';
            dot.style.top = topOffset + 'px';
            dot.title = '折叠标记 #' + (idx + 1) + '\n' + fp.text.substring(0, 30) + '\n点击移除';
            dot.addEventListener('click', function(e) {
                e.stopPropagation();
                removeFoldPoint(fp.id);
            });
            foldGutter.appendChild(dot);
        });

        // 渲染折叠区域括号
        for (var i = 0; i < foldPoints.length - 1; i++) {
            var startFp = foldPoints[i];
            var endFp = foldPoints[i + 1];
            var startEl = editor.querySelector('[data-pid="' + startFp.pid + '"]');
            var endEl = editor.querySelector('[data-pid="' + endFp.pid + '"]');
            if (!startEl || !endEl) continue;

            var startRect = startEl.getBoundingClientRect();
            var endRect = endEl.getBoundingClientRect();
            var bracketTop = startRect.top - gutterRect.top + startRect.height / 2;
            var bracketHeight = (endRect.top - startRect.top) + endRect.height / 2 - startRect.height / 2;

            var region = findFoldRegion(startFp.pid, endFp.pid);
            var isCollapsed = region && region.isFolded;

            var bracket = document.createElement('div');
            bracket.className = 'fold-bracket' + (isCollapsed ? ' collapsed' : '');
            bracket.style.top = bracketTop + 'px';
            bracket.style.height = Math.max(bracketHeight, 8) + 'px';
            bracket.title = isCollapsed ? '点击展开' : '点击折叠 (#' + (i + 1) + '→#' + (i + 2) + ')';
            bracket.addEventListener('click', function(e) {
                e.stopPropagation();
                toggleFoldRegion(startFp.pid, endFp.pid);
            });
            foldGutter.appendChild(bracket);
        }
    }

    // 查找折叠区域
    function findFoldRegion(startPid, endPid) {
        for (var i = 0; i < foldRegions.length; i++) {
            if (foldRegions[i].startPid === startPid && foldRegions[i].endPid === endPid) {
                return foldRegions[i];
            }
        }
        return null;
    }

    // 添加折叠标记点
    function addFoldPoint(el) {
        ensureParagraphIds();
        var pid = el.dataset.pid;
        if (!pid) return;
        // 检查是否已存在
        for (var i = 0; i < foldPoints.length; i++) {
            if (foldPoints[i].pid === pid) {
                // 已存在则移除
                removeFoldPoint(foldPoints[i].id);
                return;
            }
        }
        var text = el.textContent.trim().substring(0, 30) || '(空段落)';
        foldPoints.push({
            id: 'fp-' + Date.now(),
            pid: pid,
            text: text,
            createdAt: new Date().toISOString()
        });
        // 按文档顺序排序
        sortFoldPoints();
        // 清理失效的折叠区域
        cleanupFoldRegions();
        renderFoldGutter();
        showToast('已添加折叠标记 #' + foldPoints.length, 'success', 1200);
    }

    // 移除折叠标记点
    function removeFoldPoint(fpId) {
        var fp = null;
        for (var i = 0; i < foldPoints.length; i++) {
            if (foldPoints[i].id === fpId) { fp = foldPoints[i]; foldPoints.splice(i, 1); break; }
        }
        if (!fp) return;
        // 移除关联的折叠区域
        foldRegions = foldRegions.filter(function(r) {
            if (r.startPid === fp.pid || r.endPid === fp.pid) {
                if (r.isFolded) unfoldRegion(r);
                return false;
            }
            return true;
        });
        sortFoldPoints();
        renderFoldGutter();
        showToast('折叠标记已移除', 'info', 1200);
    }

    // 按文档中段落顺序排序折叠点
    function sortFoldPoints() {
        foldPoints.sort(function(a, b) {
            var aEl = editor.querySelector('[data-pid="' + a.pid + '"]');
            var bEl = editor.querySelector('[data-pid="' + b.pid + '"]');
            if (!aEl || !bEl) return 0;
            // 使用 compareDocumentPosition 比较 DOM 顺序
            var pos = aEl.compareDocumentPosition(bEl);
            return pos & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1;
        });
    }

    // 清理引用不存在段落的折叠区域
    function cleanupFoldRegions() {
        foldRegions = foldRegions.filter(function(r) {
            var sEl = editor.querySelector('[data-pid="' + r.startPid + '"]');
            var eEl = editor.querySelector('[data-pid="' + r.endPid + '"]');
            if (!sEl || !eEl) {
                if (r.isFolded) unfoldRegion(r);
                return false;
            }
            return true;
        });
    }

    // 切换折叠区域
    function toggleFoldRegion(startPid, endPid) {
        var region = findFoldRegion(startPid, endPid);
        if (region && region.isFolded) {
            unfoldRegion(region);
            showToast('已展开', 'info', 1200);
        } else {
            if (!region) {
                region = {
                    id: 'fr-' + Date.now(),
                    startPid: startPid,
                    endPid: endPid,
                    isFolded: false
                };
                foldRegions.push(region);
            }
            foldRegion(region);
            showToast('已折叠', 'success', 1200);
        }
        renderFoldGutter();
    }

    // 折叠区域
    function foldRegion(region) {
        var startEl = editor.querySelector('[data-pid="' + region.startPid + '"]');
        var endEl = editor.querySelector('[data-pid="' + region.endPid + '"]');
        if (!startEl || !endEl) return;

        // 收集两个标记点之间的所有段落
        var between = getParagraphsBetween(startEl, endEl);
        if (between.length === 0) {
            showToast('两个标记点之间没有可折叠内容', 'warning', 2000);
            return;
        }

        // 隐藏中间段落
        between.forEach(function(el) {
            el.classList.add('fold-hidden');
        });

        // 插入占位符
        var placeholder = createFoldPlaceholder(region, between.length);
        startEl.parentNode.insertBefore(placeholder, between[0]);

        region.placeholderEl = placeholder;
        region.foldedCount = between.length;
        region.foldedElements = between;
        region.isFolded = true;
    }

    // 展开区域
    function unfoldRegion(region) {
        // 移除占位符
        if (region.placeholderEl && region.placeholderEl.parentNode) {
            region.placeholderEl.parentNode.removeChild(region.placeholderEl);
        }
        // 恢复显示
        if (region.foldedElements) {
            region.foldedElements.forEach(function(el) {
                el.classList.remove('fold-hidden');
            });
        }
        region.isFolded = false;
        region.placeholderEl = null;
        region.foldedElements = null;
        region.foldedCount = 0;
    }

    // 获取两个元素之间的所有段落（不含两个端点本身）
    function getParagraphsBetween(startEl, endEl) {
        var result = [];
        var el = startEl.nextElementSibling;
        while (el && el !== endEl) {
            var tag = el.tagName ? el.tagName.toLowerCase() : '';
            if (tag === 'p' || (tag.length === 2 && tag[0] === 'h' && tag[1] >= '1' && tag[1] <= '6')) {
                result.push(el);
            } else {
                // 如果不是段落元素，也收集其中的段落子元素
                var inner = el.querySelectorAll ? el.querySelectorAll('p, h1, h2, h3, h4, h5, h6') : [];
                for (var j = 0; j < inner.length; j++) result.push(inner[j]);
            }
            el = el.nextElementSibling;
        }
        return result;
    }

    // 创建折叠占位符
    function createFoldPlaceholder(region, count) {
        var div = document.createElement('div');
        div.className = 'fold-placeholder';
        div.contentEditable = 'false';
        div.setAttribute('data-fold-region', region.id);
        div.innerHTML = '<span class="fold-icon">📁</span>' +
            '<span class="fold-text">已折叠 ' + count + ' 段内容</span>' +
            '<span class="fold-hint">点击展开</span>';
        div.addEventListener('click', function() {
            unfoldRegion(region);
            renderFoldGutter();
            showToast('已展开', 'info', 1200);
        });
        return div;
    }

    // 折叠装订线点击事件
    if (foldGutter) {
        foldGutter.addEventListener('click', function(e) {
            if (e.target.classList.contains('fold-gutter-dot') || e.target.classList.contains('fold-bracket')) return;
            ensureParagraphIds();
            var y = e.clientY - foldGutter.getBoundingClientRect().top;
            var el = findParagraphAtY(y, foldGutter);
            if (el) addFoldPoint(el);
        });
    }

    // 重新应用所有已折叠的区域（页面加载/切换标签后调用）
    function reapplyFoldRegions() {
        var regionsToApply = foldRegions.filter(function(r) { return r.isFolded; });
        // 重置状态（foldRegion 会重新设置 isFolded）
        regionsToApply.forEach(function(r) {
            r.isFolded = false;
            foldRegion(r);
        });
    }

    // ===== 📁 折叠面板 =====
    var foldPanel = document.getElementById('foldPanel');
    var foldPanelBtn = $('foldPanelBtn');
    var foldPanelClose = $('foldPanelClose');
    var foldList = $('foldList');
    var foldExpandAll = $('foldExpandAll');
    var foldCollapseAll = $('foldCollapseAll');
    var foldClearAll = $('foldClearAll');
    var foldPanelOpen = false;

    // 渲染折叠面板
    function renderFoldPanel() {
        if (!foldList) return;
        if (!foldPoints.length) {
            foldList.innerHTML = '<div class="anchor-empty">暂无折叠标记<br>点击编辑器右侧装订线添加</div>';
            return;
        }
        var html = '';
        for (var i = 0; i < foldPoints.length; i++) {
            var fp = foldPoints[i];
            var fpEl = editor.querySelector('[data-pid="' + fp.pid + '"]');
            var isHidden = fpEl && fpEl.classList.contains('fold-hidden');
            html += '<div class="anchor-item" data-fold-id="' + fp.id + '" style="' + (isHidden ? 'opacity:0.45;' : '') + '">' +
                '<span class="anchor-dot" style="color:#6366f1;">■</span>' +
                '<span class="anchor-index">#' + (i + 1) + '</span>' +
                '<span class="anchor-text" title="' + escHtmlAttr(fp.text) + '">' + escHtml(fp.text) + '</span>' +
                '<button class="btn btn-sm btn-secondary anchor-nav-btn" data-action="fold-jump" data-id="' + fp.id + '">跳转</button>' +
                '<button class="btn btn-sm btn-secondary anchor-nav-btn" data-action="fold-delete" data-id="' + fp.id + '" style="color:#ef4444;">✕</button>' +
                '</div>';

            // 如果后面还有折叠点，显示区域信息
            if (i < foldPoints.length - 1) {
                var nextFp = foldPoints[i + 1];
                var region = findFoldRegion(fp.pid, nextFp.pid);
                var isCollapsed = region && region.isFolded;
                var betweenCount = getFoldBetweenCount(fp.pid, nextFp.pid);
                html += '<div class="anchor-item" style="background:rgba(99,102,241,0.06);font-size:11px;padding:3px 12px;">' +
                    '<span style="color:#6366f1;">' + (isCollapsed ? '📁 已折叠' : '┆ 可折叠') + '</span>' +
                    '<span style="flex:1;color:var(--text-secondary);margin-left:6px;">#' + (i + 1) + ' → #' + (i + 2) + ' (' + betweenCount + '段)</span>' +
                    '<button class="btn btn-sm ' + (isCollapsed ? 'btn-primary' : 'btn-secondary') + ' anchor-nav-btn" data-action="fold-toggle" data-start="' + fp.pid + '" data-end="' + nextFp.pid + '">' + (isCollapsed ? '展开' : '折叠') + '</button>' +
                    '</div>';
            }
        }
        foldList.innerHTML = html;
    }

    // 获取两个标记点之间的段落数
    function getFoldBetweenCount(startPid, endPid) {
        var startEl = editor.querySelector('[data-pid="' + startPid + '"]');
        var endEl = editor.querySelector('[data-pid="' + endPid + '"]');
        if (!startEl || !endEl) return 0;
        return getParagraphsBetween(startEl, endEl).length;
    }

    // 折叠面板点击委托
    if (foldList) {
        foldList.addEventListener('click', function(e) {
            var btn = e.target.closest('.anchor-nav-btn');
            if (!btn) {
                // 点击条目本身跳转
                var item = e.target.closest('[data-fold-id]');
                if (item) jumpToFoldPoint(item.dataset.foldId);
                return;
            }
            var action = btn.dataset.action;
            if (action === 'fold-jump') {
                jumpToFoldPoint(btn.dataset.id);
            } else if (action === 'fold-delete') {
                removeFoldPoint(btn.dataset.id);
                renderFoldPanel();
            } else if (action === 'fold-toggle') {
                toggleFoldRegion(btn.dataset.start, btn.dataset.end);
                renderFoldGutter();
                renderFoldPanel();
                renderAnchorGutter();
            }
        });
    }

    // 跳转到折叠标记点
    function jumpToFoldPoint(fpId) {
        var fp = null;
        for (var i = 0; i < foldPoints.length; i++) {
            if (foldPoints[i].id === fpId) { fp = foldPoints[i]; break; }
        }
        if (!fp) return;
        var el = editor.querySelector('[data-pid="' + fp.pid + '"]');
        if (!el) return;
        // 如果在折叠区域内，先展开
        if (el.classList.contains('fold-hidden')) {
            // 找到包含此段落的折叠区域并展开
            for (var j = 0; j < foldRegions.length; j++) {
                var r = foldRegions[j];
                if (r.isFolded && r.foldedElements) {
                    for (var k = 0; k < r.foldedElements.length; k++) {
                        if (r.foldedElements[k] === el) {
                            unfoldRegion(r);
                            renderFoldGutter();
                            renderFoldPanel();
                            renderAnchorGutter();
                            break;
                        }
                    }
                }
            }
        }
        scrollEditorTo(el, 'center');
        el.classList.remove('anchor-flash');
        void el.offsetWidth;
        el.classList.add('anchor-flash');
    }

    // 折叠面板显示/隐藏
    function toggleFoldPanel() {
        if (foldPanelOpen) { hideFoldPanel(); } else { showFoldPanel(); }
    }

    function showFoldPanel() {
        if (!foldPanel) return;
        renderFoldPanel();
        foldPanel.classList.remove('hidden');
        foldPanelOpen = true;
        if (foldPanelBtn) foldPanelBtn.classList.add('active');
    }

    function hideFoldPanel() {
        if (!foldPanel) return;
        foldPanel.classList.add('hidden');
        foldPanelOpen = false;
        if (foldPanelBtn) foldPanelBtn.classList.remove('active');
    }

    // 折叠面板按钮
    if (foldPanelBtn) foldPanelBtn.addEventListener('click', toggleFoldPanel);
    if (foldPanelClose) foldPanelClose.addEventListener('click', hideFoldPanel);

    // 全部展开
    if (foldExpandAll) {
        foldExpandAll.addEventListener('click', function() {
            if (!foldRegions.length) { showToast('没有折叠区域', 'info', 1200); return; }
            var unfolded = 0;
            foldRegions.forEach(function(r) { if (r.isFolded) { unfoldRegion(r); unfolded++; } });
            if (unfolded) {
                renderFoldGutter();
                renderFoldPanel();
                renderAnchorGutter();
                showToast('已展开 ' + unfolded + ' 个区域', 'success', 1500);
            } else {
                showToast('所有区域已展开', 'info', 1200);
            }
        });
    }

    // 全部折叠
    if (foldCollapseAll) {
        foldCollapseAll.addEventListener('click', function() {
            if (foldPoints.length < 2) { showToast('至少需要 2 个折叠标记点', 'info', 1500); return; }
            var folded = 0;
            for (var i = 0; i < foldPoints.length - 1; i++) {
                var region = findFoldRegion(foldPoints[i].pid, foldPoints[i + 1].pid);
                if (!region) {
                    region = { id: 'fr-' + Date.now() + '-' + i, startPid: foldPoints[i].pid, endPid: foldPoints[i + 1].pid, isFolded: false };
                    foldRegions.push(region);
                }
                if (!region.isFolded) { foldRegion(region); folded++; }
            }
            if (folded) {
                renderFoldGutter();
                renderFoldPanel();
                renderAnchorGutter();
                showToast('已折叠 ' + folded + ' 个区域', 'success', 1500);
            } else {
                showToast('所有区域已折叠', 'info', 1200);
            }
        });
    }

    // 清除全部折叠
    if (foldClearAll) {
        foldClearAll.addEventListener('click', function() {
            if (!foldPoints.length) { showToast('没有折叠标记可清除', 'info', 1200); return; }
            if (confirm('确定要清除全部 ' + foldPoints.length + ' 个折叠标记吗？')) {
                // 先展开所有区域
                foldRegions.forEach(function(r) { if (r.isFolded) unfoldRegion(r); });
                foldPoints = [];
                foldRegions = [];
                renderFoldGutter();
                renderFoldPanel();
                renderAnchorGutter();
                showToast('已清除全部折叠标记', 'success', 1500);
            }
        });
    }

    // 折叠面板外点击关闭
    document.addEventListener('click', function(e) {
        if (!foldPanelOpen || !foldPanel) return;
        if (foldPanel.contains(e.target)) return;
        if (foldPanelBtn && foldPanelBtn.contains(e.target)) return;
        hideFoldPanel();
    });

    // 检测锚点是否在折叠区域内（供 jumpToAnchor 使用）
    function findFoldRegionContaining(pid) {
        for (var i = 0; i < foldRegions.length; i++) {
            var r = foldRegions[i];
            if (!r.isFolded || !r.foldedElements) continue;
            for (var j = 0; j < r.foldedElements.length; j++) {
                if (r.foldedElements[j].dataset.pid === pid) return { region: r, index: i };
            }
        }
        return null;
    }

    // 与锚点系统共享的辅助函数已在前面定义（findParagraphAtY, ensureParagraphIds）

    // ===== 文章导航快捷键 =====
    document.addEventListener('keydown', function(e) {
        if (!e.ctrlKey && !e.metaKey) return;
        var inEditor = document.activeElement && editor && editor.contains(document.activeElement);
        if (!inEditor) return;
        var container = editorContainer;
        if (!container) return;
        if (e.key === 'Home') {
            e.preventDefault();
            container.scrollTop = 0;
        } else if (e.key === 'End') {
            e.preventDefault();
            container.scrollTop = container.scrollHeight;
        }
    });

    // ===== 自定义离开确认弹窗 =====
    // 显示自定义确认弹窗，返回 Promise（用户选择留在页面则 resolve(false)，离开则 resolve(true)）
    function showLeaveConfirmModal(action) {
        // 先完整保存当前状态
        saveConfigToStorage();
        syncStateToSession();
        performEmergencyBackup();
        saveDocumentFull(); // 🔄 使用完整保存确保撤销历史不丢失

        _leaveAction = action;

        var overlay = document.getElementById('leaveConfirmOverlay');
        var messageEl = document.getElementById('leaveConfirmMessage');
        if (!overlay) {
            // 兜底：如果弹窗 HTML 未加载，直接执行离开操作
            doLeave();
            return;
        }

        // 根据操作类型设置提示文字
        var messages = {
            'reload': '您确定要刷新页面吗？<br>建议先导出或保存文档，避免数据丢失。',
            'hardReload': '您确定要强制刷新页面吗？<br>这将清除缓存并重新加载，未保存的更改可能丢失。',
            'close': '您确定要关闭此页面吗？<br>建议先导出或保存文档，避免数据丢失。'
        };
        if (messageEl) {
            messageEl.innerHTML = messages[action] || messages['reload'];
        }

        overlay.classList.remove('hidden');

        // 聚焦"留在页面"按钮（默认安全选项）
        var stayBtn = document.getElementById('leaveConfirmStay');
        if (stayBtn) { setTimeout(function() { stayBtn.focus(); }, 50); }
    }

    // 用户选择离开：设置标志位 + 执行对应操作
    function doLeave() {
        _allowNavigation = true;
        var overlay = document.getElementById('leaveConfirmOverlay');
        if (overlay) { overlay.classList.add('hidden'); }

        if (_leaveAction === 'reload') {
            window.location.reload();
        } else if (_leaveAction === 'hardReload') {
            // 强制刷新（跳过缓存）
            window.location.reload(true);
        } else if (_leaveAction === 'close') {
            // 尝试关闭窗口
            window.close();
            // 如果 window.close() 无效（非脚本打开的窗口），提示用户手动关闭
            setTimeout(function() {
                // 如果 200ms 后页面还在，说明 close() 无效
                showToast('请手动关闭标签页（Ctrl+W 或点击关闭按钮）', 'info', 5000);
            }, 200);
        }
    }

    // 重置离开标志（用户取消离开后，在编辑时重新启用拦截）
    function resetLeaveFlag() {
        if (_allowNavigation) {
            var overlay = document.getElementById('leaveConfirmOverlay');
            if (overlay && !overlay.classList.contains('hidden')) {
                // 弹窗仍显示中，不重置
                return;
            }
        }
        _allowNavigation = false;
        _leaveAction = null;
    }

    // 紧急备份（同步写入 localStorage）
    function performEmergencyBackup() {
        try {
            var allData = {
                sessions: tabManager.sessions.map(function(s) {
                    return {
                        id: s.id,
                        title: s.title,
                        customTitle: s.customTitle,
                        html: s.html,
                        imageDataMapEntries: s.imageDataMap ? Array.from(s.imageDataMap.entries()) : [],
                        undoHistory: (s.undoHistory || []).slice(),
                        undoIndex: s.undoIndex,
                        scrollTop: s.scrollTop,
                        sourceFileName: s.sourceFileName,
                        sourceImportTime: s.sourceImportTime
                    };
                }),
                activeIndex: tabManager.activeIndex,
                tabIdCounter: tabManager.tabIdCounter
            };
            localStorage.setItem('docx-editor-tabs-backup', JSON.stringify(allData));
            localStorage.setItem('docx-editor-backup-time', new Date().toISOString());
        } catch(ex) { /* ignore */ }
    }

    // 绑定离开确认弹窗按钮事件
    function bindLeaveConfirmButtons() {
        var overlay = document.getElementById('leaveConfirmOverlay');
        if (!overlay) { return; }

        var stayBtn = document.getElementById('leaveConfirmStay');
        var goBtn = document.getElementById('leaveConfirmGo');

        if (stayBtn) {
            stayBtn.addEventListener('click', function() {
                overlay.classList.add('hidden');
                _leaveAction = null;
                showToast('已取消离开，继续编辑', 'info', 2000);
            });
        }

        if (goBtn) {
            goBtn.addEventListener('click', function() {
                doLeave();
            });
        }

        // 点击遮罩关闭 → 留在页面
        overlay.addEventListener('click', function(e) {
            if (e.target === overlay) {
                overlay.classList.add('hidden');
                _leaveAction = null;
            }
        });

        // Escape 键关闭弹窗 → 留在页面
        overlay.addEventListener('keydown', function(e) {
            if (e.key === 'Escape') {
                overlay.classList.add('hidden');
                _leaveAction = null;
            }
        });
    }

    // DOM 加载完成后绑定弹窗按钮
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bindLeaveConfirmButtons);
    } else {
        bindLeaveConfirmButtons();
    }

    // beforeunload：保存全部标签 + 强制询问确认
    window.addEventListener('beforeunload', function(e) {
        // 如果用户已通过自定义弹窗确认离开，直接放行不拦截
        if (_allowNavigation) return;

        saveConfigToStorage();
        syncStateToSession();
        performEmergencyBackup();

        // 始终触发浏览器原生离开确认对话框（不判断内容是否为空）
        // 这是最后的安全网：拦截浏览器 UI 操作（关闭按钮、刷新按钮、地址栏导航等）
        e.preventDefault();
        e.returnValue = '';  // 现代浏览器显示通用提示，旧浏览器显示此文字
    });
    editorContainer.addEventListener('scroll', throttle(highlightVisibleHeading, 200));
    document.addEventListener('keydown', handleKeyboard);

    // 启动：加载配置 → 初始化标签管理器 → 恢复文档
    loadConfigFromProjectDir().then(function(fromFile) {
        if (!fromFile) {
            var fromStorage = loadConfigFromStorage();
            if (!fromStorage) { readHeadingConfig(); applyHeadingStylesToEditor(); }
        }
        // 初始化标签管理器（从 IndexedDB 恢复或创建默认文档）
        return tabManager.init().then(function() {
            var active = tabManager.getActive();
            if (active && active.html && !isPlaceholderContent(active.html)) {
                // 检查 localStorage 兜底备份（IndexedDB 可能被清除）
                applyHeadingStylesToEditor();
                if (typeof applyBodyFormatFn === 'function') applyBodyFormatFn();
                generateTOC();
                renumber();
                ensureParagraphIds();
                renderAnchorGutter();
                renderFoldGutter();
                reapplyFoldRegions();
                if (imgCenterToggle && imgCenterToggle.checked) centerAllImages();
                if (typeof applyAllCodeThemes === 'function') applyAllCodeThemes();
                // 记录初始快照
                saveUndoState('恢复文档');
                var savedAt = active.savedAt ? (' (' + new Date(active.savedAt).toLocaleString() + ')') : '';
                setStatus('已恢复上次编辑的文档' + savedAt);
                // 渲染标签
                tabManager.renderTabs();
            } else {
                // 尝试 localStorage 兜底
                var backup = localStorage.getItem('docx-editor-tabs-backup');
                if (backup) {
                    try {
                        var backupData = JSON.parse(backup);
                        if (backupData.sessions && backupData.sessions.length > 0) {
                            backupData.sessions.forEach(function(sData, i) {
                                if (i < tabManager.sessions.length) {
                                    var s = tabManager.sessions[i];
                                    s.html = sData.html || '';
                                    s.title = sData.title || s.title;
                                    s.customTitle = sData.customTitle || false;
                                    s.imageDataMap = new Map(sData.imageDataMapEntries || []);
                                    s.undoHistory = (sData.undoHistory || []).slice();
                                    s.undoIndex = sData.undoIndex != null ? sData.undoIndex : -1;
                                    s.anchors = (sData.anchors || []).slice();
                                    s.foldPoints = (sData.foldPoints || []).slice();
                                    s.foldRegions = (sData.foldRegions || []).slice();
                                    s.scrollTop = sData.scrollTop || 0;
                                    s.sourceFileName = sData.sourceFileName || null;
                                    s.sourceImportTime = sData.sourceImportTime || null;
                                }
                            });
                            tabManager.activeIndex = backupData.activeIndex || 0;
                            var activeS = tabManager.getActive();
                            if (activeS) {
                                syncStateFromSession(activeS);
                                suppressCaptureUntil = Date.now() + 500;
                                editor.innerHTML = activeS.html || '';
                                updateStats();
                                setStatus('已从本地备份恢复文档');
                            }
                            saveUndoState('恢复备份');
                            tabManager.renderTabs();
                            tabManager.saveAllToDB();
                        }
                    } catch(ex) { /* ignore */ }
                }
                if (!tabManager.getActive() || isPlaceholderContent(editor.innerHTML)) {
                    setStatus('就绪 — 点击"导入 DOCX"加载文档，或直接在此编辑');
                }
            }
        });
    }).catch(function() {
        setStatus('就绪 — 点击"导入 DOCX"加载文档，或直接在此编辑');
    });

    // ===== 🐾 写作陪伴宠物系统（三只可选：猫/狗/兔） =====
    var petPanel = document.getElementById('petPanel');
    var petBubble = document.getElementById('petBubble');
    var petBubbleText = document.getElementById('petBubbleText');
    var petEmoji = document.getElementById('petEmoji');
    var petState = 'idle';
    var petIdleTimer = null;
    var petBubbleTimer = null;

    // 默认宠物档案
    var DEFAULT_PET_PROFILES = {
        cat: {
            type: 'cat', emoji: '🐱', name: '小橘', nickname: '橘子',
            gender: '♂ 男生', age: '2岁', birthday: '2024-03-15',
            height: '25cm', weight: '4.5kg', education: '喵星小学毕业',
            hobby: '晒太阳、抓玩具老鼠、陪主人写作', personality: '活泼好动、粘人、有点小傲娇'
        },
        dog: {
            type: 'dog', emoji: '🐶', name: '旺财', nickname: '旺旺',
            gender: '♂ 男生', age: '3岁', birthday: '2023-08-08',
            height: '45cm', weight: '12kg', education: '汪汪训练营优秀学员',
            hobby: '捡球、跑步、守护主人', personality: '忠诚勇敢、热情开朗、有点憨厚'
        },
        bunny: {
            type: 'bunny', emoji: '🐰', name: '雪球', nickname: '球球',
            gender: '♀ 女生', age: '1岁', birthday: '2025-01-20',
            height: '18cm', weight: '1.8kg', education: '胡萝卜大学在读',
            hobby: '跳跳、啃胡萝卜、卖萌', personality: '温柔可爱、安静优雅、有点胆小'
        }
    };

    // 每只宠物的消息集
    var PET_MESSAGES = {
        cat: [
            '加油～✊', '好棒！🌟', '继续写～📝', '你真厉害！✨',
            '喵～🐱', '慢慢来～☕', '写得不错！👍', '休息一下？🍵',
            '专注的样子好帅！😊', '今天也要加油！🌸', '我在陪你哦～💕',
            '好无聊… 写点东西吧～', '这个文档会变很棒的！🎉', '你好！👋'
        ],
        dog: [
            '加油！汪～🐶', '好厉害！🦴', '继续冲！💪', '你最棒！⭐',
            '汪～！主人加油！', '我在守护你哦～🛡️', '写得真好！🎾', '休息一下去散步？🌳',
            '忠诚陪伴中…💖', '每天都要元气满满！☀️', '汪汪！好开心～',
            '有点困了…但会陪着你的！😴', '你的文档超棒！🏆', '嘿嘿～'
        ],
        bunny: [
            '蹦蹦跳跳～🐰', '加油呀～🌸', '好温柔的字…✨', '你真细心！🎀',
            '咕…好安静呢～', '给你胡萝卜！🥕', '写得真好呢～💕', '要休息一下吗？🍰',
            '安静地陪伴你…☁️', '今天也很棒哦～🌷', '蹭蹭～',
            '在发呆吗？嘻嘻～', '这个文档一定会闪闪发光！💎', '你好呀～'
        ],
        sleeping: {
            cat: ['zzz… 写完了叫我…😴', '好困… 你继续…💤', 'Zzz… 我在做梦写代码…'],
            dog: ['呼呼… 梦里在追球…🎾💤', 'Zzz… 好香…是肉的味道…🍖', '呼噜噜…守护中…💤'],
            bunny: ['zzz… 胡萝卜田…🥕💤', '呼… 软软的云…☁️', '梦里也在跳跳…✨💤']
        }
    };

    // 运行时状态
    var currentPetType = 'cat';
    var customPetProfiles = {};
    var petVisible = true;

    // 加载/保存宠物数据
    function loadPetData() {
        try {
            var saved = localStorage.getItem('docx-pet-profiles');
            if (saved) customPetProfiles = JSON.parse(saved);
            var t = localStorage.getItem('docx-pet-active');
            if (t && DEFAULT_PET_PROFILES[t]) currentPetType = t;
        } catch(e) {}
    }
    function savePetData() {
        try {
            localStorage.setItem('docx-pet-profiles', JSON.stringify(customPetProfiles));
            localStorage.setItem('docx-pet-active', currentPetType);
        } catch(e) {}
    }
    function getPetProfile(type) {
        var t = type || currentPetType;
        var base = DEFAULT_PET_PROFILES[t];
        var over = customPetProfiles[t] || {};
        var merged = {};
        var keys = Object.keys(base);
        for (var i = 0; i < keys.length; i++) {
            var k = keys[i];
            merged[k] = (over[k] !== undefined && over[k] !== '') ? over[k] : base[k];
        }
        return merged;
    }
    function getActivePetChar() {
        return petPanel.querySelector('.pet-character.active');
    }

    // 更新侧边栏摘要
    function updatePetInfoSummary() {
        var el = document.getElementById('petInfoSummary');
        if (!el) return;
        var p = getPetProfile();
        el.textContent = p.emoji + ' ' + p.name + ' · ' + p.gender.charAt(0) + ' · ' + p.age + ' · ' + p.weight;
    }

    // 切换宠物
    function switchPet(type) {
        if (type === currentPetType) return;
        // 重置状态
        petSetState('idle');
        if (petIdleTimer) { clearTimeout(petIdleTimer); petIdleTimer = null; }
        currentPetType = type;
        // CSS 切换
        var allChars = petPanel.querySelectorAll('.pet-character');
        for (var i = 0; i < allChars.length; i++) {
            allChars[i].classList.toggle('active', allChars[i].getAttribute('data-pet') === type);
        }
        // 选择器按钮
        var allBtns = document.querySelectorAll('.pet-select-btn');
        for (var j = 0; j < allBtns.length; j++) {
            allBtns[j].classList.toggle('active', allBtns[j].getAttribute('data-pet') === type);
        }
        // 更新面板
        updatePetInfoSummary();
        updatePetToggleBtn();
        savePetData();
        // 打招呼
        var p = getPetProfile();
        petShowBubble('你好！我是' + p.name + '～' + p.emoji, 2500);
    }

    // 更新隐藏/显示按钮文字
    function updatePetToggleBtn() {
        var btn = document.getElementById('petToggleBtn');
        if (!btn) return;
        if (!petVisible) {
            btn.textContent = '🐾 显示';
        } else {
            btn.textContent = '🙈 隐藏';
        }
    }

    // 拖拽
    (function initPetDrag() {
        var dragging = false, dx = 0, dy = 0;
        petPanel.addEventListener('mousedown', function(e) {
            if (e.target.closest('.pet-bubble') || e.target.closest('.pet-top-btns')) return;
            dragging = true;
            var rect = petPanel.getBoundingClientRect();
            dx = e.clientX - rect.left;
            dy = e.clientY - rect.top;
            petPanel.style.cursor = 'grabbing';
            e.preventDefault();
        });
        document.addEventListener('mousemove', function(e) {
            if (!dragging) return;
            petPanel.style.left = (e.clientX - dx) + 'px';
            petPanel.style.top = (e.clientY - dy) + 'px';
            petPanel.style.bottom = 'auto';
            petPanel.style.right = 'auto';
        });
        document.addEventListener('mouseup', function() {
            if (dragging) { dragging = false; petPanel.style.cursor = 'grab'; }
        });
    })();

    function petShowBubble(text, duration) {
        if (petBubbleTimer) { clearTimeout(petBubbleTimer); petBubbleTimer = null; }
        petBubbleText.textContent = text;
        petBubble.classList.remove('hidden');
        petBubbleTimer = setTimeout(function() {
            petBubble.classList.add('hidden');
            petBubbleTimer = null;
        }, duration || 3000);
    }

    function petBurstEmoji(emojis) {
        var items = emojis || ['❤️', '💕', '🌟'];
        petEmoji.innerHTML = '';
        for (var i = 0; i < 6; i++) {
            var el = document.createElement('span');
            el.className = 'burst-item';
            el.textContent = items[Math.floor(Math.random() * items.length)];
            var angle = Math.random() * Math.PI * 2;
            var dist = 30 + Math.random() * 40;
            el.style.setProperty('--bx', Math.cos(angle) * dist + 'px');
            el.style.setProperty('--by', Math.sin(angle) * dist - 20 + 'px');
            el.style.left = (40 + (i % 3 - 1) * 20) + 'px';
            el.style.top = (40) + 'px';
            petEmoji.appendChild(el);
        }
        setTimeout(function() { petEmoji.innerHTML = ''; }, 900);
    }

    function petSetState(state) {
        var active = getActivePetChar();
        if (active) {
            active.classList.remove('idle', 'typing', 'sleep', 'happy');
            active.classList.add(state);
        }
        petState = state;
    }

    function petOnInput() {
        if (petState === 'sleep') {
            petSetState('idle');
            petShowBubble('啊！你醒啦？😊', 2000);
        }
        petSetState('typing');
        if (petIdleTimer) clearTimeout(petIdleTimer);
        petIdleTimer = setTimeout(function() {
            petSetState('idle');
            petIdleTimer = null;
            petIdleTimer = setTimeout(function() {
                petSetState('sleep');
                var sleepMsgs = PET_MESSAGES.sleeping[currentPetType] || PET_MESSAGES.sleeping.cat;
                petShowBubble(sleepMsgs[Math.floor(Math.random() * sleepMsgs.length)], 4000);
                petIdleTimer = null;
            }, 60000);
        }, 3000);
    }

    function petOnClick(e) {
        if (e.target.closest('.pet-bubble') || e.target.closest('.pet-top-btns')) return;
        petSetState('happy');
        petBurstEmoji(['❤️', '💕', '✨', '🌟', '💗']);
        var msgs = PET_MESSAGES[currentPetType] || PET_MESSAGES.cat;
        var msg = msgs[Math.floor(Math.random() * msgs.length)];
        petShowBubble(msg, 2500);
        if (petIdleTimer) { clearTimeout(petIdleTimer); petIdleTimer = null; }
        setTimeout(function() {
            petSetState('idle');
            petIdleTimer = setTimeout(function() {
                petSetState('sleep');
                var sleepMsgs = PET_MESSAGES.sleeping[currentPetType] || PET_MESSAGES.sleeping.cat;
                petShowBubble(sleepMsgs[Math.floor(Math.random() * sleepMsgs.length)], 4000);
                petIdleTimer = null;
            }, 60000);
        }, 500);
    }

    // 绑定交互事件（使用事件委托，因为宠物角色会切换）
    petPanel.addEventListener('click', function(e) {
        if (e.target.closest('.pet-character') && !e.target.closest('.pet-top-btns')) {
            petOnClick(e);
        }
    });
    editor.addEventListener('input', petOnInput);

    // 随机气泡（每 2 分钟检查一次）
    function petRandomBubble() {
        if (petState === 'sleep') return;
        if (petBubbleTimer) return;
        var msgs = PET_MESSAGES[currentPetType] || PET_MESSAGES.cat;
        var msg = msgs[Math.floor(Math.random() * msgs.length)];
        petShowBubble(msg, 2500);
    }
    setInterval(function() {
        if (petState !== 'sleep' && Math.random() < 0.4) {
            petRandomBubble();
        }
    }, 120000);

    // ---- 侧边栏按钮 ----
    var petToggleBtn = document.getElementById('petToggleBtn');
    var petPetBtn = document.getElementById('petPetBtn');
    var petProfileBtn = document.getElementById('petProfileBtn');
    var petSwitchBtn = document.getElementById('petSwitchBtn');
    var petInfoBtn = document.getElementById('petInfoBtn');

    if (petToggleBtn) {
        petToggleBtn.addEventListener('click', function() {
            petVisible = !petVisible;
            petPanel.classList.toggle('hidden', !petVisible);
            updatePetToggleBtn();
        });
    }
    if (petPetBtn) {
        petPetBtn.addEventListener('click', function() {
            if (!petVisible) {
                petVisible = true;
                petPanel.classList.remove('hidden');
                updatePetToggleBtn();
            }
            petOnClick({ target: getActivePetChar() });
        });
    }

    // 宠物选择器按钮
    var petSelectBtns = document.querySelectorAll('.pet-select-btn');
    for (var si = 0; si < petSelectBtns.length; si++) {
        petSelectBtns[si].addEventListener('click', function() {
            var type = this.getAttribute('data-pet');
            if (type) switchPet(type);
        });
    }

    // 宠物面板上的切换按钮
    if (petSwitchBtn) {
        petSwitchBtn.addEventListener('click', function() {
            var types = ['cat', 'dog', 'bunny'];
            var idx = types.indexOf(currentPetType);
            var next = types[(idx + 1) % 3];
            switchPet(next);
        });
    }

    // 宠物面板上的信息按钮 → 打开档案编辑
    if (petInfoBtn) {
        petInfoBtn.addEventListener('click', function() { openPetProfileEditor(); });
    }

    // ---- 宠物档案编辑模态框 ----
    function openPetProfileEditor() {
        try {
            var p = getPetProfile();
            setVal('petProfileAvatar', p.emoji, 'textContent');
            setVal('petProfileTitle', '✏️ 编辑 ' + p.name + ' 的档案', 'textContent');
            setVal('petFieldName', p.name, 'value');
            setVal('petFieldNickname', p.nickname, 'value');
            setVal('petFieldGender', p.gender, 'value');
            setVal('petFieldAge', p.age, 'value');
            setVal('petFieldBirthday', p.birthday, 'value');
            setVal('petFieldHeight', p.height, 'value');
            setVal('petFieldWeight', p.weight, 'value');
            setVal('petFieldEducation', p.education, 'value');
            setVal('petFieldHobby', p.hobby, 'value');
            setVal('petFieldPersonality', p.personality, 'value');
            var overlay = document.getElementById('petProfileOverlay');
            if (overlay) overlay.classList.remove('hidden');
        } catch(e) {
            console.error('打开宠物档案编辑失败:', e);
            showToast('打开档案编辑失败，请刷新页面后重试', 'error');
        }
    }
    function setVal(id, val, prop) {
        var el = document.getElementById(id);
        if (el) el[prop || 'value'] = val;
    }

    function closePetProfileEditor() {
        var overlay = document.getElementById('petProfileOverlay');
        if (overlay) overlay.classList.add('hidden');
    }

    function savePetProfile() {
        try {
            var p = getPetProfile();
            var data = {
                name: getVal('petFieldName') || p.name,
                nickname: getVal('petFieldNickname') || p.nickname,
                gender: getVal('petFieldGender') || p.gender,
                age: getVal('petFieldAge') || p.age,
                birthday: getVal('petFieldBirthday') || p.birthday,
                height: getVal('petFieldHeight') || p.height,
                weight: getVal('petFieldWeight') || p.weight,
                education: getVal('petFieldEducation') || p.education,
                hobby: getVal('petFieldHobby') || p.hobby,
                personality: getVal('petFieldPersonality') || p.personality
            };
            customPetProfiles[currentPetType] = data;
            savePetData();
            updatePetInfoSummary();
            closePetProfileEditor();
            petShowBubble('档案已更新！我是' + data.name + '～' + getPetProfile().emoji, 2000);
            showToast(data.name + ' 的档案已保存', 'success');
        } catch(e) {
            console.error('保存宠物档案失败:', e);
            showToast('保存失败，请重试', 'error');
        }
    }
    function getVal(id) {
        var el = document.getElementById(id);
        return el ? el.value.trim() : '';
    }

    function resetPetProfile() {
        try {
            var def = DEFAULT_PET_PROFILES[currentPetType];
            if (!confirm('确定要恢复 ' + def.name + ' 的默认档案吗？')) return;
            delete customPetProfiles[currentPetType];
            savePetData();
            updatePetInfoSummary();
            openPetProfileEditor();
            showToast('已恢复 ' + def.name + ' 的默认档案', 'info');
        } catch(e) {
            console.error('重置宠物档案失败:', e);
            showToast('重置失败，请重试', 'error');
        }
    }

    // 档案编辑按钮（侧边栏）
    if (petProfileBtn) {
        petProfileBtn.addEventListener('click', function() { openPetProfileEditor(); });
    }

    // 模态框关闭/保存/重置
    var petProfileClose = document.getElementById('petProfileClose');
    var petProfileCancel = document.getElementById('petProfileCancel');
    var petProfileSave = document.getElementById('petProfileSave');
    var petProfileReset = document.getElementById('petProfileReset');
    if (petProfileClose) petProfileClose.addEventListener('click', closePetProfileEditor);
    if (petProfileCancel) petProfileCancel.addEventListener('click', closePetProfileEditor);
    if (petProfileSave) petProfileSave.addEventListener('click', savePetProfile);
    if (petProfileReset) petProfileReset.addEventListener('click', resetPetProfile);
    // 点击遮罩关闭
    var petProfileOverlay = document.getElementById('petProfileOverlay');
    if (petProfileOverlay) {
        petProfileOverlay.addEventListener('click', function(e) {
            if (e.target === petProfileOverlay) closePetProfileEditor();
        });
    }

    // ---- 暴露函数到 window 供内联 onclick 调用 ----
    window.__petOpenProfile = openPetProfileEditor;
    window.__petCloseProfile = closePetProfileEditor;
    window.__petSaveProfile = savePetProfile;
    window.__petResetProfile = resetPetProfile;
    window.__petSwitchTo = function(type) { switchPet(type); };
    window.__petSwitchNext = function() {
        var types = ['cat', 'dog', 'bunny'];
        var idx = types.indexOf(currentPetType);
        switchPet(types[(idx + 1) % 3]);
    };

    window.__petToggle = function() {
        petVisible = !petVisible;
        petPanel.classList.toggle('hidden', !petVisible);
        updatePetToggleBtn();
    };
    window.__petPet = function() {
        if (!petVisible) {
            petVisible = true;
            petPanel.classList.remove('hidden');
            updatePetToggleBtn();
        }
        petOnClick({ target: getActivePetChar() });
    };

    // ---- 初始化宠物系统 ----
    function initPetSystem() {
        loadPetData();
        // 显示当前宠物
        var allChars = petPanel.querySelectorAll('.pet-character');
        for (var ci = 0; ci < allChars.length; ci++) {
            allChars[ci].classList.toggle('active', allChars[ci].getAttribute('data-pet') === currentPetType);
        }
        // 高亮选择器
        var allSelBtns = document.querySelectorAll('.pet-select-btn');
        for (var sj = 0; sj < allSelBtns.length; sj++) {
            allSelBtns[sj].classList.toggle('active', allSelBtns[sj].getAttribute('data-pet') === currentPetType);
        }
        updatePetInfoSummary();
        updatePetToggleBtn();
    }

    initPetSystem();

    // ===== ⏱ 会话计时器 =====
    var sessionStartTime = Date.now();
    function updateSessionTimer() {
        var elapsed = Math.floor((Date.now() - sessionStartTime) / 1000);
        var min = String(Math.floor(elapsed / 60)).padStart(2, '0');
        var sec = String(elapsed % 60).padStart(2, '0');
        var timerEl = document.getElementById('sessionTimer');
        if (timerEl) timerEl.textContent = '⏱ ' + min + ':' + sec;
    }
    setInterval(updateSessionTimer, 1000);
    updateSessionTimer();

    // ===== 😊 表情栏 =====
    var EMOJI_CATEGORIES = [
        { id:'face',    label:'😊 笑脸',  items:['😀','😃','😄','😁','😆','😅','🤣','😂','🙂','😊','😇','🥰','😍','🤩','😘','😗','😚','😋','😛','😜','🤪','😝','🤑','🤗','🤭','🫢','🫣','🤫','🤔','🫡','🤐','🤨','😐','😑','😶','🫥','😏','😒','🙄','😬','🤥','😌','😔','😪','🤤','😴','😷','🤒','🤕','🤢','🤮','🥴','😵','🤯','🥳','🥺','😢','😭','😤','😠','😡','🤬','🤡','💩','👻','💀','☠️'] },
        { id:'hand',    label:'✋ 手势',  items:['👍','👎','👌','✌️','🤞','🫰','🤟','🤘','🤙','👈','👉','👆','👇','🫵','✋','🖐️','🖖','👋','🤚','✊','👊','🤛','🤜','👏','🙌','🫶','👐','🤲','🤝','🙏','✍️','💅','🤳','💪','🦵','🦶','👂','🦻','👃','🧠','🫀','🫁','🗣️','👤','👥','🫂'] },
        { id:'heart',   label:'❤️ 爱心',  items:['❤️','🧡','💛','💚','💙','💜','🖤','🤍','🤎','💔','❣️','💕','💞','💓','💗','💖','💘','💝','💟','♥️','💑','💏'] },
        { id:'symbol',  label:'⭐ 符号',  items:['✅','❌','❓','❔','❗','‼️','⁉️','➕','➖','➗','✖️','✔️','☑️','🔘','⭕','🔴','🟠','🟡','🟢','🔵','🟣','🟤','⚫','⚪','🔺','🔻','🔸','🔹','🔶','🔷','💠','🔲','🔳','⭐','🌟','✨','💫','🎯','🏆','🥇','🥈','🥉','🏅','🎖️','🏵️','🎗️','📌','📍','💡','🔦','🏮'] },
        { id:'animal',  label:'🐱 动物',  items:['🐱','🐶','🐰','🐭','🐹','🐻','🐼','🐨','🦊','🐸','🐵','🐔','🐧','🐦','🦆','🦅','🦉','🦇','🐺','🐗','🐴','🦄','🐝','🐛','🦋','🐌','🐞','🐜','🦟','🪰','🪱','🐢','🐍','🦎','🦖','🦕','🐙','🦑','🦐','🦞','🦀','🐡','🐠','🐟','🐳','🐋','🦈','🐊','🐅','🐆','🦓','🦍','🦧','🐘','🦛','🦏','🐪','🐫','🦒','🦘','🐃','🐂','🐄','🐎','🐖','🐏','🐑','🦙','🐐','🦌','🐕','🐩','🦮','🐕‍🦺','🐈','🐈‍⬛','🪶','🐓','🦃','🦤','🦚','🦜','🦢','🦩','🐇','🦝','🦨','🦡','🦫','🦦','🦥','🐁','🐀','🐿️','🦔','🐾'] },
        { id:'food',    label:'🍎 食物',  items:['🍎','🍐','🍊','🍋','🍌','🍉','🍇','🍓','🫐','🍈','🍒','🍑','🥭','🍍','🥥','🥝','🍅','🍆','🥑','🥦','🥬','🥒','🌶️','🫑','🌽','🥕','🧄','🧅','🥔','🍠','🥐','🍞','🥖','🥨','🧀','🥚','🍳','🧈','🥞','🧇','🥓','🥩','🍗','🍖','🦴','🌭','🍔','🍟','🍕','🫓','🥪','🥙','🧆','🌮','🌯','🥗','🥘','🫕','🥫','🍝','🍜','🍲','🍛','🍣','🍱','🥟','🦪','🍤','🍙','🍚','🍘','🍥','🥠','🥮','🍢','🍡','🍧','🍨','🍦','🥧','🧁','🍰','🎂','🍮','🍭','🍬','🍫','🍿','🍩','🍪','🌰','🥜','🍯','🥛','🍼','☕','🫖','🍵','🧃','🥤','🧋','🍶','🍺','🍻','🥂','🍷','🫗','🥃','🍸','🧉','🍾','🧊','🥄','🍴','🥣','🍽️','🔪'] },
        { id:'weather', label:'☀️ 天气',  items:['☀️','🌞','🌝','🌛','🌜','🌙','🌚','🌕','🌖','🌗','🌘','🌑','🌒','🌓','🌔','⭐','🌟','✨','💫','☁️','⛅','🌤️','🌥️','🌦️','🌧️','🌨️','🌩️','🌪️','🌫️','🌬️','☔','⚡','❄️','☃️','⛄','🔥','💧','🌊','🌈','☄️','🌋','🏔️','⛰️','🏕️'] },
        { id:'flag',    label:'🚩 旗帜',  items:['🏁','🚩','🎌','🏴','🏳️','🏳️‍🌈','🏳️‍⚧️','🏴‍☠️','🇺🇳','🇺🇸','🇬🇧','🇨🇳','🇯🇵','🇰🇷','🇩🇪','🇫🇷','🇮🇹','🇪🇸','🇷🇺','🇧🇷','🇮🇳','🇦🇺','🇨🇦','🇦🇷'] },
        { id:'number',  label:'🔢 数字',  items:['0️⃣','1️⃣','2️⃣','3️⃣','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣','9️⃣','🔟','#️⃣','*️⃣','🆒','🆓','🆔','🆕','🆖','🆗','🆘','🆙','🆚','🅰️','🅱️','🆎','🅾️'] },
    ];
    var EMOJI_STORAGE_KEY = 'docx-emoji-categories';

    // 从 localStorage 加载用户配置
    function loadEmojiConfig() {
        try {
            var saved = localStorage.getItem(EMOJI_STORAGE_KEY);
            if (saved) return JSON.parse(saved);
        } catch(e) {}
        return null;
    }
    function saveEmojiConfig(config) {
        try { localStorage.setItem(EMOJI_STORAGE_KEY, JSON.stringify(config)); } catch(e) {}
    }

    // 获取当前启用的分类列表（已排序）
    function getActiveCategories() {
        var saved = loadEmojiConfig();
        if (saved && Array.isArray(saved) && saved.length) {
            // 合并用户配置与全量数据
            var map = {};
            EMOJI_CATEGORIES.forEach(function(c) { map[c.id] = c; });
            var result = [];
            saved.forEach(function(s) {
                if (map[s.id] && s.visible !== false) {
                    result.push({ id: s.id, label: map[s.id].label, items: map[s.id].items });
                }
            });
            if (result.length) return result;
        }
        return EMOJI_CATEGORIES.slice();
    }

    var activeEmojiCats = [];
    var emojiPopupOpen = false;

    // 😊 弹出式表情选择器
    function renderEmojiPopup() {
        var tabsEl = document.getElementById('emojiTabs');
        var gridEl = document.getElementById('emojiGrid');
        if (!tabsEl || !gridEl) return;
        activeEmojiCats = getActiveCategories();
        if (!activeEmojiCats.length) return;

        // 渲染标签
        var tabsHtml = '';
        activeEmojiCats.forEach(function(cat, i) {
            tabsHtml += '<button class="emoji-tab' + (i === 0 ? ' active' : '') + '" data-idx="' + i + '">' + cat.label + '</button>';
        });
        tabsEl.innerHTML = tabsHtml;

        // 标签点击
        tabsEl.querySelectorAll('.emoji-tab').forEach(function(tab) {
            tab.addEventListener('click', function() {
                tabsEl.querySelectorAll('.emoji-tab').forEach(function(t) { t.classList.remove('active'); });
                this.classList.add('active');
                renderEmojiGrid(parseInt(this.dataset.idx));
            });
        });

        // 渲染第一个分类
        renderEmojiGrid(0);
    }

    function renderEmojiGrid(idx) {
        var gridEl = document.getElementById('emojiGrid');
        if (!gridEl || idx >= activeEmojiCats.length) return;
        var cat = activeEmojiCats[idx];
        var html = '';
        cat.items.forEach(function(emo) {
            html += '<span class="emoji-item" title="' + emo + '">' + emo + '</span>';
        });
        gridEl.innerHTML = html;
        gridEl.querySelectorAll('.emoji-item').forEach(function(el) {
            el.addEventListener('click', function() {
                var emoji = this.textContent;
                insertEmojiAtCursor(emoji);
            });
        });
    }

    function toggleEmojiPopup() {
        var popup = document.getElementById('emojiPopup');
        if (!popup) return;
        if (emojiPopupOpen) { hideEmojiPopup(); } else { showEmojiPopup(); }
    }

    function showEmojiPopup() {
        var popup = document.getElementById('emojiPopup');
        if (!popup) return;
        renderEmojiPopup();
        popup.classList.remove('hidden');
        emojiPopupOpen = true;
    }

    function hideEmojiPopup() {
        var popup = document.getElementById('emojiPopup');
        if (!popup) return;
        popup.classList.add('hidden');
        emojiPopupOpen = false;
    }

    // 工具栏按钮
    var emojiPopupBtn = document.getElementById('emojiPopupBtn');
    if (emojiPopupBtn) emojiPopupBtn.addEventListener('click', toggleEmojiPopup);

    var emojiPopupClose = document.getElementById('emojiPopupClose');
    if (emojiPopupClose) emojiPopupClose.addEventListener('click', hideEmojiPopup);

    // Ctrl+E 快捷键
    document.addEventListener('keydown', function(e) {
        if ((e.ctrlKey || e.metaKey) && (e.key === 'e' || e.key === 'E') && !e.shiftKey && !e.altKey) {
            var inEditor = document.activeElement && editor && editor.contains(document.activeElement);
            if (inEditor || emojiPopupOpen) {
                e.preventDefault();
                toggleEmojiPopup();
            }
        }
    });

    // 点击外部关闭
    document.addEventListener('click', function(e) {
        if (!emojiPopupOpen) return;
        var popup = document.getElementById('emojiPopup');
        if (!popup || popup.contains(e.target)) return;
        if (emojiPopupBtn && emojiPopupBtn.contains(e.target)) return;
        hideEmojiPopup();
    });

    function insertEmojiAtCursor(emoji) {
        saveUndoState('插入表情');
        var sel = window.getSelection();
        if (sel.rangeCount && editor.contains(sel.getRangeAt(0).commonAncestorContainer)) {
            document.execCommand('insertText', false, emoji);
        } else {
            // 光标不在编辑区，聚焦到编辑区末尾
            editor.focus();
            var r = document.createRange();
            r.setStart(editor, editor.childNodes.length);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
            document.execCommand('insertText', false, emoji);
        }
    }


    // ===== 表格功能 =====
    var MAX_TABLE_SIZE = 10;
    var tablePickerVisible = false;
    var pickerRows = 1, pickerCols = 1;
    var insertMarkerId = null;

    function createInsertMarker() {
        var id = 'ins-marker-' + Date.now() + '-' + Math.random().toString(36).substr(2, 4);
        editor.focus();
        var sel = window.getSelection();
        // 无选区时在末尾建一个
        if (!sel.rangeCount || !editor.contains(sel.getRangeAt(0).commonAncestorContainer)) {
            var r = document.createRange();
            r.setStart(editor, editor.childNodes.length);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
        }
        try {
            document.execCommand('insertHTML', false, '<span id="' + id + '" style="display:inline-block;width:1px;height:1px;">&nbsp;</span>');
        } catch(e) {
            // 回退：直接 append
            var sp = document.createElement('span');
            sp.id = id;
            sp.style.display = 'inline-block';
            sp.style.width = '1px';
            sp.style.height = '1px';
            sp.innerHTML = '&nbsp;';
            editor.appendChild(sp);
        }
        insertMarkerId = id;
        return id;
    }

    function getInsertMarker() {
        if (!insertMarkerId) return null;
        var el = document.getElementById(insertMarkerId);
        return el;
    }

    function removeInsertMarker() {
        var el = getInsertMarker();
        if (el && el.parentNode) el.parentNode.removeChild(el);
        insertMarkerId = null;
    }

    function buildTablePickerGrid() {
        var grid = $('tablePickerGrid');
        grid.innerHTML = '';
        for (var r = 0; r < MAX_TABLE_SIZE; r++) {
            for (var c = 0; c < MAX_TABLE_SIZE; c++) {
                var cell = document.createElement('div');
                cell.className = 'table-picker-cell';
                cell.dataset.row = r + 1;
                cell.dataset.col = c + 1;
                cell.addEventListener('mouseenter', onPickerCellEnter);
                cell.addEventListener('click', onPickerCellClick);
                grid.appendChild(cell);
            }
        }
    }

    function onPickerCellEnter(e) {
        var row = parseInt(e.target.dataset.row);
        var col = parseInt(e.target.dataset.col);
        pickerRows = row;
        pickerCols = col;
        $('tablePickerInfo').textContent = row + ' 行 × ' + col + ' 列';
        var cells = $('tablePickerGrid').querySelectorAll('.table-picker-cell');
        cells.forEach(function(cell) {
            var cr = parseInt(cell.dataset.row);
            var cc = parseInt(cell.dataset.col);
            cell.classList.remove('active', 'selected');
            if (cr <= row && cc <= col) {
                cell.classList.add(cr === row && cc === col ? 'selected' : 'active');
            }
        });
    }

    function onPickerCellClick(e) {
        insertTableAtCursor(pickerRows, pickerCols);
        hideTablePicker();
    }

    function showTablePicker() {
        if (tablePickerVisible) return;
        createInsertMarker(); // 在光标位置插入标记
        tablePickerVisible = true;
        pickerRows = 1;
        pickerCols = 1;
        $('tablePickerInfo').textContent = '1 行 × 1 列';
        var cells = $('tablePickerGrid').querySelectorAll('.table-picker-cell');
        cells.forEach(function(cell) { cell.classList.remove('active', 'selected'); });
        if (cells.length) cells[0].classList.add('selected');
        $('tablePickerOverlay').classList.remove('hidden');
    }

    function hideTablePicker() {
        tablePickerVisible = false;
        $('tablePickerOverlay').classList.add('hidden');
        removeInsertMarker();
    }

    function insertTableAtCursor(rows, cols) {
        saveUndoState('插入表格'); // 记录插入表格前状态
        try {
            var padding = parseInt($('tableCellPadding').value) || 8;
            var html = '<table style="border-collapse:collapse;width:100%;margin:0.5em 0;"><thead><tr>';
            for (var c = 0; c < cols; c++) {
                html += '<th style="border:1px solid #999;padding:' + padding + 'px;min-width:40px;vertical-align:top;background:#f0f0f0;font-weight:bold;">&nbsp;</th>';
            }
            html += '</tr></thead>';
            if (rows > 1) {
                html += '<tbody>';
                for (var r = 1; r < rows; r++) {
                    html += '<tr>';
                    for (var c = 0; c < cols; c++) {
                        html += '<td style="border:1px solid #999;padding:' + padding + 'px;min-width:40px;vertical-align:top;">&nbsp;</td>';
                    }
                    html += '</tr>';
                }
                html += '</tbody>';
            }
            html += '</table>';

            var marker = getInsertMarker();
            if (marker) {
                // 在标记位置插入表格
                var table = document.createElement('table');
                table.style.borderCollapse = 'collapse';
                table.style.width = '100%';
                table.style.margin = '0.5em 0';
                var innerHtml = '<thead><tr>';
                for (var c = 0; c < cols; c++) {
                    innerHtml += '<th style="border:1px solid #999;padding:' + padding + 'px;min-width:40px;vertical-align:top;background:#f0f0f0;font-weight:bold;">&nbsp;</th>';
                }
                innerHtml += '</tr></thead>';
                if (rows > 1) {
                    innerHtml += '<tbody>';
                    for (var r = 1; r < rows; r++) {
                        innerHtml += '<tr>';
                        for (var c = 0; c < cols; c++) {
                            innerHtml += '<td style="border:1px solid #999;padding:' + padding + 'px;min-width:40px;vertical-align:top;">&nbsp;</td>';
                        }
                        innerHtml += '</tr>';
                    }
                    innerHtml += '</tbody>';
                }
                table.innerHTML = innerHtml;
                // 用表格替换标记
                marker.parentNode.insertBefore(table, marker);
                marker.parentNode.removeChild(marker);
                insertMarkerId = null;
                // 光标放入第一个单元格
                var fc = table.querySelector('th, td');
                if (fc) {
                    var sel = window.getSelection();
                    var r2 = document.createRange();
                    r2.setStart(fc, 0);
                    r2.collapse(true);
                    sel.removeAllRanges();
                    sel.addRange(r2);
                }
            } else {
                // 无标记，直接 execCommand
                document.execCommand('insertHTML', false, html);
            }

            showToast('已插入 ' + rows + '×' + cols + ' 表格', 'success');
            setStatus('Table inserted: ' + rows + '×' + cols);
        } catch(err) {
            console.error('insertTableAtCursor error:', err);
            showToast('插入表格失败: ' + err.message, 'error');
        }
    }

    // 表格 Tab 导航 — 通过事件委托监听 editor 的 keydown
    function handleTableKeydown(e) {
        if (e.key !== 'Tab') return;
        var sel = window.getSelection();
        if (!sel.rangeCount) return;
        var node = sel.getRangeAt(0).commonAncestorContainer;
        // 向上找最近的 td/th
        var cell = null, p = node;
        while (p && p !== editor) {
            if (p.nodeType === Node.ELEMENT_NODE && (p.tagName === 'TD' || p.tagName === 'TH')) {
                cell = p; break;
            }
            p = p.parentNode;
        }
        if (!cell) return;
        e.preventDefault();
        var table = cell.closest ? cell.closest('table') : null;
        if (!table) return;

        var allCells = table.querySelectorAll('th, td');
        var idx = Array.prototype.indexOf.call(allCells, cell);
        var nextIdx = e.shiftKey ? idx - 1 : idx + 1;
        if (nextIdx >= 0 && nextIdx < allCells.length) {
            var nextCell = allCells[nextIdx];
            var r = document.createRange();
            r.setStart(nextCell, 0);
            r.collapse(true);
            sel.removeAllRanges();
            sel.addRange(r);
        } else if (!e.shiftKey && nextIdx >= allCells.length) {
            // 在最后一个单元格按 Tab → 添加新行
            addRowToTable(table, 'bottom');
            var rows = table.querySelectorAll('tr');
            var newCells = rows[rows.length - 1].querySelectorAll('td');
            if (newCells.length) {
                sel.removeAllRanges();
                var r = document.createRange();
                r.setStart(newCells[0], 0);
                r.collapse(true);
                sel.addRange(r);
            }
        }
    }

    // ===== 动态增删行列 =====
    function addRowToTable(table, position) {
        saveUndoState('添加表格行');
        var rows = table.querySelectorAll('tr');
        var colCount = rows.length ? rows[0].querySelectorAll('th, td').length : 1;
        var newRow = document.createElement('tr');
        var isHeader = position === 'top';
        for (var c = 0; c < colCount; c++) {
            var cell = isHeader ? document.createElement('th') : document.createElement('td');
            cell.style.border = '1px solid #999';
            cell.style.padding = '8px';
            cell.style.minWidth = '40px';
            cell.style.verticalAlign = 'top';
            if (isHeader) {
                cell.style.background = '#f0f0f0';
                cell.style.fontWeight = 'bold';
            }
            cell.innerHTML = '&nbsp;';
            newRow.appendChild(cell);
        }
        if (position === 'top') {
            // 在表头最上方插入
            var thead = table.querySelector('thead');
            if (thead) {
                thead.insertBefore(newRow, thead.firstChild);
            } else {
                thead = document.createElement('thead');
                thead.appendChild(newRow);
                table.insertBefore(thead, table.firstChild);
            }
        } else if (position === 'bottom') {
            // 在表格最下方追加
            var tbody = table.querySelector('tbody');
            if (!tbody) {
                tbody = document.createElement('tbody');
                table.appendChild(tbody);
            }
            tbody.appendChild(newRow);
        } else {
            // "above" / "below" — 基于选中行
            var selCell = table.querySelector('.selected-cell');
            var targetRow = selCell ? selCell.closest('tr') : null;
            if (!targetRow) {
                // 无选中行，追加到最后
                var tbody2 = table.querySelector('tbody');
                if (!tbody2) {
                    tbody2 = document.createElement('tbody');
                    table.appendChild(tbody2);
                }
                tbody2.appendChild(newRow);
            } else if (targetRow.parentNode.tagName.toLowerCase() === 'thead' && position === 'above') {
                targetRow.parentNode.insertBefore(newRow, targetRow);
            } else if (targetRow.parentNode.tagName.toLowerCase() === 'thead' && position === 'below') {
                // thead 结束后插入 → 创建/获取 tbody
                var tb = table.querySelector('tbody');
                if (!tb) {
                    tb = document.createElement('tbody');
                    table.appendChild(tb);
                }
                tb.insertBefore(newRow, tb.firstChild);
            } else {
                targetRow.parentNode.insertBefore(newRow, position === 'above' ? targetRow : targetRow.nextSibling);
            }
        }
        showToast('已添加行', 'success');
    }

    function addColumnToTable(table, position) {
        saveUndoState('添加表格列');
        var rows = table.querySelectorAll('tr');
        if (!rows.length) return;
        rows.forEach(function(row) {
            var cell = document.createElement('td');
            cell.style.border = '1px solid #999';
            cell.style.padding = '8px';
            cell.style.minWidth = '40px';
            cell.style.verticalAlign = 'top';
            cell.innerHTML = '&nbsp;';

            // 如果是 thead 中的行，使用 th
            var isHeadRow = row.parentNode && row.parentNode.tagName.toLowerCase() === 'thead';
            if (isHeadRow) {
                cell.style.background = '#f0f0f0';
                cell.style.fontWeight = 'bold';
            }

            var refCell = row.querySelector('.selected-cell');
            if (refCell) {
                if (position === 'left') {
                    row.insertBefore(cell, refCell);
                } else {
                    row.insertBefore(cell, refCell.nextSibling);
                }
            } else {
                row.appendChild(cell);
            }
        });
        showToast('已添加列', 'success');
    }

    function deleteRowFromTable(table) {
        saveUndoState('删除表格行');
        var selectedCell = table.querySelector('.selected-cell');
        var targetRow = selectedCell ? selectedCell.closest('tr') : null;
        if (!targetRow || table.querySelectorAll('tr').length <= 1) {
            showToast('至少保留一行', 'warning');
            return;
        }
        targetRow.parentNode.removeChild(targetRow);
        showToast('已删除行', 'info');
    }

    function deleteColumnFromTable(table) {
        saveUndoState('删除表格列');
        var selectedCell = table.querySelector('.selected-cell');
        if (!selectedCell) return;
        var colIdx = Array.prototype.indexOf.call(selectedCell.parentNode.children, selectedCell);
        var rows = table.querySelectorAll('tr');
        if (rows[0] && rows[0].querySelectorAll('th, td').length <= 1) {
            showToast('至少保留一列', 'warning');
            return;
        }
        rows.forEach(function(row) {
            var cells = row.querySelectorAll('th, td');
            if (cells[colIdx]) cells[colIdx].parentNode.removeChild(cells[colIdx]);
        });
        showToast('已删除列', 'info');
    }

    // ===== 表格右键菜单 =====
    var tableContextMenu = null;
    var activeTable = null;

    function showTableContextMenu(e, table) {
        e.preventDefault();
        hideTableContextMenu();
        activeTable = table;

        // 清除旧选中
        table.querySelectorAll('.selected-cell').forEach(function(c) { c.classList.remove('selected-cell'); });
        var targetCell = e.target.closest ? e.target.closest('td, th') : null;
        if (targetCell) targetCell.classList.add('selected-cell');

        var menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = e.clientX + 'px';
        menu.style.top = e.clientY + 'px';

        var items = [
            { label: '⬆ 在上方插入行', action: 'rowAbove' },
            { label: '⬇ 在下方插入行', action: 'rowBelow' },
            { label: '⬅ 在左侧插入列', action: 'colLeft' },
            { label: '➡ 在右侧插入列', action: 'colRight' },
            { label: '', divider: true },
            { label: '🗑 删除行', action: 'delRow' },
            { label: '🗑 删除列', action: 'delCol' },
            { label: '', divider: true },
            { label: '❌ 删除表格', action: 'delTable' },
        ];

        items.forEach(function(item) {
            if (item.divider) {
                var div = document.createElement('div');
                div.className = 'context-menu-divider';
                menu.appendChild(div);
                return;
            }
            var el = document.createElement('div');
            el.className = 'context-menu-item';
            el.textContent = item.label;
            el.addEventListener('click', function() {
                handleTableAction(item.action, table);
                hideTableContextMenu();
            });
            menu.appendChild(el);
        });

        document.body.appendChild(menu);
        tableContextMenu = menu;

        // 调整边界
        var mr = menu.getBoundingClientRect();
        if (mr.right > window.innerWidth) menu.style.left = (window.innerWidth - mr.width - 10) + 'px';
        if (mr.bottom > window.innerHeight) menu.style.top = (window.innerHeight - mr.height - 10) + 'px';

        setTimeout(function() {
            document.addEventListener('click', hideTableContextMenu, { once: true });
        }, 0);
    }

    function hideTableContextMenu() {
        if (tableContextMenu) {
            tableContextMenu.remove();
            tableContextMenu = null;
        }
        if (activeTable) {
            activeTable.querySelectorAll('.selected-cell').forEach(function(c) { c.classList.remove('selected-cell'); });
            activeTable = null;
        }
    }

    function handleTableAction(action, table) {
        switch (action) {
            case 'rowAbove': addRowToTable(table, 'above'); break;
            case 'rowBelow': addRowToTable(table, 'below'); break;
            case 'colLeft': addColumnToTable(table, 'left'); break;
            case 'colRight': addColumnToTable(table, 'right'); break;
            case 'delRow': deleteRowFromTable(table); break;
            case 'delCol': deleteColumnFromTable(table); break;
            case 'delTable':
                if (confirm('确定删除整个表格？')) {
                    saveUndoState('删除表格'); // 记录删除前状态
                    table.parentNode.removeChild(table);
                    showToast('表格已删除', 'info');
                }
                break;
        }
    }

    // ===== 编辑器右键菜单（表格 + 文本格式化） =====
    var textContextMenu = null;
    var activeSubmenu = null;

    // 颜色色板
    var COLOR_PALETTE = [
        ['#000000', '#333333', '#666666', '#999999', '#CCCCCC', '#E0E0E0', '#FFFFFF', '#FF0000'],
        ['#FF6600', '#FFCC00', '#FFFF00', '#00CC00', '#00BFFF', '#0000FF', '#9900FF', '#FF00FF'],
        ['#8B0000', '#B22222', '#CD853F', '#B8860B', '#006400', '#008B8B', '#000080', '#4B0082'],
        ['#DC143C', '#FF6347', '#FFD700', '#32CD32', '#008080', '#4169E1', '#8A2BE2', '#FF69B4']
    ];

    // 背景色色板（浅色系为主）
    var BG_COLOR_PALETTE = [
        ['#FFFF00', '#FFD700', '#FFA500', '#90EE90', '#87CEEB', '#FFB6C1', '#DDA0DD', '#F0E68C'],
        ['#FFFACD', '#FFEFD5', '#FFE4E1', '#E0FFFF', '#E6E6FA', '#F0FFF0', '#FFF0F5', '#F5F5DC'],
        ['#FFFFC8', '#FFE0B2', '#FFCDD2', '#BBDEFB', '#C8E6C9', '#D1C4E9', '#F8BBD0', '#B3E5FC'],
        ['#FFFFFF', '#F5F5F5', '#E8E8E8', '#DCDCDC', '#D3D3D3', '#C0C0C0', '#A9A9A9', '#808080']
    ];

    var FONT_SIZES = ['9pt', '10pt', '10.5pt', '12pt', '14pt', '15pt', '16pt', '18pt', '20pt', '22pt', '24pt', '28pt', '36pt', '48pt'];
    var FONT_FAMILIES = ['宋体', '黑体', '微软雅黑', '仿宋', '楷体', 'Arial', 'Times New Roman', 'Consolas', 'Courier New', 'Georgia', 'Verdana'];
    var savedTextRange = null; // 保存右键时的选区，供格式应用时恢复

    function showTextContextMenu(e) {
        e.preventDefault();
        hideTextContextMenu();
        hideTableContextMenu();

        var sel = window.getSelection();
        if (!sel.rangeCount || sel.isCollapsed) return; // 没有选中文字不显示

        // 确保选区在编辑器内
        var range = sel.getRangeAt(0);
        if (!editor.contains(range.commonAncestorContainer)) return;

        // 保存选区（右键菜单会导致选区丢失，需要在应用格式前恢复）
        savedTextRange = range.cloneRange();

        // 查询当前格式状态
        var isBold = document.queryCommandState('bold');
        var isItalic = document.queryCommandState('italic');
        var isUnderline = document.queryCommandState('underline');
        var currentColor = document.queryCommandValue('foreColor') || '#000000';
        var currentBg = document.queryCommandValue('backColor') || 'transparent';
        // 规范化颜色值
        if (currentColor === 'rgb(0, 0, 0)') currentColor = '#000000';
        if (currentBg === 'rgba(0, 0, 0, 0)' || currentBg === 'transparent') currentBg = '';

        var menu = document.createElement('div');
        menu.className = 'context-menu text-context-menu';
        menu.style.left = Math.min(e.clientX, window.innerWidth - 220) + 'px';
        menu.style.top = Math.min(e.clientY, window.innerHeight - 400) + 'px';

        // 加粗
        menu.appendChild(buildMenuItem('加粗', 'bold', isBold, 'Ctrl+B'));
        // 斜体
        menu.appendChild(buildMenuItem('斜体', 'italic', isItalic, 'Ctrl+I'));
        // 下划线
        menu.appendChild(buildMenuItem('下划线', 'underline', isUnderline, 'Ctrl+U'));

        // 分隔线
        var div = document.createElement('div');
        div.className = 'context-menu-divider';
        menu.appendChild(div);

        // 字体颜色（含色板子菜单）
        menu.appendChild(buildColorMenuItem('字体颜色', 'fontColor', currentColor, COLOR_PALETTE, false));

        // 背景颜色（含色板子菜单）
        menu.appendChild(buildColorMenuItem('背景颜色', 'bgColor', currentBg, BG_COLOR_PALETTE, true));

        // 分隔线
        div = document.createElement('div');
        div.className = 'context-menu-divider';
        menu.appendChild(div);

        // 字号子菜单
        menu.appendChild(buildFontSizeMenuItem());

        // 字体子菜单
        menu.appendChild(buildFontFamilyMenuItem());

        document.body.appendChild(menu);
        textContextMenu = menu;

        // 调整边界
        var mr = menu.getBoundingClientRect();
        if (mr.right > window.innerWidth) menu.style.left = (window.innerWidth - mr.width - 10) + 'px';
        if (mr.bottom > window.innerHeight) menu.style.top = (window.innerHeight - mr.height - 10) + 'px';

        setTimeout(function() {
            document.addEventListener('click', hideTextContextMenu, { once: true });
        }, 0);
    }

    function buildMenuItem(label, action, active, shortcut) {
        var el = document.createElement('div');
        el.className = 'context-menu-item' + (active ? ' active' : '');
        el.innerHTML = '<span class="menu-icon">' + getActionIcon(action) + '</span>' +
                       '<span>' + label + '</span>' +
                       (shortcut ? '<span class="shortcut">' + shortcut + '</span>' : '');
        el.addEventListener('click', function(e) {
            e.stopPropagation();
            applyTextFormat(action);
            hideTextContextMenu();
        });
        return el;
    }

    function getActionIcon(action) {
        switch (action) {
            case 'bold': return '<b>B</b>';
            case 'italic': return '<i>I</i>';
            case 'underline': return '<u>U</u>';
            default: return '';
        }
    }

    function buildColorMenuItem(label, action, currentColor, palette, isBg) {
        var el = document.createElement('div');
        el.className = 'context-menu-item has-submenu';
        var swatchHtml = '';
        if (currentColor && currentColor !== 'transparent' && currentColor !== 'rgba(0, 0, 0, 0)') {
            swatchHtml = '<span class="color-swatch" style="background-color:' + currentColor + ';border:1px solid var(--border);"></span>';
        } else {
            swatchHtml = '<span class="color-swatch color-swatch-empty"></span>';
        }
        el.innerHTML = swatchHtml + '<span>' + label + '</span><span class="submenu-arrow">▸</span>';

        // 子菜单
        var sub = document.createElement('div');
        sub.className = 'context-submenu color-submenu';
        sub.addEventListener('click', function(e) { e.stopPropagation(); });

        // 色板网格
        var grid = document.createElement('div');
        grid.className = 'color-grid';
        for (var r = 0; r < palette.length; r++) {
            for (var c = 0; c < palette[r].length; c++) {
                (function(color) {
                    var cell = document.createElement('div');
                    cell.className = 'color-cell';
                    cell.style.backgroundColor = color;
                    if (color === '#FFFFFF' || color === '#F5F5F5' || color === '#FFFACD' || color === '#FFFFC8' || color === '#F0FFF0' || color === '#FFF0F5' || color === '#F5F5DC') {
                        cell.style.border = '1px solid var(--border)';
                    }
                    cell.title = color;
                    cell.addEventListener('click', function() {
                        applyColorFormat(action, color, isBg);
                        hideTextContextMenu();
                    });
                    grid.appendChild(cell);
                })(palette[r][c]);
            }
        }
        sub.appendChild(grid);

        // 自定义颜色按钮
        var customBtn = document.createElement('div');
        customBtn.className = 'context-menu-item custom-color-btn';
        customBtn.innerHTML = '🎨 自定义颜色...';
        customBtn.addEventListener('click', function() {
            var input = document.createElement('input');
            input.type = 'color';
            input.value = currentColor || '#000000';
            input.style.position = 'fixed';
            input.style.left = '-9999px';
            document.body.appendChild(input);
            input.click();
            input.addEventListener('change', function() {
                applyColorFormat(action, input.value, isBg);
                document.body.removeChild(input);
                hideTextContextMenu();
            });
            input.addEventListener('cancel', function() {
                document.body.removeChild(input);
            });
        });
        sub.appendChild(customBtn);

        // 清除颜色按钮（仅背景色需要）
        if (isBg) {
            var clearBtn = document.createElement('div');
            clearBtn.className = 'context-menu-item clear-color-btn';
            clearBtn.innerHTML = '✕ 清除背景色';
            clearBtn.addEventListener('click', function() {
                applyColorFormat(action, 'transparent', true);
                hideTextContextMenu();
            });
            sub.appendChild(clearBtn);
        }

        el.appendChild(sub);

        // 悬停显示子菜单
        el.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        el.addEventListener('mouseleave', function(e) {
            if (!sub.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });
        sub.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        sub.addEventListener('mouseleave', function(e) {
            if (!el.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });

        return el;
    }

    function buildFontSizeMenuItem() {
        var el = document.createElement('div');
        el.className = 'context-menu-item has-submenu';
        el.innerHTML = '<span class="menu-icon">Ⓐ</span><span>字号</span><span class="submenu-arrow">▸</span>';

        var sub = document.createElement('div');
        sub.className = 'context-submenu';
        sub.addEventListener('click', function(e) { e.stopPropagation(); });

        for (var i = 0; i < FONT_SIZES.length; i++) {
            (function(size) {
                var item = document.createElement('div');
                item.className = 'context-menu-item';
                item.textContent = size;
                item.style.fontSize = size;
                item.addEventListener('click', function() {
                    applyFontSize(size);
                    hideTextContextMenu();
                });
                sub.appendChild(item);
            })(FONT_SIZES[i]);
        }

        el.appendChild(sub);

        el.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        el.addEventListener('mouseleave', function(e) {
            if (!sub.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });
        sub.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        sub.addEventListener('mouseleave', function(e) {
            if (!el.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });

        return el;
    }

    function buildFontFamilyMenuItem() {
        var el = document.createElement('div');
        el.className = 'context-menu-item has-submenu';
        el.innerHTML = '<span class="menu-icon">ƒ</span><span>字体</span><span class="submenu-arrow">▸</span>';

        var sub = document.createElement('div');
        sub.className = 'context-submenu';
        sub.addEventListener('click', function(e) { e.stopPropagation(); });

        for (var i = 0; i < FONT_FAMILIES.length; i++) {
            (function(family) {
                var item = document.createElement('div');
                item.className = 'context-menu-item';
                item.textContent = family;
                item.style.fontFamily = family;
                item.addEventListener('click', function() {
                    applyFontFamily(family);
                    hideTextContextMenu();
                });
                sub.appendChild(item);
            })(FONT_FAMILIES[i]);
        }

        el.appendChild(sub);

        el.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        el.addEventListener('mouseleave', function(e) {
            if (!sub.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });
        sub.addEventListener('mouseenter', function() { showSubmenu(sub, el); });
        sub.addEventListener('mouseleave', function(e) {
            if (!el.contains(e.relatedTarget)) { hideSubmenu(sub); }
        });

        return el;
    }

    function showSubmenu(sub, parent) {
        if (activeSubmenu === sub) return;
        hideAllSubmenus();
        var pr = parent.getBoundingClientRect();
        sub.style.display = 'block';
        // 定位在父菜单项的右侧
        var left = pr.right + 2;
        var top = pr.top;
        // 边界调整
        var subRect = sub.getBoundingClientRect();
        if (left + subRect.width > window.innerWidth) left = pr.left - subRect.width - 2;
        if (top + subRect.height > window.innerHeight) top = Math.max(0, window.innerHeight - subRect.height - 10);
        if (top < 0) top = 0;
        sub.style.left = left + 'px';
        sub.style.top = top + 'px';
        activeSubmenu = sub;
    }

    function hideSubmenu(sub) {
        sub.style.display = 'none';
        if (activeSubmenu === sub) activeSubmenu = null;
    }

    function hideAllSubmenus() {
        if (activeSubmenu) {
            activeSubmenu.style.display = 'none';
            activeSubmenu = null;
        }
    }

    // ===== 文本格式化应用函数 =====

    // 恢复右键菜单显示前保存的选区
    function restoreTextSelection() {
        if (savedTextRange) {
            var sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(savedTextRange);
            // 恢复后清除，避免误用过期选区
            savedTextRange = null;
            return true;
        }
        return false;
    }

    function applyTextFormat(action) {
        restoreTextSelection();
        switch (action) {
            case 'bold':
                saveUndoState('加粗');
                document.execCommand('styleWithCSS', false, false);
                document.execCommand('bold');
                break;
            case 'italic':
                saveUndoState('斜体');
                document.execCommand('styleWithCSS', false, false);
                document.execCommand('italic');
                break;
            case 'underline':
                saveUndoState('下划线');
                document.execCommand('styleWithCSS', false, false);
                document.execCommand('underline');
                break;
        }
    }

    function applyColorFormat(action, color, isBg) {
        restoreTextSelection();
        var desc = isBg ? '背景色' : '文字颜色';
        saveUndoState(desc);
        // 不使用 execCommand('foreColor'/'backColor')，因为 Chrome 会生成
        // <font color="..."> 标签而非 <span style="color:...">，导致导出丢失颜色。
        // 统一使用 wrapSelectionWithSpan 确保生成 span 元素。
        if (isBg) {
            if (color === 'transparent') {
                // 清除背景色：用白色覆盖（execCommand backColor 在这里还可以接受）
                document.execCommand('styleWithCSS', false, true);
                document.execCommand('backColor', false, '#FFFFFF');
            } else {
                wrapSelectionWithSpan({ backgroundColor: color });
            }
        } else {
            wrapSelectionWithSpan({ color: color });
        }
    }

    function applyFontSize(size) {
        restoreTextSelection();
        saveUndoState('字号');
        // 不使用 execCommand fontSize（它用 1-7 的 HTML 字号）
        wrapSelectionWithSpan({ fontSize: size });
    }

    function applyFontFamily(family) {
        restoreTextSelection();
        saveUndoState('字体');
        // 不使用 execCommand fontName（可能生成 <font face="...">）
        wrapSelectionWithSpan({ fontFamily: family });
    }

    // 将当前选区包装在带指定样式的 span 中
    function wrapSelectionWithSpan(styleProps) {
        var sel = window.getSelection();
        if (!sel.rangeCount || sel.isCollapsed) return;
        var range = sel.getRangeAt(0);
        if (!editor.contains(range.commonAncestorContainer)) return;

        // 尝试 surroundContents（简单情况：选区完整包含在单个文本节点中）
        try {
            var span = document.createElement('span');
            for (var key in styleProps) {
                if (styleProps.hasOwnProperty(key)) {
                    span.style[key] = styleProps[key];
                }
            }
            range.surroundContents(span);
            // 恢复选区
            sel.removeAllRanges();
            var newRange = document.createRange();
            newRange.selectNodeContents(span);
            sel.addRange(newRange);
            return;
        } catch(e) {
            // 跨元素选择：使用 extractContents + insertNode
        }

        // 复杂情况：提取片段，包装，插入
        try {
            var fragment = range.extractContents();
            var span = document.createElement('span');
            for (var key in styleProps) {
                if (styleProps.hasOwnProperty(key)) {
                    span.style[key] = styleProps[key];
                }
            }
            span.appendChild(fragment);
            range.insertNode(span);
            sel.removeAllRanges();
            var newRange = document.createRange();
            newRange.selectNodeContents(span);
            sel.addRange(newRange);
        } catch(e2) {
            console.warn('wrapSelectionWithSpan failed:', e2);
        }
    }

    function hideTextContextMenu() {
        if (textContextMenu) {
            textContextMenu.remove();
            textContextMenu = null;
        }
        hideAllSubmenus();
        savedTextRange = null; // 清除保存的选区
    }

    // ===== 编辑器右键事件统一处理 =====
    editor.addEventListener('contextmenu', function(e) {
        // 检查是否右键点击了图片
        var img = e.target.closest ? e.target.closest('img') : null;
        if (img && editor.contains(img)) {
            hideTableContextMenu();
            hideTextContextMenu();
            hideImageContextMenu();
            showImageContextMenu(e, img);
            return;
        }

        var table = e.target.closest ? e.target.closest('table') : null;
        if (table && editor.contains(table)) {
            hideTextContextMenu();
            hideImageContextMenu();
            showTableContextMenu(e, table);
            return;
        }

        // 检查是否有选中文字
        var sel = window.getSelection();
        if (sel.rangeCount && !sel.isCollapsed && editor.contains(sel.getRangeAt(0).commonAncestorContainer)) {
            hideTableContextMenu();
            hideImageContextMenu();
            showTextContextMenu(e);
            return;
        }

        // 无表格、无选中文字：关闭所有菜单，让浏览器显示默认菜单
        hideTableContextMenu();
        hideTextContextMenu();
        hideImageContextMenu();
    });

    // ===== 图片右键菜单 =====
    var imageContextMenu = null;
    var imageContextTarget = null;

    function showImageContextMenu(e, img) {
        e.preventDefault();
        imageContextTarget = img;

        var menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.left = e.clientX + 'px';
        menu.style.top = e.clientY + 'px';

        // 查看图片
        var viewItem = document.createElement('div');
        viewItem.className = 'context-menu-item img-context-item';
        viewItem.innerHTML = '<span>🖼️</span> 查看图片';
        viewItem.addEventListener('click', function() {
            hideImageContextMenu();
            window.openImageEditor(img, 'view');
        });

        // 编辑图片
        var editItem = document.createElement('div');
        editItem.className = 'context-menu-item img-context-item';
        editItem.innerHTML = '<span>✏️</span> 编辑图片';
        editItem.addEventListener('click', function() {
            hideImageContextMenu();
            window.openImageEditor(img, 'edit');
        });

        // 裁剪图片
        var cropItem = document.createElement('div');
        cropItem.className = 'context-menu-item img-context-item';
        cropItem.innerHTML = '<span>✂️</span> 裁剪图片';
        cropItem.addEventListener('click', function() {
            hideImageContextMenu();
            window.openImageEditor(img, 'crop');
        });

        menu.appendChild(viewItem);
        menu.appendChild(editItem);
        menu.appendChild(cropItem);

        document.body.appendChild(menu);
        imageContextMenu = menu;

        // 调整菜单位置防止溢出
        var menuRect = menu.getBoundingClientRect();
        if (menuRect.right > window.innerWidth) {
            menu.style.left = (e.clientX - menuRect.width) + 'px';
        }
        if (menuRect.bottom > window.innerHeight) {
            menu.style.top = (e.clientY - menuRect.height) + 'px';
        }

        document.addEventListener('click', hideImageContextMenu, { once: true });
    }

    function hideImageContextMenu() {
        if (imageContextMenu) {
            imageContextMenu.remove();
            imageContextMenu = null;
        }
        imageContextTarget = null;
    }

    // 点击其他位置关闭所有右键菜单
    document.addEventListener('click', function(e) {
        if (tableContextMenu && !tableContextMenu.contains(e.target)) {
            hideTableContextMenu();
        }
        if (textContextMenu && !textContextMenu.contains(e.target)) {
            hideTextContextMenu();
        }
        if (imageContextMenu && !imageContextMenu.contains(e.target)) {
            hideImageContextMenu();
        }
    });

    // 表格 Tab 导航（事件委托）
    editor.addEventListener('keydown', handleTableKeydown);

    // ===== 表格按钮绑定 =====
    $('insertTableBtn').addEventListener('click', showTablePicker);
    $('tablePickerInsert').addEventListener('click', function() {
        if (pickerRows > 0 && pickerCols > 0) {
            insertTableAtCursor(pickerRows, pickerCols);
            hideTablePicker();
        }
    });
    $('tablePickerCancel').addEventListener('click', hideTablePicker);
    $('tablePickerOverlay').addEventListener('click', function(e) {
        if (e.target === this) hideTablePicker();
    });
    // 按 ESC 关闭
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && tablePickerVisible) {
            hideTablePicker();
        }
    });
    // 初始化网格
    buildTablePickerGrid();

    window.__editor = editor;
    window.__editorGetContent = function() { return editor.innerHTML; };
    window.__editorGetImageData = function() { return imageDataMap; };
    window.__editorGetHeadingConfig = function() { return headingConfig; };
    window.__editorGetBodyFormat = function() { return { font: bodyFont.value, size: bodySize.value, lineHeight: bodyLineHeight.value }; };
    window.__editorShowLoading = showLoading;
    window.__editorHideLoading = hideLoading;
    window.__editorShowToast = showToast;
    window.__saveDocument = saveDocumentState;
    window.__saveDocumentFull = saveDocumentFull;
    window.__loadDocument = loadDocumentState;

    // 强制保存（供导出前调用确保内容最新，使用完整保存）
    window.__flushSave = function() {
        if (autoSaveTimer) { clearTimeout(autoSaveTimer); autoSaveTimer = null; }
        return saveDocumentFull();
    };
    // 暴露标签管理器（供调试和导出使用）
    window.__tabManager = tabManager;
    window.__getActiveSession = function() { return tabManager.getActive(); };

    // ===== 🎨 绘图工具（支持流程图、思维导图） =====
    (function initDrawing() {
        var drawPanel = document.getElementById('drawPanel');
        var drawHeader = document.getElementById('drawPanelHeader');
        var drawClose = document.getElementById('drawPanelClose');
        var drawCanvas = document.getElementById('drawCanvas');
        var drawStatus = document.getElementById('drawStatus');
        var drawColorPicker = document.getElementById('drawColorPicker');
        var drawFillColorPicker = document.getElementById('drawFillColorPicker');
        var drawLineWidth = document.getElementById('drawLineWidth');
        var drawFillToggle = document.getElementById('drawFillToggle');
        var drawUndoBtn = document.getElementById('drawUndoBtn');
        var drawDeleteBtn = document.getElementById('drawDeleteBtn');
        var drawClearBtn = document.getElementById('drawClearBtn');
        var drawCopyBtn = document.getElementById('drawCopyBtn');
        var drawCanvasHint = document.getElementById('drawCanvasHint');
        var drawTextInput = document.getElementById('drawTextInput');
        var drawBtn = document.getElementById('drawBtn');
        var drawPanelOpen = false;

        if (!drawCanvas) return;

        var ctx = drawCanvas.getContext('2d');

        // 离屏 Canvas：仅存储自由手绘笔触，不包含形状
        var offscreen = document.createElement('canvas');
        var offCtx = offscreen.getContext('2d');
        var currentTool = 'pencil';
        var isDrawing = false;
        var startX = 0, startY = 0;
        var lastX = 0, lastY = 0;

        // === 形状对象系统 ===
        var shapes = [];            // { id, type, x, y, w, h, text, fill, stroke, strokeW, fontSize }
        var connectors = [];        // { id, fromId, toId, text, color, strokeW, fromX, fromY, toX, toY }
        var selectedId = null;
        var selectedConnId = null;
        var shapeIdCounter = 0;
        var dragState = null;       // { mode:'move'|'resize', shapeId, offX, offY, handle, startX, startY, startW, startH }
        var connectState = null;    // { fromId, tempX, tempY }
        var textEditShapeId = null;
        // 撤销栈：每个条目 { pixels: ImageData, shapes: clone, connectors: clone }
        var undoStack = [];
        var maxUndo = 30;
        var canvasInitialized = false;
        var CANVAS_W = 800, CANVAS_H = 500;

        // 工具分组
        var FREEHAND_TOOLS = { pencil:1, line:1, arrow:1, eraser:1 };
        var SHAPE_TOOLS = { roundrect:1, process:1, diamond:1, parallelogram:1, document:1, hexagon:1, topic:1, subtopic:1, textbox:1, triangle:1, star:1, pentagon:1, heart:1, cloud:1, cylinder:1 };

        if (drawUndoBtn) drawUndoBtn.disabled = true;

        // ===== Canvas 尺寸管理 =====
        function resizeCanvas() {
            var wrap = drawCanvas.parentElement;
            if (!wrap) return;
            var w = wrap.clientWidth - 16;
            var h = wrap.clientHeight - 16;
            if (w < 200) w = 200;
            if (h < 150) h = 150;
            var ratio = CANVAS_W / CANVAS_H;
            var displayW = w;
            var displayH = displayW / ratio;
            if (displayH > h) { displayH = h; displayW = displayH * ratio; }
            drawCanvas.style.width = Math.floor(displayW) + 'px';
            drawCanvas.style.height = Math.floor(displayH) + 'px';
        }

        // ===== 撤销系统 =====
        function saveState() {
            var pixels = offCtx.getImageData(0, 0, CANVAS_W, CANVAS_H);
            var state = {
                pixels: pixels,
                shapes: JSON.parse(JSON.stringify(shapes)),
                connectors: JSON.parse(JSON.stringify(connectors)),
                selectedId: selectedId,
                selectedConnId: selectedConnId
            };
            if (undoStack.length >= maxUndo) undoStack.shift();
            undoStack.push(state);
            if (drawUndoBtn) drawUndoBtn.disabled = false;
        }

        function undo() {
            if (undoStack.length <= 1) { if (drawUndoBtn) drawUndoBtn.disabled = true; return; }
            if (textEditShapeId) { finishTextEdit(); }
            undoStack.pop();
            var state = undoStack[undoStack.length - 1];
            offCtx.putImageData(state.pixels, 0, 0);
            shapes = JSON.parse(JSON.stringify(state.shapes));
            connectors = JSON.parse(JSON.stringify(state.connectors));
            selectedId = state.selectedId || null;
            selectedConnId = state.selectedConnId || null;
            renderShapes();
            if (undoStack.length <= 1 && drawUndoBtn) drawUndoBtn.disabled = true;
            setDrawStatus('已撤销');
        }

        // ===== 画布初始化 =====
        function initCanvas() {
            drawCanvas.width = CANVAS_W;
            drawCanvas.height = CANVAS_H;
            offscreen.width = CANVAS_W;
            offscreen.height = CANVAS_H;
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, CANVAS_W, CANVAS_H);
            offCtx.fillStyle = '#ffffff';
            offCtx.fillRect(0, 0, CANVAS_W, CANVAS_H);
            shapes = []; connectors = []; selectedId = null; selectedConnId = null; undoStack = [];
            saveState();
            resizeCanvas();
        }

        // ===== 渲染所有内容 =====
        function renderShapes() {
            // 每次完全重绘：先清空，画离屏（仅手绘），再画形状
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, CANVAS_W, CANVAS_H);
            // 将离屏 Canvas（仅自由手绘）绘制到主画布
            ctx.drawImage(offscreen, 0, 0);
            // 画连接线
            connectors.forEach(function(c) { drawConnector(c); });
            // 画形状
            shapes.forEach(function(s) { drawShape(s); });
            // 画选中状态
            if (selectedId) drawSelectionHandles(findShape(selectedId));
            if (selectedConnId) {
                var sc = connectors.find(function(c){ return c.id === selectedConnId; });
                if (sc) drawConnectorSelection(sc);
            }
            // 连接中临时线
            if (connectState) {
                ctx.save();
                ctx.strokeStyle = '#2563eb';
                ctx.lineWidth = 2;
                ctx.setLineDash([5,5]);
                ctx.beginPath();
                var from = findShape(connectState.fromId);
                if (from) {
                    var fc = getShapeCenter(from);
                    ctx.moveTo(fc.x, fc.y);
                    ctx.lineTo(connectState.tempX, connectState.tempY);
                }
                ctx.stroke();
                ctx.restore();
            }
        }

        function findShape(id) { return shapes.find(function(s){ return s.id === id; }); }
        function getShapeCenter(s) { return { x: s.x + s.w/2, y: s.y + s.h/2 }; }

        // ===== 绘制各形状 =====
        function drawShape(s) {
            ctx.save();
            ctx.strokeStyle = s.stroke || '#1e293b';
            ctx.lineWidth = s.strokeW || 2;
            ctx.fillStyle = s.fill || 'transparent';
            ctx.font = (s.fontSize || 14) + 'px -apple-system, BlinkMacSystemFont, sans-serif';
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            var cx = s.x + s.w/2, cy = s.y + s.h/2;

            switch (s.type) {
                case 'roundrect':
                    var r = Math.min(s.w, s.h) * 0.25;
                    roundRect(ctx, s.x, s.y, s.w, s.h, r);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'process':
                    ctx.fillRect(s.x, s.y, s.w, s.h);
                    ctx.strokeRect(s.x, s.y, s.w, s.h);
                    break;
                case 'diamond':
                    ctx.beginPath();
                    ctx.moveTo(cx, s.y); ctx.lineTo(s.x + s.w, cy);
                    ctx.lineTo(cx, s.y + s.h); ctx.lineTo(s.x, cy);
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'parallelogram':
                    var skew = s.w * 0.2;
                    ctx.beginPath();
                    ctx.moveTo(s.x + skew, s.y); ctx.lineTo(s.x + s.w, s.y);
                    ctx.lineTo(s.x + s.w - skew, s.y + s.h); ctx.lineTo(s.x, s.y + s.h);
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'document':
                    var docR = s.h * 0.15;
                    ctx.beginPath();
                    ctx.moveTo(s.x, s.y);
                    ctx.lineTo(s.x + s.w, s.y);
                    ctx.lineTo(s.x + s.w, s.y + s.h - docR);
                    ctx.quadraticCurveTo(s.x + s.w*0.75, s.y + s.h + docR, s.x + s.w*0.5, s.y + s.h - docR);
                    ctx.quadraticCurveTo(s.x + s.w*0.25, s.y + s.h - docR*3, s.x, s.y + s.h - docR);
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'hexagon':
                    var hx = s.w * 0.25;
                    ctx.beginPath();
                    ctx.moveTo(s.x + hx, s.y); ctx.lineTo(s.x + s.w - hx, s.y);
                    ctx.lineTo(s.x + s.w, cy); ctx.lineTo(s.x + s.w - hx, s.y + s.h);
                    ctx.lineTo(s.x + hx, s.y + s.h); ctx.lineTo(s.x, cy);
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'topic':
                case 'subtopic':
                case 'circle':
                    var rx = s.w/2, ry = s.type === 'circle' ? s.w/2 : s.h/2;
                    ctx.beginPath();
                    ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI*2);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'oval':
                    ctx.beginPath();
                    ctx.ellipse(cx, cy, s.w/2, s.h/2, 0, 0, Math.PI*2);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'textbox':
                    // 文本框：只有文字，无边框（选中时边框由 handles 显示）
                    break;
                case 'triangle':
                    ctx.beginPath();
                    ctx.moveTo(cx, s.y); ctx.lineTo(s.x+s.w, s.y+s.h); ctx.lineTo(s.x, s.y+s.h);
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'star':
                    var sp = 5, outerR = Math.min(s.w,s.h)/2, innerR = outerR*0.4;
                    ctx.beginPath();
                    for (var si = 0; si < sp*2; si++) {
                        var sa = (si * Math.PI / sp) - Math.PI/2;
                        var sr = si % 2 === 0 ? outerR : innerR;
                        var sx = cx + sr * Math.cos(sa), sy2 = cy + sr * Math.sin(sa);
                        si === 0 ? ctx.moveTo(sx, sy2) : ctx.lineTo(sx, sy2);
                    }
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'pentagon':
                    ctx.beginPath();
                    for (var pi = 0; pi < 5; pi++) {
                        var pa = (pi * 2 * Math.PI / 5) - Math.PI/2;
                        var px = cx + Math.min(s.w,s.h)/2 * Math.cos(pa);
                        var py = cy + Math.min(s.w,s.h)/2 * Math.sin(pa);
                        pi === 0 ? ctx.moveTo(px, py) : ctx.lineTo(px, py);
                    }
                    ctx.closePath();
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'heart':
                    var hh = Math.min(s.w, s.h) / 2;
                    ctx.beginPath();
                    var htx = cx, hty = s.y + hh * 1.1;
                    ctx.moveTo(htx, hty + hh * 0.3);
                    ctx.bezierCurveTo(htx - hh*1.3, hty - hh*0.7, htx - hh*0.3, hty - hh*1.5, htx, hty - hh*0.7);
                    ctx.bezierCurveTo(htx + hh*0.3, hty - hh*1.5, htx + hh*1.3, hty - hh*0.7, htx, hty + hh*0.3);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'cloud':
                    var cr = Math.min(s.w, s.h) * 0.22;
                    ctx.beginPath();
                    ctx.arc(cx - s.w*0.2, cy + s.h*0.05, cr*1.1, 0, Math.PI*2);
                    ctx.arc(cx + s.w*0.2, cy + s.h*0.05, cr*1.1, 0, Math.PI*2);
                    ctx.arc(cx - s.w*0.05, cy - s.h*0.2, cr*1.2, 0, Math.PI*2);
                    ctx.arc(cx + s.w*0.1, cy - s.h*0.15, cr*1.0, 0, Math.PI*2);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                case 'cylinder':
                    var cycR = s.w * 0.4;
                    var cycTopY = s.y + s.h * 0.15;
                    var cycBotY = s.y + s.h - s.h * 0.1;
                    // 主体
                    ctx.beginPath();
                    ctx.moveTo(cx - cycR, cycTopY);
                    ctx.lineTo(cx - cycR, cycBotY);
                    ctx.quadraticCurveTo(cx - cycR, cycBotY + s.h*0.1, cx, cycBotY + s.h*0.1);
                    ctx.quadraticCurveTo(cx + cycR, cycBotY + s.h*0.1, cx + cycR, cycBotY);
                    ctx.lineTo(cx + cycR, cycTopY);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    // 顶部椭圆
                    ctx.beginPath();
                    ctx.ellipse(cx, cycTopY, cycR, s.h*0.12, 0, 0, Math.PI*2);
                    if (s.fill !== 'transparent') ctx.fill();
                    ctx.stroke();
                    break;
                default:
                    ctx.fillRect(s.x, s.y, s.w, s.h);
                    ctx.strokeRect(s.x, s.y, s.w, s.h);
            }

            // 文字（如果有）
            if (s.text) {
                ctx.fillStyle = '#1e293b';
                ctx.font = (s.fontSize || 14) + 'px -apple-system, BlinkMacSystemFont, sans-serif';
                ctx.textAlign = 'center';
                ctx.textBaseline = 'middle';
                wrapText(ctx, s.text, cx, cy, s.w - 10, s.fontSize || 14);
            }
            ctx.restore();
        }

        function roundRect(ctx, x, y, w, h, r) {
            ctx.beginPath();
            ctx.moveTo(x + r, y);
            ctx.lineTo(x + w - r, y);
            ctx.quadraticCurveTo(x + w, y, x + w, y + r);
            ctx.lineTo(x + w, y + h - r);
            ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
            ctx.lineTo(x + r, y + h);
            ctx.quadraticCurveTo(x, y + h, x, y + h - r);
            ctx.lineTo(x, y + r);
            ctx.quadraticCurveTo(x, y, x + r, y);
            ctx.closePath();
        }

        function wrapText(ctx, text, x, y, maxWidth, fontSize) {
            var lines = text.split('\n');
            var lh = fontSize + 4;
            var startY = y - ((lines.length - 1) * lh) / 2;
            lines.forEach(function(line, i) {
                ctx.fillText(line, x, startY + i * lh);
            });
        }

        // ===== 绘制连接线 =====
        function drawConnector(c) {
            var from = findShape(c.fromId);
            var to = findShape(c.toId);
            if (!from && !to) return;
            ctx.save();
            ctx.strokeStyle = c.color || '#64748b';
            ctx.lineWidth = c.strokeW || 2;
            ctx.beginPath();
            if (from && to) {
                var f = getConnectionPoint(from, to);
                var t = getConnectionPoint(to, from);
                ctx.moveTo(f.x, f.y);
                // 贝塞尔曲线
                var dx = Math.abs(t.x - f.x) * 0.4;
                ctx.bezierCurveTo(f.x + dx, f.y, t.x - dx, t.y, t.x, t.y);
            } else if (from) {
                var f2 = getShapeCenter(from);
                ctx.moveTo(f2.x, f2.y);
                ctx.lineTo(c.toX || f2.x + 60, c.toY || f2.y + 60);
            } else if (to) {
                var t2 = getShapeCenter(to);
                ctx.moveTo(c.fromX || t2.x - 60, c.fromY || t2.y - 60);
                ctx.lineTo(t2.x, t2.y);
            }
            ctx.stroke();
            // 箭头
            if (to) {
                var tp = getConnectionPoint(to, from || { x: c.fromX, y: c.fromY });
                var angle = Math.atan2(
                    (from ? getConnectionPoint(from, to).y : (c.fromY || 0)) - tp.y,
                    (from ? getConnectionPoint(from, to).x : (c.fromX || 0)) - tp.x
                );
                drawArrowhead(ctx, tp.x, tp.y, angle, ctx.lineWidth * 2.5);
            }
            // 标签
            if (c.text) {
                var fromC = from ? getShapeCenter(from) : { x: c.fromX || 0, y: c.fromY || 0 };
                var toC = to ? getShapeCenter(to) : { x: c.toX || 0, y: c.toY || 0 };
                var mx = (fromC.x + toC.x) / 2, my = (fromC.y + toC.y) / 2;
                ctx.fillStyle = '#1e293b';
                ctx.font = '12px sans-serif';
                ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
                ctx.fillText(c.text, mx, my - 4);
            }
            ctx.restore();
        }

        function getConnectionPoint(from, to) {
            var fc = getShapeCenter(from);
            var tc = getShapeCenter(to);
            var dx = tc.x - fc.x, dy = tc.y - fc.y;
            if (dx === 0 && dy === 0) return fc;
            // 对矩形形状，计算边界交点
            var absDx = Math.abs(dx), absDy = Math.abs(dy);
            var hw = from.w / 2, hh = from.h / 2;
            // 对于椭圆/圆形
            if (from.type === 'topic' || from.type === 'subtopic' || from.type === 'circle' || from.type === 'oval') {
                var rx = hw, ry = from.type === 'circle' ? hw : hh;
                var angle = Math.atan2(dy, dx);
                return { x: fc.x + rx * Math.cos(angle), y: fc.y + ry * Math.sin(angle) };
            }
            // 对于菱形
            if (from.type === 'diamond') {
                var dAngle = Math.atan2(dy, dx);
                var cosA = Math.cos(dAngle), sinA = Math.sin(dAngle);
                var scale = Math.max(Math.abs(cosA)/hw + Math.abs(sinA)/hh, 0.001);
                return { x: fc.x + cosA/scale, y: fc.y + sinA/scale };
            }
            // 默认矩形
            var scaleX = absDx / hw, scaleY = absDy / hh;
            var scale = Math.max(scaleX, scaleY);
            if (scale < 0.001) return fc;
            return { x: fc.x + dx/scale, y: fc.y + dy/scale };
        }

        function drawArrowhead(ctx, x, y, angle, size) {
            ctx.save();
            ctx.translate(x, y);
            ctx.rotate(angle);
            ctx.beginPath();
            ctx.moveTo(0, 0);
            ctx.lineTo(-size, -size/2.5);
            ctx.lineTo(-size, size/2.5);
            ctx.closePath();
            ctx.fill();
            ctx.restore();
        }

        // ===== 选中高亮 =====
        function drawSelectionHandles(s) {
            if (!s) return;
            ctx.save();
            var hs = 8, hh = 4;
            ctx.strokeStyle = '#2563eb';
            ctx.lineWidth = 2;
            ctx.setLineDash([4,3]);
            ctx.strokeRect(s.x, s.y, s.w, s.h);
            ctx.setLineDash([]);
            ctx.fillStyle = '#ffffff';
            ctx.strokeStyle = '#2563eb';
            ctx.lineWidth = 2;
            // 8 个控制点
            var handles = [
                { x: s.x - hh, y: s.y - hh }, { x: s.x + s.w/2 - hh, y: s.y - hh }, { x: s.x + s.w - hh, y: s.y - hh },
                { x: s.x - hh, y: s.y + s.h/2 - hh }, { x: s.x + s.w - hh, y: s.y + s.h/2 - hh },
                { x: s.x - hh, y: s.y + s.h - hh }, { x: s.x + s.w/2 - hh, y: s.y + s.h - hh }, { x: s.x + s.w - hh, y: s.y + s.h - hh }
            ];
            handles.forEach(function(h) {
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(h.x, h.y, hs, hs);
                ctx.strokeRect(h.x, h.y, hs, hs);
            });
            ctx.restore();
        }

        function drawConnectorSelection(c) {
            if (!c) return;
            ctx.save();
            ctx.strokeStyle = '#2563eb';
            ctx.lineWidth = 2;
            ctx.setLineDash([4,3]);
            var from = findShape(c.fromId), to = findShape(c.toId);
            if (from && to) {
                var f = getConnectionPoint(from, to), t = getConnectionPoint(to, from);
                ctx.beginPath(); ctx.moveTo(f.x, f.y);
                ctx.bezierCurveTo(f.x + Math.abs(t.x-f.x)*0.4, f.y, t.x - Math.abs(t.x-f.x)*0.4, t.y, t.x, t.y);
                ctx.stroke();
            }
            ctx.setLineDash([]);
            ctx.restore();
        }

        // ===== 碰撞检测 =====
        function hitTestShape(x, y) {
            for (var i = shapes.length - 1; i >= 0; i--) {
                var s = shapes[i];
                if (x >= s.x && x <= s.x + s.w && y >= s.y && y <= s.y + s.h) return s;
            }
            return null;
        }

        function hitTestConnector(x, y) {
            for (var i = connectors.length - 1; i >= 0; i--) {
                var c = connectors[i];
                var from = findShape(c.fromId), to = findShape(c.toId);
                if (!from || !to) continue;
                var f = getConnectionPoint(from, to);
                var t = getConnectionPoint(to, from);
                // 近似碰撞
                var mx = (f.x + t.x) / 2, my = (f.y + t.y) / 2;
                if (Math.abs(x - mx) < 8 && Math.abs(y - my) < 8) return c;
            }
            return null;
        }

        function getResizeHandle(x, y, s) {
            var hs = 10;
            var handles = [
                { name:'nw', x:s.x, y:s.y }, { name:'n', x:s.x+s.w/2, y:s.y }, { name:'ne', x:s.x+s.w, y:s.y },
                { name:'w', x:s.x, y:s.y+s.h/2 }, { name:'e', x:s.x+s.w, y:s.y+s.h/2 },
                { name:'sw', x:s.x, y:s.y+s.h }, { name:'s', x:s.x+s.w/2, y:s.y+s.h }, { name:'se', x:s.x+s.w, y:s.y+s.h }
            ];
            for (var i = 0; i < handles.length; i++) {
                if (Math.abs(x - handles[i].x) < hs && Math.abs(y - handles[i].y) < hs) return handles[i].name;
            }
            return null;
        }

        // ===== 创建形状 =====
        function createShape(type, x, y) {
            var id = 'shape_' + (++shapeIdCounter);
            var w = 100, h = 60;
            if (type === 'topic') { w = 120; h = 70; }
            if (type === 'subtopic') { w = 100; h = 50; }
            if (type === 'diamond') { w = 100; h = 80; }
            if (type === 'textbox') { w = 120; h = 50; }
            if (type === 'star') { w = 80; h = 80; }
            if (type === 'heart') { w = 80; h = 72; }
            if (type === 'cloud') { w = 130; h = 80; }
            if (type === 'cylinder') { w = 100; h = 80; }
            if (type === 'connector') return null;
            var defaultFill = drawFillToggle && drawFillToggle.checked ? (drawFillColorPicker ? drawFillColorPicker.value : '#dbeafe') : 'transparent';
            // 文本框默认无边框、白色半透明背景
            if (type === 'textbox') {
                defaultFill = 'rgba(255,255,255,0.01)'; // 几乎透明，但能选中
            }
            return {
                id: id, type: type,
                x: x - w/2, y: y - h/2,
                w: w, h: h,
                text: getDefaultText(type),
                fill: defaultFill,
                stroke: type === 'textbox' ? 'transparent' : (drawColorPicker ? drawColorPicker.value : '#1e293b'),
                strokeW: type === 'textbox' ? 0 : parseInt(drawLineWidth ? drawLineWidth.value : '2'),
                fontSize: type === 'topic' ? 18 : (type === 'subtopic' ? 15 : 14)
            };
        }

        function getDefaultText(type) {
            var map = {
                roundrect: '开始',
                process: '处理',
                diamond: '判断',
                parallelogram: '输入',
                document: '文档',
                hexagon: '预定义',
                topic: '中心主题',
                subtopic: '子主题',
                textbox: '文本框',
                triangle: '三角形',
                star: '星形',
                pentagon: '五边形',
                heart: '心形',
                cloud: '云形',
                cylinder: '数据库'
            };
            return map[type] || '';
        }

        function addShapeToCanvas(shape) {
            shapes.push(shape);
            selectedId = shape.id;
            selectedConnId = null;
            saveState();
            renderShapes();
            setDrawStatus('已添加 ' + getShapeLabel(shape.type));
        }

        function getShapeLabel(type) {
            var map = { roundrect:'开始/结束', process:'处理', diamond:'判断', parallelogram:'输入/输出', document:'文档', hexagon:'预定义', topic:'主题', subtopic:'子主题', textbox:'文本框', triangle:'三角形', star:'星形', pentagon:'五边形', heart:'心形', cloud:'云形', cylinder:'圆柱' };
            return map[type] || type;
        }

        // ===== 鼠标事件 =====
        function getCanvasCoords(e) {
            var rect = drawCanvas.getBoundingClientRect();
            return {
                x: (e.clientX - rect.left) * (CANVAS_W / rect.width),
                y: (e.clientY - rect.top) * (CANVAS_H / rect.height)
            };
        }

        function isShapeTool(tool) { return SHAPE_TOOLS[tool] || false; }
        function isFreehandTool(tool) { return FREEHAND_TOOLS[tool] || false; }

        function startDrawing(e) {
            if (e.button !== 0) return;
            e.preventDefault();
            if (textEditShapeId) { finishTextEdit(); }
            var coords = getCanvasCoords(e);

            // === 选择工具 ===
            if (currentTool === 'select') {
                // 先检测控制点
                if (selectedId) {
                    var sel = findShape(selectedId);
                    if (sel) {
                        var handle = getResizeHandle(coords.x, coords.y, sel);
                        if (handle) {
                            dragState = { mode:'resize', shapeId:selectedId, handle:handle,
                                startX:coords.x, startY:coords.y,
                                startW:sel.w, startH:sel.h, startSx:sel.x, startSy:sel.y };
                            return;
                        }
                    }
                }
                // 检测形状
                var hit = hitTestShape(coords.x, coords.y);
                if (hit) {
                    selectedId = hit.id;
                    selectedConnId = null;
                    renderShapes();
                    dragState = { mode:'move', shapeId:hit.id, offX:coords.x - hit.x, offY:coords.y - hit.y };
                    return;
                }
                // 检测连接线
                var connHit = hitTestConnector(coords.x, coords.y);
                if (connHit) {
                    selectedConnId = connHit.id;
                    selectedId = null;
                    renderShapes();
                    return;
                }
                selectedId = null; selectedConnId = null;
                renderShapes();
                return;
            }

            // === 连接器工具 ===
            if (currentTool === 'connector') {
                var hit = hitTestShape(coords.x, coords.y);
                if (!hit) return;
                if (!connectState) {
                    connectState = { fromId: hit.id, tempX: coords.x, tempY: coords.y };
                    setDrawStatus('选择目标形状');
                } else {
                    if (hit.id !== connectState.fromId) {
                        connectors.push({
                            id: 'conn_' + (++shapeIdCounter),
                            fromId: connectState.fromId,
                            toId: hit.id,
                            text: '',
                            color: drawColorPicker ? drawColorPicker.value : '#64748b',
                            strokeW: parseInt(drawLineWidth ? drawLineWidth.value : '2')
                        });
                        saveState();
                        renderShapes();
                        setDrawStatus('已连接');
                    }
                    connectState = null;
                }
                renderShapes();
                return;
            }

            // === 形状工具 ===
            if (isShapeTool(currentTool)) {
                var shape = createShape(currentTool, coords.x, coords.y);
                if (shape) addShapeToCanvas(shape);
                return;
            }

            // === 自由手绘工具 ===
            if (isFreehandTool(currentTool)) {
                isDrawing = true;
                startX = coords.x; startY = coords.y;
                lastX = coords.x; lastY = coords.y;
                if (drawCanvasHint) drawCanvasHint.classList.add('hidden');
                if (currentTool === 'pencil' || currentTool === 'eraser') {
                    ctx.beginPath();
                    ctx.moveTo(coords.x, coords.y);
                }
            }
        }

        function draw(e) {
            if (!isDrawing) return;
            e.preventDefault();
            var coords = getCanvasCoords(e);
            var color = drawColorPicker ? drawColorPicker.value : '#1e293b';
            var lw = parseInt(drawLineWidth ? drawLineWidth.value : '2');

            if (currentTool === 'pencil') {
                ctx.strokeStyle = color;
                ctx.lineWidth = lw;
                ctx.lineCap = 'round'; ctx.lineJoin = 'round';
                ctx.globalCompositeOperation = 'source-over';
                ctx.lineTo(coords.x, coords.y);
                ctx.stroke();
                ctx.beginPath();
                ctx.moveTo(coords.x, coords.y);
                // 同步到离屏
                offCtx.strokeStyle = color;
                offCtx.lineWidth = lw;
                offCtx.lineCap = 'round'; offCtx.lineJoin = 'round';
                offCtx.globalCompositeOperation = 'source-over';
                offCtx.lineTo(coords.x, coords.y);
                offCtx.stroke();
                offCtx.beginPath();
                offCtx.moveTo(coords.x, coords.y);
                lastX = coords.x; lastY = coords.y;
            } else if (currentTool === 'eraser') {
                ctx.strokeStyle = '#ffffff';
                ctx.lineWidth = lw * 3;
                ctx.lineCap = 'round'; ctx.lineJoin = 'round';
                ctx.globalCompositeOperation = 'source-over';
                ctx.lineTo(coords.x, coords.y);
                ctx.stroke();
                ctx.beginPath();
                ctx.moveTo(coords.x, coords.y);
                offCtx.strokeStyle = '#ffffff';
                offCtx.lineWidth = lw * 3;
                offCtx.lineCap = 'round'; offCtx.lineJoin = 'round';
                offCtx.globalCompositeOperation = 'source-over';
                offCtx.lineTo(coords.x, coords.y);
                offCtx.stroke();
                offCtx.beginPath();
                offCtx.moveTo(coords.x, coords.y);
                lastX = coords.x; lastY = coords.y;
            } else if (currentTool === 'line' || currentTool === 'arrow') {
                lastX = coords.x; lastY = coords.y;
                // 预览：重绘场景 + 临时线
                renderShapes();
                ctx.save();
                ctx.strokeStyle = color;
                ctx.lineWidth = lw;
                ctx.lineCap = 'round';
                ctx.globalCompositeOperation = 'source-over';
                ctx.beginPath();
                ctx.moveTo(startX, startY);
                ctx.lineTo(coords.x, coords.y);
                ctx.stroke();
                if (currentTool === 'arrow') {
                    var angle = Math.atan2(coords.y - startY, coords.x - startX);
                    drawArrowhead(ctx, coords.x, coords.y, angle, lw * 3);
                }
                ctx.restore();
            }
        }

        function endDrawing(e) {
            if (!isDrawing) return;
            isDrawing = false;
            if (currentTool === 'pencil' || currentTool === 'eraser') {
                ctx.closePath();
                offCtx.closePath();
                saveState();
                renderShapes();
                setDrawStatus('');
            } else if (currentTool === 'line' || currentTool === 'arrow') {
                // 最终线条写入离屏
                var color = drawColorPicker ? drawColorPicker.value : '#1e293b';
                var lw = parseInt(drawLineWidth ? drawLineWidth.value : '2');
                offCtx.save();
                offCtx.strokeStyle = color;
                offCtx.lineWidth = lw;
                offCtx.lineCap = 'round';
                offCtx.globalCompositeOperation = 'source-over';
                offCtx.beginPath();
                offCtx.moveTo(startX, startY);
                offCtx.lineTo(lastX || startX, lastY || startY);
                offCtx.stroke();
                if (currentTool === 'arrow') {
                    var angle = Math.atan2((lastY||startY) - startY, (lastX||startX) - startX);
                    offCtx.translate((lastX||startX), (lastY||startY));
                    offCtx.rotate(angle);
                    offCtx.beginPath();
                    offCtx.moveTo(0, 0);
                    offCtx.lineTo(-lw*3, -lw*1.2);
                    offCtx.lineTo(-lw*3, lw*1.2);
                    offCtx.closePath();
                    offCtx.fill();
                }
                offCtx.restore();
                saveState();
                renderShapes();
                setDrawStatus('');
            }
        }

        // 选择工具：拖拽移动/缩放
        function handleSelectDrag(e) {
            if (!dragState) return;
            e.preventDefault();
            var coords = getCanvasCoords(e);
            var s = findShape(dragState.shapeId);
            if (!s) return;

            if (dragState.mode === 'move') {
                // 移动时保持在画布内
                var newX = coords.x - dragState.offX;
                var newY = coords.y - dragState.offY;
                newX = Math.max(0, Math.min(newX, CANVAS_W - s.w));
                newY = Math.max(0, Math.min(newY, CANVAS_H - s.h));
                s.x = newX;
                s.y = newY;
                renderShapes();
            } else if (dragState.mode === 'resize') {
                var dx = coords.x - dragState.startX;
                var dy = coords.y - dragState.startY;
                var h = dragState.handle;
                var newX = dragState.startSx, newY = dragState.startSy, newW = dragState.startW, newH = dragState.startH;
                if (h.indexOf('e') >= 0) newW = Math.max(40, dragState.startW + dx);
                if (h.indexOf('w') >= 0) { newW = Math.max(40, dragState.startW - dx); newX = dragState.startSx + dragState.startW - newW; }
                if (h.indexOf('s') >= 0) newH = Math.max(30, dragState.startH + dy);
                if (h.indexOf('n') >= 0) { newH = Math.max(30, dragState.startH - dy); newY = dragState.startSy + dragState.startH - newH; }
                s.x = newX; s.y = newY; s.w = newW; s.h = newH;
                renderShapes();
            }
        }

        function endSelectDrag(e) {
            if (!dragState) return;
            dragState = null;
            saveState();
            renderShapes();
        }

        // 连接器拖拽临时线
        function handleConnectorMove(e) {
            if (!connectState) return;
            e.preventDefault();
            var coords = getCanvasCoords(e);
            connectState.tempX = coords.x;
            connectState.tempY = coords.y;
            renderShapes();
        }


        // ===== 文字编辑 =====
        function startTextEdit(shape) {
            if (!shape || !drawTextInput) return;
            textEditShapeId = shape.id;
            var rect = drawCanvas.getBoundingClientRect();
            var scaleX = rect.width / CANVAS_W;
            var scaleY = rect.height / CANVAS_H;
            var pad = 4;
            drawTextInput.style.left = (rect.left + (shape.x + pad) * scaleX) + 'px';
            drawTextInput.style.top = (rect.top + (shape.y + pad) * scaleY) + 'px';
            drawTextInput.style.width = Math.max(20, (shape.w - pad*2) * scaleX) + 'px';
            drawTextInput.style.height = Math.max(20, (shape.h - pad*2) * scaleY) + 'px';
            drawTextInput.style.fontSize = (shape.fontSize || 14) + 'px';
            drawTextInput.textContent = shape.text || '';
            drawTextInput.classList.remove('hidden');
            setTimeout(function() { drawTextInput.focus(); }, 10);
        }

        // 文字输入框失焦时完成编辑
        if (drawTextInput) {
            drawTextInput.addEventListener('blur', function() {
                // 短暂延迟让点击事件先触发
                setTimeout(function() {
                    if (textEditShapeId) finishTextEdit();
                }, 100);
            });
        }

        function finishTextEdit() {
            if (!textEditShapeId || !drawTextInput) return;
            var s = findShape(textEditShapeId);
            if (s) {
                s.text = drawTextInput.textContent || '';
                saveState();
                renderShapes();
            }
            drawTextInput.classList.add('hidden');
            drawTextInput.textContent = '';
            textEditShapeId = null;
        }

        // ===== 清空 =====
        function clearCanvas() {
            if (textEditShapeId) finishTextEdit();
            if (!confirm('确定清空画布？将清除所有绘图和形状。')) return;
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, CANVAS_W, CANVAS_H);
            offCtx.fillStyle = '#ffffff';
            offCtx.fillRect(0, 0, CANVAS_W, CANVAS_H);
            shapes = []; connectors = []; selectedId = null; selectedConnId = null; connectState = null;
            saveState();
            renderShapes();
            setDrawStatus('已清空');
        }

        // ===== 删除选中 =====
        function deleteSelected() {
            if (textEditShapeId) { finishTextEdit(); return; }
            if (selectedId) {
                var idx = shapes.findIndex(function(s){ return s.id === selectedId; });
                if (idx >= 0) {
                    // 删除相关连接
                    connectors = connectors.filter(function(c){ return c.fromId !== selectedId && c.toId !== selectedId; });
                    shapes.splice(idx, 1);
                    selectedId = null;
                    saveState();
                    renderShapes();
                    setDrawStatus('已删除形状');
                }
            } else if (selectedConnId) {
                var cidx = connectors.findIndex(function(c){ return c.id === selectedConnId; });
                if (cidx >= 0) { connectors.splice(cidx, 1); selectedConnId = null; saveState(); renderShapes(); setDrawStatus('已删除连线'); }
            }
        }

        // ===== 复制到剪贴板 =====
        function copyToClipboard() {
            drawCanvas.toBlob(function(blob) {
                if (!blob) { showToast('复制失败', 'error'); return; }
                try {
                    navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })])
                        .then(function() { showToast('✅ 图片已复制到剪贴板，可粘贴到编辑器中', 'success', 3000); setDrawStatus('已复制'); })
                        .catch(function(err) {
                            var url = URL.createObjectURL(blob);
                            var a = document.createElement('a');
                            a.href = url; a.download = 'drawing.png';
                            document.body.appendChild(a); a.click(); document.body.removeChild(a);
                            URL.revokeObjectURL(url);
                            showToast('已导出为 PNG 文件', 'info', 3000);
                        });
                } catch(e) { showToast('复制失败', 'error'); }
            }, 'image/png');
        }

        // ===== 工具切换 =====
        function setTool(tool) {
            if (textEditShapeId) finishTextEdit();
            if (connectState) { connectState = null; renderShapes(); }
            currentTool = tool;
            drawPanel.querySelectorAll('.draw-tool-btn.active').forEach(function(b) { b.classList.remove('active'); });
            var btn = drawPanel.querySelector('.draw-tool-btn[data-tool="' + tool + '"]');
            if (btn) btn.classList.add('active');

            if (tool === 'select') {
                drawCanvas.style.cursor = 'default';
                setDrawStatus('点击选择形状，拖拽移动');
            } else if (isShapeTool(tool) || tool === 'connector') {
                drawCanvas.style.cursor = 'crosshair';
                selectedId = null; selectedConnId = null; connectState = null;
                renderShapes();
                setDrawStatus('单击添加 ' + (tool==='connector'?'连接线':getShapeLabel(tool)));
            } else if (tool === 'eraser') {
                drawCanvas.style.cursor = 'cell';
                setDrawStatus('橡皮擦');
            } else {
                drawCanvas.style.cursor = 'crosshair';
                if (tool === 'pencil') setDrawStatus('画笔');
                else if (tool === 'line') setDrawStatus('直线');
                else if (tool === 'arrow') setDrawStatus('箭头');
            }
        }

        function setDrawStatus(msg) {
            if (drawStatus) drawStatus.textContent = msg || '就绪';
        }

        // ===== 面板控制 =====
        function showDrawPanel() {
            if (drawPanelOpen) return;
            drawPanelOpen = true;
            drawPanel.classList.remove('hidden');
            drawPanel.style.left = 'auto';
            drawPanel.style.right = '40px';
            drawPanel.style.top = '80px';
            resizeCanvas();
            if (!canvasInitialized) {
                initCanvas();
                canvasInitialized = true;
            }
            setDrawStatus('就绪');
            // 恢复工具状态
            setTool(currentTool);
        }

        function hideDrawPanel() {
            if (textEditShapeId) finishTextEdit();
            drawPanelOpen = false;
            drawPanel.classList.add('hidden');
        }

        if (drawBtn) drawBtn.addEventListener('click', showDrawPanel);
        if (drawClose) drawClose.addEventListener('click', hideDrawPanel);
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && drawPanelOpen) hideDrawPanel();
        });

        // ===== 拖动面板 =====
        if (drawHeader) {
            drawHeader.addEventListener('mousedown', function(e) {
                if (e.button !== 0 || e.target.closest('button')) return;
                var rect = drawPanel.getBoundingClientRect();
                var dx = e.clientX - rect.left, dy = e.clientY - rect.top;
                function onMove(ev) { drawPanel.style.left = (ev.clientX - dx) + 'px'; drawPanel.style.top = (ev.clientY - dy) + 'px'; drawPanel.style.right = 'auto'; }
                function onUp() { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
                document.addEventListener('mousemove', onMove); document.addEventListener('mouseup', onUp);
                e.preventDefault();
            });
        }

        // ===== 画布事件 =====
        drawCanvas.addEventListener('mousedown', startDrawing);
        drawCanvas.addEventListener('mousemove', function(e) {
            if (currentTool === 'select') { handleSelectDrag(e); return; }
            if (currentTool === 'connector' && connectState) { handleConnectorMove(e); return; }
            draw(e);
        });
        drawCanvas.addEventListener('mouseup', function(e) {
            if (currentTool === 'select') { endSelectDrag(e); return; }
            endDrawing(e);
        });
        drawCanvas.addEventListener('mouseleave', function(e) {
            if (currentTool === 'select' && dragState) { endSelectDrag(e); return; }
            endDrawing(e);
        });

        // 双击编辑文字
        drawCanvas.addEventListener('dblclick', function(e) {
            if (currentTool !== 'select') return;
            var coords = getCanvasCoords(e);
            var hit = hitTestShape(coords.x, coords.y);
            if (hit) startTextEdit(hit);
        });

        // 键盘删除
        document.addEventListener('keydown', function delKeyHandler(e) {
            if (!drawPanelOpen) return;
            if ((e.key === 'Delete' || e.key === 'Backspace') && (selectedId || selectedConnId)) {
                e.preventDefault();
                deleteSelected();
            }
            // 回车完成文字编辑
            if (e.key === 'Enter' && textEditShapeId && !e.shiftKey) {
                e.preventDefault();
                finishTextEdit();
            }
        });

        // 触摸支持
        drawCanvas.addEventListener('touchstart', function(e) {
            e.preventDefault();
            var touch = e.touches[0];
            startDrawing({ button: 0, clientX: touch.clientX, clientY: touch.clientY, preventDefault: function(){}, target: drawCanvas });
        });
        drawCanvas.addEventListener('touchmove', function(e) {
            e.preventDefault();
            var touch = e.touches[0];
            if (currentTool === 'select') { handleSelectDrag({ preventDefault:function(){}, clientX:touch.clientX, clientY:touch.clientY }); return; }
            draw({ preventDefault: function(){}, clientX: touch.clientX, clientY: touch.clientY });
        });
        drawCanvas.addEventListener('touchend', function(e) {
            e.preventDefault();
            if (currentTool === 'select') { endSelectDrag({ preventDefault:function(){} }); return; }
            endDrawing({ preventDefault: function(){}, clientX: lastX, clientY: lastY });
        });

        // ===== 工具按钮事件 =====
        drawPanel.querySelectorAll('.draw-tool-btn[data-tool]').forEach(function(btn) {
            btn.addEventListener('click', function() { setTool(this.dataset.tool); });
        });

        if (drawUndoBtn) drawUndoBtn.addEventListener('click', undo);
        if (drawDeleteBtn) drawDeleteBtn.addEventListener('click', function() { if (selectedId || selectedConnId) deleteSelected(); else showToast('请先选中要删除的组件', 'info', 1500); });
        if (drawClearBtn) drawClearBtn.addEventListener('click', clearCanvas);
        if (drawCopyBtn) drawCopyBtn.addEventListener('click', copyToClipboard);

        // ===== 窗口变化 =====
        window.addEventListener('resize', function() {
            if (drawPanelOpen) setTimeout(resizeCanvas, 200);
        });

        // ===== 键盘快捷键 =====
        document.addEventListener('keydown', function drawKeyHandler(e) {
            if (!drawPanelOpen) return;
            var toolMap = {
                's':'select', 'p':'pencil', 'l':'line', 'a':'arrow', 'e':'eraser',
                'r':'roundrect', 'd':'diamond', 'h':'hexagon',
                't':'topic', 'u':'subtopic', 'c':'connector',
                '1':'textbox', '2':'triangle', '3':'star', '4':'pentagon', '5':'heart', '6':'cloud', '7':'cylinder'
            };
            var tool = toolMap[e.key.toLowerCase()];
            if (tool && !e.ctrlKey && !e.metaKey && !e.altKey && !e.shiftKey) {
                e.preventDefault();
                setTool(tool);
                return;
            }
            if ((e.ctrlKey || e.metaKey) && (e.key === 'z' || e.key === 'Z')) { e.preventDefault(); undo(); }
        });

        // ===== 观察面板显示 =====
        var panelObserver = new MutationObserver(function() {
            if (!drawPanel.classList.contains('hidden')) setTimeout(resizeCanvas, 50);
        });
        panelObserver.observe(drawPanel, { attributes: true, attributeFilter: ['class'] });

        // ===== 初始设置 =====
        setTool('pencil');
        console.log('Drawing tool initialized with shapes');
    })();

    // ===== 🖌️ 格式刷（完整版 — 捕获并应用所有文本/段落格式） =====
    (function initFormatPainter() {
        var painterBtn = document.getElementById('formatPainterBtn');
        if (!painterBtn) return;

        var picking = false;
        var applying = false;
        var copiedStyle = null;   // 完整格式快照
        var painterActive = false;

        // ===== 要捕获的样式属性清单 =====
        // 文本级
        var TEXT_PROPS = [
            'fontFamily','fontSize','fontWeight','fontStyle',
            'color','backgroundColor',
            'textDecorationLine','textDecorationStyle','textDecorationColor','textDecoration',
            'textTransform','letterSpacing','wordSpacing','verticalAlign'
        ];
        // 段落/块级
        var BLOCK_PROPS = [
            'textAlign','textIndent','lineHeight','whiteSpace',
            'borderTopWidth','borderTopStyle','borderTopColor',
            'borderBottomWidth','borderBottomStyle','borderBottomColor',
            'borderLeftWidth','borderLeftStyle','borderLeftColor',
            'borderRightWidth','borderRightStyle','borderRightColor',
            'paddingTop','paddingBottom','paddingLeft','paddingRight',
            'marginTop','marginBottom','marginLeft','marginRight',
            'listStyleType'
        ];
        var ALL_PROPS = TEXT_PROPS.concat(BLOCK_PROPS);

        // ===== 重置 =====
        function resetPainter() {
            picking = false;
            applying = false;
            copiedStyle = null;
            painterActive = false;
            painterBtn.classList.remove('active');
            editor.style.cursor = '';
            painterBtn.title = '格式刷 — 复制格式';
            document.querySelectorAll('.format-painter-highlight').forEach(function(el) { el.classList.remove('format-painter-highlight'); });
            setStatus('就绪');
        }

        // ===== 从 DOM 元素提取完整样式 =====
        function extractStyle(el) {
            var style = {};
            style.tag = el.tagName;
            style.isHeading = /^H[1-6]$/i.test(el.tagName);

            var cs = window.getComputedStyle(el);

            // 批量读取所有属性（先取 computed，再被 inline 覆盖）
            ALL_PROPS.forEach(function(prop) {
                style[prop] = cs[prop] || '';
            });

            // 是否为粗体（强化判断）
            var fw = cs.fontWeight;
            style.isBold = (fw === 'bold' || fw === '700' || fw === '800' || fw === '900' || parseInt(fw) >= 600);

            // 是否有下划线
            var deco = (cs.textDecorationLine || cs.textDecoration || '').toLowerCase();
            style.isUnderline = deco.indexOf('underline') >= 0;
            style.isLineThrough = deco.indexOf('line-through') >= 0;
            style.isOverline = deco.indexOf('overline') >= 0;

            // inline 样式优先覆盖（用户手动设置过的才真正算数）
            var inline = el.style;
            if (inline.fontFamily) style.fontFamily = inline.fontFamily;
            if (inline.fontSize) style.fontSize = inline.fontSize;
            if (inline.fontWeight) style.fontWeight = inline.fontWeight;
            if (inline.fontStyle) style.fontStyle = inline.fontStyle;
            if (inline.color) style.color = inline.color;
            if (inline.backgroundColor) style.backgroundColor = inline.backgroundColor;
            if (inline.textAlign) style.textAlign = inline.textAlign;
            if (inline.lineHeight) style.lineHeight = inline.lineHeight;
            if (inline.textIndent) style.textIndent = inline.textIndent;
            if (inline.letterSpacing) style.letterSpacing = inline.letterSpacing;
            if (inline.verticalAlign) style.verticalAlign = inline.verticalAlign;
            // textDecoration — inline 中可能是组合字符串
            if (inline.textDecoration || inline.textDecorationLine) {
                var td = (inline.textDecoration || inline.textDecorationLine || '').toLowerCase();
                style.isUnderline = td.indexOf('underline') >= 0;
                style.isLineThrough = td.indexOf('line-through') >= 0;
            }
            // border — 四个方向
            ['Top','Bottom','Left','Right'].forEach(function(side) {
                var w = inline['border' + side + 'Width'];
                var s = inline['border' + side + 'Style'];
                var c = inline['border' + side + 'Color'];
                if (w) style['border' + side + 'Width'] = w;
                if (s) style['border' + side + 'Style'] = s;
                if (c) style['border' + side + 'Color'] = c;
            });
            // padding
            ['Top','Bottom','Left','Right'].forEach(function(side) {
                if (inline['padding' + side]) style['padding' + side] = inline['padding' + side];
            });

            return style;
        }

        // ===== 应用样式到目标元素 =====
        function applyStyle(el) {
            if (!copiedStyle) return;
            saveUndoState('格式刷'); // 记录格式刷应用前状态

            // 标题复制 → 目标转为同级别标题
            if (copiedStyle.isHeading) {
                var block = el;
                while (block && block !== editor && !/^H[1-6]$|^P$|^DIV$|^LI$|^TD$/i.test(block.tagName)) block = block.parentNode;
                if (block && block !== editor) {
                    var tag = copiedStyle.tag;
                    var newEl = document.createElement(tag);
                    newEl.innerHTML = block.innerHTML;
                    if (block.id) newEl.id = block.id;
                    block.parentNode.replaceChild(newEl, block);
                    applyAllInlineStyle(newEl);
                    renumber();
                    generateTOC();
                    showToast('格式已应用 (→ ' + tag + ')', 'success', 1500);
                    return;
                }
            }

            // 普通块：尝试应用到最近的块级元素
            var block = el;
            while (block && block !== editor && !/^H[1-6]$|^P$|^DIV$|^LI$|^TD$/i.test(block.tagName)) block = block.parentNode;
            if (block && block !== editor) {
                applyAllInlineStyle(block);
                showToast('格式已应用', 'success', 1500);
            }
        }

        // ===== 将所有捕获的样式写为目标元素 inline style =====
        function applyAllInlineStyle(el) {
            var s = copiedStyle;
            if (!s) return;

            // 字体族
            if (s.fontFamily) el.style.fontFamily = s.fontFamily;
            // 字号
            if (s.fontSize) el.style.fontSize = s.fontSize;
            // 字重
            if (s.isBold !== undefined) el.style.fontWeight = s.isBold ? 'bold' : 'normal';
            // 斜体
            if (s.fontStyle) el.style.fontStyle = s.fontStyle;
            // 颜色
            if (s.color) el.style.color = s.color;
            // 背景色
            if (s.backgroundColor && s.backgroundColor !== 'rgba(0, 0, 0, 0)' && s.backgroundColor !== 'transparent') {
                el.style.backgroundColor = s.backgroundColor;
            }
            // 行高
            if (s.lineHeight && s.lineHeight !== 'normal') el.style.lineHeight = s.lineHeight;
            // 对齐
            if (s.textAlign) el.style.textAlign = s.textAlign;
            // 缩进
            if (s.textIndent && s.textIndent !== '0px') el.style.textIndent = s.textIndent;
            // 字间距
            if (s.letterSpacing && s.letterSpacing !== 'normal') el.style.letterSpacing = s.letterSpacing;
            // 词间距
            if (s.wordSpacing && s.wordSpacing !== 'normal' && s.wordSpacing !== '0px') el.style.wordSpacing = s.wordSpacing;
            // 垂直对齐
            if (s.verticalAlign && s.verticalAlign !== 'baseline') el.style.verticalAlign = s.verticalAlign;
            // 文本转换
            if (s.textTransform && s.textTransform !== 'none') el.style.textTransform = s.textTransform;

            // 下划线 / 删除线 / 上划线（合并为 textDecoration）
            var decoParts = [];
            if (s.isUnderline) decoParts.push('underline');
            if (s.isLineThrough) decoParts.push('line-through');
            if (s.isOverline) decoParts.push('overline');
            if (decoParts.length) {
                var decoVal = decoParts.join(' ');
                if (s.textDecorationStyle && s.textDecorationStyle !== 'solid') decoVal += ' ' + s.textDecorationStyle;
                if (s.textDecorationColor) decoVal += ' ' + s.textDecorationColor;
                el.style.textDecoration = decoVal;
            }

            // 边框（四个方向）
            ['Top','Bottom','Left','Right'].forEach(function(side) {
                var w = s['border' + side + 'Width'];
                var st = s['border' + side + 'Style'];
                var c = s['border' + side + 'Color'];
                if (st && st !== 'none' && st !== '') {
                    el.style['border' + side + 'Width'] = w || '1px';
                    el.style['border' + side + 'Style'] = st;
                    el.style['border' + side + 'Color'] = c || '#000';
                }
            });

            // padding
            ['Top','Bottom','Left','Right'].forEach(function(side) {
                if (s['padding' + side] && s['padding' + side] !== '0px') {
                    el.style['padding' + side] = s['padding' + side];
                }
            });

            // margin 有值才设
            if (s.marginTop && s.marginTop !== '0px') el.style.marginTop = s.marginTop;
            if (s.marginBottom && s.marginBottom !== '0px') el.style.marginBottom = s.marginBottom;
        }

        // ===== 编辑器点击 =====
        function painterClickHandler(e) {
            if (!painterActive) return;

            // 第一步：拾取样式
            if (picking && !copiedStyle) {
                // 点击的元素可以是行内元素，我们取最近的块包装
                var src = e.target;
                // 但如果选择的是文本选区，捕获选区内的文本样式
                var sel = window.getSelection();
                if (sel.rangeCount && !sel.isCollapsed && editor.contains(sel.getRangeAt(0).commonAncestorContainer)) {
                    // 用户有文本选择 — 取选区起始的父元素样式
                    var range = sel.getRangeAt(0);
                    src = range.startContainer;
                    while (src && src.nodeType === Node.TEXT_NODE) src = src.parentNode;
                }
                // 确保在编辑区内
                if (!src || !editor.contains(src)) return;
                e.preventDefault();

                copiedStyle = extractStyle(src);
                picking = false;
                applying = true;
                painterBtn.title = '🖌️ 点击目标应用格式，右键取消';
                setStatus('格式已复制 — 点击目标段落应用格式');
                showToast('已复制 ' + (copiedStyle.isHeading ? copiedStyle.tag + ' ' : '') + '格式', 'info', 2000);
                // 高亮来源元素
                var hl = src;
                while (hl && hl !== editor && !/^H[1-6]$|^P$|^DIV$|^LI$|^TD$/i.test(hl.tagName)) hl = hl.parentNode;
                if (hl && hl !== editor) hl.classList.add('format-painter-highlight');
                editor.style.cursor = 'copy';
                return;
            }

            // 第二步：应用样式
            if (applying && copiedStyle) {
                var target = e.target;
                while (target && target !== editor && !/^H[1-6]$|^P$|^DIV$|^LI$/i.test(target.tagName)) target = target.parentNode;
                if (!target || target === editor) return;
                e.preventDefault();

                applyStyle(target);
                document.querySelectorAll('.format-painter-highlight').forEach(function(el) { el.classList.remove('format-painter-highlight'); });
                resetPainter();
                painterBtn.classList.remove('active');
                painterActive = false;
            }
        }

        // ===== 格式刷按钮 =====
        painterBtn.addEventListener('click', function() {
            if (painterActive) {
                resetPainter();
                painterBtn.classList.remove('active');
                painterActive = false;
                showToast('格式刷已关闭', 'info', 1000);
                return;
            }
            painterActive = true;
            picking = true;
            applying = false;
            copiedStyle = null;
            painterBtn.classList.add('active');
            painterBtn.title = '🖌️ 点击源元素复制格式';
            setStatus('格式刷 — 请点击要复制格式的段落或标题');
            editor.style.cursor = 'copy';
            showToast('格式刷已开启，点击源元素复制全部格式', 'info', 2000);
        });

        editor.addEventListener('click', painterClickHandler);

        // ESC 取消
        document.addEventListener('keydown', function(e) {
            if (e.key === 'Escape' && painterActive) {
                resetPainter();
                painterBtn.classList.remove('active');
                painterActive = false;
                showToast('格式刷已取消', 'info', 1000);
            }
        });

        // 右键取消
        document.addEventListener('contextmenu', function(e) {
            if (painterActive) {
                resetPainter();
                painterBtn.classList.remove('active');
                painterActive = false;
            }
        });
    })();

    // ===== 🖼️ 图片编辑器 =====
    (function() {
        var overlay = $('imageEditorOverlay');
        var editorCanvas = $('imageEditorCanvas');
        var canvasWrap = $('imageEditorCanvasWrap');
        var editorTitle = $('imageEditorTitle');
        var editorStatus = $('imageEditorStatus');
        var zoomLabel = $('imageZoomLabel');
        var cropBar = $('imageCropBar');
        var editorHint = $('imageEditorHint');
        var colorPicker = $('imgEditorColor');
        var fillColorPicker = $('imgEditorFillColor');
        var lineWidthSelect = $('imgEditorLineWidth');
        var undoBtn = $('imgEditorUndo');
        var redoBtn = $('imgEditorRedo');

        if (!overlay || !editorCanvas) return;

        var ctx = editorCanvas.getContext('2d');

        // ===== 状态 =====
        var baseImage = null;           // 原始图片 Image 对象
        var annotations = [];           // 标注数组
        var undoStack = [];             // 撤销栈（annotations 快照）
        var redoStack = [];             // 重做栈
        var currentTool = 'select';     // 当前工具
        var editorMode = 'edit';        // 模式: 'view' | 'edit' | 'crop'
        var zoomLevel = 1;              // 缩放比例
        var targetImg = null;           // 正在编辑的 <img> 元素
        var originalSrc = '';           // 原始 src（用于放弃修改）
        var originalDataImgId = '';     // 原始 data-img-id
        var isDrawing = false;          // 是否正在绘制
        var drawStart = null;           // 绘制起始点
        var currentPath = null;         // 当前绘制的路径
        var tempAnnotation = null;      // 临时标注（绘制中）
        var isPanning = false;          // 是否正在平移
        var panStart = null;            // 平移起始点
        var panScrollStart = null;      // 平移前滚动位置
        var cropMode = false;           // 裁剪模式
        var cropRect = null;            // 裁剪矩形（canvas 坐标）
        var cropStart = null;           // 裁剪起始点
        var cropDragging = false;       // 是否正在拖拽裁剪区域
        var cropResizing = null;        // 正在调整大小的边角
        var textEditAnnotation = null;  // 正在编辑的文字标注
        var textEditDiv = null;         // 文字输入 div
        var imageWasModified = false;   // 图片是否被修改（裁剪或标注）

        // 标注类型工具映射
        var DRAW_TOOLS = { pen: true, eraser: true, text: true, rect: true, line: true, arrow: true };
        var SHAPE_TOOLS = { rect: true, line: true, arrow: true };

        // document 级事件绑定标记（确保鼠标在画布外松开也能完成操作）
        var docMoveBound = false;
        var docUpBound = false;

        function bindDocumentEvents() {
            if (!docMoveBound) {
                document.addEventListener('mousemove', onCanvasMouseMove);
                docMoveBound = true;
            }
            if (!docUpBound) {
                document.addEventListener('mouseup', onCanvasMouseUp);
                docUpBound = true;
            }
        }

        function unbindDocumentEvents() {
            if (docMoveBound) {
                document.removeEventListener('mousemove', onCanvasMouseMove);
                docMoveBound = false;
            }
            if (docUpBound) {
                document.removeEventListener('mouseup', onCanvasMouseUp);
                docUpBound = false;
            }
        }

        // ===== 初始化事件 =====
        $('imageEditorClose').addEventListener('click', closeImageEditor);
        $('imageEditorApply').addEventListener('click', applyChanges);
        $('imageEditorDiscard').addEventListener('click', discardChanges);
        $('imageCropConfirm').addEventListener('click', confirmCrop);
        $('imageCropCancel').addEventListener('click', cancelCrop);

        // 工具栏按钮
        overlay.querySelectorAll('.img-editor-tool-btn[data-tool]').forEach(function(btn) {
            btn.addEventListener('click', function() {
                var tool = this.dataset.tool;
                if (tool === 'undo') { imgEditorUndo(); return; }
                if (tool === 'redo') { imgEditorRedo(); return; }
                if (tool === 'zoomin') { imgEditorZoomIn(); return; }
                if (tool === 'zoomout') { imgEditorZoomOut(); return; }
                if (tool === 'fit') { imgEditorFitToScreen(); return; }
                setImageEditTool(tool);
            });
        });

        // 颜色/线宽变化
        colorPicker.addEventListener('input', function() {
            if (currentTool === 'pen' || SHAPE_TOOLS[currentTool]) {
                // 颜色变化时更新当前工具状态
            }
        });

        // 画布鼠标事件
        editorCanvas.addEventListener('mousedown', onCanvasMouseDown);
        editorCanvas.addEventListener('mousemove', onCanvasMouseMove);
        editorCanvas.addEventListener('mouseup', onCanvasMouseUp);
        editorCanvas.addEventListener('mouseleave', onCanvasMouseLeave);
        editorCanvas.addEventListener('wheel', onCanvasWheel, { passive: false });

        // 键盘快捷键
        document.addEventListener('keydown', onImageEditorKeydown);

        // ===== 打开/关闭 =====
        function openImageEditor(img, mode) {
            if (!img) return;
            targetImg = img;
            editorMode = mode || 'edit';
            originalSrc = img.getAttribute('src') || '';
            originalDataImgId = img.getAttribute('data-img-id') || '';

            // 重置状态
            annotations = [];
            undoStack = [];
            redoStack = [];
            zoomLevel = 1;
            currentTool = 'select';
            cropMode = false;
            cropRect = null;
            textEditAnnotation = null;
            imageWasModified = false;
            removeTextEditDiv();

            // 更新 UI
            overlay.classList.remove('hidden');
            cropBar.classList.add('hidden');

            if (editorMode === 'view') {
                editorTitle.textContent = '🖼️ 查看图片';
                setViewModeUI(true);
                $('imageEditorApply').style.display = 'none';
                $('imageEditorDiscard').textContent = '❌ 关闭';
            } else {
                editorTitle.textContent = editorMode === 'crop' ? '✂️ 裁剪图片' : '🖼️ 编辑图片';
                setViewModeUI(false);
                $('imageEditorApply').style.display = '';
                $('imageEditorDiscard').textContent = '❌ 放弃修改';
                if (editorMode === 'crop') {
                    setImageEditTool('crop');
                }
            }

            updateUndoRedoUI();
            setStatusText('正在加载图片...');

            // 加载图片
            loadBaseImage(img);
            overlay.focus();
        }

        function setViewModeUI(isView) {
            // 隐藏/显示绘制工具
            var drawToolBtns = overlay.querySelectorAll('.img-editor-tool-btn[data-tool="pen"],' +
                '.img-editor-tool-btn[data-tool="eraser"],.img-editor-tool-btn[data-tool="text"],' +
                '.img-editor-tool-btn[data-tool="rect"],.img-editor-tool-btn[data-tool="line"],' +
                '.img-editor-tool-btn[data-tool="arrow"],.img-editor-tool-btn[data-tool="crop"],' +
                '.img-editor-tool-btn[data-tool="undo"],.img-editor-tool-btn[data-tool="redo"]');
            drawToolBtns.forEach(function(btn) {
                btn.style.display = isView ? 'none' : '';
            });
            // 隐藏颜色和线宽
            var drawOpts = overlay.querySelectorAll('.img-editor-color,.img-editor-line-width,.img-editor-tool-divider');
            // 只隐藏绘制相关的 divider（第2、4、5、6个）
            var dividers = overlay.querySelectorAll('.img-editor-tool-divider');
            if (dividers.length >= 5) {
                dividers[1].style.display = isView ? 'none' : ''; // 视图后
                dividers[2].style.display = isView ? 'none' : ''; // 裁剪前
                dividers[3].style.display = isView ? 'none' : ''; // 颜色前
                dividers[4].style.display = isView ? 'none' : ''; // 撤销前
            }
            colorPicker.style.display = isView ? 'none' : '';
            fillColorPicker.style.display = isView ? 'none' : '';
            lineWidthSelect.style.display = isView ? 'none' : '';
        }

        function closeImageEditor() {
            removeTextEditDiv();
            overlay.classList.add('hidden');
            targetImg = null;
            baseImage = null;
            annotations = [];
            undoStack = [];
            redoStack = [];
            cropMode = false;
            cropRect = null;
            imageWasModified = false;
        }

        function loadBaseImage(img) {
            var image = new Image();
            image.onload = function() {
                baseImage = image;
                // 设置画布内部分辨率
                editorCanvas.width = image.naturalWidth;
                editorCanvas.height = image.naturalHeight;
                // 适应屏幕
                imgEditorFitToScreen();
                // 重绘
                redrawCanvas();
                updateZoomLabel();
                setStatusText('就绪 - ' + image.naturalWidth + '×' + image.naturalHeight + 'px');
                saveUndoState(); // 保存初始状态
            };
            image.onerror = function() {
                setStatusText('加载图片失败');
                showToast('加载图片失败', 'error');
            };
            // 优先使用 data-img-id 从 imageDataMap 加载
            var dataImgId = img.getAttribute('data-img-id');
            if (dataImgId && imageDataMap && imageDataMap.has(dataImgId)) {
                var imgData = imageDataMap.get(dataImgId);
                image.src = 'data:' + imgData.contentType + ';base64,' + imgData.base64;
            } else {
                image.src = img.getAttribute('src') || '';
            }
        }

        // ===== 画布渲染 =====
        function redrawCanvas() {
            if (!baseImage) return;
            ctx.clearRect(0, 0, editorCanvas.width, editorCanvas.height);
            // 绘制底图
            ctx.drawImage(baseImage, 0, 0, editorCanvas.width, editorCanvas.height);

            // 绘制标注
            for (var i = 0; i < annotations.length; i++) {
                drawAnnotation(ctx, annotations[i]);
            }

            // 绘制临时标注
            if (tempAnnotation) {
                drawAnnotation(ctx, tempAnnotation);
            }

            // 绘制裁剪遮罩
            if (cropMode && cropRect) {
                drawCropOverlay(ctx, cropRect);
            }
        }

        function drawAnnotation(context, ann) {
            context.save();
            switch (ann.type) {
                case 'path':
                    if (!ann.points || ann.points.length < 2) break;
                    context.strokeStyle = ann.color || '#ff0000';
                    context.lineWidth = ann.lineWidth || 2;
                    context.lineCap = 'round';
                    context.lineJoin = 'round';
                    context.globalCompositeOperation = ann.erase ? 'destination-out' : 'source-over';
                    context.beginPath();
                    context.moveTo(ann.points[0].x, ann.points[0].y);
                    for (var i = 1; i < ann.points.length; i++) {
                        context.lineTo(ann.points[i].x, ann.points[i].y);
                    }
                    context.stroke();
                    break;

                case 'text':
                    context.fillStyle = ann.color || '#ff0000';
                    context.font = (ann.fontSize || 18) + 'px "Microsoft YaHei", sans-serif';
                    context.textBaseline = 'top';
                    context.fillText(ann.text || '', ann.x, ann.y);
                    break;

                case 'rect':
                    context.strokeStyle = ann.color || '#ff0000';
                    context.lineWidth = ann.lineWidth || 2;
                    if (ann.fillColor && ann.fillColor !== 'transparent') {
                        context.fillStyle = ann.fillColor;
                        context.fillRect(ann.x, ann.y, ann.w, ann.h);
                    }
                    context.strokeRect(ann.x, ann.y, ann.w, ann.h);
                    break;

                case 'line':
                    context.strokeStyle = ann.color || '#ff0000';
                    context.lineWidth = ann.lineWidth || 2;
                    context.lineCap = 'round';
                    context.beginPath();
                    context.moveTo(ann.x1, ann.y1);
                    context.lineTo(ann.x2, ann.y2);
                    context.stroke();
                    break;

                case 'arrow':
                    context.strokeStyle = ann.color || '#ff0000';
                    context.fillStyle = ann.color || '#ff0000';
                    context.lineWidth = ann.lineWidth || 2;
                    context.lineCap = 'round';
                    context.beginPath();
                    context.moveTo(ann.x1, ann.y1);
                    context.lineTo(ann.x2, ann.y2);
                    context.stroke();
                    // 画箭头
                    var angle = Math.atan2(ann.y2 - ann.y1, ann.x2 - ann.x1);
                    var arrowLen = 12 + (ann.lineWidth || 2) * 2;
                    context.beginPath();
                    context.moveTo(ann.x2, ann.y2);
                    context.lineTo(
                        ann.x2 - arrowLen * Math.cos(angle - Math.PI / 6),
                        ann.y2 - arrowLen * Math.sin(angle - Math.PI / 6)
                    );
                    context.lineTo(
                        ann.x2 - arrowLen * Math.cos(angle + Math.PI / 6),
                        ann.y2 - arrowLen * Math.sin(angle + Math.PI / 6)
                    );
                    context.closePath();
                    context.fill();
                    break;
            }
            context.restore();
        }

        function drawCropOverlay(context, rect) {
            // 半透明遮罩
            context.fillStyle = 'rgba(0,0,0,0.5)';
            context.fillRect(0, 0, editorCanvas.width, rect.y);
            context.fillRect(0, rect.y, rect.x, rect.h);
            context.fillRect(rect.x + rect.w, rect.y, editorCanvas.width - rect.x - rect.w, rect.h);
            context.fillRect(0, rect.y + rect.h, editorCanvas.width, editorCanvas.height - rect.y - rect.h);
            // 裁剪框
            context.strokeStyle = '#fff';
            context.lineWidth = 2;
            context.setLineDash([6, 3]);
            context.strokeRect(rect.x, rect.y, rect.w, rect.h);
            context.setLineDash([]);
            // 九宫格线
            context.strokeStyle = 'rgba(255,255,255,0.4)';
            context.lineWidth = 1;
            context.setLineDash([4, 4]);
            context.beginPath();
            context.moveTo(rect.x + rect.w / 3, rect.y);
            context.lineTo(rect.x + rect.w / 3, rect.y + rect.h);
            context.moveTo(rect.x + rect.w * 2 / 3, rect.y);
            context.lineTo(rect.x + rect.w * 2 / 3, rect.y + rect.h);
            context.moveTo(rect.x, rect.y + rect.h / 3);
            context.lineTo(rect.x + rect.w, rect.y + rect.h / 3);
            context.moveTo(rect.x, rect.y + rect.h * 2 / 3);
            context.lineTo(rect.x + rect.w, rect.y + rect.h * 2 / 3);
            context.stroke();
            context.setLineDash([]);
        }

        // ===== 坐标转换 =====
        function getCanvasCoords(e) {
            var rect = editorCanvas.getBoundingClientRect();
            var scaleX = editorCanvas.width / rect.width;
            var scaleY = editorCanvas.height / rect.height;
            return {
                x: (e.clientX - rect.left) * scaleX,
                y: (e.clientY - rect.top) * scaleY
            };
        }

        // ===== 工具切换 =====
        function setImageEditTool(tool) {
            if (editorMode === 'view' && tool !== 'zoomin' && tool !== 'zoomout' && tool !== 'fit') return;
            currentTool = tool;

            // 更新按钮状态
            overlay.querySelectorAll('.img-editor-tool-btn.active').forEach(function(b) {
                b.classList.remove('active');
            });
            var activeBtn = overlay.querySelector('.img-editor-tool-btn[data-tool="' + tool + '"]');
            if (activeBtn) activeBtn.classList.add('active');

            // 退出裁剪模式
            if (cropMode && tool !== 'crop') {
                cropMode = false;
                cropRect = null;
                cropBar.classList.add('hidden');
                redrawCanvas();
            }

            // 进入裁剪模式
            if (tool === 'crop') {
                cropMode = true;
                cropRect = null;
                cropBar.classList.remove('hidden');
                setStatusText('✂️ 在图片上拖拽选择裁剪区域');
                redrawCanvas();
            }

            // 更新光标
            if (tool === 'select') {
                editorCanvas.style.cursor = 'default';
            } else if (tool === 'text') {
                editorCanvas.style.cursor = 'text';
            } else if (tool === 'crop') {
                editorCanvas.style.cursor = 'crosshair';
            } else if (tool === 'eraser') {
                editorCanvas.style.cursor = 'cell';
            } else if (DRAW_TOOLS[tool]) {
                editorCanvas.style.cursor = 'crosshair';
            }

            if (editorHint) {
                var hints = {
                    pen: '✏️ 画笔：按住拖拽自由绘制',
                    eraser: '🧹 橡皮：擦除标注（涂抹路径上的标注将被删除）',
                    text: '📝 文本框：点击图片放置文字',
                    rect: '⬜ 矩形：按住拖拽绘制矩形',
                    line: '📏 直线：按住拖拽绘制直线',
                    arrow: '➡️ 箭头：按住拖拽绘制箭头',
                    crop: '✂️ 裁剪：拖拽选择裁剪区域，确认后裁剪',
                    select: '🖱️ 选择：拖拽平移画布，滚轮缩放'
                };
                editorHint.textContent = hints[tool] || '';
            }

            setStatusText('工具：' + (activeBtn ? activeBtn.title : tool));
        }

        // ===== 鼠标事件 =====
        function onCanvasMouseDown(e) {
            if (!baseImage) return;
            if (e.button !== 0) return; // 只处理左键

            var coords = getCanvasCoords(e);

            // 裁剪模式
            if (cropMode) {
                // 检查是否点击了裁剪框的边角
                if (cropRect) {
                    var handle = getCropHandle(coords, cropRect);
                    if (handle) {
                        cropResizing = handle;
                        cropStart = coords;
                        bindDocumentEvents();
                        e.preventDefault();
                        return;
                    }
                    // 检查是否在裁剪框内
                    if (coords.x >= cropRect.x && coords.x <= cropRect.x + cropRect.w &&
                        coords.y >= cropRect.y && coords.y <= cropRect.y + cropRect.h) {
                        cropDragging = true;
                        cropStart = coords;
                        bindDocumentEvents();
                        e.preventDefault();
                        return;
                    }
                }
                // 开始新的裁剪选区
                cropDragging = false;
                cropResizing = null;
                cropStart = coords;
                cropRect = null;
                bindDocumentEvents();
                e.preventDefault();
                return;
            }

            // 文字工具
            if (currentTool === 'text') {
                finishTextEdit();
                placeTextAnnotation(coords);
                e.preventDefault();
                return;
            }

            // 选择工具 - 平移
            if (currentTool === 'select') {
                isPanning = true;
                panStart = { x: e.clientX, y: e.clientY };
                panScrollStart = { left: canvasWrap.scrollLeft, top: canvasWrap.scrollTop };
                editorCanvas.style.cursor = 'grabbing';
                bindDocumentEvents();
                e.preventDefault();
                return;
            }

            // 绘制工具
            if (DRAW_TOOLS[currentTool]) {
                isDrawing = true;
                drawStart = coords;
                currentPath = [];
                tempAnnotation = null;
                bindDocumentEvents();

                if (currentTool === 'pen') {
                    currentPath = [coords];
                    tempAnnotation = { type: 'path', points: currentPath, color: colorPicker.value, lineWidth: parseInt(lineWidthSelect.value) };
                } else if (currentTool === 'eraser') {
                    // 橡皮不创建临时标注，在 mouseup 时处理
                    currentPath = [coords];
                } else if (currentTool === 'rect') {
                    tempAnnotation = { type: 'rect', x: coords.x, y: coords.y, w: 0, h: 0, color: colorPicker.value, lineWidth: parseInt(lineWidthSelect.value), fillColor: fillColorPicker.value };
                } else if (currentTool === 'line') {
                    tempAnnotation = { type: 'line', x1: coords.x, y1: coords.y, x2: coords.x, y2: coords.y, color: colorPicker.value, lineWidth: parseInt(lineWidthSelect.value) };
                } else if (currentTool === 'arrow') {
                    tempAnnotation = { type: 'arrow', x1: coords.x, y1: coords.y, x2: coords.x, y2: coords.y, color: colorPicker.value, lineWidth: parseInt(lineWidthSelect.value) };
                }
                e.preventDefault();
            }
        }

        function onCanvasMouseMove(e) {
            if (!baseImage) return;

            var coords = getCanvasCoords(e);

            // 平移
            if (isPanning) {
                var dx = e.clientX - panStart.x;
                var dy = e.clientY - panStart.y;
                canvasWrap.scrollLeft = panScrollStart.left - dx;
                canvasWrap.scrollTop = panScrollStart.top - dy;
                return;
            }

            // 裁剪调整大小
            if (cropMode && cropResizing && cropRect) {
                resizeCropRect(coords);
                redrawCanvas();
                return;
            }

            // 裁剪拖拽
            if (cropMode && cropDragging && cropRect) {
                var moveDx = coords.x - cropStart.x;
                var moveDy = coords.y - cropStart.y;
                cropRect.x += moveDx;
                cropRect.y += moveDy;
                // 限制在画布内
                cropRect.x = Math.max(0, Math.min(cropRect.x, editorCanvas.width - cropRect.w));
                cropRect.y = Math.max(0, Math.min(cropRect.y, editorCanvas.height - cropRect.h));
                cropStart = coords;
                redrawCanvas();
                return;
            }

            // 裁剪初始拖拽（持续更新选区大小）
            if (cropMode && cropStart && !cropResizing && !cropDragging) {
                var rx = Math.min(cropStart.x, coords.x);
                var ry = Math.min(cropStart.y, coords.y);
                var rw = Math.abs(coords.x - cropStart.x);
                var rh = Math.abs(coords.y - cropStart.y);
                cropRect = { x: rx, y: ry, w: rw, h: rh };
                redrawCanvas();
                return;
            }

            // 绘制
            if (isDrawing && tempAnnotation) {
                if (currentTool === 'pen') {
                    currentPath.push(coords);
                    // 只重绘最新线段以提高性能
                    redrawCanvas();
                } else if (currentTool === 'rect') {
                    tempAnnotation.w = coords.x - drawStart.x;
                    tempAnnotation.h = coords.y - drawStart.y;
                    redrawCanvas();
                } else if (currentTool === 'line' || currentTool === 'arrow') {
                    tempAnnotation.x2 = coords.x;
                    tempAnnotation.y2 = coords.y;
                    redrawCanvas();
                }
            } else if (isDrawing && currentTool === 'eraser') {
                // 橡皮实时擦除
                currentPath.push(coords);
                eraseAnnotationsAlongPath(currentPath, parseInt(lineWidthSelect.value) + 4);
                redrawCanvas();
            }
        }

        function onCanvasMouseUp(e) {
            // 平移结束
            if (isPanning) {
                isPanning = false;
                panStart = null;
                editorCanvas.style.cursor = 'default';
                unbindDocumentEvents();
                return;
            }

            // 裁剪结束
            if (cropMode && cropStart && cropRect) {
                if (cropResizing) {
                    cropResizing = null;
                    cropStart = null;
                    unbindDocumentEvents();
                    return;
                }
                if (cropDragging) {
                    cropDragging = false;
                    cropStart = null;
                    unbindDocumentEvents();
                    return;
                }
                // 最小裁剪区域检查
                if (cropRect.w < 10 || cropRect.h < 10) {
                    cropRect = null;
                    redrawCanvas();
                }
                cropStart = null;
                unbindDocumentEvents();
                return;
            }

            // 绘制结束
            if (isDrawing) {
                isDrawing = false;
                unbindDocumentEvents();

                if (currentTool === 'eraser') {
                    // 橡皮：删除路径上的标注
                    eraseAnnotationsAlongPath(currentPath, parseInt(lineWidthSelect.value) + 8);
                    currentPath = null;
                    saveUndoState();
                    redrawCanvas();
                    drawStart = null;
                    return;
                }

                if (tempAnnotation) {
                    // 规范化矩形（处理负宽高）
                    if (tempAnnotation.type === 'rect') {
                        if (tempAnnotation.w < 0) {
                            tempAnnotation.x += tempAnnotation.w;
                            tempAnnotation.w = -tempAnnotation.w;
                        }
                        if (tempAnnotation.h < 0) {
                            tempAnnotation.y += tempAnnotation.h;
                            tempAnnotation.h = -tempAnnotation.h;
                        }
                        // 过滤太小的矩形
                        if (Math.abs(tempAnnotation.w) < 3 && Math.abs(tempAnnotation.h) < 3) {
                            tempAnnotation = null;
                            redrawCanvas();
                            drawStart = null;
                            return;
                        }
                    }

                    // 过滤太短的线
                    if ((tempAnnotation.type === 'line' || tempAnnotation.type === 'arrow')) {
                        var lineLen = Math.sqrt(
                            Math.pow(tempAnnotation.x2 - tempAnnotation.x1, 2) +
                            Math.pow(tempAnnotation.y2 - tempAnnotation.y1, 2)
                        );
                        if (lineLen < 3) {
                            tempAnnotation = null;
                            redrawCanvas();
                            drawStart = null;
                            return;
                        }
                    }

                    saveUndoState();
                    annotations.push(tempAnnotation);
                    imageWasModified = true;
                    tempAnnotation = null;
                    redrawCanvas();
                    updateUndoRedoUI();
                    setStatusText('标注数：' + annotations.length);
                }

                drawStart = null;
                currentPath = null;
            }
        }

        // mouseleave 仅结束绘制操作，不终止裁剪（裁剪需要精确 mouseup）
        function onCanvasMouseLeave(e) {
            if (!baseImage) return;

            // 平移结束
            if (isPanning) {
                isPanning = false;
                panStart = null;
                editorCanvas.style.cursor = 'default';
                unbindDocumentEvents();
                return;
            }

            // 绘制工具结束时保留标注
            if (isDrawing) {
                onCanvasMouseUp(e);
                return;
            }

            // 裁剪操作不在此终止，用户可以回到画布继续调整
        }

        function onCanvasWheel(e) {
            if (!baseImage) return;
            e.preventDefault();
            var delta = e.deltaY > 0 ? -0.1 : 0.1;
            var newZoom = zoomLevel + delta;
            newZoom = Math.max(0.1, Math.min(5, newZoom));
            setZoom(newZoom);
        }

        function getCropHandle(coords, rect) {
            var handleSize = 10;
            var handles = {
                nw: { x: rect.x, y: rect.y },
                n: { x: rect.x + rect.w / 2, y: rect.y },
                ne: { x: rect.x + rect.w, y: rect.y },
                e: { x: rect.x + rect.w, y: rect.y + rect.h / 2 },
                se: { x: rect.x + rect.w, y: rect.y + rect.h },
                s: { x: rect.x + rect.w / 2, y: rect.y + rect.h },
                sw: { x: rect.x, y: rect.y + rect.h },
                w: { x: rect.x, y: rect.y + rect.h / 2 }
            };
            for (var key in handles) {
                if (Math.abs(coords.x - handles[key].x) < handleSize &&
                    Math.abs(coords.y - handles[key].y) < handleSize) {
                    return key;
                }
            }
            return null;
        }

        function resizeCropRect(coords) {
            if (!cropRect || !cropResizing) return;
            var r = cropRect;
            switch (cropResizing) {
                case 'nw': r.w += r.x - coords.x; r.h += r.y - coords.y; r.x = coords.x; r.y = coords.y; break;
                case 'n': r.h += r.y - coords.y; r.y = coords.y; break;
                case 'ne': r.w = coords.x - r.x; r.h += r.y - coords.y; r.y = coords.y; break;
                case 'e': r.w = coords.x - r.x; break;
                case 'se': r.w = coords.x - r.x; r.h = coords.y - r.y; break;
                case 's': r.h = coords.y - r.y; break;
                case 'sw': r.w += r.x - coords.x; r.h = coords.y - r.y; r.x = coords.x; break;
                case 'w': r.w += r.x - coords.x; r.x = coords.x; break;
            }
            if (r.w < 20) r.w = 20;
            if (r.h < 20) r.h = 20;
            r.x = Math.max(0, Math.min(r.x, editorCanvas.width - 1));
            r.y = Math.max(0, Math.min(r.y, editorCanvas.height - 1));
            if (r.x + r.w > editorCanvas.width) r.w = editorCanvas.width - r.x;
            if (r.y + r.h > editorCanvas.height) r.h = editorCanvas.height - r.y;
        }

        // ===== 橡皮擦实现 =====
        function eraseAnnotationsAlongPath(path, radius) {
            if (!path || path.length < 2) return;
            var toRemove = [];
            for (var i = 0; i < annotations.length; i++) {
                var ann = annotations[i];
                if (annotationIntersectsPath(ann, path, radius)) {
                    toRemove.push(i);
                }
            }
            // 从后往前删除
            for (var j = toRemove.length - 1; j >= 0; j--) {
                annotations.splice(toRemove[j], 1);
            }
            if (toRemove.length > 0) {
                imageWasModified = true;
                setStatusText('擦除了 ' + toRemove.length + ' 个标注');
            }
        }

        function annotationIntersectsPath(ann, path, radius) {
            for (var i = 0; i < path.length; i++) {
                var px = path[i].x;
                var py = path[i].y;
                switch (ann.type) {
                    case 'path':
                        for (var j = 0; j < ann.points.length; j++) {
                            if (pointDistance(px, py, ann.points[j].x, ann.points[j].y) < radius) return true;
                        }
                        break;
                    case 'text':
                        if (px >= ann.x - radius && px <= ann.x + 100 + radius &&
                            py >= ann.y - radius && py <= ann.y + (ann.fontSize || 18) + radius) return true;
                        break;
                    case 'rect':
                        if (px >= ann.x - radius && px <= ann.x + ann.w + radius &&
                            py >= ann.y - radius && py <= ann.y + ann.h + radius) return true;
                        break;
                    case 'line':
                    case 'arrow':
                        if (pointToSegmentDistance(px, py, ann.x1, ann.y1, ann.x2, ann.y2) < radius) return true;
                        break;
                }
            }
            return false;
        }

        function pointDistance(x1, y1, x2, y2) {
            return Math.sqrt(Math.pow(x2 - x1, 2) + Math.pow(y2 - y1, 2));
        }

        function pointToSegmentDistance(px, py, x1, y1, x2, y2) {
            var dx = x2 - x1;
            var dy = y2 - y1;
            var lenSq = dx * dx + dy * dy;
            if (lenSq === 0) return pointDistance(px, py, x1, y1);
            var t = Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / lenSq));
            return pointDistance(px, py, x1 + t * dx, y1 + t * dy);
        }

        // ===== 文字标注 =====
        function placeTextAnnotation(coords) {
            removeTextEditDiv();
            textEditAnnotation = { type: 'text', x: coords.x, y: coords.y, text: '', color: colorPicker.value, fontSize: 18 };

            // 创建输入框
            var div = document.createElement('div');
            div.className = 'image-text-input';
            div.contentEditable = 'true';
            div.spellcheck = false;
            div.textContent = '';

            // 计算放置位置（相对于 canvasWrap，考虑滚动偏移）
            var canvasRect = editorCanvas.getBoundingClientRect();
            var wrapRect = canvasWrap.getBoundingClientRect();
            var scaleX = canvasRect.width / editorCanvas.width;
            var scaleY = canvasRect.height / editorCanvas.height;
            var left = canvasRect.left - wrapRect.left + canvasWrap.scrollLeft + coords.x * scaleX;
            var top = canvasRect.top - wrapRect.top + canvasWrap.scrollTop + coords.y * scaleY;

            div.style.position = 'absolute';
            div.style.left = left + 'px';
            div.style.top = top + 'px';
            div.style.color = colorPicker.value;
            div.style.fontSize = '18px';
            div.style.zIndex = '10';

            canvasWrap.appendChild(div);
            textEditDiv = div;

            div.focus();

            // 失焦时完成编辑
            div.addEventListener('blur', function() {
                finishTextEdit();
            });

            // Enter 完成编辑
            div.addEventListener('keydown', function(ev) {
                if (ev.key === 'Enter' && !ev.shiftKey) {
                    ev.preventDefault();
                    finishTextEdit();
                }
                if (ev.key === 'Escape') {
                    ev.preventDefault();
                    cancelTextEdit();
                }
            });
        }

        function finishTextEdit() {
            if (!textEditDiv || !textEditAnnotation) return;
            var text = textEditDiv.textContent.trim();
            if (text) {
                textEditAnnotation.text = text;
                // 计算实际位置（div 在 canvasWrap 中的像素 → canvas 内部坐标）
                var canvasRect = editorCanvas.getBoundingClientRect();
                var wrapRect = canvasWrap.getBoundingClientRect();
                var scaleX = editorCanvas.width / canvasRect.width;
                var scaleY = editorCanvas.height / canvasRect.height;
                var divLeft = parseFloat(textEditDiv.style.left) || 0;
                var divTop = parseFloat(textEditDiv.style.top) || 0;
                // div 在 wrap 内坐标 → 画布 CSS 显示坐标 → 画布内部坐标
                textEditAnnotation.x = (divLeft - canvasWrap.scrollLeft - (canvasRect.left - wrapRect.left)) * scaleX;
                textEditAnnotation.y = (divTop - canvasWrap.scrollTop - (canvasRect.top - wrapRect.top)) * scaleY;
                textEditAnnotation.color = colorPicker.value;
                textEditAnnotation.fontSize = parseInt(textEditDiv.style.fontSize) || 18;

                saveUndoState();
                annotations.push(textEditAnnotation);
                imageWasModified = true;
                redrawCanvas();
                updateUndoRedoUI();
                setStatusText('已添加文字标注');
            }
            removeTextEditDiv();
            textEditAnnotation = null;
        }

        function cancelTextEdit() {
            removeTextEditDiv();
            textEditAnnotation = null;
        }

        function removeTextEditDiv() {
            if (textEditDiv) {
                textEditDiv.remove();
                textEditDiv = null;
            }
        }

        // ===== 撤销/重做 =====
        function saveUndoState() {
            undoStack.push(JSON.parse(JSON.stringify(annotations)));
            if (undoStack.length > 50) undoStack.shift(); // 限制 50 步
            redoStack = [];
            updateUndoRedoUI();
        }

        function imgEditorUndo() {
            if (undoStack.length <= 1) return;
            // 当前状态移到 redo
            redoStack.push(JSON.parse(JSON.stringify(annotations)));
            // 恢复上一个状态
            undoStack.pop();
            annotations = JSON.parse(JSON.stringify(undoStack[undoStack.length - 1]));
            redrawCanvas();
            updateUndoRedoUI();
            setStatusText('撤销 - 标注数：' + annotations.length);
        }

        function imgEditorRedo() {
            if (redoStack.length === 0) return;
            undoStack.push(JSON.parse(JSON.stringify(redoStack[redoStack.length - 1])));
            annotations = JSON.parse(JSON.stringify(redoStack.pop()));
            redrawCanvas();
            updateUndoRedoUI();
            setStatusText('重做 - 标注数：' + annotations.length);
        }

        function updateUndoRedoUI() {
            if (undoBtn) undoBtn.disabled = (undoStack.length <= 1);
            if (redoBtn) redoBtn.disabled = (redoStack.length === 0);
        }

        // ===== 缩放 =====
        function imgEditorZoomIn() {
            setZoom(Math.min(5, zoomLevel + 0.25));
        }

        function imgEditorZoomOut() {
            setZoom(Math.max(0.1, zoomLevel - 0.25));
        }

        function imgEditorFitToScreen() {
            if (!baseImage) return;
            var wrapW = canvasWrap.clientWidth - 20;
            var wrapH = canvasWrap.clientHeight - 20;
            var scaleW = wrapW / baseImage.naturalWidth;
            var scaleH = wrapH / baseImage.naturalHeight;
            zoomLevel = Math.min(scaleW, scaleH, 1); // 最大 100%
            setZoom(zoomLevel);
        }

        function setZoom(level) {
            zoomLevel = Math.max(0.1, Math.min(5, level));
            if (baseImage) {
                editorCanvas.style.width = Math.round(baseImage.naturalWidth * zoomLevel) + 'px';
                editorCanvas.style.height = Math.round(baseImage.naturalHeight * zoomLevel) + 'px';
            }
            updateZoomLabel();
        }

        function updateZoomLabel() {
            if (zoomLabel) zoomLabel.textContent = Math.round(zoomLevel * 100) + '%';
        }

        function setStatusText(msg) {
            if (editorStatus) editorStatus.textContent = msg;
        }

        // ===== 裁剪确认/取消 =====
        function confirmCrop() {
            if (!baseImage || !cropRect || cropRect.w < 10 || cropRect.h < 10) {
                showToast('请先拖拽选择裁剪区域', 'warning');
                return;
            }

            // 创建裁剪后的图片
            var tempCanvas = document.createElement('canvas');
            tempCanvas.width = Math.round(cropRect.w);
            tempCanvas.height = Math.round(cropRect.h);
            var tempCtx = tempCanvas.getContext('2d');

            // 先绘制底图
            tempCtx.drawImage(baseImage,
                cropRect.x, cropRect.y, cropRect.w, cropRect.h,
                0, 0, cropRect.w, cropRect.h
            );

            // 调整标注坐标
            var adjustedAnnotations = [];
            for (var i = 0; i < annotations.length; i++) {
                var ann = JSON.parse(JSON.stringify(annotations[i]));
                if (ann.type === 'path' && ann.points) {
                    ann.points = ann.points.map(function(p) {
                        return { x: p.x - cropRect.x, y: p.y - cropRect.y };
                    }).filter(function(p) { return p.x >= 0 && p.y >= 0 && p.x <= cropRect.w && p.y <= cropRect.h; });
                    if (ann.points.length < 2) continue;
                } else if (ann.type === 'text') {
                    ann.x -= cropRect.x;
                    ann.y -= cropRect.y;
                    if (ann.x < -50 || ann.y < -20 || ann.x > cropRect.w + 50 || ann.y > cropRect.h + 20) continue;
                } else if (ann.type === 'rect') {
                    ann.x -= cropRect.x;
                    ann.y -= cropRect.y;
                    if (ann.x + ann.w < 0 || ann.y + ann.h < 0 || ann.x > cropRect.w || ann.y > cropRect.h) continue;
                } else if (ann.type === 'line' || ann.type === 'arrow') {
                    ann.x1 -= cropRect.x;
                    ann.y1 -= cropRect.y;
                    ann.x2 -= cropRect.x;
                    ann.y2 -= cropRect.y;
                }
                adjustedAnnotations.push(ann);
            }

            // 绘制标注到裁剪画布
            for (var j = 0; j < adjustedAnnotations.length; j++) {
                drawAnnotation(tempCtx, adjustedAnnotations[j]);
            }

            // 保存裁剪尺寸（onload 回调中 cropRect 会被清空）
            var croppedW = Math.round(cropRect.w);
            var croppedH = Math.round(cropRect.h);

            // 更新底图
            var newImg = new Image();
            newImg.onload = function() {
                baseImage = newImg;
                editorCanvas.width = baseImage.naturalWidth;
                editorCanvas.height = baseImage.naturalHeight;
                annotations = adjustedAnnotations;
                imageWasModified = true;
                saveUndoState();
                undoStack = [JSON.parse(JSON.stringify(annotations))];
                redoStack = [];
                cropMode = false;
                cropRect = null;
                cropBar.classList.add('hidden');
                setImageEditTool('select');
                imgEditorFitToScreen();
                redrawCanvas();
                updateUndoRedoUI();
                setStatusText('裁剪完成 - ' + croppedW + '×' + croppedH + 'px');
                showToast('裁剪完成', 'success');
            };
            newImg.src = tempCanvas.toDataURL('image/png');
        }

        function cancelCrop() {
            cropMode = false;
            cropRect = null;
            cropBar.classList.add('hidden');
            setImageEditTool('select');
            redrawCanvas();
            setStatusText('已取消裁剪');
        }

        // ===== 应用/放弃修改 =====
        function applyChanges() {
            if (!targetImg) return;
            if (editorMode === 'view') { closeImageEditor(); return; }

            // 检查是否有实质性修改（标注或裁剪）
            if (!imageWasModified) {
                closeImageEditor();
                return;
            }

            // 将画布导出为 data URL
            var exportCanvas = document.createElement('canvas');
            exportCanvas.width = editorCanvas.width;
            exportCanvas.height = editorCanvas.height;
            var exportCtx = exportCanvas.getContext('2d');

            // 绘制底图
            exportCtx.drawImage(baseImage, 0, 0);

            // 绘制所有标注
            for (var i = 0; i < annotations.length; i++) {
                drawAnnotation(exportCtx, annotations[i]);
            }

            var dataURL = exportCanvas.toDataURL('image/png');
            var m = dataURL.match(/^data:([^;]+);base64,(.+)$/);

            // 更新编辑器中的图片
            targetImg.setAttribute('src', dataURL);

            // 更新 imageDataMap
            if (m) {
                var newId = 'img-' + Date.now();
                targetImg.setAttribute('data-img-id', newId);
                if (imageDataMap) {
                    // 移除旧条目
                    if (originalDataImgId) imageDataMap.delete(originalDataImgId);
                    imageDataMap.set(newId, {
                        contentType: m[1],
                        base64: m[2],
                        altText: targetImg.getAttribute('alt') || ''
                    });
                }
            }

            // 标记图片已变更
            var session = tabManager ? tabManager.getActive() : null;
            if (session) {
                session._imagesChanged = true;
                session._lastSavedHtml = ''; // 强制下次保存
            }

            // 保存撤销状态
            saveUndoState('图片编辑');
            triggerAutoSave();

            setStatusText('修改已应用');
            showToast('图片修改已应用', 'success', 1500);
            closeImageEditor();
        }

        function discardChanges() {
            if (editorMode === 'view') { closeImageEditor(); return; }

            // 还原原始 src 和 data-img-id
            if (targetImg && originalSrc) {
                targetImg.setAttribute('src', originalSrc);
                if (originalDataImgId) {
                    targetImg.setAttribute('data-img-id', originalDataImgId);
                }
            }
            closeImageEditor();
            setStatusText('图片修改已放弃');
            showToast('图片修改已放弃', 'info', 1500);
        }

        // ===== 键盘快捷键 =====
        function onImageEditorKeydown(e) {
            if (overlay.classList.contains('hidden')) return;

            // Ctrl+Z 撤销
            if (e.ctrlKey && e.key === 'z' && !e.shiftKey) {
                e.preventDefault();
                e.stopPropagation();
                if (textEditDiv) return; // 文字编辑中不处理
                imgEditorUndo();
                return;
            }

            // Ctrl+Y 或 Ctrl+Shift+Z 重做
            if ((e.ctrlKey && e.key === 'y') || (e.ctrlKey && e.shiftKey && e.key === 'z')) {
                e.preventDefault();
                e.stopPropagation();
                if (textEditDiv) return;
                imgEditorRedo();
                return;
            }

            // Ctrl+S 保存（在图片编辑器中拦截，避免触发浏览器保存）
            if (e.ctrlKey && e.key === 's') {
                e.preventDefault();
                e.stopPropagation();
                return;
            }

            // 工具快捷键
            if (e.ctrlKey || e.metaKey) return;
            if (textEditDiv) return;

            var toolKeys = {
                's': 'select', 'p': 'pen', 'e': 'eraser', 't': 'text',
                'r': 'rect', 'l': 'line', 'a': 'arrow', 'c': 'crop'
            };
            var tool = toolKeys[e.key.toLowerCase()];
            if (tool && editorMode !== 'view') {
                e.preventDefault();
                e.stopPropagation();
                setImageEditTool(tool);
            }

            // +/- 缩放
            if (e.key === '+' || e.key === '=') { e.preventDefault(); e.stopPropagation(); imgEditorZoomIn(); }
            if (e.key === '-') { e.preventDefault(); e.stopPropagation(); imgEditorZoomOut(); }
            if (e.key === '0') { e.preventDefault(); e.stopPropagation(); imgEditorFitToScreen(); }

            // Escape 关闭
            if (e.key === 'Escape') {
                e.preventDefault();
                e.stopPropagation();
                if (cropMode) {
                    cancelCrop();
                } else if (textEditDiv) {
                    cancelTextEdit();
                } else {
                    closeImageEditor();
                }
            }
        }

        // ===== 暴露全局接口 =====
        window.openImageEditor = openImageEditor;
        window.closeImageEditor = closeImageEditor;
    })();

    console.log('DOCX Editor initialized');

})();
