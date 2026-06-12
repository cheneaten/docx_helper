/* ============================================================
   DOCX 编辑器 - DOCX 导出模块 (完整重写)
   使用 JSZip 构建 OOXML 格式，完整支持图片/表格/格式
   ============================================================ */

(function() {
    'use strict';

    // ===== 命名空间 =====
    const NS = {
        w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
        pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
        rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
        mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
        v: 'urn:schemas-microsoft-com:vml',
        wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        wpc: 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
    };

    // ===== 工具函数 =====
    function escXml(str) {
        if (typeof str !== 'string') return '';
        return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
                  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
    }

    function ptToHalfPoint(ptStr) {
        const num = parseFloat(ptStr);
        return isNaN(num) ? 24 : Math.round(num * 2);
    }

    // CSS 颜色值转 OOXML 十六进制（不含 # 前缀）
    function cssColorToHex(color) {
        if (!color || typeof color !== 'string') return '000000';
        color = color.trim().toLowerCase();
        if (color === 'transparent' || color === 'rgba(0, 0, 0, 0)' || color === 'auto') return 'auto';
        // 已是标准 hex
        if (/^#[0-9a-f]{6}$/.test(color)) return color.replace('#', '').toUpperCase();
        if (/^#[0-9a-f]{3}$/.test(color)) {
            const h = color.replace('#', '');
            return (h[0]+h[0]+h[1]+h[1]+h[2]+h[2]).toUpperCase();
        }
        // rgb(r, g, b)
        let m = color.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
        if (m) {
            return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3])]
                .map(v => { const h = v.toString(16); return h.length === 1 ? '0' + h : h; })
                .join('').toUpperCase();
        }
        // rgba(r, g, b, ...)
        m = color.match(/rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
        if (m) {
            return [parseInt(m[1]), parseInt(m[2]), parseInt(m[3])]
                .map(v => { const h = v.toString(16); return h.length === 1 ? '0' + h : h; })
                .join('').toUpperCase();
        }
        // 常见命名颜色
        const namedColors = {
            black:'000000', white:'FFFFFF', red:'FF0000', green:'008000', blue:'0000FF',
            yellow:'FFFF00', orange:'FFA500', purple:'800080', pink:'FFC0CB', gray:'808080',
            grey:'808080', silver:'C0C0C0', maroon:'800000', navy:'000080', teal:'008080',
            aqua:'00FFFF', lime:'00FF00', fuchsia:'FF00FF', olive:'808000', brown:'A52A2A',
            darkred:'8B0000', darkgreen:'006400', darkblue:'00008B', lightgray:'D3D3D3',
            darkgray:'A9A9A9', lightgrey:'D3D3D3', darkgrey:'A9A9A9', darkorange:'FF8C00',
            gold:'FFD700', indigo:'4B0082', violet:'EE82EE', tomato:'FF6347',
            cornsilk:'FFF8DC', wheat:'F5DEB3', lavender:'E6E6FA'
        };
        if (namedColors[color]) return namedColors[color];
        return '000000';
    }

    // ===== 全局状态 (导出过程中的共享数据) =====
    let EXPORT = {
        images: [],        // {fileName, contentType, base64, rId, docPrId}
        imgIdMap: new Map(), // data-img-id -> {rId, docPrId}
        nextDocPrId: 1,
        headingConfig: {},
        bodyFormat: {}
    };

    // ===== 主导出函数 =====
    window.exportDocumentAsDocx = window.exportDocx = async function() {
        const editor = window.__editor;
        if (!editor) throw new Error('编辑器未初始化');

        const content = editor.innerHTML;
        if (!content || content.includes('class="placeholder"')) {
            throw new Error('文档为空，请先导入或输入内容');
        }

        EXPORT.headingConfig = window.__editorGetHeadingConfig ? window.__editorGetHeadingConfig() : {};
        EXPORT.bodyFormat = window.__editorGetBodyFormat ? window.__editorGetBodyFormat() : { font: '宋体', size: '10.5pt', lineHeight: '1.5' };
        EXPORT.images = [];
        EXPORT.imgIdMap = new Map();
        EXPORT.nextDocPrId = 1;

        // 克隆内容
        const tempDiv = document.createElement('div');
        tempDiv.innerHTML = content;

        // 清理锚点和折叠相关的编辑属性（不影响导出内容）
        tempDiv.querySelectorAll('[data-pid]').forEach(function(el) { el.removeAttribute('data-pid'); });
        // 移除折叠占位符
        tempDiv.querySelectorAll('.fold-placeholder').forEach(function(el) { el.parentNode.removeChild(el); });
        // 恢复被折叠隐藏的段落
        tempDiv.querySelectorAll('.fold-hidden').forEach(function(el) { el.classList.remove('fold-hidden'); });

        // === 第一步：扫描所有图片，建立映射 ===
        scanImages(tempDiv);

        // === 第二步：创建 ZIP ===
        const zip = new JSZip();

        zip.file('[Content_Types].xml', buildContentTypes());
        zip.file('_rels/.rels', buildRels());
        zip.file('word/_rels/document.xml.rels', buildDocumentRels());
        zip.file('word/document.xml', buildDocumentXml(tempDiv));
        zip.file('word/styles.xml', buildStyles());
        zip.file('word/numbering.xml', buildNumbering());

        // 主题（Word 需要）
        zip.file('word/theme/theme1.xml', buildTheme());

        // 图片文件
        EXPORT.images.forEach(img => {
            zip.file(`word/media/${img.fileName}`, img.base64, { base64: true });
        });

        zip.file('docProps/core.xml', buildCoreProps());
        zip.file('docProps/app.xml', buildAppProps());

        // === 第三步：生成并下载 ===
        const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });

        // 用当前文档名或默认名
        const fileName = 'edited_document.docx';
        saveAs(blob, fileName);
        return true;
    };

    // ===== 扫描图片 =====
    function scanImages(root) {
        const imgs = root.querySelectorAll('img');
        const imageData = window.__editorGetImageData ? window.__editorGetImageData() : new Map();

        imgs.forEach((img, idx) => {
            const imgId = img.dataset.imgId || img.dataset['img-id'] || '';
            let src = img.getAttribute('src') || '';
            let contentType = 'image/png';
            let base64Data = '';

            if (src.startsWith('data:')) {
                const m = src.match(/^data:([^;]+);base64,(.+)$/);
                if (m) { contentType = m[1]; base64Data = m[2]; }
            } else if (imgId && imageData && imageData.has(imgId)) {
                const d = imageData.get(imgId);
                contentType = d.contentType;
                base64Data = d.base64;
            }

            if (base64Data) {
                const ext = contentType.split('/')[1] || 'png';
                const fileName = `image${idx + 1}.${ext}`;
                const rId = 'rIdImg' + (idx + 1);
                const docPrId = EXPORT.nextDocPrId++;

                EXPORT.images.push({ fileName, contentType, base64: base64Data, rId, docPrId });
                if (imgId) {
                    EXPORT.imgIdMap.set(imgId, { rId, docPrId, fileName });
                }
                // 在 img 上记录 rId 供后续转换使用
                img.dataset.exportRId = rId;
                img.dataset.exportDocPrId = docPrId;
            } else {
                img.dataset.exportRId = '';
            }
        });
    }

    // ===== Content Types =====
    function buildContentTypes() {
        let imgOverrides = '';
        EXPORT.images.forEach(img => {
            imgOverrides += `\n    <Override PartName="/word/media/${img.fileName}" ContentType="${img.contentType}"/>`;
        });
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Default Extension="png" ContentType="image/png"/>
    <Default Extension="jpg" ContentType="image/jpeg"/>
    <Default Extension="jpeg" ContentType="image/jpeg"/>
    <Default Extension="gif" ContentType="image/gif"/>
    <Default Extension="bmp" ContentType="image/bmp"/>
    <Default Extension="tiff" ContentType="image/tiff"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
    <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>${imgOverrides}
</Types>`;
    }

    // ===== _rels/.rels =====
    function buildRels() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
    }

    // ===== word/_rels/document.xml.rels =====
    function buildDocumentRels() {
        let rels = `
    <Relationship Id="rIdStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rIdNumbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;
        EXPORT.images.forEach(img => {
            rels += `\n    <Relationship Id="${img.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.fileName}"/>`;
        });
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}
</Relationships>`;
    }

    // ===== word/document.xml =====
    function buildDocumentXml(root) {
        let body = '';
        root.childNodes.forEach(child => {
            body += convertNode(child);
        });

        // 图片相关命名空间（有图时需要；纯文本时仅声明也不影响）
        const imgNs = EXPORT.images.length > 0
            ? ` xmlns:wp="${NS.wp}" xmlns:a="${NS.a}" xmlns:pic="${NS.pic}"`
            : '';
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${NS.w}" xmlns:r="${NS.r}"${imgNs}>
    <w:body>
        ${body || '<w:p><w:r><w:t> </w:t></w:r></w:p>'}
    </w:body>
</w:document>`;
    }

    // ===== 节点转换核心 =====
    function convertNode(node) {
        if (!node) return '';
        if (node.nodeType === Node.TEXT_NODE) {
            const t = node.textContent;
            if (!t.trim()) return '';
            return `<w:r><w:t>${escXml(t)}</w:t></w:r>`;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return '';

        const tag = node.tagName.toLowerCase();

        if (tag.match(/^h[1-6]$/)) return convertHeading(node, tag);
        if (tag === 'p') return convertParagraph(node);
        if (tag === 'pre') return convertPre(node);
        if (tag === 'table') return convertTable(node);
        if (tag === 'img') return convertImage(node);
        if (tag === 'div' && node.classList && node.classList.contains('code-block')) {
            return convertCodeBlock(node);
        }
        // Chromium contenteditable 使用 <div> 作为段落容器，和 <p> 一样处理
        if (tag === 'div') return convertParagraph(node);
        if (['section', 'article', 'main', 'blockquote', 'figure', 'figcaption'].includes(tag)) {
            let r = '';
            node.childNodes.forEach(c => r += convertNode(c));
            return r;
        }
        if (tag === 'ul' || tag === 'ol') return convertList(node, tag);
        if (tag === 'li') return convertListItem(node);
        // 内联
        if (['b', 'strong', 'i', 'em', 'u', 'span', 'a', 'sub', 'sup', 's', 'del', 'code', 'mark', 'font'].includes(tag)) {
            return convertInline(node);
        }
        if (tag === 'br') return '<w:r><w:br/></w:r>';
        // 未知 -> 递归
        let r = '';
        node.childNodes.forEach(c => r += convertNode(c));
        return r;
    }

    // ===== 段落 =====
    function convertParagraph(node) {
        // 检查是否包含 img 的段落
        const imgs = node.querySelectorAll(':scope > img');
        if (imgs.length > 0) {
            let result = '';
            node.childNodes.forEach(child => {
                if (child.nodeType === Node.ELEMENT_NODE && child.tagName.toLowerCase() === 'img') {
                    const isCenter = child.classList.contains('img-center') ||
                        child.style.display === 'block';
                    result += convertImage(child, isCenter);
                } else {
                    result += convertNode(child);
                }
            });
            return result;
        }

        let runs = '';
        node.childNodes.forEach(c => runs += convertNode(c));
        if (!runs.trim()) {
            runs = `<w:r><w:t> </w:t></w:r>`;
        }

        // 纯文本段落，无段落布局属性（段落柄）
        return `<w:p>${runs}</w:p>`;
    }

    // ===== 代码块 =====
    function convertPre(node) {
        // 获取纯文本（保留换行）
        var text = '';
        node.childNodes.forEach(function(ch) {
            if (ch.nodeType === Node.TEXT_NODE) {
                text += ch.textContent;
            } else if (ch.nodeType === Node.ELEMENT_NODE) {
                if (ch.tagName.toLowerCase() === 'code' || ch.tagName.toLowerCase() === 'samp') {
                    text += ch.textContent;
                } else if (ch.tagName.toLowerCase() === 'br') {
                    text += '\n';
                } else {
                    text += ch.textContent;
                }
            }
        });

        // 拆分行
        var lines = text.split('\n');
        var result = '';
        var codeFont = 'Consolas';
        var codeSize = 18; // 9pt in half-points

        lines.forEach(function(line, idx) {
            // 跳过首尾空行
            if (idx === 0 && !line.trim()) return;
            // 行末不能为空格（Word 会 trim，用 &nbsp; 替代尾部空格）
            var trimmedLine = line.replace(/\t/g, '    '); // tab → 4 spaces
            // 修正：用 run 避免 Word 吃掉空格
            var runs = '<w:r><w:rPr>' + codeRPr(codeFont, codeSize) + '</w:rPr><w:t xml:space="preserve">' + escXml(trimmedLine || ' ') + '</w:t></w:r>';
            // 代码块行属性：灰色底色、等宽字体、左缩进
            var pPr = '<w:pBdr><w:left w:val="single" w:sz="6" w:space="4" w:color="999999"/></w:pBdr>' +
                      '<w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>' +
                      '<w:ind w:left="240" w:right="240" w:firstLine="0"/>' +
                      '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
            result += '<w:p><w:pPr>' + pPr + '</w:pPr>' + runs + '</w:p>';
        });

        if (!result) {
            result = '<w:p><w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/></w:pPr><w:r><w:rPr>' + codeRPr(codeFont, codeSize) + '</w:rPr><w:t> </w:t></w:r></w:p>';
        }

        return result;
    }

    // ===== 代码块 div → 导出为 1×1 表格 =====
    function convertCodeBlock(node) {
        // 优先使用原始表格（data-otable）
        var encoded = node.getAttribute && node.getAttribute('data-otable');
        if (encoded) {
            try {
                var decoded = decodeURIComponent(escape(atob(encoded)));
                var tempDiv = document.createElement('div');
                tempDiv.innerHTML = decoded;
                var table = tempDiv.querySelector('table');
                if (table) return convertTable(table);
            } catch(e) {
                console.warn('Failed to decode code block table:', e);
            }
        }
        // 没有原始表格（手动插入的代码块），现场构建 1×1 表格
        var code = node.querySelector('code') || node.querySelector('pre');
        var text = code ? code.textContent : node.textContent;
        // 清理首尾换行
        text = text.replace(/^\n+/, '').replace(/\n+$/, '');
        // 构建 1×1 表格的 HTML
        var lines = text.split('\n');
        var cellHtml = '';
        for (var i = 0; i < lines.length; i++) {
            cellHtml += (i > 0 ? '<w:br/>' : '') + escXml(lines[i]);
        }
        var padding = 8;
        var colWidth = 9000;
        var tblPr = '<w:tblPr><w:tblStyle w:val="TableGrid"/>' +
            '<w:tblW w:w="9000" w:type="dxa"/>' +
            '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="333333"/>' +
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="333333"/>' +
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="333333"/>' +
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="333333"/>' +
            '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>';
        var grid = '<w:gridCol w:w="' + colWidth + '"/>';
        var tcPr = '<w:tcW w:w="9000" w:type="dxa"/>' +
            '<w:vAlign w:val="center"/>';
        var run = '<w:r><w:rPr>' + codeRPr('Consolas', 18) + '</w:rPr>' +
            '<w:t xml:space="preserve">' + cellHtml + '</w:t></w:r>';
        var content = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>' +
            '<w:ind w:left="120" w:right="120" w:firstLine="0"/></w:pPr>' + run + '</w:p>';
        return '<w:tbl>' + tblPr + '<w:tblGrid>' + grid + '</w:tblGrid>' +
            '<w:tr><w:tc>' + tcPr + content + '</w:tc></w:tr></w:tbl>';
    }

    // ===== 标题（纯格式，无段落柄，无编号） =====
    function convertHeading(node, tag) {
        var level = parseInt(tag[1]);
        var cfg = EXPORT.headingConfig[level] || { family: '黑体', size: '16pt', bold: true };
        var family = cfg.family || '黑体';
        var size = ptToHalfPoint(cfg.size || '16pt');
        var bold = cfg.bold !== false;
        // 颜色转换：#1e40af → 1E40AF
        var colorHex = (cfg.color || '#000000').replace('#', '').toUpperCase();

        var runs = '';
        node.childNodes.forEach(function(child) {
            if (child.nodeType === Node.TEXT_NODE) {
                var t = child.textContent;
                if (t.trim()) {
                    runs += '<w:r><w:rPr>' + headingRPr(family, size, bold, colorHex) + '</w:rPr><w:t>' + escXml(t) + '</w:t></w:r>';
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                var ctag = child.tagName.toLowerCase();
                if (ctag === 'img') {
                    runs += convertImage(child);
                } else if (child.classList && child.classList.contains('heading-number')) {
                    var t = child.textContent;
                    if (t.trim()) {
                        runs += '<w:r><w:rPr>' + headingRPr(family, size, bold, colorHex) + '</w:rPr><w:t>' + escXml(t) + '</w:t></w:r>';
                    }
                } else if (['b', 'strong', 'i', 'em', 'u'].indexOf(ctag) !== -1) {
                    var rb = bold, ri = false, ru = false;
                    if (ctag === 'b' || ctag === 'strong') rb = true;
                    if (ctag === 'i' || ctag === 'em') ri = true;
                    if (ctag === 'u') ru = true;
                    var rp = headingRPr(family, size, rb, colorHex);
                    if (ri) rp += '<w:i/>';
                    if (ru) rp += '<w:u w:val="single"/>';
                    var t = child.textContent;
                    if (t.trim()) {
                        runs += '<w:r><w:rPr>' + rp + '</w:rPr><w:t>' + escXml(t) + '</w:t></w:r>';
                    }
                } else {
                    runs += convertNode(child);
                }
            }
        });

        if (!runs) {
            runs = '<w:r><w:rPr>' + headingRPr(family, size, bold, colorHex) + '</w:rPr><w:t>' + escXml(node.textContent || ' ') + '</w:t></w:r>';
        }

        // 纯文本标题，标题/正文区分仅靠 <w:rPr>（字体/字号/颜色），无段落布局
        return '<w:p>' + runs + '</w:p>';
    }

    // ===== 表格 (完整实现) =====
    function convertTable(node) {
        const rows = node.rows || node.querySelectorAll('tr');
        if (!rows || rows.length === 0) return '';

        // 计算列数
        let colCount = 0;
        const firstRow = rows[0];
        if (firstRow) {
            const cells = firstRow.querySelectorAll('td, th');
            cells.forEach(c => {
                colCount += parseInt(c.getAttribute('colspan')) || 1;
            });
        }
        if (colCount === 0) colCount = 1;

        // 网格列
        let grid = '';
        const colWidth = Math.floor(9000 / colCount);
        for (let c = 0; c < colCount; c++) {
            grid += `<w:gridCol w:w="${colWidth}"/>`;
        }

        // 行
        let rowsXml = '';
        Array.from(rows).forEach(row => {
            const cells = row.querySelectorAll('td, th');
            let rowXml = '<w:tr>';

            cells.forEach(cell => {
                const isH = cell.tagName.toLowerCase() === 'th';
                const colspan = parseInt(cell.getAttribute('colspan')) || 1;
                const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

                let content = '';
                const childNodes = cell.childNodes;
                if (childNodes.length === 0) {
                    content = '<w:p><w:r><w:t> </w:t></w:r></w:p>';
                } else {
                    childNodes.forEach(ch => {
                        if (ch.nodeType === Node.TEXT_NODE && ch.textContent.trim()) {
                            content += `<w:p><w:r><w:t>${escXml(ch.textContent)}</w:t></w:r></w:p>`;
                        } else if (ch.nodeType === Node.ELEMENT_NODE) {
                            const ct = ch.tagName.toLowerCase();
                            if (ct === 'p') content += convertParagraph(ch);
                            else if (ct.match(/^h[1-6]$/)) content += convertHeading(ch, ct);
                            else if (ct === 'img') content += convertImage(ch);
                            else if (['b','strong','i','em','u','span'].includes(ct)) {
                                content += '<w:p>' + convertInline(ch) + '</w:p>';
                            } else content += convertNode(ch);
                        }
                    });
                }

                let tcPr = '';
                const cellW = Math.floor((9000 / colCount) * colspan);
                tcPr += `<w:tcW w:w="${cellW}" w:type="dxa"/>`;
                if (colspan > 1) tcPr += `<w:gridSpan w:val="${colspan}"/>`;
                if (isH) tcPr += '<w:shd w:val="clear" w:color="auto" w:fill="F0F0F0"/>';
                tcPr += '<w:vAlign w:val="center"/>';

                rowXml += `<w:tc>${tcPr ? `<w:tcPr>${tcPr}</w:tcPr>` : ''}${content}</w:tc>`;
            });

            rowXml += '</w:tr>';
            rowsXml += rowXml;
        });

        const tblPr = `
            <w:tblPr>
                <w:tblStyle w:val="TableGrid"/>
                <w:tblW w:w="9000" w:type="dxa"/>
                <w:tblBorders>
                    <w:top w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                    <w:left w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                    <w:bottom w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                    <w:right w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                    <w:insideH w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                    <w:insideV w:val="single" w:sz="4" w:space="0" w:color="333333"/>
                </w:tblBorders>
                <w:tblLayout w:type="fixed"/>
            </w:tblPr>`;

        return `<w:tbl>${tblPr}<w:tblGrid>${grid}</w:tblGrid>${rowsXml}</w:tbl>`;
    }

    // ===== 列表 =====
    function convertList(node, tag) {
        let r = '';
        node.childNodes.forEach((ch, i) => {
            if (ch.nodeType === Node.ELEMENT_NODE && ch.tagName.toLowerCase() === 'li') {
                r += convertListItem(ch, tag === 'ol', i + 1);
            } else r += convertNode(ch);
        });
        return r;
    }

    function convertListItem(node, isOrdered, numId) {
        let runs = '';
        node.childNodes.forEach(c => runs += convertNode(c));
        if (!runs.trim()) {
            runs = `<w:r><w:t> </w:t></w:r>`;
        }
        const lh = EXPORT.bodyFormat.lineHeight || '1.5';
        return `<w:p>
            <w:pPr>
                <w:pStyle w:val="ListParagraph"/>
                <w:numPr><w:ilvl w:val="0"/><w:numId w:val="${isOrdered ? 1 : 2}"/></w:numPr>
                <w:spacing w:line="${Math.round(parseFloat(lh) * 240)}" w:lineRule="auto"/>
            </w:pPr>
            ${runs}
        </w:p>`;
    }

    // ===== 内联元素 =====
    function convertInline(node) {
        const tag = node.tagName.toLowerCase();
        let rPr = '';
        if (['b','strong'].includes(tag)) rPr += '<w:b/>';
        if (['i','em'].includes(tag)) rPr += '<w:i/>';
        if (tag === 'u') rPr += '<w:u w:val="single"/>';
        if (tag === 's' || tag === 'del') rPr += '<w:strike/>';
        if (tag === 'sub') rPr += '<w:vertAlign w:val="subscript"/>';
        if (tag === 'sup') rPr += '<w:vertAlign w:val="superscript"/>';
        if (tag === 'code' || tag === 'samp') {
            // 用等宽字体覆盖
            rPr += '<w:rFonts w:ascii="Consolas" w:eastAsia="Consolas"/>';
            rPr += '<w:sz w:val="18"/>';
        }

        // 兼容 <font> 标签（部分浏览器 execCommand 的遗留产物）
        if (tag === 'font') {
            const fc = node.getAttribute('color') || '';
            if (fc) {
                const fHex = cssColorToHex(fc);
                if (fHex !== '000000' && fHex !== 'auto') rPr += `<w:color w:val="${fHex}"/>`;
            }
            const ff = node.getAttribute('face') || '';
            if (ff) {
                const fArr = ff.split(',')[0].trim().replace(/["']/g, '');
                if (fArr) rPr += `<w:rFonts w:ascii="${fArr}" w:eastAsia="${fArr}"/>`;
            }
            const fs = node.getAttribute('size') || '';
            if (fs) {
                // HTML font size 1-7 映射到近似 pt
                const sizeMap = { '1':'7.5pt', '2':'10pt', '3':'12pt', '4':'13.5pt', '5':'18pt', '6':'24pt', '7':'36pt' };
                const mapped = sizeMap[fs] || (parseInt(fs) * 3 + 'pt');
                rPr += `<w:sz w:val="${ptToHalfPoint(mapped)}"/>`;
            }
        }

        if (tag === 'span' && node.style) {
            if (node.style.fontWeight === 'bold' || node.style.fontWeight === '700') rPr += '<w:b/>';
            if (node.style.fontStyle === 'italic') rPr += '<w:i/>';
            if (node.style.textDecorationLine === 'underline' || node.style.textDecoration === 'underline') rPr += '<w:u w:val="single"/>';
            if (node.style.fontFamily) {
                const f = node.style.fontFamily.replace(/["']/g, '');
                rPr += `<w:rFonts w:ascii="${f}" w:eastAsia="${f}"/>`;
            }
            if (node.style.fontSize) {
                rPr += `<w:sz w:val="${ptToHalfPoint(node.style.fontSize)}"/>`;
            }
            // 文字颜色
            if (node.style.color) {
                const colorHex = cssColorToHex(node.style.color);
                if (colorHex !== '000000') {
                    rPr += `<w:color w:val="${colorHex}"/>`;
                }
            }
            // 背景颜色
            if (node.style.backgroundColor && node.style.backgroundColor !== 'transparent' &&
                node.style.backgroundColor !== 'rgba(0, 0, 0, 0)') {
                const bgHex = cssColorToHex(node.style.backgroundColor);
                if (bgHex !== 'auto' && bgHex !== 'FFFFFF' && bgHex !== '000000') {
                    rPr += `<w:shd w:val="clear" w:color="auto" w:fill="${bgHex}"/>`;
                }
            }
        }

        let result = '';
        node.childNodes.forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                const t = child.textContent;
                if (t.trim()) {
                    result += `<w:r><w:rPr>${rPr}</w:rPr><w:t>${escXml(t)}</w:t></w:r>`;
                }
            } else if (child.nodeType === Node.ELEMENT_NODE) {
                const ct = child.tagName.toLowerCase();
                if (ct === 'img') result += convertImage(child);
                else if (ct === 'br') result += '<w:r><w:br/></w:r>';
                else if (['b','strong','i','em','u','span'].includes(ct)) result += convertInline(child);
                else result += convertNode(child);
            }
        });

        if (!result) {
            const t = node.textContent || '';
            if (t.trim()) {
                result = `<w:r><w:rPr>${rPr}</w:rPr><w:t>${escXml(t)}</w:t></w:r>`;
            }
        }
        return result;
    }

    // ===== 图片 (使用预扫描的 rId) =====
    function convertImage(node, forceCenter) {
        const rId = node.dataset.exportRId;
        const docPrId = node.dataset.exportDocPrId || EXPORT.nextDocPrId++;
        const isCenter = forceCenter || node.classList.contains('img-center') ||
            node.style.display === 'block';

        if (!rId) {
            // 没有图片数据，输出占位
            return `<w:p><w:r><w:rPr>${bodyRPr()}</w:rPr><w:t>【图片】</w:t></w:r></w:p>`;
        }

        // 尝试获取图片实际尺寸
        let cx = 3657600, cy = 2052000; // 默认 ~4x2.25 英寸
        const naturalW = node.naturalWidth || node.width;
        const naturalH = node.naturalHeight || node.height;
        if (naturalW && naturalH) {
            // EMU: 1pt = 12700 EMU, 1px ≈ 9525 EMU (at 96dpi)
            const maxWidth = 6000000; // ~6.5 inch
            const scale = Math.min(1, maxWidth / (naturalW * 9525));
            cx = Math.round(naturalW * 9525 * scale);
            cy = Math.round(naturalH * 9525 * scale);
        }

        const jcAttr = isCenter ? 'center' : '';
        const pPr = jcAttr ? `<w:pPr><w:jc w:val="${jcAttr}"/></w:pPr>` : '';

        return `<w:p>${pPr}
            <w:r>
                <w:drawing>
                    <wp:inline distT="0" distB="0" distL="0" distR="0">
                        <wp:extent cx="${cx}" cy="${cy}"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:docPr id="${docPrId}" name="Image ${docPrId}"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic>
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="0" name="Image ${docPrId}"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="${rId}"/>
                                        <a:stretch><a:fillRect/></a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                                        <a:prstGeom prst="rect"/>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>`;
    }

    // ===== 运行属性辅助 =====
    function bodyRPr() {
        return rPr(EXPORT.bodyFormat.font || '宋体', ptToHalfPoint(EXPORT.bodyFormat.size || '10.5pt'), false);
    }

    function rPr(family, halfPt, bold) {
        let p = '';
        p += `<w:rFonts w:ascii="${family}" w:eastAsia="${family}"/>`;
        p += `<w:sz w:val="${halfPt}"/>`;
        if (bold) p += '<w:b/>';
        return p;
    }

    // 标题用 run 属性：加粗 + 大字 + 指定字体 + 颜色
    function headingRPr(family, halfPt, bold, colorHex) {
        let p = '';
        p += `<w:rFonts w:ascii="${family}" w:eastAsia="${family}"/>`;
        p += `<w:sz w:val="${halfPt}"/>`;
        if (bold) p += '<w:b/>';
        p += `<w:color w:val="${colorHex || '000000'}"/>`;
        return p;
    }

    // 代码等宽字体 run 属性
    function codeRPr(family, halfPt) {
        let p = '';
        p += `<w:rFonts w:ascii="${family}" w:eastAsia="${family}"/>`;
        p += `<w:sz w:val="${halfPt}"/>`;
        p += '<w:color w:val="333333"/>';
        return p;
    }

    // ===== Styles =====
    function buildStyles() {
        const bf = EXPORT.bodyFormat;
        const bodyFont = bf.font || '宋体';
        const bodySize = ptToHalfPoint(bf.size || '10.5pt');

        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:r="${NS.r}">
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
        <w:name w:val="Normal"/>
        <w:rPr>
            <w:rFonts w:ascii="${bodyFont}" w:eastAsia="${bodyFont}"/>
            <w:sz w:val="${bodySize}"/>
        </w:rPr>
    </w:style>
    <w:style w:type="paragraph" w:styleId="ListParagraph">
        <w:name w:val="List Paragraph"/>
        <w:basedOn w:val="Normal"/>
        <w:pPr><w:ind w:left="720"/></w:pPr>
    </w:style>
    <w:style w:type="table" w:styleId="TableGrid">
        <w:name w:val="Table Grid"/>
        <w:basedOn w:val="TableNormal"/>
        <w:uiPriority w:val="59"/>
        <w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>
        <w:tblPr>
            <w:tblBorders>
                <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
                <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
            </w:tblBorders>
        </w:tblPr>
    </w:style>
    <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
        <w:name w:val="Default Paragraph Font"/>
        <w:uiPriority w:val="1"/>
        <w:semiHidden/>
    </w:style>
</w:styles>`;
    }

    // ===== Theme =====
    function buildTheme() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Default Theme">
    <a:themeElements>
        <a:clrScheme name="Default">
            <a:dk1><a:srgbClr val="000000"/></a:dk1>
            <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
            <a:dk2><a:srgbClr val="44546A"/></a:dk2>
            <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
            <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
            <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
            <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
            <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
            <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
            <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
            <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
            <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
        </a:clrScheme>
        <a:fontScheme name="Default">
            <a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
            <a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
        </a:fontScheme>
        <a:fmtScheme name="Default">
            <a:fillStyleLst>
                <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
                <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill>
            </a:fillStyleLst>
            <a:lnStyleLst>
                <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
                <a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
            </a:lnStyleLst>
            <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
            <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
        </a:fmtScheme>
    </a:themeElements>
</a:theme>`;
    }

    // ===== Numbering (列表编号) =====
    function buildNumbering() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${NS.w}" xmlns:mc="${NS.mc}" xmlns:r="${NS.r}">
    <!-- 有序列表 -->
    <w:abstractNum w:abstractNumId="0">
        <w:nsid w:val="FFFFFFFF"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="B5A2E4B0"/>
        <w:lvl w:ilvl="0" w:tplc="B5A2E4B0">
            <w:start w:val="1"/>
            <w:numFmt w:val="decimal"/>
            <w:lvlText w:val="%1."/>
            <w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
            <w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>
        </w:lvl>
    </w:abstractNum>
    <!-- 无序列表 -->
    <w:abstractNum w:abstractNumId="1">
        <w:nsid w:val="FFFFFFFE"/>
        <w:multiLevelType w:val="hybridMultilevel"/>
        <w:tmpl w:val="B5A2E4B1"/>
        <w:lvl w:ilvl="0" w:tplc="B5A2E4B1">
            <w:start w:val="1"/>
            <w:numFmt w:val="bullet"/>
            <w:lvlText w:val="●"/>
            <w:lvlJc w:val="left"/>
            <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
            <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="eastAsia"/></w:rPr>
        </w:lvl>
    </w:abstractNum>
    <w:num w:numId="1">
        <w:abstractNumId w:val="0"/>
    </w:num>
    <w:num w:numId="2">
        <w:abstractNumId w:val="1"/>
    </w:num>
</w:numbering>`;
    }

    // ===== docProps =====
    function buildCoreProps() {
        const now = new Date().toISOString();
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>DOCX Editor</dc:creator>
    <cp:lastModifiedBy>DOCX Editor</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
    }

    function buildAppProps() {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>DOCX Editor</Application>
    <DocSecurity>0</DocSecurity>
    <Lines>1</Lines>
    <Paragraphs>1</Paragraphs>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs/>
    <TitlesOfParts/>
    <Template>Normal.dotm</Template>
    <TotalTime>0</TotalTime>
    <AppVersion>16.0000</AppVersion>
</Properties>`;
    }

    console.log('DOCX 导出模块 (v2) 已初始化');

})();
