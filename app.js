// app.js
const express = require('express');
const path = require('path');
const multer = require('multer');
const xlsx = require('xlsx');

// 使用内存存储上传的文件（不用落地到磁盘）
const upload = multer({ storage: multer.memoryStorage() });

const app = express();
const PORT = 3000;

// 静态资源（前端页面）
app.use(express.static(path.join(__dirname, 'public')));

// 处理 Excel 的路由
app.post('/process', upload.single('file'), (req, res) => {
    const { cacheR, mode, suffixA, suffixB, suffixC } = req.body || {};

    if (!req.file) {
        return res.status(400).send('未上传文件');
    }
    if (!cacheR || !mode) {
        return res.status(400).send('参数不足：请填写小区/道路名和模式');
    }

    // 读取上传的 xlsx
    let workbook;
    try {
        workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    } catch (e) {
        console.error(e);
        return res.status(400).send('无法解析 Excel 文件，请确认是 .xlsx 格式');
    }

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // A 列：地址（格式：XXX-YYY-ZZZ 后面可以跟备注）
    // B 列：电话1
    // C 列：电话2（可以为空，但要保留）
    const pattern = /^([^-]+)-([^-]+)-([^-]+)(.*)$/; // XXX-YYY-ZZZ + 可能的后缀备注

    const out = [];
    // 输出表头：只保留你要的那几列
    out.push(['生成地址', '尾部备注', '电话1(原B列)', '电话2(原C列)']);

    rows.forEach((row, index) => {
        if (!row || row.length === 0) return;

        const addrRaw = row[0];          // 原始地址
        const phone1 = row[1] || '';     // 原 B 列
        const phone2 = row[2] || '';     // 原 C 列（即使为空也要有一列）

        let generated = '';  // 生成地址
        let tail = '';       // A 列中拆出来的备注

        if (typeof addrRaw === 'string') {
            const m = addrRaw.trim().match(pattern);
            if (m) {
                let [, XXX, YYY, ZZZ, rest] = m;
                XXX = (XXX || '').trim();
                YYY = (YYY || '').trim();
                ZZZ = (ZZZ || '').trim();
                tail = (rest || '').trim();

                const sA = suffixA || '';
                const sB = suffixB || '';
                const sC = suffixC || '';

                if (mode === 'community') {
                    // 小区开头：缓存R + YYY + 变量B + ZZZ + 变量C
                    generated = `${cacheR}${YYY}${sB}${ZZZ}${sC}`;
                } else {
                    // 道路开头：缓存R + XXX + 变量A + YYY + 变量B + ZZZ + 变量C
                    generated = `${cacheR}${XXX}${sA}${YYY}${sB}${ZZZ}${sC}`;
                }
            }
        }

        out.push([generated, tail, phone1, phone2]);
    });

    // 把结果写回新的 Excel 工作簿
    const outBook = xlsx.utils.book_new();
    const outSheet = xlsx.utils.aoa_to_sheet(out);
    xlsx.utils.book_append_sheet(outBook, outSheet, '结果');

    const buffer = xlsx.write(outBook, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader(
        'Content-Disposition',
        'attachment; filename="电话表整理结果.xlsx"'
    );
    res.setHeader(
        'Content-Type',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.send(buffer);
});

// 启动服务
app.listen(PORT, () => {
    console.log(`Server running at http://localhost:${PORT}`);
});