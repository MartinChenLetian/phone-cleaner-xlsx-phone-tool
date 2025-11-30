// src/index.js
import * as XLSX from "xlsx";

// 生成首页 HTML（带表单）
function renderFormHtml() {
    return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8" />
  <title>电话表整理小工具</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f5f6fa;
      margin: 0;
      padding: 0;
    }
    .container {
      max-width: 720px;
      margin: 40px auto;
      background: #ffffff;
      padding: 24px 28px 32px;
      border-radius: 12px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.06);
    }
    h1 {
      margin-top: 0;
      font-size: 24px;
      text-align: center;
    }
    .desc {
      font-size: 14px;
      color: #555;
      line-height: 1.6;
    }
    fieldset {
      border: 1px solid #ddd;
      border-radius: 8px;
      margin-top: 18px;
      padding: 16px 18px 12px;
    }
    legend {
      padding: 0 8px;
      font-size: 14px;
      font-weight: 600;
    }
    label {
      display: block;
      margin-top: 10px;
      font-size: 14px;
    }
    input[type="text"],
    input[type="file"] {
      display: block;
      margin-top: 6px;
      padding: 6px 8px;
      width: 100%;
      max-width: 380px;
      border-radius: 6px;
      border: 1px solid #ccc;
      font-size: 14px;
      box-sizing: border-box;
    }
    .mode {
      margin-top: 6px;
      font-size: 14px;
    }
    .mode label {
      display: inline-block;
      margin-right: 16px;
    }
    .options {
      margin-top: 10px;
      font-size: 14px;
    }
    .hidden { display: none; }
    .formula {
      background: #f7f7ff;
      border-radius: 6px;
      padding: 6px 8px;
      font-size: 13px;
      margin: 8px 0;
    }
    .tip {
      font-size: 12px;
      color: #888;
    }
    code {
      background: #f0f0f0;
      padding: 1px 4px;
      border-radius: 4px;
    }
    .submit-btn {
      margin-top: 20px;
      width: 100%;
      padding: 10px 16px;
      border-radius: 8px;
      border: none;
      font-size: 15px;
      font-weight: 600;
      cursor: pointer;
      background: #4c6fff;
      color: #ffffff;
    }
    .submit-btn:hover { background: #3c59d4; }
    .footer {
      margin-top: 16px;
      font-size: 12px;
      text-align: center;
      color: #888;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>电话表整理小工具</h1>
    <p class="desc">
      使用说明：<br/>
      1. A 列为地址（格式：<code>417-21-103</code>，后面可以跟备注）<br/>
      2. B 列为电话1，C 列为电话2（C 即使为空也会保留）<br/>
    </p>

    <form id="form" action="/process" method="post" enctype="multipart/form-data">
      <fieldset>
        <legend>步骤 1：填写街道 / 小区名（缓存R）</legend>
        <label>
          您当前制作的街道 / 小区名是？
          <input
            type="text"
            name="cacheR"
            onkeyup="document.getElementById('rx').innerHTML = this.value;console.log('乾溪新村');"
            required
            placeholder="例如：甘泉二村 / 环镇北路"
          />
        </label>
      </fieldset>

      <fieldset>
        <legend>步骤 2：上传原始电话表（.xlsx）</legend>
        <label>
          请选择文件（第一张表：A=地址，B=电话1，C=电话2）：
          <input type="file" name="file" accept=".xlsx" required />
        </label>
      </fieldset>

      <fieldset>
        <legend>步骤 3：生成规则设置</legend>
        <div class="mode">
          <span>生成片段开头：</span>
          <label>
            <input type="radio" name="mode" value="community" checked />
            以 <strong>小区</strong> 开头
          </label>
          <label>
            <input type="radio" name="mode" value="road" />
            以 <strong>道路</strong> 开头
          </label>
        </div>

        <div id="community-options" class="options">
          <p>当前模式：小区开头</p>
          <p class="formula">
            生成地址 = <code>
            <span id="rx">R</span>  22  <span id="hao">A</span>  303  <span id="shi">B</span> </code>
          </p>
          <label>
            变量 B（例如：号）：
            <input type="text" name="suffixB" placeholder="例如：号" onkeyup="document.getElementById('hao').innerHTML = this.value;console.log('21');" required/>
          </label>
          <label>
            变量 C（例如：室）：
            <input type="text" name="suffixC" placeholder="例如：室" onkeyup="document.getElementById('shi').innerHTML = this.value;console.log('103');" required/>
          </label>
        </div>

        <div id="road-options" class="options hidden">
          <p>当前模式：道路开头</p>
          <p class="formula">
            生成地址 = <code>
            <span id="rx">R</span>  400  <span id="nong2">A</span>  22  <span id="hao2">B</span>  303  <span id="shi2">C</span> </code>
          </p>
          <label>
            变量 A（例如：弄）：
            <input type="text" name="suffixA" placeholder="例如：弄" onkeyup="document.getElementById('nong2').innerHTML = this.value;console.log('417');" required />
          </label>
          <label>
            变量 B（例如：号）：
            <input type="text" name="suffixB_road" placeholder="例如：号" onkeyup="document.getElementById('hao2').innerHTML = this.value;console.log('21');" required />
          </label>
          <label>
            变量 C（例如：室）：
            <input type="text" name="suffixC_road" placeholder="例如：室" onkeyup="document.getElementById('shi2').innerHTML = this.value;console.log('103');" required />
          </label>
          <p class="tip">
            提交前会自动把这里的 B/C 写回通用字段，后端统一处理。
          </p>
        </div>
      </fieldset>

      <button type="submit" class="submit-btn">开始整理并下载结果</button>
    </form>

    <p class="footer">手机号整理专用小工具 · Cloudflare Workers 陈乐天发布专属尊贵版</p>
  </div>

  <script>
    const modeRadios = document.querySelectorAll('input[name="mode"]');
    const communityOptions = document.getElementById('community-options');
    const roadOptions = document.getElementById('road-options');
    const form = document.getElementById('form');

    modeRadios.forEach((radio) => {
      radio.addEventListener('change', () => {
        if (radio.value === 'community' && radio.checked) {
          communityOptions.classList.remove('hidden');
          roadOptions.classList.add('hidden');
        } else if (radio.value === 'road' && radio.checked) {
          communityOptions.classList.add('hidden');
          roadOptions.classList.remove('hidden');
        }
      });
    });

    form.addEventListener('submit', (e) => {
      const mode = document.querySelector('input[name="mode"]:checked').value;
      if (mode === 'road') {
        const suffixAInput = form.querySelector('input[name="suffixA"]');
        const suffixBRoad = form.querySelector('input[name="suffixB_road"]');
        const suffixCRoad = form.querySelector('input[name="suffixC_road"]');

        let suffixBCommon = form.querySelector('input[name="suffixB"]');
        let suffixCCommon = form.querySelector('input[name="suffixC"]');

        if (!suffixBCommon) {
          suffixBCommon = document.createElement('input');
          suffixBCommon.type = 'hidden';
          suffixBCommon.name = 'suffixB';
          form.appendChild(suffixBCommon);
        }
        if (!suffixCCommon) {
          suffixCCommon = document.createElement('input');
          suffixCCommon.type = 'hidden';
          suffixCCommon.name = 'suffixC';
          form.appendChild(suffixCCommon);
        }

        suffixBCommon.value = suffixBRoad.value || '';
        suffixCCommon.value = suffixCRoad.value || '';
        suffixAInput.value = suffixAInput.value || '';
      }
    });
  </script>
</body>
</html>`;
}

// Worker 主逻辑
export default {
    async fetch(request, env, ctx) {
        const url = new URL(request.url);

        // GET / => 返回网页
        if (request.method === "GET" && url.pathname === "/") {
            return new Response(renderFormHtml(), {
                status: 200,
                headers: {
                    "content-type": "text/html; charset=utf-8"
                }
            });
        }

        // POST /process => 处理 Excel 并返回新文件
        if (request.method === "POST" && url.pathname === "/process") {
            const formData = await request.formData();

            const cacheR = (formData.get("cacheR") || "").toString().trim();
            const mode = (formData.get("mode") || "community").toString();
            const suffixA = (formData.get("suffixA") || "").toString();
            const suffixB = (formData.get("suffixB") || "").toString();
            const suffixC = (formData.get("suffixC") || "").toString();

            const file = formData.get("file");
            if (!(file instanceof File)) {
                return new Response("未找到上传文件", { status: 400 });
            }

            const arrayBuffer = await file.arrayBuffer();

            let workbook;
            try {
                workbook = XLSX.read(arrayBuffer, { type: "array" });
            } catch (e) {
                return new Response("无法解析 Excel 文件，请确认是 .xlsx 格式", {
                    status: 400
                });
            }

            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            const pattern = /^([^-]+)-([^-]+)-([^-]+)(.*)$/;

            const out = [];
            out.push(["地址", "备注", "电话1", "电话2"]);

            for (const row of rows) {
                if (!row || row.length === 0) continue;

                const addrRaw = row[0];
                const phone1 = row[1] || "";
                const phone2 = row[2] || "";

                let generated = "";
                let tail = "";

                if (typeof addrRaw === "string") {
                    const m = addrRaw.trim().match(pattern);
                    if (m) {
                        let [, XXX, YYY, ZZZ, rest] = m;
                        XXX = (XXX || "").trim();
                        YYY = (YYY || "").trim();
                        ZZZ = (ZZZ || "").trim();

                        // ===== 关键新增逻辑：把 ZZZ 拆成「门牌号 + 备注的一部分」 =====
                        let zRaw = ZZZ;
                        let door = zRaw;        // 真正的门牌号
                        let extraRemark = "";   // 从 ZZZ 里拆出来的备注

                        // 匹配「数字开头 + 其他东西」的情况，例如：102租户 / 104 女 / 302 Z A翠
                        const mDoor = zRaw.match(/^(\d+)([\s\S]*)$/);
                        if (mDoor) {
                            door = mDoor[1];                      // 前面的数字
                            extraRemark = (mDoor[2] || "").trim(); // 后面的内容当备注
                        }

                        const restTrim = (rest || "").trim();

                        // 汇总备注：ZZZ 里多出来的 + 第四段 rest
                        const remarkParts = [];
                        if (extraRemark) remarkParts.push(extraRemark);
                        if (restTrim) remarkParts.push(restTrim);
                        tail = remarkParts.join(" ");           // tail 就是我们要写到「尾部备注」那一列

                        // ===== 按新的 door（纯门牌号）来生成地址 =====
                        if (mode === "community") {
                            // 小区开头：缓存R + YYY + 变量B + door + 变量C
                            generated = `${cacheR}${YYY}${suffixB}${door}${suffixC}`;
                        } else {
                            // 道路开头：缓存R + XXX + 变量A + YYY + 变量B + door + 变量C
                            generated = `${cacheR}${XXX}${suffixA}${YYY}${suffixB}${door}${suffixC}`;
                        }
                    }
                }

                out.push([generated, tail, phone1, phone2]);
            }

            const outBook = XLSX.utils.book_new();
            const outSheet = XLSX.utils.aoa_to_sheet(out);
            XLSX.utils.book_append_sheet(outBook, outSheet, "结果");

            const outArrayBuffer = XLSX.write(outBook, {
                bookType: "xlsx",
                type: "array"
            });

            return new Response(outArrayBuffer, {
                status: 200,
                headers: {
                    "Content-Type":
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "Content-Disposition":
                        'attachment; filename="{{ resource R }}.xlsx"'
                }
            });
        }

        return new Response("Not found", { status: 404 });
    }
};