// docs/app.js

const pattern = /^([^-]+)-([^-]+)-([^-]+)(.*)$/;

document.addEventListener("DOMContentLoaded", () => {
    const form = document.getElementById("form");
    const fileInput = document.getElementById("fileInput");

    const modeRadios = document.querySelectorAll('input[name="mode"]');
    const communityOptions = document.getElementById('community-options');
    const roadOptions = document.getElementById('road-options');

    // 和 Worker 页面一样的模式切换逻辑
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

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        const cacheR = (form.cacheR.value || "").trim();
        if (!cacheR) {
            alert("请先填写街道 / 小区名（缓存R）");
            return;
        }

        if (!fileInput.files || fileInput.files.length === 0) {
            alert("请先选择一个 .xlsx 文件");
            return;
        }

        const mode = document.querySelector('input[name="mode"]:checked').value;

        // 道路模式：把 suffixB_road / suffixC_road 写回通用字段
        if (mode === "road") {
            const suffixAInput = form.querySelector('input[name="suffixA"]');
            const suffixBRoad = form.querySelector('input[name="suffixB_road"]');
            const suffixCRoad = form.querySelector('input[name="suffixC_road"]');

            if (
                !suffixAInput.value.trim() ||
                !suffixBRoad.value.trim() ||
                !suffixCRoad.value.trim()
            ) {
                alert("道路模式下，变量 A / B / C 不能为空");
                return;
            }


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
        } else if (mode === "community") {
            const suffixBInput = form.querySelector('input[name="suffixB"]');
            const suffixCInput = form.querySelector('input[name="suffixC"]');
            if (!suffixBInput.value.trim() || !suffixCInput.value.trim()) {
                alert("小区模式下，变量 B / C 不能为空");
                return;
            }
        }

        const suffixA = (form.suffixA && form.suffixA.value) || "";
        const suffixB = (form.suffixB && form.suffixB.value) || "";
        const suffixC = (form.suffixC && form.suffixC.value) || "";

        

        const file = fileInput.files[0];

        try {
            const arrayBuffer = await readFileAsArrayBuffer(file);

            const workbook = XLSX.read(arrayBuffer, { type: "array" });
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

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

                        let zRaw = ZZZ;
                        let door = zRaw;
                        let extraRemark = "";

                        const mDoor = zRaw.match(/^(\d+)([\s\S]*)$/);
                        if (mDoor) {
                            door = mDoor[1];
                            extraRemark = (mDoor[2] || "").trim();
                        }

                        const restTrim = (rest || "").trim();

                        const remarkParts = [];
                        if (extraRemark) remarkParts.push(extraRemark);
                        if (restTrim) remarkParts.push(restTrim);
                        tail = remarkParts.join(" ");

                        if (mode === "community") {
                            generated = `${cacheR}${YYY}${suffixB}${door}${suffixC}`;
                        } else {
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

            const blob = new Blob([outArrayBuffer], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            });

            const a = document.createElement("a");
            const url = URL.createObjectURL(blob);
            const fileName = (cacheR || "结果") + ".xlsx";
            a.href = url;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            a.remove();
            URL.revokeObjectURL(url);
        } catch (err) {
            console.error(err);
            alert("解析或生成 Excel 时出错，请检查文件格式。");
        }
    });
});

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsArrayBuffer(file);
    });
}