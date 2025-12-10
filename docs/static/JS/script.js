document.addEventListener("DOMContentLoaded", () => {
    const downloadBtn = document.getElementById("downloadBtn");
    if (!downloadBtn) return;

    downloadBtn.addEventListener("click", async function (e) {
        e.preventDefault();
        this.disabled = true;

        console.log("💾 Excelダウンロード処理開始");

        const data = {
            search_park: document.getElementById("search_park")?.value || "",
            inspection_year: document.getElementById("inspection_year")?.value || "",
            install_year_num: document.getElementById("install_year_num")?.value || "",
            inspection_sections: window.inspection_sections ?? [
  {
    "section": "柱・梁（本体）",
    "items": [
      { "name": "pillar_wobble", "label": "ぐらつき",
        "excel": {
            "B": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "pillar_damage", "label": "破損",
        "excel": {
            "B": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "pillar_deform", "label": "変形",
        "excel": {
            "B": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "pillar_corrosion", "label": "腐食（腐朽）",
        "excel": {
            "B": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "pillar_joint_loose", "label": "〔接合部・ボルト〕緩み",
        "excel": {
            "B": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D6", "dx": 2, "dy": 2, "icon" : "none.png" } } }
    ]
  },

  {
    "section": "接合部（継ぎ手）",
    "items": [
      { "name": "joint_damage", "label": "破損",
        "excel": {
            "B": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "joint_deform", "label": "変形",
        "excel": {
            "B": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "joint_corrosion", "label": "腐食",
        "excel": {
            "B": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "joint_bolt_loose", "label": "ボルトの緩み",
        "excel": {
            "B": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "joint_missing", "label": "欠落",
        "excel": {
            "B": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D7", "dx": 2, "dy": 2, "icon" : "none.png" } } }
    ]
  },

  {
    "section": "吊金具",
    "items": [
      { "name": "hanger_damage", "label": "破損",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_deform", "label": "変形",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_corrosion", "label": "腐食",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_noise", "label": "異音",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_shift", "label": "金具本体のずれ",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_wear_13", "label": "摩耗（×：1/3以上）",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_wear_12", "label": "摩耗（×：1/2以上 使用禁止）",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "hanger_bolt", "label": "ボルトの緩み／欠落",
        "excel": {
            "B": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D8", "dx": 2, "dy": 2, "icon" : "none.png" } } }
    ]
  },

  {
    "section": "揺動部（チェーン・ロープ）",
    "items": [
      { "name": "chain_twist", "label": "ねじれ",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_deform", "label": "変形",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_damage", "label": "破損",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_fray", "label": "ほつれ",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_break", "label": "断線",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_wear_13", "label": "摩耗（×：1/3以上）",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } },

      { "name": "chain_wear_12", "label": "摩耗（×：1/2以上 使用禁止）",
        "excel": {
            "B": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2 , "icon" : "triangle.png"},
            "C": { "type": "icon", "cell": "D9", "dx": 2, "dy": 2, "icon" : "none.png" } } }
    ]
  },
  {
  "section": "揺動部（座板・座面）",
  "items": [
    {
      "name": "seat_crack","label": "ヒビ",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_break","label": "割れ",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_deform","label": "湾曲等変形",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_damage","label": "破損",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_rot","label": "腐朽",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_metal_wear_13","label": "金具の摩耗（×：1/3以上）",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_metal_wear_12","label": "金属の摩耗（×：1/2以上 使用禁止）",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_bolt_loose",
      "label": "ボルト・袋ナットの緩み",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "seat_bolt_missing","label": "欠落",
      "excel": {
        "B": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D10", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    }
  ]
},
{
  "section": "安全柵",
  "items": [
    {
      "name": "fence_wobble","label": "ぐらつき",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "fence_damage","label": "破損",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "fence_deform","label": "変形",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "fence_corrosion","label": "腐食",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "fence_joint_loose","label": "〔接合部・ボルト〕緩み",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "fence_missing","label": "欠落",
      "excel": {
        "B": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D11", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    }
  ]
},
{
  "section": "その他",
  "items": [
    {
      "name": "other_sharp","label": "異物",
      "excel": {
        "B": { "type": "icon", "cell": "D12", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D12", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "other_sign","label": "落書き",
      "excel": {
        "B": { "type": "icon", "cell": "D12", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D12", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    }
  ]
},
{
  "section": "基礎",
  "items": [
    {
      "name": "base_sink","label": "基礎の露出",
      "excel": {
        "B": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "base_crack","label": "亀裂",
      "excel": {
        "B": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "base_expose","label": "破損",
      "excel": {
        "B": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D13", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    }
  ]
},
{
  "section": "地表部・安全柵内",
  "items": [
    {
      "name": "ground_uneven","label": "大きな凹凸",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "ground_exposed_stone_root","label": "石や根の露出",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "ground_foreign_object","label": "異物",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "ground_mat_flip","label": "マットのめくれ",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "ground_mat_damage","label": "マットの破損",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    },
    {
      "name": "ground_tree_branch","label": "樹木の枝",
      "excel": {
        "B": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "triangle.png"},
        "C": { "type": "icon", "cell": "D14", "dx": 2, "dy": 2, "icon" : "none.png" }
      }
    }
  ]
}



        ],
            items: []
        };

        // === 1. tbody 点検結果 ===
        const inspectionResults = {};
        document.querySelectorAll("tbody tr").forEach(tr => {
            const radioChecked = tr.querySelector("input[type='radio']:checked");
            if (radioChecked && radioChecked.name) {
                inspectionResults[radioChecked.name] = radioChecked.value;
            }
        });

        data.inspection_sections.forEach(section => {
            section.items.forEach(item => {
                const result = inspectionResults[item.name] || "A";
                if (result === "A") return;
                const excelDef = item.excel?.[result];
                if (!excelDef) return;
                data.items.push({
                    type: "checkbox",
                    name: item.name,
                    value: result,
                    cell: excelDef.cell,
                    dx: excelDef.dx ?? 0,
                    dy: excelDef.dy ?? 0,
                    icon: excelDef.icon,
                    text: excelDef.text ?? ""
                });
            });
        });

        // === 2. CheckSheet_measure_area 入力 ===
        const measureArea = document.querySelector(".CheckSheet_measure_area");
        if (measureArea) {
            measureArea.querySelectorAll("input,textarea").forEach(input => {
                if ((input.type === "checkbox" || input.type === "radio") && input.checked) {
                    data.items.push({ type: input.type, name: input.name, value: input.value });
                } else if ((input.type === "text" || input.type === "number") && input.value.trim()) {
                    data.items.push({ type: input.type, name: input.name, value: input.value.trim() });
                }
            });
        }

        // === 3. 点検時に実施した措置 ===
        const actionMap = {
            "action_grease": { cell: "F6", dx: 0, dy: 0 },
            "action_bolt": { cell: "F6", dx: 20, dy: 0 },
            "action_hanger": { cell: "F6", dx: 0, dy: 20, inputId: "hanger_count" },
            "action_chain": { cell: "F6", dx: 0, dy: 40, inputId: "chain_count" },
            "action_seat": { cell: "F6", dx: 0, dy: 60, inputId: "seat_count" },
            "action_removal": { cell: "F6", dx: 20, dy: 20 },
            "action_other": { cell: "F6", dx: 20, dy: 40, inputId: "action_other_detail" }
        };

        Object.keys(actionMap).forEach(name => {
            const cb = document.getElementById(name);
            if (cb && cb.checked) {
                const cfg = actionMap[name];
                const item = {
                    type: "checkbox",
                    name: name,
                    value: cb.value,
                    cell: cfg.cell,
                    dx: cfg.dx,
                    dy: cfg.dy,
                    icon: "icons/check.png"
                };
                if (cfg.inputId) {
                    const input = document.getElementById(cfg.inputId);
                    if (input && input.value) item.text = input.value;
                }
                data.items.push(item);
            }
        });

        // === 4. 所見 ===
        const observations = document.getElementById("observations");
        if (observations && observations.value.trim()) {
            data.items.push({
                type: "text",
                name: "observations",
                value: observations.value.trim(),
                cell: "F10",
                text: observations.value.trim()
            });
        }

        // === 5. 総合結果 ===
        const overallMap = {
            "overall_a": { cell: "F13", dx: 0, dy: 0 },
            "overall_b": { cell: "F13", dx: 0, dy: 20 },
            "overall_c": { cell: "F13", dx: 0, dy: 40 },
            "overall_d": { cell: "F13", dx: 0, dy: 60, inputId: "overall_d_detail" }
        };

        Object.keys(overallMap).forEach(id => {
            const radio = document.getElementById(id);
            if (radio && radio.checked) {
                const cfg = overallMap[id];
                const item = {
                    type: "radio",
                    name: "overall_result",
                    value: radio.value,
                    cell: cfg.cell,
                    dx: cfg.dx,
                    dy: cfg.dy,
                    icon: "icons/check.png"
                };
                if (cfg.inputId) {
                    const input = document.getElementById(cfg.inputId);
                    if (input && input.value) item.text = input.value;
                }
                data.items.push(item);
            }
        });

        // === 6. 対応方針 ===
        const planMap = {
            "plan_maintenance": { cell: "H6", dx: 0, dy: 0 },
            "plan_repair": { cell: "H6", dx: 20, dy: 0 },
            "plan_improvement": { cell: "H6", dx: 0, dy: 20 },
            "plan_precision": { cell: "H6", dx: 20, dy: 20 },
            "plan_removal": { cell: "H6", dx: 0, dy: 40 },
            "plan_other": { cell: "H6", dx: 20, dy: 40, inputId: "plan_other_detail" }
        };

        Object.keys(planMap).forEach(name => {
            const cb = document.getElementById(name);
            if (cb && cb.checked) {
                const cfg = planMap[name];
                const item = {
                    type: "checkbox",
                    name: name,
                    value: cb.value,
                    cell: cfg.cell,
                    dx: cfg.dx,
                    dy: cfg.dy,
                    icon: "icons/check.png"
                };
                if (cfg.inputId) {
                    const input = document.getElementById(cfg.inputId);
                    if (input && input.value) item.text = input.value;
                }
                data.items.push(item);
            }
        });

        // === 7. 対応予定時期 ===
        const month = document.getElementById("response_month")?.value;
        ["period_early","period_mid","period_late"].forEach(id => {
            const radio = document.getElementById(id);
            if (radio && radio.checked) {
                data.items.push({
                    type: "radio",
                    name: "period",
                    value: radio.value,
                    cell: "H10",
                    dx: 0,
                    dy: 0,
                    icon: "icons/check.png",
                    text: month ? `${month}月 ${radio.nextElementSibling.textContent}` : radio.nextElementSibling.textContent
                });
            }
        });

        // === 8. 本格的使用禁止措置 ===
        const prohibition_cb = document.getElementById("prohibition_measure");
        if (prohibition_cb && prohibition_cb.checked) {
            data.items.push({
                type: "checkbox",
                name: "prohibition_measure",
                value: prohibition_cb.value,
                cell: "H11",
                dx: 0,
                dy: 0,
                icon: "icons/check.png"
            });
        }
        const prohibition_date = document.getElementById("prohibition_date")?.value;
        const prohibition_status = ["prohibition_done","prohibition_planned"].find(id => document.getElementById(id)?.checked);
        if (prohibition_date || prohibition_status) {
            data.items.push({
                type: "text",
                name: "prohibition_detail",
                cell: "H11",
                text: `${prohibition_date || ""} ${prohibition_status ? document.getElementById(prohibition_status).nextElementSibling.textContent : ""}`
            });
        }

        // === 9. 備考 ===
        const remarks = document.getElementById("remarks");
        if (remarks && remarks.value.trim()) {
            data.items.push({
                type: "text",
                name: "remarks",
                value: remarks.value.trim(),
                cell: "F12",
                text: remarks.value.trim()
            });
        }

        console.log("=== Excelに送信される items ===", data.items);

        // === 10. Flask API へ POST ===
        try {
            const response = await fetch("http://127.0.0.1:5000/api/generate_excel", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(data)
            });

            if (!response.ok) throw new Error("Excel生成に失敗しました");

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = "点検チェックシート.xlsx";
            document.body.appendChild(a);
            a.click();
            a.remove();
            window.URL.revokeObjectURL(url);

        } catch (error) {
            alert(error.message);
            console.error(error);
        } finally {
            this.disabled = false;
        }
    });
});