document.addEventListener("DOMContentLoaded", () => {
    const downloadBtn = document.getElementById("downloadBtn");
    if (!downloadBtn) return;

    downloadBtn.addEventListener("click", async (e) => {
        e.preventDefault();
        downloadBtn.disabled = true;
        console.log("💾 Excelダウンロード処理開始");

        const data = {
            search_park: document.getElementById("search_park")?.value || "",
            inspection_year: document.getElementById("inspection_year")?.value || "",
            install_year_num: document.getElementById("install_year_num")?.value || "",
            inspection_sections: window.inspection_sections || [
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
// 点検結果取得
        const inspectionResults = {};
        document.querySelectorAll("tbody tr").forEach(tr => {
            const radioChecked = tr.querySelector("input[type='radio']:checked");
            if (radioChecked && radioChecked.name) {
                inspectionResults[radioChecked.name] = radioChecked.value;
            }
        });

        // Excel items 作成（dx, dy の微調整対応）
        data.inspection_sections.forEach(section => {
            section.items.forEach(item => {
                const result = inspectionResults[item.name] || "A";
                if (result === "A") return;
                const excelDef = item.excel?.[result];
                if (!excelDef) return;

                // 微調整のため dxInput / dyInput がある場合は取得
                let dxAdj = 0, dyAdj = 0;
                const dxEl = document.getElementById(item.name + "_dx");
                const dyEl = document.getElementById(item.name + "_dy");
                if (dxEl) dxAdj = parseInt(dxEl.value) || 0;
                if (dyEl) dyAdj = parseInt(dyEl.value) || 0;

                data.items.push({
                    type: "checkbox",
                    name: item.name,
                    value: result,
                    cell: excelDef.cell,
                    dx: (excelDef.dx ?? 0) + dxAdj,
                    dy: (excelDef.dy ?? 0) + dyAdj,
                    icon: excelDef.icon,
                    text: excelDef.text ?? ""
                });
            });
        });

        // CheckSheet_measure_area 入力反映（dx/dy微調整も対応）
        document.querySelectorAll(".CheckSheet_measure_area input, .CheckSheet_measure_area textarea").forEach(input => {
            let dxAdj = 0, dyAdj = 0;
            const dxEl = document.getElementById(input.name + "_dx");
            const dyEl = document.getElementById(input.name + "_dy");
            if (dxEl) dxAdj = parseInt(dxEl.value) || 0;
            if (dyEl) dyAdj = parseInt(dyEl.value) || 0;

            if ((input.type === "checkbox" || input.type === "radio") && input.checked) {
                data.items.push({
                    type: input.type,
                    name: input.name,
                    value: input.value,
                    dx: dxAdj,
                    dy: dyAdj
                });
            } else if ((input.type === "text" || input.type === "number") && input.value.trim()) {
                data.items.push({
                    type: input.type,
                    name: input.name,
                    value: input.value.trim(),
                    dx: dxAdj,
                    dy: dyAdj
                });
            }
        });

        
        // === 措置・所見・総合結果・対応方針・対応予定時期・禁止措置・備考 ===
        const appendItems = (map, type="checkbox") => {
            Object.keys(map).forEach(name => {
                const el = document.getElementById(name);
                if (el && (el.checked || el.type === "radio")) {
                    const cfg = map[name];
                    const item = {
                        type: el.type,
                        name: el.type === "radio" ? "overall_result" : name,
                        value: el.value,
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
        };

        // ==========================
        // ●点検時に実施した措置 (F6:G9)
        // ==========================
        let actionText = "●点検時に実施した措置\n";
        const actionCounts = {
            "吊金具": document.getElementById("hanger_count")?.value || 0,
            "チェーン": document.getElementById("chain_count")?.value || 0,
            "座板": document.getElementById("seat_count")?.value || 0
        };
        actionText += `□グリース・オイル等の注入※1\n`;
        actionText += `□ボルト・ナットの増し締め・交換 \n`;
        actionText += `□吊金具の交換 (${actionCounts["吊金具"]}箇所)\n`;
        actionText += `□チェーンの交換 (${actionCounts["チェーン"]}箇所)\n`;
        actionText += `□座板の交換 (${actionCounts["座板"]}箇所)\n`;
        actionText += `□石・異物の除去、枝の剪定\n`;
        const otherText = document.getElementById("action_other_detail")?.value || "";
        if (otherText) actionText += `□その他 (${otherText})\n`;


        data.items.push({
            type: "text",
            name: "action_text",
            value: actionText,
            cell: "F6",
            text: actionText
        });

        // チェックボックスに対応する check.png を配置
        const actionChecks = [
            {id: "check_grease", dx:0, dy:0},
            {id: "check_bolt", dx:0, dy:120},
            {id: "check_hanger", dx:0, dy:240},
            {id: "check_chain", dx:0, dy:360},
            {id: "check_seat", dx:0, dy:480},
            {id: "check_other", dx:0, dy:600},
        ];
        actionChecks.forEach(chk=>{
            const el=document.getElementById(chk.id);
            if(el && el.checked){
                data.items.push({
                    type:"icon",
                    cell:"F6",
                    icon:"check.png",
                    dx:chk.dx,
                    dy:chk.dy
                });
            }
        });

        // ==========================
        // ●所見 (F10:G12)
        // ==========================
        const observations = document.getElementById("observations")?.value.trim();
        if(observations){
            data.items.push({
                type:"text",
                name:"observations",
                value:"●所見\r\n"+observations,
                cell:"F10",
                text:"●所見\r\n"+observations
            });
        }

        // ==========================
        // ●総合結果 (F13:G15)
        // ==========================
        const totalResultRadios = document.querySelectorAll('input[name="total_result"]');
        totalResultRadios.forEach(r=>{
            if(r.checked){
                const val = r.value;
                let dxVal = val==="A"?0:val==="B"?120:val==="C"?240:360;
                let totalText="";
                if(val==="D"){
                    const dDetail=document.getElementById("total_result_D_text")?.value || "";
                    const totalText=`D:使用禁止措置\r\n（${dDetail}）`;
                    data.items.push({
                        type:"text",
                        name:"total_result_text",
                        value:totalText,
                        cell:"F13",
                        text:totalText
                    });
                    data.items.push({
                        type:"icon",
                        cell:"F13",
                        icon:"check.png",
                        dx:360, // D の列位置に合わせる
                        dy:0
                    });
                }

            }
        });

        // ==========================
        // ●対応方針・対応予定時期 (H6:H10)
        // ==========================

        // 固定文ベース
        let policyText = "●対応方針\n";
        policyText += "□ 整備班で対応予定\n";
        policyText += "□ 修繕・修繕工事で対応予定\n";
        policyText += "□ 施設改良工事で対応予定\n";
        policyText += "□ 精密点検予定\n";
        policyText += "□ 撤去予定\n";

        // チェックボックス → Excel アイコン描画設定
        // dx,dy はあなたの Excel レイアウトに合わせて自由に調整可能
        const policyChecks = [
            { id: "plan_maintenance", label: "整備班で対応予定", dx: 0,  dy: 0   },
            { id: "plan_repair",      label: "修繕・修繕工事で対応予定", dx: 0,  dy: 120 },
            { id: "plan_improvement", label: "施設改良工事で対応予定", dx: 0,  dy: 240 },
            { id: "plan_precision",   label: "精密点検予定", dx: 0,  dy: 360 },
            { id: "plan_removal",     label: "撤去予定", dx: 0,  dy: 480 },
            { id: "plan_other",       label: "その他", dx: 0,  dy: 600 }
        ];

        // 各チェックボックスがチェックされていたら → icon を置く
        policyChecks.forEach(chk => {
            const el = document.getElementById(chk.id);
            if (el && el.checked) {
                data.items.push({
                    type: "icon",
                    cell: "H6",
                    icon: "check.png",
                    dx: chk.dx,
                    dy: chk.dy
                });
            }
        });

        // 「その他」の詳細テキスト
        const policyOtherDetail = document.getElementById("plan_other_detail")?.value || "";
        if (document.getElementById("plan_other")?.checked && policyOtherDetail) {
            policyText += `□ その他 (${policyOtherDetail})\n`;
        }

        // --------------------------
        // ●対応予定時期
        // --------------------------
        policyText += "●対応予定時期\n";

        const month = document.getElementById("response_month")?.value || "";
        let periodText = "";
        let periodElement = null;

        // 上旬
        if (document.getElementById("period_early")?.checked) {
            periodText = "上旬頃";
            periodElement = { dx: 160, dy: 0 };
        }
        // 中旬
        else if (document.getElementById("period_mid")?.checked) {
            periodText = "中旬頃";
            periodElement = { dx: 160, dy: 120 };
        }
        // 下旬
        else if (document.getElementById("period_late")?.checked) {
            periodText = "下旬頃";
            periodElement = { dx: 160, dy: 240 };
        }

        policyText += `　${month} 月 ${periodText}\n`;

        // 予定時期のラジオに応じて check.png を配置
        if (periodElement) {
            data.items.push({
                type: "icon",
                cell: "H6",
                icon: "check.png",
                dx: periodElement.dx,
                dy: periodElement.dy
            });
        }

        // 最後に Excel にテキストを挿入
        data.items.push({
            type: "text",
            name: "policy_text",
            value: policyText,
            cell: "H6",
            text: policyText
        });



        // ==========================
        // □本格的な使用禁止措置 (H11)
        // ==========================
        const prohibitedDate=document.getElementById("prohibited_date")?.value;
        const prohibitedStatus=document.querySelector('input[name="prohibited_status"]:checked')?.value || "";
        if(prohibitedDate && prohibitedStatus){
            const d=new Date(prohibitedDate);
            const txt=`□本格的な使用禁止措置\n　${d.getMonth()+1}月 ${d.getDate()}日 上・中・下　旬　頃`;
            data.items.push({
                type:"text",
                name:"prohibited_action",
                value:txt,
                cell:"H11",
                text:txt
            });
            data.items.push({
                type:"icon",
                cell:"H11",
                icon:"check.png",
                dx:0,
                dy:0
            });
        }


        // ==========================
        // ●備考 (F12:G15)
        // ==========================
        const remarks=document.getElementById("remarks")?.value.trim();
        if(remarks){
            data.items.push({
                type:"text",
                name:"remarks",
                value:"●備考\r\n"+remarks,
                cell:"H12",
                text:"●備考\r\n"+remarks
            });
        }

        // ==========================
        // Flask API へ送信
        // ==========================
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
            downloadBtn.disabled = false;
        }
    });
});