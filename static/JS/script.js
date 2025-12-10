// ============================
// Excel ダウンロード処理
// ============================

const downloadBtn = document.getElementById("downloadExcelBtn");
if (downloadBtn) {
    downloadBtn.type = "button"; // ★これが最重要っぽい
}

downloadBtn.addEventListener("click", async function (e) {
    e.preventDefault();
    e.stopPropagation();

    const btn = this;   // ★thisを固定
    btn.disabled = true;
    console.log("💾 Excelダウンロード処理開始");

    // ============================
    // 判定結果をHTMLから収集
    // ============================
    const inspectionResults = {};
    document.querySelectorAll("input[type='radio']:checked").forEach(el => {
        if (["B","C"].includes(el.value)) {
            inspectionResults[el.name] = el.value;
        }
    });

    console.log("=== inspectionResults ===", inspectionResults);

    // ============================
    // データ構造作成
    // ============================
    const data = {
        search_park: document.getElementById("search_park")?.value || "",
        inspection_year: document.getElementById("inspection_year")?.value || "",
        install_year_num: document.getElementById("install_year_num")?.value || "",
        inspection_sections: [
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
        items:[]
    };

        // Excel に反映する項目
    data.items.push(
            // ==============================
            // 実施措置（F6:G9）
            // ==============================
            {
                "type": "icon",
                "cell": "F6",
                "dx": 2,
                "dy": 4,
                "icon": "check.png"
            },
            {
                "type": "text",
                "cell": "F8",
                "dx": 35,
                "dy": 0,
                "text": "2"   // 吊金具交換 箇所数
            },

            // ==============================
            // 所見（F10:G12）
            // ==============================
            {
                "type": "text",
                "cell": "F10",
                "dx": 4,
                "dy": 18,
                "text": "吊金具に摩耗が見られる"
            },

            // ==============================
            // 総合結果（F13:G15）
            // ==============================
            {
                "type": "icon",
                "cell": "F14",
                "dx": 2,
                "dy": 3,
                "icon": "check.png"   // B:経過観察
            },
            {
                "type": "text",
                "cell": "F15",
                "dx": 22,
                "dy": 0,
                "text": "落下防止のため使用注意"
            },

            // ==============================
            // 対応方針（H6:H10）
            // ==============================
            {
                "type": "icon",
                "cell": "H7",
                "dx": 1,
                "dy": 3,
                "icon": "check.png"
            },
            {
                "type": "text",
                "cell": "H10",
                "dx": 14,
                "dy": 0,
                "text": "部品調達後対応"
            },

            // ==============================
            // 対応予定時期（H10）
            // ==============================
            {
                "type": "text",
                "cell": "H11",
                "dx": 8,
                "dy": 0,
                "text": "6"
            },
            {
                "type": "icon",
                "cell": "H11",
                "dx": 30,
                "dy": 3,
                "icon": "circle.png"   // 上旬
            },

            // ==============================
            // 本格的使用禁止（H11）
            // ==============================
            {
                "type": "icon",
                "cell": "H11",
                "dx": 55,
                "dy": 3,
                "icon": "check.png"   // 実施予定
            },

            // ==============================
            // 備考（H12:H15）
            // ==============================
            {
                "type": "text",
                "cell": "H12",
                "dx": 2,
                "dy": 18,
                "text": "次回点検時に重点確認"
            }
    );
    
    
    // ============================
    // B / C → Excel Items 変換
    // ============================
    data.inspection_sections.forEach(section => {
        section.items.forEach(item => {
            const result = inspectionResults[item.name];
            if (!["B","C"].includes(result)) return;

            const excelDef = item.excel?.[result];
            if (!excelDef) return;

            data.items.push({
                type: "icon",
                cell: excelDef.cell,
                dx: excelDef.dx ?? 0,
                dy: excelDef.dy ?? 0,
                icon: excelDef.icon
            });
        });
    });

    console.log("=== Excelに送信される items ===", data.items);


    // ============================
    // Flask API に POST
    // ============================

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
        btn.disabled = false;   // ✅ 必ず元に戻す
    }
});
