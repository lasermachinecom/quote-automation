/**
 * excel-export.js v3.0
 * テンプレート流し込み方式（ExcelJS使用）
 * addRow() 完全廃止 / 既存セルへの値書き込みのみ
 * テンプレートは template.xlsx をfetchで読み込む
 */

/**
 * メイン関数（非同期）
 * @param {Object} quote  - 計算結果オブジェクト
 * @param {Object} meta   - 帳票メタ情報
 */
async function exportQuoteToExcel(quote, meta) {
          meta = meta || {};

  /* ===== 定数 ===== */
  const COMPANY = {
              name1: "サンマックスレーザーFL/GS/LTシリーズ",
              name2: "株式会社リンシュンドウ",
              zip:   "〒502-0013",
              addr:  "岐阜市中川原4-47",
              reg:   "登録番号T6200001005823",
              bank:  "十六銀行 柳ヶ瀬支店",
              acct:  "普通 0777577 カ)リンシュンドウ",
              staff: meta.staffName || "田中 健吾"
  };

  /* ===== タイトル選択 ===== */
  const DOC_TITLES = {
              "見積書":  "御 見 積 書",
              "納品書":  "納 品 書",
              "請求書":  "請 求 書",
              "注文請書": "注 文 請 書"
  };
          const docType  = meta.docType  || "見積書";
          const docTitle = DOC_TITLES[docType] || "御 見 積 書";

  /* ===== 日付整形（令和） ===== */
  function fmtDate(iso) {
              if (!iso) return "";
              const d = new Date(iso);
              if (isNaN(d)) return iso;
              const era = d.getFullYear() - 2018;
              return "令和" + era + "年" + (d.getMonth() + 1) + "月" + d.getDate() + "日";
  }

  /* ===== 見積番号生成 ===== */
  function genNo() {
              const n = new Date();
              return "FL-" + n.getFullYear()
                + String(n.getMonth() + 1).padStart(2, "0")
                + String(n.getDate()).padStart(2, "0")
                + "-" + String(Math.floor(Math.random() * 900) + 100);
  }

  /* ===== 変数展開 ===== */
  const clientName    = meta.clientName    || "";
          const clientAddress = meta.clientAddress || "";
          const quoteNumber   = meta.quoteNumber   || genNo();
          const quoteDate     = fmtDate(meta.quoteDate || new Date().toISOString().slice(0, 10));
          const venue         = meta.venue || (quote.region + "（会場未定）");
          const demoDate      = meta.demoDate ? fmtDate(meta.demoDate) : "";
          const userNotes     = meta.notes || "";

  /* ===== 金額フォーマット ===== */
  function yen(n) { return (n > 0) ? n.toLocaleString() : ""; }
          function yenNum(n) { return (n > 0) ? n : 0; }

  /* ===== 明細行を構築 ===== */
  const bd = quote.breakdown;
          const details = [];

  if (bd.transportation && bd.transportation.amount > 0) {
              details.push({
                            name: bd.transportation.label,
                            qty:  "一式",
                            unit: bd.transportation.amount,
                            amt:  bd.transportation.amount
              });
  }
          if (bd.dayAllowance && bd.dayAllowance.amount > 0) {
                      details.push({
                                    name: bd.dayAllowance.label,
                                    qty:  "一式",
                                    unit: Math.round(bd.dayAllowance.amount / (quote.days * quote.people)),
                                    amt:  bd.dayAllowance.amount
                      });
          }
          if (bd.accommodation && bd.accommodation.amount > 0) {
                      details.push({
                                    name: bd.accommodation.label,
                                    qty:  "一式",
                                    unit: 15000,
                                    amt:  bd.accommodation.amount
                      });
          }
          // デモ実施場所（情報行）
  details.push({
              name: "デモ実施場所 " + (quote.distanceKm > 0 ? "片道" + quote.distanceKm + "キロ" : venue),
              qty:  "",
              unit: "",
              amt:  "",
              info: true
  });
          if (bd.lcwSetup && bd.lcwSetup.amount > 0) {
                      details.push({ name: bd.lcwSetup.label, qty: "一式", unit: 20000, amt: bd.lcwSetup.amount });
          }
          if (bd.consumables && bd.consumables.amount > 0) {
                      details.push({ name: bd.consumables.label, qty: "一式", unit: 7500, amt: bd.consumables.amount });
          }
          if (bd.nitrogen && bd.nitrogen.amount > 0) {
                      details.push({ name: bd.nitrogen.label, qty: String(quote.days) + "日", unit: 5000, amt: bd.nitrogen.amount });
          }
          if (bd.partition && bd.partition.amount > 0) {
                      details.push({ name: bd.partition.label, qty: "一式", unit: 5000, amt: bd.partition.amount });
          }
          if (bd.workTable && bd.workTable.amount > 0) {
                      details.push({ name: bd.workTable.label, qty: "一式", unit: 5000, amt: bd.workTable.amount });
          }

  /* ===== ExcelJS でテンプレートを読み込む ===== */
  // テンプレートXLSXをfetchで取得（同一オリジン）
  const templateUrl = "template.xlsx";
          let workbook;
          try {
                      const resp = await fetch(templateUrl);
                      if (!resp.ok) throw new Error("template.xlsx not found: " + resp.status);
                      const arrayBuffer = await resp.arrayBuffer();
                      workbook = new ExcelJS.Workbook();
                      await workbook.xlsx.load(arrayBuffer);
          } catch (e) {
                      alert("テンプレート読み込みエラー: " + e.message);
                      return;
          }

  const sheet = workbook.getWorksheet(1);

  /* ========================================
             テンプレートのセルマップ
             （template.xlsx の固定セル番地に値を流し込む）
             ======================================== */

  // ── ヘッダー部 ──
  // B2: 帳票タイトル
  sheet.getCell("B2").value = docTitle;

  // B4: 宛先顧客名
  sheet.getCell("B4").value = clientName ? clientName + " 御中" : "";

  // B5: 顧客住所
  sheet.getCell("B5").value = clientAddress;

  // F2: 見積番号
  sheet.getCell("F2").value = "No. " + quoteNumber;

  // F3: 日付
  sheet.getCell("F3").value = quoteDate;

  // ── 発行者情報 ──
  // E5: 会社名1
  sheet.getCell("E5").value = COMPANY.name1;
          // E6: 会社名2
  sheet.getCell("E6").value = COMPANY.name2;
          // E7: 住所
  sheet.getCell("E7").value = COMPANY.zip + " " + COMPANY.addr;
          // E8: 登録番号
  sheet.getCell("E8").value = COMPANY.reg;
          // F9: 担当者名
  sheet.getCell("F9").value = COMPANY.staff;

  // ── 金額サマリ ──
  // C11: 税込合計（数値）
  sheet.getCell("C11").value = quote.total;

  // ── 明細（14行目〜27行目、最大14行） ──
  const DETAIL_START_ROW = 14;
          const DETAIL_MAX_ROWS  = 14;

  for (let i = 0; i < DETAIL_MAX_ROWS; i++) {
              const row = DETAIL_START_ROW + i;
              const d   = details[i];
              if (d) {
                            sheet.getCell("B" + row).value = d.name;
                            if (!d.info) {
                                            sheet.getCell("D" + row).value = d.qty;
                                            sheet.getCell("E" + row).value = (typeof d.unit === "number" && d.unit > 0) ? d.unit : "";
                                            sheet.getCell("F" + row).value = (typeof d.amt  === "number" && d.amt  > 0) ? d.amt  : "";
                            }
              } else {
                            // 空行はクリア（テンプレートの既存値を消す）
                sheet.getCell("B" + row).value = "";
                            sheet.getCell("D" + row).value = "";
                            sheet.getCell("E" + row).value = "";
                            sheet.getCell("F" + row).value = "";
              }
  }

  // ── 集計行 ──
  // F29: 小計（税抜）
  sheet.getCell("F29").value = quote.discount > 0
            ? quote.subtotalAfterDiscount
              : quote.subtotal;

  // F30: 消費税
  sheet.getCell("F30").value = quote.tax;

  // F31: 値引き（あれば）
  sheet.getCell("F31").value = quote.discount > 0 ? -quote.discount : "";

  // F32: 合計（税込）
  sheet.getCell("F32").value = quote.total;

  // ── 摘要・備考 ──
  const notesText = userNotes
            ? userNotes
              : venue + (demoDate ? "　デモ予定日：" + demoDate : "");
          sheet.getCell("B34").value = notesText;

  // ── 振込先 ──
  sheet.getCell("B35").value = "【振込先】" + COMPANY.bank + "　" + COMPANY.acct;

  // ── ワークシート名を帳票タイトルに ──
  sheet.name = docTitle.replace(/[\s　]/g, "").slice(0, 31);

  /* ===== ダウンロード ===== */
  const buf = await workbook.xlsx.writeBuffer();
          const blob = new Blob([buf], {
                      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          });
          const url = URL.createObjectURL(blob);
          const a   = document.createElement("a");
          const safeName = (clientName || "見積書").replace(/[\s\/\\:*?"<>|]/g, "_");
          const safeDate = (meta.quoteDate || "").replace(/[\/\-]/g, "");
          a.href     = url;
          a.download = safeName + "_" + quoteNumber + "_" + safeDate + ".xlsx";
          a.click();
          URL.revokeObjectURL(url);
          return a.download;
}

if (typeof module !== "undefined" && module.exports) {
          module.exports = { exportQuoteToExcel };
}
