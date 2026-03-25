/**
 * excel-export.js
 * 出張見積もり自動化システム - Excel出力モジュール
 *
 * 依存: SheetJS (xlsx) CDN版
 * <script src="https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js"></script>
 *
 * 使い方:
 *   const quote = new QuoteCalculator({...}).getQuote();
 *   exportQuoteToExcel(quote, { clientName: "株式会社〇〇 様", quoteNumber: "FL2601001", ... });
 */

/**
 * 見積書をExcelファイルとして出力する
 * @param {Object} quote       - QuoteCalculator.getQuote() の戻り値
 * @param {Object} meta        - 見積書メタ情報
 * @param {string} meta.clientName    - 顧客名（例: "大塚刷毛製造株式会社 様"）
 * @param {string} meta.clientAddress - 顧客住所
 * @param {string} meta.quoteNumber   - 見積番号（例: "FL2601001"）
 * @param {string} meta.quoteDate     - 見積日（例: "2026/1/15"）
 * @param {string} meta.demoDate      - デモ予定日（例: "2026/2/10"）
 * @param {string} meta.venue         - デモ会場（例: "群馬県高崎市内（会場未定）"）
 * @param {string} meta.notes         - 摘要・注意事項
 */
function exportQuoteToExcel(quote, meta = {}) {
      // ---- 定数 ----
  const COMPANY = {
          name:           "株式会社リンシュンドウ / サンマックスレーザー",
          address:        "〒502-0013 岐阜市中川原サンマックスビル",
          email:          "lasermachine.com@gmail.com",
          bank:           "十六銀行 柳ヶ瀬支店",
          account_type:   "普通",
          account_number: "0777577",
          account_name:   "カ)リンシュンドウ",
          invoice_number: "T6200001005823",
  };

  const clientName    = meta.clientName    || "";
      const clientAddress = meta.clientAddress || "";
      const quoteNumber   = meta.quoteNumber   || "FL" + Date.now().toString().slice(-7);
      const quoteDate     = meta.quoteDate     || new Date().toLocaleDateString("ja-JP");
      const demoDate      = meta.demoDate      || "";
      const venue         = meta.venue         || `${quote.region}（会場未定）`;
      const notes         = meta.notes         || "デモ機の搬入・回収時は男性のお手伝いをお願い致します。\n電源御社手配をお願い致します。\n※交通費等は自動計算にてお預かりいたします。";
      const validUntil    = meta.validUntil    || "";

  // ---- ワークシートデータ（行の配列） ----
  // セルは { v: 値, t: 型, s: スタイル } で表す（スタイルはxlsxの書き込みのみ対応）
  const rows = [
          // 1行目: タイトル
          ["御見積書", "", "", "", "見積番号", quoteNumber],
          // 2行目: 顧客名
          [clientName, "", "", "", "見積日", quoteDate],
          // 3行目: 顧客住所
          [clientAddress, "", "", "", "", ""],
          // 4行目: 件名
          [`${venue} デモ機操作・講習見積書（到着IC：${quote.region}IC）`, "", "", "", "支払条件", "末締め末現金払い"],
          // 5行目: 有効期限
          ["", "", "", "", "見積有効期限", "発行後30日"],
          // 6行目: デモ予定日
          [demoDate ? `デモ予定日：${demoDate}` : "", "", "", "", "", validUntil],
          // 7行目: 会場
          [venue, "", "", "", "", ""],
          // 8行目: 空行
          [],
          // 9行目: ヘッダー
          ["商品名", "", "単価（円）", "数量", "金額（円）", "計算方法・備考"],
        ];

  // ---- 明細行 ----
  const bd = quote.breakdown;
      const detailRows = [
              [
                        `交通費往復（岐阜-${quote.region}）社用車使用＠150`,
                        "",
                        `${(quote.distanceKm * 2 * 150).toLocaleString()}`,
                        "1式",
                        bd.transportation.amount.toLocaleString(),
                        `距離${quote.distanceKm}km×往復÷社用車利用`
                      ],
              [
                        `日当 ${quote.days}日${quote.people}名分（${quote.isWeekend ? "休日" : "平日"}）`,
                        "",
                        (quote.isWeekend ? "75,000" : "50,000"),
                        `${quote.days}日×${quote.people}名`,
                        bd.dayAllowance.amount.toLocaleString(),
                        "平日／人数×日数×日当たり"
                      ],
            ];

  if (quote.nights > 0) {
          detailRows.push([
                    `宿泊費（${quote.nights}泊${quote.people}名）`,
                    "",
                    "15,000",
                    `${quote.nights}泊×${quote.people}名`,
                    bd.accommodation.amount.toLocaleString(),
                    "人数×宿泊数×宿泊料（1泊）"
                  ]);
  }

  if (bd.lcwSetup.amount > 0) {
          detailRows.push([
                    "LCWデモ機展示設置費（1.5kW相当）",
                    "", "20,000", "1式",
                    bd.lcwSetup.amount.toLocaleString(), "設置費"
                  ]);
  }

  detailRows.push([
          "デモ用消耗品費", "", "7,500", "1式",
          bd.consumables.amount.toLocaleString(), "ワイヤー・試材"
        ]);

  if (bd.nitrogen.amount > 0) {
          detailRows.push([
                    `窒素使用（${quote.days}日×5,000円/日）`,
                    "", "5,000", `${quote.days}日`,
                    bd.nitrogen.amount.toLocaleString(), "窒素使用（1日×5000円/日）"
                  ]);
  }

  if (bd.partition.amount > 0) {
          detailRows.push([
                    "パーテーション", "", "5,000", "1式",
                    bd.partition.amount.toLocaleString(), "展示会やデモ会場用の簡易遮光パネル"
                  ]);
  }

  if (bd.workTable.amount > 0) {
          detailRows.push([
                    "保護メガネ・手袋・溶接台", "", "5,000", "1式",
                    bd.workTable.amount.toLocaleString(), "デモ時に使用する専用作業台"
                  ]);
  }

  // 明細をrowsに追加
  detailRows.forEach(r => rows.push(r));

  // ---- 集計行 ----
  rows.push([]); // 空行
  if (quote.discount > 0) {
          rows.push(["", "", "", "小計", quote.formatted.subtotal, ""]);
          rows.push(["", "", "", "値引き", `-${quote.formatted.discount}`, ""]);
          rows.push(["", "", "", "値引き後小計", quote.formatted.subtotalAfterDiscount, ""]);
  } else {
          rows.push(["", "", "", "小計", quote.formatted.subtotal, ""]);
  }
      rows.push(["", "", "", "消費税(10%)", quote.formatted.tax, ""]);
      rows.push(["", "", "", "税込合計", quote.formatted.total, ""]);

  // ---- 摘要 ----
  rows.push([]);
      rows.push(["摘要（注意事項）", "", "", "", "", ""]);
      rows.push([notes, "", "", "", "", ""]);

  // ---- 振込先・発行元 ----
  rows.push([]);
      rows.push(["振込先", COMPANY.bank, "", "登録番号", COMPANY.invoice_number, ""]);
      rows.push(["", `${COMPANY.account_type} ${COMPANY.account_number}`, "", "", "", ""]);
      rows.push(["", COMPANY.account_name, "", "", "", ""]);
      rows.push([]);
      rows.push(["発行元", COMPANY.name, "", "", "", ""]);
      rows.push(["所在地", COMPANY.address, "", "", "", ""]);
      rows.push(["連絡先", `メール:${COMPANY.email}`, "", "", "", ""]);

  // ---- SheetJS でワークシート作成 ----
  const ws = XLSX.utils.aoa_to_sheet(rows);

  // 列幅の設定
  ws["!cols"] = [
      { wch: 45 }, // A: 商品名
      { wch: 10 }, // B
      { wch: 14 }, // C: 単価
      { wch: 12 }, // D: 数量
      { wch: 14 }, // E: 金額
      { wch: 30 }, // F: 備考
        ];

  // ---- ワークブック作成・ダウンロード ----
  const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "御見積書");

  // ファイル名: 顧客名_見積番号_日付.xlsx
  const safeClient = (clientName || "見積書").replace(/[\s\/\\:*?"<>|]/g, "_");
      const safeDate   = quoteDate.replace(/\//g, "");
      const filename   = `${safeClient}_${quoteNumber}_${safeDate}.xlsx`;

  XLSX.writeFile(wb, filename);

  console.log(`[excel-export] ダウンロード完了: ${filename}`);
      return filename;
}

// ブラウザ・Node.js 両対応
if (typeof module !== "undefined" && module.exports) {
      module.exports = { exportQuoteToExcel };
}
