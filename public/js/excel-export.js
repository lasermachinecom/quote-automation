/**
 * excel-export.js  v2.0
 * 帳票形式Excel出力（御見積書 / 納品書 / 請求書 / 注文請書）
 * 依存: SheetJS (xlsx) CDN版
 */

function exportQuoteToExcel(quote, meta) {
        meta = meta || {};

  /* ===== 定数 ===== */
  const COMPANY = {
            name1: "サンマックスレーザーFL/GS/LTシリーズ",
            name2: "株式会社リンシュンドウ",
            zip:   "〒502-0013",
            addr:  "岐阜市中川原4-47",
            reg:   "登録番号T6200001005823",
            bank:  "十六銀行 柳ヶ瀬支店",
            acct:  "普通 0777577　カ)リンシュンドウ",
            staff: meta.staffName || "田中 健吾"
  };

  const DOC_TITLES = {
            "見積書":  "御 見 積 書",
            "納品書":  "納　品　書",
            "請求書":  "請　求　書",
            "注文請書": "注 文 請 書"
  };
        const docType  = meta.docType || "見積書";
        const docTitle = DOC_TITLES[docType] || "御 見 積 書";

  /* ===== 日付整形 ===== */
  function fmtDate(iso) {
            if (!iso) return "";
            const d = new Date(iso);
            if (isNaN(d)) return iso;
            const era = d.getFullYear() - 2018;  // 令和
          return "令和　" + era + "年　" + (d.getMonth()+1) + "月　" + d.getDate() + "日";
  }

  /* ===== 見積番号 ===== */
  function genNo() {
            const n = new Date();
            return "FL-" + n.getFullYear()
              + String(n.getMonth()+1).padStart(2,"0")
              + String(n.getDate()).padStart(2,"0")
              + "-" + String(Math.floor(Math.random()*900)+100);
  }

  const clientName    = meta.clientName    || "";
        const clientAddress = meta.clientAddress || "";
        const quoteNumber   = meta.quoteNumber   || genNo();
        const quoteDate     = fmtDate(meta.quoteDate || new Date().toISOString().slice(0,10));
        const venue         = meta.venue         || (quote.region + "（会場未定）");
        const demoDate      = meta.demoDate      ? fmtDate(meta.demoDate) : "";
        const userNotes     = meta.notes         || "";
        const validDays     = "見積書有効期限　発効後30日";

  /* ===== 金額フォーマット ===== */
  function yen(n) { return n > 0 ? n.toLocaleString() : ""; }

  /* ===== 明細行を構築 ===== */
  const bd = quote.breakdown;
        const details = [];

  // 交通費
  if (bd.transportation && bd.transportation.amount > 0) {
            details.push({
                        name: bd.transportation.label,
                        qty:  "一式",
                        unit: yen(bd.transportation.amount),
                        amt:  yen(bd.transportation.amount)
            });
  }
        // 日当
  if (bd.dayAllowance && bd.dayAllowance.amount > 0) {
            details.push({
                        name: bd.dayAllowance.label,
                        qty:  "一式",
                        unit: yen(bd.dayAllowance.amount / (quote.days * quote.people)),
                        amt:  yen(bd.dayAllowance.amount)
            });
  }
        // 宿泊費
  if (bd.accommodation && bd.accommodation.amount > 0) {
            details.push({
                        name: bd.accommodation.label,
                        qty:  "一式",
                        unit: "15,000",
                        amt:  yen(bd.accommodation.amount)
            });
  }
        // デモ実施場所（地域情報）
  details.push({
            name: "デモ実施場所　" + (quote.distanceKm > 0 ? "片道" + quote.distanceKm + "キロ" : venue),
            qty:  "一式",
            unit: "",
            amt:  "",
            info: true
  });
        // LCW設置
  if (bd.lcwSetup && bd.lcwSetup.amount > 0) {
            details.push({
                        name: bd.lcwSetup.label,
                        qty:  "一式",
                        unit: "20,000",
                        amt:  yen(bd.lcwSetup.amount)
            });
  }
        // 消耗品
  if (bd.consumables && bd.consumables.amount > 0) {
            details.push({
                        name: bd.consumables.label,
                        qty:  "一式",
                        unit: "7,500",
                        amt:  yen(bd.consumables.amount)
            });
  }
        // 窒素
  if (bd.nitrogen && bd.nitrogen.amount > 0) {
            details.push({
                        name: bd.nitrogen.label,
                        qty:  String(quote.days) + "日",
                        unit: "5,000",
                        amt:  yen(bd.nitrogen.amount)
            });
  }
        // パーテーション
  if (bd.partition && bd.partition.amount > 0) {
            details.push({
                        name: bd.partition.label,
                        qty:  "一式",
                        unit: "5,000",
                        amt:  yen(bd.partition.amount)
            });
  }
        // 保護用品
  if (bd.workTable && bd.workTable.amount > 0) {
            details.push({
                        name: bd.workTable.label,
                        qty:  "一式",
                        unit: "5,000",
                        amt:  yen(bd.workTable.amount)
            });
  }

  /* ===== セル結合リスト ===== */
  const merges = [];
        function addMerge(r1,c1,r2,c2){ merges.push({s:{r:r1,c:c1},e:{r:r2,c:c2}}); }

  /* ===== スタイル定義 ===== */
  // border helpers
  const bThin = { style:"thin", color:{rgb:"000000"} };
        const bMed  = { style:"medium", color:{rgb:"000000"} };
        const bAll  = { top:bThin, bottom:bThin, left:bThin, right:bThin };
        const bMedAll = { top:bMed, bottom:bMed, left:bMed, right:bMed };
        const bHdr  = { top:bMed, bottom:bMed, left:bMed, right:bMed };

  function cs(v, opts) {
            opts = opts || {};
            return {
                        v: v,
                        t: typeof v === "number" ? "n" : "s",
                        s: {
                                      font:      { name: opts.font || "MS Pゴシック", sz: opts.sz || 10,
                                                                       bold: opts.bold || false, color:{rgb: opts.color || "000000"} },
                                      alignment: { horizontal: opts.ha || "left", vertical: "center",
                                                                       wrapText: opts.wrap || false },
                                      fill:      opts.fill ? { patternType:"solid", fgColor:{rgb: opts.fill} } : { patternType:"none" },
                                      border:    opts.border || {}
                        }
            };
  }

  /* ===== ワークシートデータ構築 ===== */
  // 列構成: A(品名)=0, B(数量)=1, C(税別単価)=2, D(金額)=3, E(備考)=4
  // 全体は6列(A-F) A=品名幅広, B=数量, C=税別単価, D=金額, E=備考
  // 実際の配置: 0-4の5列
  // ヘッダー等は後でセル座標で直接配置する

  const ws = {};
        let ROW = 0; // 0-indexed

  function setCell(c, r, cell) {
            const addr = XLSX.utils.encode_cell({c:c, r:r});
            ws[addr] = cell;
  }
        function setV(c, r, v, opts) { setCell(c, r, cs(v, opts)); }

  /* --- ROW 0: タイトル行 --- */
  // A0: 左上住所ブロック（宛先住所）
  setV(0, ROW, clientAddress || "　", {sz:9});
        addMerge(ROW,0,ROW,1);
        // C0: タイトル
  setV(2, ROW, docTitle, {sz:20, bold:true, ha:"center", font:"MS P明朝"});
        addMerge(ROW,2,ROW,3);
        // E0: No.
  setV(4, ROW, "No.　" + quoteNumber, {sz:10, ha:"left"});

  ROW++; /* ROW=1 */
  // 住所2行目
  setV(0, ROW, "", {});
        addMerge(ROW,0,ROW,1);
        // 日付
  setV(2, ROW, quoteDate, {sz:11, ha:"center"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=2 空行 */
  ROW++; /* ROW=3 */
  // 顧客名
  setV(0, ROW, clientName, {sz:14, bold:true, font:"MS P明朝"});
        addMerge(ROW,0,ROW,1);
        // 社名1
  setV(2, ROW, COMPANY.name1, {sz:10, ha:"center"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=4 */
  // 社名2
  setV(2, ROW, COMPANY.name2, {sz:12, bold:true, ha:"center", font:"MS P明朝"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=5 */
  setV(2, ROW, COMPANY.zip + "  " + COMPANY.addr, {sz:9, ha:"center"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=6 */
  setV(2, ROW, COMPANY.reg, {sz:9, ha:"center"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=7 */
  // 受渡期日など
  setV(0, ROW, "受渡期日　ご相談", {sz:10, border:{bottom:bThin}});
        addMerge(ROW,0,ROW,1);
        setV(2, ROW, "担当者　　" + COMPANY.staff, {sz:10, ha:"center"});
        addMerge(ROW,2,ROW,4);

  ROW++; /* ROW=8 */
  setV(0, ROW, "受渡場所　貴社指定場所", {sz:10, border:{bottom:bThin}});
        addMerge(ROW,0,ROW,1);

  ROW++; /* ROW=9 */
  setV(0, ROW, "現金支払　未締め末現金払い", {sz:10, border:{bottom:bThin}});
        addMerge(ROW,0,ROW,1);

  ROW++; /* ROW=10 */
  // 税込合計ボックス
  setV(0, ROW, "税込合計", {sz:13, bold:true, border:bAll});
        addMerge(ROW,0,ROW,0);
        setV(1, ROW, "￥" + quote.total.toLocaleString(), {sz:18, bold:true, ha:"center", font:"MS P明朝", border:bAll});
        addMerge(ROW,1,ROW,2);

  ROW++; /* ROW=11 */
  setV(0, ROW, validDays, {sz:9});
        addMerge(ROW,0,ROW,4);

  ROW++; /* ROW=12: 明細ヘッダー */
  const hdrFill = "C6EFCE";  // 緑ヘッダー
  const hdrStyle = {sz:10, bold:true, ha:"center", fill:hdrFill, border:bHdr};
        setV(0, ROW, "品　　名", hdrStyle);
        setV(1, ROW, "数　量",  hdrStyle);
        setV(2, ROW, "税別単価", hdrStyle);
        setV(3, ROW, "金　額",  hdrStyle);
        setV(4, ROW, "備　考",  hdrStyle);

  ROW++; /* ROW=13: 明細開始 */
  const DETAIL_ROWS = 14; // 明細行数（空行込み）
  const detailStart = ROW;

  for (let i = 0; i < DETAIL_ROWS; i++) {
            const d = details[i];
            if (d) {
                        const nameStyle = {sz:10, border:{top:bThin,bottom:bThin,left:bThin,right:bThin}, wrap:true};
                        const numStyle  = {sz:10, ha:"right", border:bAll};
                        const infoStyle = {sz:10, border:bAll};
                        setV(0, ROW, d.name,       d.info ? infoStyle : nameStyle);
                        setV(1, ROW, d.info ? "" : d.qty,  {sz:10, ha:"right", border:bAll});
                        setV(2, ROW, d.info ? "" : d.unit, numStyle);
                        setV(3, ROW, d.info ? "" : d.amt,  numStyle);
                        setV(4, ROW, "",           {sz:10, border:bAll});
            } else {
                        // 空行
              for (let c=0; c<5; c++) setV(c, ROW, "", {border:bAll});
            }
            ROW++;
  }

  /* --- 消費税・小計行 --- */
  const taxFill = "C6EFCE";
        setV(0, ROW, "消費税10%",       {sz:10, bold:true, ha:"center", fill:taxFill, border:bAll});
        setV(1, ROW, "￥" + yen(quote.tax), {sz:10, ha:"right", fill:taxFill, border:bAll});
        setV(2, ROW, "小　計",           {sz:10, bold:true, ha:"center", fill:taxFill, border:bAll});
        setV(3, ROW, "",                 {sz:10, border:bAll});
        setV(4, ROW, yen(quote.discount > 0 ? quote.subtotalAfterDiscount : quote.subtotal), {sz:10, ha:"right", border:bAll});

  ROW++; /* 合計行 */
  setV(0, ROW, "合計金額",         {sz:11, bold:true, ha:"center", fill:taxFill, border:bMedAll});
        addMerge(ROW,0,ROW,2);
        setV(3, ROW, "￥" + quote.total.toLocaleString(), {sz:14, bold:true, ha:"right", fill:taxFill, border:bMedAll, font:"MS P明朝"});
        addMerge(ROW,3,ROW,4);

  ROW++; /* 値引き行（値引きがある場合） */
  if (quote.discount > 0) {
            setV(0, ROW, "値引き",      {sz:10, ha:"center", fill:taxFill, border:bAll});
            setV(1, ROW, "",            {sz:10, border:bAll});
            setV(2, ROW, "値引き後小計", {sz:10, ha:"center", border:bAll});
            setV(3, ROW, yen(quote.subtotalAfterDiscount), {sz:10, ha:"right", border:bAll});
            setV(4, ROW, "-" + yen(quote.discount), {sz:10, ha:"right", border:bAll});
            ROW++;
  }

  /* --- 摘要行 --- */
  const notesText = userNotes
          ? "摘要：" + userNotes
            : ("摘要：" + venue + (demoDate ? "　デモ予定日：" + demoDate : ""));
        setV(0, ROW, notesText, {sz:9, border:bAll, wrap:true});
        addMerge(ROW,0,ROW,4);
        ROW++;

  // 振込先
  setV(0, ROW, "【振込先】" + COMPANY.bank + "　" + COMPANY.acct, {sz:9, border:{bottom:bThin}});
        addMerge(ROW,0,ROW,4);
        ROW++;
        setV(0, ROW, "　", {sz:8});
        ROW++;

  /* ===== ワークシートの範囲を設定 ===== */
  ws["!ref"] = XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:ROW,c:4}});
        ws["!merges"] = merges;
        ws["!cols"] = [
              {wch:32},  // A: 品名
              {wch:10},  // B: 数量
              {wch:13},  // C: 税別単価
              {wch:13},  // D: 金額
              {wch:18}   // E: 備考
                ];
        ws["!rows"] = [];
        for (let i=0; i<ROW; i++) {
                  ws["!rows"][i] = {hpx: (i===10) ? 30 : (i===12) ? 20 : 18};
        }

  /* ===== ワークブック作成・ダウンロード ===== */
  const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, docTitle.replace(/[\/\\:*?"<>|　]/g,"").slice(0,31));

  const safeName = (clientName||"見積書").replace(/[\s\/\\:*?"<>|]/g,"_");
        const safeDate = (meta.quoteDate||"").replace(/[\/\-]/g,"");
        const filename = safeName + "_" + quoteNumber + "_" + safeDate + ".xlsx";
        XLSX.writeFile(wb, filename);
        return filename;
}

if (typeof module !== "undefined" && module.exports) {
        module.exports = { exportQuoteToExcel };
}
