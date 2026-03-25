/**
 * quote-calculator.js
 * 出張見積もり自動計算エンジン
 * pricing-config.json の価格ルールを使用して見積金額を算出する
 */

// pricing-config.json の内容をインラインで定義（ブラウザ環境用）
// ※ Node.js環境では fetch や require で読み込む
const PRICING = {
      day_allowance: { weekday: 50000, weekend: 75000 },
      accommodation: { per_night: 15000 },
      transportation: { cost_per_km: 150, from: "Gifu" },
      lcw_setup: { cost: 20000 },
      consumables: { cost: 7500 },
      nitrogen: { daily_cost: 5000 },
      equipment: { partition: 5000, work_table: 5000 },
      tax_rate: 0.10,
      major_regions: {
              "東京": { km: 350 }, "大阪": { km: 600 }, "名古屋": { km: 200 },
              "福岡": { km: 790 }, "広島": { km: 700 }, "札幌": { km: 1100 },
              "仙台": { km: 500 }, "高崎": { km: 360 }, "千葉": { km: 380 },
              "土浦": { km: 400 }, "岡崎": { km: 80 },  "四日市": { km: 90 },
              "函館": { km: 1200 }, "行田": { km: 340 }, "岐阜": { km: 0 }
      }
};

class QuoteCalculator {
      /**
       * @param {Object} params - 見積もりパラメータ
       * @param {string} params.region       - 出張先地域名（例: "高崎"）
       * @param {number} params.customKm     - 地域一覧にない場合の手入力km
       * @param {number} params.days         - 作業日数
       * @param {number} params.nights       - 宿泊数
       * @param {boolean} params.isWeekend   - 休日作業かどうか
       * @param {boolean} params.includeLcw  - LCWセットアップを含むか
       * @param {boolean} params.includeNitrogen - 窒素を使用するか
       * @param {boolean} params.includePartition - パーテーションを使用するか
       * @param {boolean} params.includeWorkTable - 作業台を使用するか
       * @param {number}  params.people      - 人数（デフォルト1名）
       * @param {number}  params.discount    - 値引き額（円）
       */
  constructor(params = {}) {
          this.region         = params.region || "";
          this.customKm       = params.customKm || 0;
          this.days           = params.days || 1;
          this.nights         = params.nights || 0;
          this.isWeekend      = params.isWeekend || false;
          this.includeLcw     = params.includeLcw !== false; // デフォルトtrue
        this.includeNitrogen    = params.includeNitrogen || false;
          this.includePartition   = params.includePartition || false;
          this.includeWorkTable   = params.includeWorkTable || false;
          this.people         = params.people || 1;
          this.discount       = params.discount || 0;
  }

  /** 片道距離(km)を取得 */
  getDistanceKm() {
          if (this.region && PRICING.major_regions[this.region]) {
                    return PRICING.major_regions[this.region].km;
          }
          return this.customKm;
  }

  /** 交通費（往復）を計算 */
  calcTransportation() {
          return this.getDistanceKm() * 2 * PRICING.transportation.cost_per_km;
  }

  /** 日当を計算 */
  calcDayAllowance() {
          const rate = this.isWeekend
            ? PRICING.day_allowance.weekend
                    : PRICING.day_allowance.weekday;
          return rate * this.days * this.people;
  }

  /** 宿泊費を計算 */
  calcAccommodation() {
          return PRICING.accommodation.per_night * this.nights * this.people;
  }

  /** オプション費用を計算 */
  calcOptions() {
          let total = 0;
          if (this.includeLcw)        total += PRICING.lcw_setup.cost;
          total += PRICING.consumables.cost; // 消耗品は常に含む
        if (this.includeNitrogen)   total += PRICING.nitrogen.daily_cost * this.days;
          if (this.includePartition)  total += PRICING.equipment.partition;
          if (this.includeWorkTable)  total += PRICING.equipment.work_table;
          return total;
  }

  /** 小計（税抜・値引き前）を計算 */
  calcSubtotal() {
          return (
                    this.calcTransportation() +
                    this.calcDayAllowance() +
                    this.calcAccommodation() +
                    this.calcOptions()
                  );
  }

  /** 消費税を計算 */
  calcTax() {
          const subtotalAfterDiscount = this.calcSubtotal() - this.discount;
          return Math.floor(subtotalAfterDiscount * PRICING.tax_rate);
  }

  /** 税込合計を計算 */
  calcTotal() {
          const subtotalAfterDiscount = this.calcSubtotal() - this.discount;
          return subtotalAfterDiscount + this.calcTax();
  }

  /**
       * 見積もり明細オブジェクトを返す
       * @returns {Object} 見積もり明細
       */
  getQuote() {
          const transportation    = this.calcTransportation();
          const dayAllowance      = this.calcDayAllowance();
          const accommodation     = this.calcAccommodation();
          const lcwSetup          = this.includeLcw ? PRICING.lcw_setup.cost : 0;
          const consumables       = PRICING.consumables.cost;
          const nitrogen          = this.includeNitrogen  ? PRICING.nitrogen.daily_cost * this.days : 0;
          const partition         = this.includePartition ? PRICING.equipment.partition : 0;
          const workTable         = this.includeWorkTable ? PRICING.equipment.work_table : 0;
          const subtotal          = this.calcSubtotal();
          const discount          = this.discount;
          const subtotalAfterDiscount = subtotal - discount;
          const tax               = this.calcTax();
          const total             = this.calcTotal();

        return {
                  // 入力パラメータ
                  region:     this.region || `カスタム(${this.customKm}km)`,
                  distanceKm: this.getDistanceKm(),
                  days:       this.days,
                  nights:     this.nights,
                  isWeekend:  this.isWeekend,
                  people:     this.people,

                  // 明細
                  breakdown: {
                              transportation:  { label: `交通費往復（岐阜-${this.region}）社用車使用＠150円`, amount: transportation },
                              dayAllowance:    { label: `日当 ${this.days}日${this.people}名分（${this.isWeekend ? "休日" : "平日"}）`, amount: dayAllowance },
                              accommodation:   { label: `宿泊費（${this.nights}泊${this.people}名）`, amount: accommodation },
                              lcwSetup:        { label: "LCWデモ機展示設置費（1.5kW相当）", amount: lcwSetup },
                              consumables:     { label: "デモ用消耗品費（ワイヤー・試材）", amount: consumables },
                              nitrogen:        { label: `窒素使用（${this.days}日×5,000円/日）`, amount: nitrogen },
                              partition:       { label: "パーテーション", amount: partition },
                              workTable:       { label: "保護メガネ・手袋・溶接台", amount: workTable },
                  },

                  // 集計
                  subtotal:               subtotal,
                  discount:               discount,
                  subtotalAfterDiscount:  subtotalAfterDiscount,
                  tax:                    tax,
                  total:                  total,

                  // フォーマット済み（表示用）
                  formatted: {
                              transportation:         transportation.toLocaleString(),
                              dayAllowance:           dayAllowance.toLocaleString(),
                              accommodation:          accommodation.toLocaleString(),
                              lcwSetup:               lcwSetup.toLocaleString(),
                              consumables:            consumables.toLocaleString(),
                              nitrogen:               nitrogen.toLocaleString(),
                              partition:              partition.toLocaleString(),
                              workTable:              workTable.toLocaleString(),
                              subtotal:               subtotal.toLocaleString(),
                              discount:               discount.toLocaleString(),
                              subtotalAfterDiscount:  subtotalAfterDiscount.toLocaleString(),
                              tax:                    tax.toLocaleString(),
                              total:                  total.toLocaleString(),
                  }
        };
  }
}

// ブラウザ・Node.js 両対応でエクスポート
if (typeof module !== "undefined" && module.exports) {
      module.exports = { QuoteCalculator, PRICING };
}
