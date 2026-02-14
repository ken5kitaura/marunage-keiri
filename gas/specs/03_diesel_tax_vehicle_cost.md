# 軽油税による車両費自動判定

## 背景と課題

### 現状の問題

1. **Costcoガスステーションの問題**
   - `STORE_ACCOUNT_MAP`で「Costco → 消耗品費」と定義されている
   - ガスステーションで軽油を入れても「消耗品費」になってしまう
   - 本来は「車両費」であるべき

2. **軽油税の特性**
   - 軽油税がある = ガソリンスタンドで軽油を給油した
   - これは100%車両費であり、例外がない
   - しかし現状は店名マッピングが優先されてしまう

### ゴール

**軽油税があるレシートは、店名に関係なく「車両費」にする**

ただし、不課税列には入湯税・宿泊税も含まれるため、単純に「不課税 > 0 → 車両費」とすると旅館が車両費になってしまう。店名がガソリンスタンド系かどうかのチェックも必要。

---

## 設計

### 判定ロジック

```
if (dieselTax > 0 AND 店名がガソリンスタンド系) {
  return '車両費';  // 最優先、他のルールより前
}
```

### ガソリンスタンド系の判定

既存の `isGasStation_()` 関数を拡張、または以下のキーワードで判定：
- ENEOS, 出光, コスモ, Shell, 石油, ガソリン, GS
- Costco, コストコ（ガスステーション併設）
- 給油, 軽油, SS

### 影響範囲

| ファイル | 修正内容 |
|----------|----------|
| Logic_Accounting.gs | `inferAccountTitleFromStore`の軽油税判定を修正（店名チェック追加） |
| Tools.gs | （既に修正済み）不課税列をdieselTaxとして渡す |
| _Main.gs | （既に修正済み）OCR結果のdieselTaxを渡す |

---

## 実装詳細

### Logic_Accounting.gs の修正

`inferAccountTitleFromStore` 関数の冒頭の軽油税判定を以下に変更：

```javascript
// ── 最優先: 軽油税がある場合は車両費 ──
// 軽油税はガソリンスタンドでの軽油給油時のみ発生するため
// ただし不課税列には入湯税等も含まれるので、店名チェックも行う
if (accountingData) {
  const dieselTax = accountingData._subtotalInfo?.dieselTax || 
                    accountingData.dieselTax || 
                    accountingData.rawNonTaxable || 0;
  if (dieselTax > 0 && isGasStationForDiesel_(storeStr)) {
    return '車両費';
  }
}
```

### 新規関数 `isGasStationForDiesel_`

```javascript
/**
 * 軽油税判定用のガソリンスタンド判定（広め）
 * 通常の isGasStation_ より判定を広くし、Costco等も含める
 * @param {string} storeStr - 小文字化済みの店舗名
 * @return {boolean}
 */
function isGasStationForDiesel_(storeStr) {
  return /eneos|出光|idemitsu|コスモ|cosmo|shell|シェル|石油|ガソリン|給油|軽油|costco|コストコ|gs|ss$|スタンド|キグナス|solato|太陽石油|宇佐美|usappy|apollo|アポロ/.test(storeStr);
}
```

---

## 実装ステップ

1. `Logic_Accounting.gs`の軽油税判定を修正（店名チェック追加）
2. `isGasStationForDiesel_`関数を追加
3. テスト：Costcoレシートで一括再判定→車両費になるか確認
4. テスト：旅館の入湯税レシートが車両費にならないか確認

---

## 成功基準

1. **Costcoガスステーションのレシート** → 軽油税があれば「車両費」
2. **旅館・ホテルのレシート** → 入湯税・宿泊税があっても「旅費交通費」のまま
3. **通常のCostco買い物** → 「消耗品費」のまま（軽油税がないので）

---

## 実装完了記録

| 日付 | 内容 | 担当 |
|------|------|------|
| 2026-02-06 | dieselTax引数の追加（_Main.gs, Tools.gs） | Claude Code ✅ |
| 2026-02-07 | 軽油税判定に店名チェック追加 | Claude Code ✅ |
| 2026-02-07 | テスト（Costco→車両費を確認） | 人間 ✅ |
