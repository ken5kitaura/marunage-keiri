# 自動検証の継続実行機能

## 背景

`runAutoVerification` は4.5分でタイムアウトするが、GPT-5検証が1件あたり60〜70秒かかるため、1回の実行で4件程度しか処理できない。

残りの未処理行がある場合、自動で継続実行したい。

---

## 現状の動作

1. `runAutoVerification()` が実行される
2. 対象行（CHECK/ERROR/COMPOUND で17列目が空欄）を検出
3. 1件ずつGPT-5で検証
4. 4.5分経過でタイムアウト → 処理中断
5. **残りは放置される** ← これを改善したい

---

## 目標の動作

1. `runAutoVerification()` が実行される
2. 対象行を検出して1件ずつ検証
3. 4.5分経過でタイムアウト
4. **未処理行が残っている場合、1分後に自動で再実行するトリガーを設定**
5. トリガーで再実行 → 続きから処理
6. **全件完了したらトリガーを削除**

---

## 実装内容

### 1. 継続用トリガーの作成

`runAutoVerification()` のタイムアウト時、未処理行が残っている場合：

```javascript
// 既存の継続トリガーがあれば削除
deleteContinuationTrigger_('runAutoVerification');

// 1分後に再実行するトリガーを作成
ScriptApp.newTrigger('runAutoVerification')
  .timeBased()
  .after(1 * 60 * 1000)  // 1分後
  .create();

console.log('継続トリガーを設定しました（1分後に再実行）');
```

### 2. 全件完了時のトリガー削除

全ての対象行を処理完了した場合：

```javascript
// 継続トリガーを削除
deleteContinuationTrigger_('runAutoVerification');
console.log('全件処理完了。継続トリガーを削除しました。');
```

### 3. ヘルパー関数

```javascript
/**
 * 指定した関数名の時間ベーストリガーを削除
 * @param {string} functionName
 */
function deleteContinuationTrigger_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName &&
        trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      console.log('既存の継続トリガーを削除: ' + functionName);
    }
  }
}
```

---

## 変更箇所

`Service_Verification.gs` の `runAutoVerification()` 関数内：

### タイムアウト時の処理（既存コードの後に追加）

```javascript
// タイムアウト判定後
if (isTimeout) {
  console.log('runAutoVerification: タイムアウト（' + processedCount + '/' + targetRows.length + '件処理済み）');
  
  // 未処理行が残っている場合、継続トリガーを設定
  if (processedCount < targetRows.length) {
    deleteContinuationTrigger_('runAutoVerification');
    ScriptApp.newTrigger('runAutoVerification')
      .timeBased()
      .after(1 * 60 * 1000)
      .create();
    ss.toast('タイムアウト。1分後に自動で続きを実行します。', '継続予定', 5);
  }
  break;
}
```

### 全件完了時の処理（関数の最後に追加）

```javascript
// 全件処理完了の場合、継続トリガーを削除
if (processedCount >= targetRows.length || targetRows.length === 0) {
  deleteContinuationTrigger_('runAutoVerification');
}
```

---

## 注意事項

- `processReceipts` の継続処理と同じ仕組みを使用
- 対象行の判定ロジック（17列目が空欄）は既に正しく動作しているので変更不要
- 手動で再実行した場合も、17列目が空欄の行のみが対象になるので問題なし
- 夜間バッチ（毎日1時）で実行された場合も同様に継続処理される

---

## 実装完了記録

| 日付 | 内容 | 担当 |
|------|------|------|
| | runAutoVerification に継続トリガー機能を追加 | Claude Code |
| | deleteContinuationTrigger_ ヘルパー関数を追加 | Claude Code |
