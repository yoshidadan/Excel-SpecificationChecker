# アーキテクチャ & 設計書

## 全体構成図

```
Excel仕様書差分チェックツール v1.0
│
├─ 【入力フェーズ】
│  ├─ ファイル選択ダイアログ
│  ├─ シート選択（複数シート対応）
│  └─ 列範囲選択＋プレビュー
│
├─ 【メモリ読み込み】
│  ├─ 仕様書①をVARIANT配列に読み込み
│  ├─ 仕様書②をVARIANT配列に読み込み
│  └─ 最大100,000行対応
│
├─ 【ペアリング処理】
│  ├─ 第1段階：スコア1.0先行確定
│  │  └─ ExecuteCompleteStringAnalysisFastScore1Only()
│  │
│  ├─ 第2段階：詳細マッチング＆スコア計算
│  │  ├─ NormalizeStringForMatchingFast() : 文字列正規化
│  │  ├─ CalculateLevenshteinDistance() : 編集距離計算
│  │  ├─ CalculateSimilarityScore() : 類似度スコア
│  │  └─ ExecuteCompleteStringAnalysisAdvanced() : 詳細判定
│  │
│  └─ 第3段階：ハンガリアンアルゴリズム実行
│     ├─ ApplyCompleteHungarianAlgorithmV4_2()
│     └─ BuildOptimalPairingResults() : 結果構築
│
├─ 【差分検出処理】
│  ├─ LCS（最長共通部分列）計算
│  │  └─ CalculateLCS()
│  │
│  ├─ 差分範囲特定
│  │  └─ CalculateDiffRangesLCS()
│  │
│  └─ ExecuteDetailedDifferenceCheckFastLCS()
│
├─ 【出力フェーズ】
│  ├─ 新規Excelブック作成
│  ├─ 「マッチング結果」シート生成
│  ├─ 条件付き書式（7色）適用
│  ├─ CreateMainResultSheetWithHighlightFast()
│  └─ SaveAsSpec3ToDownloadsPhase2()
│
└─ 【赤文字インライン処理】
   ├─ ユーザー選択ダイアログ
   ├─ ApplyRedTextHighlightToResultWorkbookLCS()
   └─ ApplyLCSDiffHighlight() : セル内文字着色
```

---

## Type定義詳細

### DiffRange
```vb
Public Type DiffRange
    startPos As Long        ' 差分開始位置
    endPos As Long          ' 差分終了位置
    diffType As String      ' "ADDED" or "DELETED"
End Type
```

### SpecSheetInfo
```vb
Public Type SpecSheetInfo
    FilePath As String              ' ファイルパス
    Workbook As Workbook            ' Workbookオブジェクト
    Worksheet As Worksheet          ' Worksheetオブジェクト
    TargetColumn As Range           ' チェック対象列
    DataArray As Variant            ' データ配列(1次元化)
    ActualRows As Long              ' 実データ行数
End Type
```

### AdvancedMatchResult
```vb
Public Type AdvancedMatchResult
    Score As Double                 ' スコア(0.0～1.0)
    matchType As String             ' マッチタイプ(7種類)
    Direction As String             ' 包含方向
    Details As String               ' 詳細情報
    EditDistance As Long            ' Levenshtein距離
    InclusionRate As Double         ' 包含率
    CommonWords As String           ' 共通部分
    DifferencePoints As String      ' 差異ポイント
End Type
```

### OptimalPairingInfo
```vb
Public Type OptimalPairingInfo
    spec1Row As Long                ' 仕様書①の行番号
    Spec2Row As Long                ' 仕様書②の行番号
    matchResult As AdvancedMatchResult  ' マッチ結果
    IsPaired As Boolean             ' ペアリング済みフラグ
    PairingRank As Long             ' ペアリング順序
    AlternativeCandidates As String ' 代替候補
    DifferenceDetailString As String  ' 差分詳細
    DifferenceData As String        ' 差分データ(エンコード済)
End Type
```

---

## アルゴリズム詳細

### 1. ハンガリアンアルゴリズム（最適割当問題）

**目的**: N×M行列のスコアから、最大スコアの1対1マッピングを検出

**処理フロー**:
```
入力: similarityMatrix(rows1, rows2)  ← 類似度スコア行列
      ↓
[ステップ1] スコア1.0の行は先行確定
      ↓
[ステップ2] 残りの候補ペアをスコア順にソート
      ↓
[ステップ3] 貪欲法で最大スコアペアを順次割当
      ↓
[ステップ4] 重複なしでペアリング完成
      ↓
出力: assignments(rows1)  ← 各①行に対する②行の対応番号
```

**計算量**: O(N² × M)

---

### 2. LCS（最長共通部分列）

**目的**: 2つの文字列の差異部分を特定

**動的計画法で実装**:
```
dp[i][j] = 文字列1[0:i]と文字列2[0:j]の最長共通部分列長

遷移式:
  - str1[i] = str2[j] → dp[i][j] = dp[i-1][j-1] + 1
  - else            → dp[i][j] = max(dp[i-1][j], dp[i][j-1])

時間計算量: O(len1 × len2)
空間計算量: O(len1 × len2)  ※最適化で O(len2)も可能
```

**差分検出ロジック**:
```
1. LCS を計算して最長共通部分列を抽出
2. LCS に含まれない文字を「ADDED」として抽出
3. 位置情報（startPos, endPos）をエンコード
4. セル内でその位置のみ赤太字で表示
```

---

### 3. Levenshtein距離（編集距離）

**目的**: 2つの文字列の違いを定量化

**定義**:
```
文字列AをBに変換するのに必要な最小編集操作数
 - 挿入: 1操作
 - 削除: 1操作
 - 置換: 1操作
```

**動的計画法実装**:
```
dp[i][j] = str1[0:i]をstr2[0:j]に変換する最小距離

遷移式:
  - i = 0: dp[0][j] = j
  - j = 0: dp[i][0] = i
  - else:
      if str1[i-1] = str2[j-1]:
        dp[i][j] = dp[i-1][j-1]
      else:
        dp[i][j] = 1 + min(
          dp[i-1][j],      # 削除
          dp[i][j-1],      # 挿入
          dp[i-1][j-1]     # 置換
        )

類似度スコア = 1 - (距離 / max(len1, len2))
```

**時間計算量**: O(len1 × len2)
**空間計算量**: O(len2)  ※最適化版

---

## 定数定義（最適化パラメータ）

```vb
Private Const MAX_ROWS As Long = 100000
    → 最大処理行数。超える場合は警告＆制限

Private Const PERFECT_MATCH_THRESHOLD As Double = 0.99
    → 完全一致とみなすスコア閾値

Private Const INCLUSION_MATCH_THRESHOLD As Double = 0.75
    → 包含一致を検出するスコア下限

Private Const SIMILARITY_MATCH_THRESHOLD As Double = 0.6
    → 高類似度とみなすスコア下限

Private Const WEAK_SIMILARITY_THRESHOLD As Double = 0.4
    → 弱類似度とみなすスコア下限

Private Const MINIMUM_MATCH_THRESHOLD As Double = 0.2
    → ペアリング候補とする最小スコア

Private Const HUNGARIAN_MAX_SIZE As Long = 5000
    → ハンガリアンアルゴリズム最大サイズ

Private Const STATUS_UPDATE_INTERVAL As Long = 500
    → ステータスバー更新間隔（行数）
```

---

## パフォーマンス最適化

### 1. メモリ最適化
- **配列読み込み**: セルの参照ではなく一括VARIANT配列で読み込み
- **スクリーン更新OFF**: `Application.ScreenUpdating = False`
- **計算モード手動**: `Application.Calculation = xlCalculationManual`

### 2. 処理速度最適化
- **スコア1.0先行確定**: 完全一致は高速な比較で先に処理
- **貪欲法ソート**: ハンガリアンアルゴリズム前にスコアでソート
- **ステータスバー更新**: 500行ごと更新で画面描画負荷を軽減

### 3. エラーハンドリング
- **On Error Resume Next**: 個別セルの読み込みエラーを許容
- **型変換**: CStr()で安全に文字列化

---

## クラス図（Type関係図）

```
SpecSheetInfo
├─ FilePath: String
├─ Workbook: Workbook
├─ Worksheet: Worksheet
├─ TargetColumn: Range
├─ DataArray: Variant
└─ ActualRows: Long

OptimalPairingInfo (配列)
├─ spec1Row: Long
├─ Spec2Row: Long
├─ matchResult: AdvancedMatchResult  ← ここで使用
├─ IsPaired: Boolean
├─ PairingRank: Long
├─ AlternativeCandidates: String
├─ DifferenceDetailString: String
└─ DifferenceData: String

AdvancedMatchResult
├─ Score: Double
├─ matchType: String
├─ Direction: String
├─ Details: String
├─ EditDistance: Long
├─ InclusionRate: Double
├─ CommonWords: String
└─ DifferencePoints: String

DiffRange (配列内で使用)
├─ startPos: Long
├─ endPos: Long
└─ diffType: String
```

---

## エラーハンドリング戦略

### メインルーチン
```vb
Sub Excel差分チェックツール()
    On Error GoTo ErrorHandler
    ' 処理...
    Exit Sub
ErrorHandler:
    MsgBox "Phase2処理エラー: " & Err.Description
    GoTo CleanExit
End Sub
```

### ファイル読み込み
```vb
On Error Resume Next
    データ読み込み処理
On Error GoTo 0
```

### セル単位アクセス
```vb
On Error Resume Next
    cellValue = CStr(array(i, 1))
On Error GoTo 0
```

---

## 今後の拡張性

### 1. Command Line Tool化（Python）
```python
python spec_checker.py spec1.xlsx spec2.xlsx
```

### 2. Web UI（Streamlit）
```
Streamlit上でドラッグ&ドロップ対応
```

### 3. 差分レポート自動生成
```
PDF出力で経営層向け報告書作成
```

### 4. 多言語対応
```
英語・中国語・日本語の自動翻訳比較
```

---

**作成日**: 2025年10月26日
**バージョン**: v1.0
