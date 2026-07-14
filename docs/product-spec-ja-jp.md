# CommonModules ライブラリ 仕様書

## 位置づけ

この文書は、`xls-common-modules\CommonModules\src\CommonModules` にある共通モジュールのライブラリ仕様である。

詳細な API リファレンス:

- 入口: `xls-common-modules\CommonModules\doc\html\index.html`
- クラス/標準モジュール別のメンバー一覧: `xls-common-modules\CommonModules\doc\html\annotated.html`、`namespace_class.html`、`namespace_standard.html`

## 全体像

CommonModules は Excel VBA ツール群に共通する基盤ライブラリである。通常のツール実装では、入口処理で `InitializeCommonService` を呼び出し、`WbSrv`、`WsSrv`、`New_RangeBounds`、`New_ObjectList`、`New_ObjectSet`、`New_ObjectDictionary`、`WorksheetVirtualTable` などを利用する。公開 UDF は `Fx_*.bas` に置き、非 UDF の公開 API は `Option Private Module` を持つ標準モジュールに置く。

| 分類 | 主なモジュール | 主な責務 | 代表的な利用場面 |
| --- | --- | --- | --- |
| 一般 | `Lib_Common`、`Lib_CommonConstructor`、`Fx_Common`、`WorkbookService`、`WorksheetService`、`WorksheetRangeBounds`、`WorksheetRangeBoundsEnumerator`、`WorksheetVirtualTable`、`WorksheetVirtualTableEnumerator`、`ApplicationScreenUpdateManager`、`CommonRunStateManager`、`DebugInformation`、`ProgressStatus` | Excel ブック/シート/セル範囲操作、範囲値オブジェクト、仮想表、共通ファクトリ、共通サービス初期化、文字列/配列/Excel アドレス補助、公開 UDF、進捗/診断 | 通常の業務マクロ、シート読み書き、データ変換、範囲列挙、ワークシート数式、GUI 実行状態管理 |
| コレクション・列挙 | `ObjectList`、`ObjectSet`、`ObjectDictionary`、`Counter`、`CounterSet`、`Enumerator`、`ArrayObject`、`IEnumerator`、`IElementTypeProvider`、`IComparable`、`IEquatable`、`IDuplicateCheckable`、`IStringable` | 型付きリスト/集合/辞書、カウンター、列挙子、内部配列、要素型自己申告、比較・重複・ソート・文字列化補助 | データ集合の保持、重複管理、キー付き参照、列挙、並べ替え、キー生成 |
| 入力画面 | `Lib_InputSheet`、`IUserInputSheet`、`UserInputSheet`、`UserInputSheetTestDouble` | 入力シート生成、見出し検索、入力値の範囲解決、テスト差し替え | 見出し付き入力シートから設定値や対象範囲を取得する処理 |
| IPv4 | `Lib_IPv4` | IPv4 アドレス、ネットワーク、マスク長、マスク値の変換と判定 | FW/ネットワーク関連ツールのアドレス正規化、ネットワーク分割/集約 |
| ファイル操作 | `Lib_FileSystem`、`FileSystemService`、`Lib_TextFile`、`TextFileService`、`TextFileEntity` | ファイルシステム、パス解決、ファイル/ディレクトリ操作、テキストファイル入出力 | 設定ファイル生成、バックアップ、ファイル一覧取得、テキスト出力 |
| テスト支援 | `Lib_UnitTest`、`UnitTestAssert`、各 `*TestDouble`、`TestDoubleBehaviorStore`、`TestDoubleVariantKeyBuilder` | Excel VBA 用ユニットテスト、assert、サービス差し替え、スタブ戻り値/呼び出し履歴/エラー注入 | CommonModules と各ツールのユニットテスト |

## 基本利用方法

### 共通サービス初期化

入口処理では `InitializeCommonService` を呼ぶ。既存サービスを本番サービスへ再生成する場合は `Force:=True` を指定する。

```vb
Call InitializeCommonService

Dim target_cell As WorksheetRangeBounds
Set target_cell = New_RangeBounds(Row:=1, Column:=1, Sheet:="INPUT")

Call WsSrv.WriteCell(target_cell, "123")
```

`InitializeCommonService` は次のように動作する。

| 対象 | 動作 |
| --- | --- |
| `Lib_Common.WbSrv` | `Force:=True` の場合は `WorkbookService` を生成し直す。`Force:=False` の場合は `Nothing` のときだけ `WorkbookService` を生成する |
| `Lib_Common.WsSrv` | `Force:=True` の場合は `WorksheetService` を生成し直す。`Force:=False` の場合は `Nothing` のときだけ `WorksheetService` を生成する |
| `Lib_FileSystem.FsSrv` | `InitializeFileSystemService` が同じブックに存在する場合だけ `Force` を渡して呼び出す。`Force:=True` の場合は `FileSystemService` を生成し直し、`Force:=False` の場合は `Nothing` のときだけ `FileSystemService` を生成する |
| `Lib_TextFile.TfSrv` | `InitializeTextFileService` が同じブックに存在する場合だけ `Force` を渡して呼び出す。`Force:=True` の場合は `TextFileService` を生成し直し、`Force:=False` の場合は `Nothing` のときだけ `TextFileService` を生成する |

`Force` は省略時 `False` とする。既にテストダブル等へ差し替え済みのサービスは、`Force:=False` では上書きしない。`Force:=True` では全対象サービスを本番サービスへ再生成する。`InitializeFileSystemService` または `InitializeTextFileService` が同じブックに存在しない場合は、`Force:=True` でも従来どおり無視する。

```vb
Set Lib_Common.WsSrv = New WorksheetServiceTestDouble
Call InitializeCommonService  ' WsSrv は差し替え済みインスタンスのまま

Call InitializeCommonService(Force:=True)  ' 全対象サービスを本番サービスへ再生成する
```

Excel ワークシート関数として実行される UDF 入口では、`InitializeCommonService` を直接呼ばず、`InitializeUdfCommonService` を呼ぶ。`InitializeCommonService` は optional service の初期化で `Application.Run` を使うため、セル再計算中には `Err.Number = 1004`、`HRESULT = 0x800A03EC`、`'Run' メソッドは失敗しました: '_Application' オブジェクト` が発生しうる。

```vb
' UDF 入口では通常 Force を省略する
Call InitializeUdfCommonService
```

`InitializeUdfCommonService` は `WbSrv` と `WsSrv` だけを対象にし、`FsSrv` と `TfSrv` は初期化しない。個別に初期化したい場合は、`InitializeWorkbookService` または `InitializeWorksheetService` を使う。これらの API も `Force:=False` では差し替え済みサービスを維持し、`Force:=True` では本番サービスへ再生成する。`DIFFSTR` などの公開 UDF は `Fx_Common.bas` に置き、`Fx_*.bas` には `Option Private Module` を置かない。公開 UDF が増えて責務が分かれる段階で、`Fx_...` を分割する。`Fx_*.bas` と `Test_*.bas` 以外の標準モジュールには `Option Private Module` を置き、Excel の数式候補に非 UDF の `Public Function` が表示されることを抑える。非 UDF の戻り値 API は、この目的だけで `Sub` + `ByRef` へ変更しない。

### 範囲操作の基本形

Excel のセルや範囲は `WorksheetRangeBounds` で表す。

```vb
Call InitializeCommonService

Dim input_cell As WorksheetRangeBounds
Set input_cell = New_RangeBounds(Row:=2, Column:=3, Sheet:="INPUT")

Dim value_text As String
value_text = WsSrv.ReadCell(input_cell)

Call WsSrv.WriteCell(input_cell.Shift(Column:=1), value_text, NumberFormat:="@")
```

### 使用範囲の列挙

```vb
Call InitializeCommonService

Dim used_range As WorksheetRangeBounds
Set used_range = WsSrv.GetUsedRangeBounds( _
        New_RangeBounds(Row:=1, Column:=1, FinishRow:=G_ROW_MAX, FinishColumn:=G_COL_MAX, Sheet:="DATA"))

used_range.EnumerationMode = G_RANGE_ENUM_MODE_ROWS

Dim row_bounds As WorksheetRangeBounds
For Each row_bounds In used_range
    Debug.Print row_bounds.ToString(CellOnly:=True)
Next row_bounds
```

`ObjectList`、`ObjectSet`、`WorksheetRangeBounds`、`WorksheetVirtualTable` は標準の `For Each` に対応する。`ObjectDictionary` の標準 `For Each` は key を列挙し、value を列挙したい場合は `Items` を使う。`For Each` は列挙開始時点のスナップショットを列挙する。標準 `For Each` の列挙中に元のコレクションや範囲の列挙設定を変更した場合、実行中のループの継続は保証しない。更新や削除を伴う列挙では、`ObjectList` / `ObjectSet` / `WorksheetRangeBounds` / `WorksheetVirtualTable` は `GetEnumerator` を使い、`ObjectDictionary` は `Keys` / `Items` の copy と明示 API を使う。

`WorksheetRangeBounds.GetRows()` と `GetColumns()` は、行単位または列単位の `WorksheetRangeBounds` を `ObjectList("WorksheetRangeBounds")` として返す。複数の範囲形状を揃える場合は `ExpandRangeBoundsToMax` を使う。

### 仮想表

```vb
Dim table_bounds As WorksheetRangeBounds
Set table_bounds = New_RangeBounds(Row:=1, Column:=1, FinishRow:=10, FinishColumn:=3, Sheet:="INPUT")

Dim vtable As WorksheetVirtualTable
Set vtable = New_WorksheetVirtualTableFromRangeBounds( _
        TableRange:=table_bounds, _
        TreatFirstRowAsHeader:=True)

Dim row_dict As ObjectDictionary
For Each row_dict In vtable
    Debug.Print row_dict.Item("Name").ToString(CellOnly:=True)
Next row_dict
```

`WorksheetVirtualTable` は、複数の `WorksheetRangeBounds` を header と相対行で対応付ける。各行は header を key、1 行分の `WorksheetRangeBounds` を value に持つ `ObjectDictionary` として返る。短い field の不足行は、元 field の列幅を維持した `RowCount = 0` の空範囲として返る。

### 重複なし集合

```vb
Dim names As ObjectSet
Set names = New ObjectSet

Call names.Add("alpha")
Call names.Add("beta")
Call names.Add("alpha", ErrorIfExists:=False)

If names.Exists("alpha") Then
    Debug.Print JoinStringSet(names, ",")
End If
```

### キー付き型付き辞書

```vb
Dim row_values As ObjectDictionary
Set row_values = New_ObjectDictionary("WorksheetRangeBounds")
row_values.CompareMode = vbTextCompare

Call row_values.Add("Name", New_RangeBounds(Row:=2, Column:=1, Sheet:="INPUT"))
Call row_values.AddOrUpdate("Value", New_RangeBounds(Row:=2, Column:=2, Sheet:="INPUT"))

If row_values.Exists("name") Then
    Debug.Print row_values.Item("name").ToString(CellOnly:=True)
End If
```

`ObjectDictionary.Item(Key)` の getter は、存在しない key を暗黙追加せずエラーにする。追加は `Add`、更新は `Update`、追加または更新は `AddOrUpdate` または `Item(Key)` への代入で行う。

### テキストファイル出力

```vb
Call InitializeCommonService
Call InitializeTextFileService

Dim entity As ITextFileEntity
Set entity = TfSrv.GetTextFileEntity("output.txt")

Call entity.OpenFile(AsWrite:=True, Force:=True)
Call entity.WriteLine("first line")
Call entity.CloseFile
```

### テストでのサービス差し替え

```vb
Dim ws_double As WorksheetServiceTestDouble
Set ws_double = New WorksheetServiceTestDouble

Set Lib_Common.WsSrv = ws_double
Call InitializeCommonService

Dim target As WorksheetRangeBounds
Set target = New_RangeBounds(Row:=1, Column:=1, Sheet:="TEST")

Dim return_value As Variant
return_value = "stubbed"
Call ws_double.Store.SetReturn("ReadCell", return_value, target, "", "", False)

Debug.Print WsSrv.ReadCell(target)
```

## 分類別仕様

### 一般

- 共通サービス
  - 主なモジュール: `Lib_Common`、`Lib_CommonConstructor`
  - 仕様概要: `WbSrv` / `WsSrv` をグローバルサービスとして保持し、通常入口では `InitializeCommonService`、UDF 入口では `InitializeUdfCommonService` で遅延初期化する。`InitializeWorkbookService` と `InitializeWorksheetService` による個別初期化も提供する。`Force:=True` では対象サービスを本番サービスへ再生成する。`New_RangeBounds`、`New_RangeBoundsFromAddress`、`New_ObjectList`、`New_ObjectSet`、`New_ObjectDictionary`、`New_WorksheetVirtualTable`、`New_WorksheetVirtualTableFromRangeBounds` を提供する。
  - 重要な制限・副作用: グローバル状態を使うため、テストでは明示的な差し替えと復元が必要。差し替え済みサービスは `Force:=False` の初期化では上書きされないが、`Force:=True` では本番サービスへ置き換わる。optional service の有無判定には VBIDE を使わない。
- Excel ブック
  - 主なモジュール: `WorkbookService`、`IWorkbookService`
  - 仕様概要: ブック存在判定、一覧取得、オープン、保存、クローズ、シート追加/削除/コピー、シート検索、VBA コンポーネント削除を扱う。
  - 重要な制限・副作用: `SaveWorkbook` は `Application.Calculation`、`DisplayAlerts`、ウィンドウ表示状態を一時変更する。`RemoveVBComponents` は VBIDE アクセス権限が必要。
- Excel ワークシート
  - 主なモジュール: `WorksheetService`、`IWorksheetService`
  - 仕様概要: セル/範囲の読み書き、検索、並べ替え、コピー、UsedRange 算出、重複削除、行挿入/削除、表示/書式/色/配置変更を扱う。`WriteCell` は型変換、数式、空値処理、表示形式指定に対応する。
  - 重要な制限・副作用: 対象セル・範囲を実際に変更する。`WriteCell` の文字列書き込みは既存の `WrapText` を保持し、書き込み失敗時はセル内容、表示形式、`WrapText` の復元を試みる。`ActivateRange` はアクティブブック/シート/セルを変える。コピー系は `CutCopyMode` やクリップボードに影響する可能性がある。
- 範囲値オブジェクト
  - 主なモジュール: `WorksheetRangeBounds`、`WorksheetRangeBoundsEnumerator`
  - 仕様概要: ブック名、シート名、開始/終了行列を保持し、セル、矩形、行全体、列全体、シート全体、空範囲を表す。範囲抽出、行/列/セル取得、`GetRows` / `GetColumns` による一覧化、変形、シフト、交差、文字列化、列挙に対応する。標準列挙の方式は `EnumerationMode` と `EnumerationDescending` で切り替える。
  - 重要な制限・副作用: 初期化は一度だけ。`Sheet` は空文字不可。複数エリア範囲は `InitializeFromAddress` でエラー。行列上限は Excel 上限に依存する。
- 仮想表
  - 主なモジュール: `WorksheetVirtualTable`、`WorksheetVirtualTableEnumerator`
  - 仕様概要: `ObjectList("WorksheetRangeBounds")` または 1 つの `WorksheetRangeBounds` から、header と相対行で対応付けた仮想表を作る。各行は header を key、対応する `WorksheetRangeBounds` を value に持つ `ObjectDictionary` として取得・列挙できる。
  - 重要な制限・副作用: `ColumnRangeList` は `WorksheetRangeBounds` の要素型契約を持つ必要がある。header 件数不一致や重複 header は初期化エラー。短い field の不足行は `RowCount = 0` の空範囲として返す。必須 header、空行除外、キー列必須、重複キー検出などの業務検証は呼び出し側が行う。
- 汎用関数
  - 主なモジュール: `Lib_Common`、`Fx_Common`
  - 仕様概要: `Lib_Common` はパス結合/分解、配列変換、文字列処理、キー生成、範囲形状の拡張、Excel アドレス生成/解析、ビット演算、Excel エラー値変換などを提供する。`Fx_Common` はワークシート数式から利用する意図のある公開 UDF を提供する。
  - 重要な制限・副作用: 一部関数は Excel オブジェクトや `WsSrv` に依存する。公開 UDF は、現在の実装でサービスを直接使わない場合でも入口で `InitializeUdfCommonService` を `Force` なしで呼ぶ。`Fx_Common.DIFFSTR` はこの対象であり、`FsSrv` / `TfSrv` は初期化しない。`DIFFSTR` の `ExtractType` 全仕様は未確認。`Fx_*.bas` と `Test_*.bas` 以外の標準モジュールは `Option Private Module` を持ち、同一 VBA プロジェクト内で使う公開 API として扱う。
- UI/状態補助
  - 主なモジュール: `ApplicationScreenUpdateManager`、`CommonRunStateManager`、`DebugInformation`、`ProgressStatus`
  - 仕様概要: Excel アプリケーション設定退避/復元、1 回の GUI 実行に紐づく `DbgInfo` / `ProgStat` の生成・破棄、デバッグ用タスクスタック、StatusBar 進捗表示を提供する。
  - 重要な制限・副作用: `ApplicationScreenUpdateManager.Restore` は退避なしではエラー。`CommonRunStateManager` は生成時または `Initialize` で `DbgInfo` / `ProgStat` を新規化し、破棄時または `Clear` で `Nothing` に戻す。`ProgressStatus` は `Application.StatusBar` を変更する。

### コレクション・列挙

- 型付き要素コレクション
  - 主なモジュール: `ObjectList`、`ObjectSet`、`ObjectDictionary`、`Enumerator`、`ArrayObject`、`IElementTypeProvider`
  - 仕様概要: 型を固定するリスト/集合/辞書、列挙子、内部配列を提供する。標準 `For Each` と明示的な `IEnumerator` による列挙に対応し、`IElementTypeProvider` による要素型自己申告、`IEquatable`、`IDuplicateCheckable`、`IComparable`、`IStringable` による比較・重複・ソート・文字列化に対応する。
  - 重要な制限・副作用: `Empty` / `Null` は要素不可。最初の非 `Nothing` 要素または `Initialize` / `New_Object...` の指定で要素型契約が固定され、異型はエラーになる。`Nothing` はオブジェクト型契約の空参照として扱う。大量要素での性能上限は未確認。
- カウンター
  - 主なモジュール: `Counter`、`CounterSet`
  - 仕様概要: 件数、ステップ、最大値を持つ単体カウンターと、名前付きカウンター集合を提供する。
  - 重要な制限・副作用: `StopWhenMax` と `MaxCount` の設定に従って進行可否が変わる。`CounterSet` は内部で `Counter` を保持する。

### 入力画面

- 入力シート
  - 主なモジュール: `Lib_InputSheet`、`IUserInputSheet`、`UserInputSheet`、`UserInputSheetTestDouble`
  - 仕様概要: `Lib_InputSheet` は `New_InputSheet` を提供する。`UserInputSheet` は入力シートの見出しを検索し、見出しに対応するセルや範囲を解決する。`IUserInputSheet` と `UserInputSheetTestDouble` により、入力シート読み取りをテストで差し替えられる。
  - 重要な制限・副作用: `UserInputSheet` は初期化時に対象ブック/シートの存在を確認する。読み取り対象は見出しの配置と `WbSrv` / `WsSrv` に依存する。

### IPv4

- アドレス変換
  - 主な API: `ConvertFromIpAddress`、`ConvertToIpAddress`
  - 仕様概要: ドットデシマル IPv4 文字列と 32bit 値を相互変換する。
  - 重要な制限: VBA の `Long` は符号付きのため、`128.0.0.0` 以上の値は負値になる。
- マスク変換
  - 主な API: `ConvertFromMaskLength`、`ConvertToMaskLength`、`InvertMaskValue`、`GetMaskValue`
  - 仕様概要: マスク長、マスク値、ビット範囲マスクを変換する。
  - 重要な制限: マスク長は 0-32。非連続マスク値はエラー。
- ネットワーク判定
  - 主な API: `IsNetwork`、`IsValidMaskValue`、`GetHostAddress`、`GetNetworkAddress`
  - 仕様概要: IP 値とマスク値からネットワーク/ホスト部を判定・抽出する。
  - 重要な制限: 不正マスク値に対する全関数の網羅挙動は未確認。
- ネットワーク変形
  - 主な API: `ExpandNetwork`、`NarrowNetwork`
  - 仕様概要: マスク長を 1 短くした親ネットワーク、または 1 長い 2 つのサブネットを返す。
  - 重要な制限: `ExpandNetwork` は `/0` でエラー。`NarrowNetwork` は実装上 `/30` から `/32` を「これ以上狭くできない」としてエラーにする。
- パース/整形
  - 主な API: `WellFormedAddress`、`ParseIpAddressAndMask`、`G_IPV4_*_RE`
  - 仕様概要: `_` を `/` に置換し、単体 IPv4 には `/32` を付ける。`IP/マスク` は CIDR 長とドットデシマルマスクを受け付ける。
  - 重要な制限: `ParseIpAddressAndMask` はマスク省略形式を受け付けないと推定される。正規表現定数の外部互換性保証範囲は未確認。

### ファイル システム アクセス

- ファイルサービス初期化
  - 主なモジュール: `Lib_FileSystem`
  - 仕様概要: `FsSrv` を保持し、`InitializeFileSystemService` で未設定時だけ `FileSystemService` を生成する。`Force:=True` では既存値に関わらず `FileSystemService` を生成し直す。
  - 重要な制限・副作用: グローバル状態のため、テストでは差し替えを意識する必要がある。
- パス・ファイルシステム
  - 主なモジュール: `FileSystemService`、`IFileSystemService`
  - 仕様概要: 相対パスの絶対化、ローカルパスの OS 環境変数展開、OS ユーザー一時フォルダー取得、一時ディレクトリ作成、ファイル/ディレクトリ一覧、存在判定、更新日時、作成、移動、コピー、削除、バックアップ作成を扱う。
  - 重要な制限・副作用: URL ではないローカルパスは `GetAbsolutePath` の絶対パス判定前に `WScript.Shell.ExpandEnvironmentStrings` で環境変数を展開する。URL 文字列は環境変数展開しない。相対パスは `WbSrv.GetThisWorkbookDirectoryPath` 基準。`GetTemporaryDirectoryPath` は末尾 `\` なしの絶対 Windows パスを返す。`CreateTemporaryDirectory` は OS ユーザー一時フォルダー直下へ実ディレクトリを作成する。実ファイルシステムを変更する。ネットワークパス、長いパス、ACL、読み取り専用属性の網羅挙動は未確認。
- テキストファイル初期化
  - 主なモジュール: `Lib_TextFile`
  - 仕様概要: `TfSrv` を保持し、`InitializeTextFileService` で未設定時だけ `TextFileService` を生成する。`Force:=True` では既存値に関わらず `TextFileService` を生成し直す。
  - 重要な制限・副作用: グローバル状態のため、テストでは差し替えを意識する必要がある。
- テキストファイル
  - 主なモジュール: `TextFileService`、`TextFileEntity`、`ITextFileService`、`ITextFileEntity`
  - 仕様概要: ファイルパスを指定して `TextFileEntity` を生成し、読み取り、上書き、追記、ロック、行単位入出力を扱う。
  - 重要な制限・副作用: 文字コード指定 API はない。VBA の `Open` / `Line Input` / `Print` の既定挙動に依存する。`WriteLine` は末尾改行を付ける。

### テスト

- テストランナー
  - 主なモジュール: `Lib_UnitTest`
  - 仕様概要: `UnitTestMain` が `Test_*.bas` 内の `Test_...` 手続きを検出し、`UNIT_TEST_SHEET` に結果を出す。
  - 重要な制限・副作用: VBIDE アクセスが必要。実行用一時 VBA モジュールを作成・削除する。結果シートを作成/更新する。
- assert
  - 主なモジュール: `UnitTestAssert`
  - 仕様概要: スカラー、数値、真偽、型、`Nothing`、`Empty`、エラー発生、配列の assert を提供する。
  - 重要な制限・副作用: assert 失敗は通常エラーではなく `IsFailed` と `ResultMessage` に保持される。最初に失敗した後の追加 assert は処理されない。
- テストダブル基盤
  - 主なモジュール: `TestDoubleBehaviorStore`、`TestDoubleCallRecord`、`TestDoubleVariantKeyBuilder`
  - 仕様概要: メソッド名と引数キーごとの戻り値、ByRef 出力、エラー、呼び出し履歴を管理する。
  - 重要な制限・副作用: 引数キーは `GetTypedValueKey` 系に依存する。`DefaultObjectKeyMode` やメソッド別キー化モードは、キー生成後に変更できないケースがある。
- サービステストダブル
  - 主なモジュール: `WorkbookServiceTestDouble`、`WorksheetServiceTestDouble`、`FileSystemServiceTestDouble`、`TextFileServiceTestDouble`、`TextFileEntityTestDouble`、`UserInputSheetTestDouble`
  - 仕様概要: 各サービスまたは入力シートのインターフェイスを実装し、`Store` からスタブ値やエラーを返し、呼び出しを記録する。
  - 重要な制限・副作用: メソッドごとに既定動作が異なる。分類は下記の「サービステストダブルの既定動作分類」に従う。

#### サービステストダブルの既定動作分類

| 分類 | 契約 | 呼び出し履歴 |
| --- | --- | --- |
| `strict stub` | 戻り値、ByRef 出力、または空結果も含めて `Store` へ明示登録する。未登録の場合はテスト不備としてエラーにする。 | 正常に成立した呼び出しだけ記録する。未登録エラーと `Store.SetError` による明示エラーは記録しない。 |
| `safe default` | 未登録時に保守的な既定値を返す。存在確認系は原則 `False`、ThisWorkbook 系はテスト用の既定値を返す。 | 正常に成立した呼び出しだけ記録する。 |
| `spy only` | 戻り値を持たない操作を実行せず、呼び出しだけを記録する。失敗を検証したい場合は原則 `Store.SetError` を使う。 | 検証したい payload がある場合はそれを記録値にし、特にない場合だけ `True` を記録値にする。 |
| `lightweight fake` | 外部リソースには触らず、テストに必要な軽い状態遷移だけを実装する。 | 正常に成立した呼び出しだけ記録する。状態違反エラーは記録しない。 |

`WorkbookServiceTestDouble` の分類:

| 分類 | メソッド |
| --- | --- |
| `safe default` | `GetThisWorkbookName`, `GetThisWorkbookDirectoryPath`, `ExistsWorkbook`, `GetAllWorkbook`, `GetOtherWorkbook`, `ExistsWorksheet`, `IsSaved`, `HasPath` |
| `strict stub` | `GetAllWorksheet`, `GetOtherWorksheet`, `OpenWorkbook`, `SaveWorkbook`, `Find`, `AddWorksheet`, `CopyWorksheet` |
| `spy only` | `CloseWorkbook`, `RemoveWorksheet`, `ActivateWorksheet`, `RemoveVBComponents` |

`WorksheetServiceTestDouble` の分類:

| 分類 | メソッド |
| --- | --- |
| `strict stub` | `Find`, `ReadCell`, `ReadRange`, `EvaluateFormula`, `XLookup`, `IsEmptyCell`, `HasFormula`, `GetUsedRangeBounds` |
| `spy only` | `Sort`, `ActivateRange`, `WriteCell`, `WriteRange`, `WriteArrayFormula`, `CopyCell`, `CopyRange`, `ClearRange`, `RemoveDuplicates`, `InsertRows`, `DeleteRows`, `SetAllDataVisible`, `SetSheetOutlineLevel`, `SetSheetTabColor`, `SetRangeColor`, `SetWrapText`, `SetShrinkToFit`, `SetAlignment` |

`FileSystemServiceTestDouble` の分類:

| 分類 | メソッド |
| --- | --- |
| `safe default` | `PathExists`, `IsFile`, `IsDirectory` |
| `strict stub` | `GetAbsolutePath`, `GetTemporaryDirectoryPath`, `CreateTemporaryDirectory`, `GetFileList`, `GetDirectoryList`, `GetLastModified`, `CreateDirectory`, `RemoveDirectory`, `RemoveFile`, `GetNewestFile`, `CreateBackupFile` |
| `spy only` | `MoveDirectory`, `CopyDirectory`, `MoveFile`, `CopyFile` |

`TextFileServiceTestDouble` / `TextFileEntityTestDouble` の分類:

| クラス | 分類 | メソッド |
| --- | --- | --- |
| `TextFileServiceTestDouble` | `lightweight fake` | `GetTextFileEntity` |
| `TextFileEntityTestDouble` | `lightweight fake` | `OpenFile`, `CloseFile` |
| `TextFileEntityTestDouble` | `strict stub` | `ReadLine` |
| `TextFileEntityTestDouble` | `spy only` | `Initialize`, `WriteLine` |

`TextFileEntityTestDouble` の `FilePath`、`AsRead`、`AsWrite`、`AsAppend`、`GetReadLock`、`GetWriteLock`、`IsOpen`、`IsEndOfFile` は fake 状態を問い合わせる読み取りプロパティであり、上表のメソッド分類の対象外とする。

## 重要な制限事項

### Excel 状態への副作用

次の API は Excel アプリケーションまたはブック状態を変更する。

- `ApplicationScreenUpdateManager`
  - 副作用: `ScreenUpdating`、`EnableEvents`、`DisplayAlerts`、`Calculation` を変更・復元する
- `WorkbookService.OpenWorkbook` / `SaveWorkbook` / `CloseWorkbook`
  - 副作用: ブック、ウィンドウ表示、アクティブブック、保存先ファイルに影響する
- `WorksheetService` の書き込み/コピー/整形 API
  - 副作用: セル内容、数式、表示形式、色、行、重複行、フィルタ、アウトライン、タブ色に影響する
- `WorksheetService.ActivateRange`
  - 副作用: アクティブブック、アクティブシート、選択セルを変更する
- `ProgressStatus`
  - 副作用: `Application.StatusBar` を変更する
- `SetClipboard` / `GetClipboard` / `PasteFormulas`
  - 副作用: クリップボードまたは貼り付け先に影響する
- `Lib_UnitTest.UnitTestMain`
  - 副作用: `UNIT_TEST_SHEET`、一時 VBA モジュール、画面更新状態に影響する

状態復元を試みる実装はあるが、Excel 由来エラー、保護状態、外部要因を含めた完全な復元保証は未確認である。

### `WorksheetRangeBounds` の制約

- 初期化
  - 制約: 一度だけ可能。再初期化はエラー
- シート名
  - 制約: 空文字不可
- ブック名
  - 制約: 空文字なら `WbSrv.GetThisWorkbookName` 由来の既定値を使う
- 行列
  - 制約: 開始行/列は 1 以上で Excel 上限内。終了行/列に負数は不可
- 空範囲
  - 制約: 開始が終了を超える場合、エラーではなく空範囲として扱う
- 複数エリア
  - 制約: `InitializeFromAddress` では複数範囲をエラーにする
- 比較
  - 制約: ブック名/シート名の大小文字違いは同一扱い。`EnumerationMode` と `EnumerationDescending` は同一性に影響しない
- 標準列挙
  - 制約: `EnumerationMode` は `G_RANGE_ENUM_MODE_ROWS`、`G_RANGE_ENUM_MODE_COLUMNS`、`G_RANGE_ENUM_MODE_CELLS_HORIZONTAL`、`G_RANGE_ENUM_MODE_CELLS_VERTICAL` のいずれか。`For Each` は列挙開始時点のスナップショットを列挙するが、列挙中に列挙設定を変更した場合、実行中のループの継続は保証しない

### `ObjectList` / `ObjectSet` / `ObjectDictionary` の制約

- 型
  - 制約: 最初の非 `Nothing` 要素、または明示初期化で要素型契約が固定される。以降の異型はエラーになる
- 要素型自己申告
  - 制約: オブジェクトが `IElementTypeProvider` を実装する場合、`ElementTypeKey` を要素型契約名として使う。`ElementTypeKey` はクラス モジュール名に使える形式である必要がある
- 要求能力
  - 制約: `RequireComparable=True` では `IComparable` 実装が必要。同一性/重複判定は `G_OBJECT_KEY_MODE_REFERENCE`、`G_OBJECT_KEY_MODE_I_EQUATABLE`、`G_OBJECT_KEY_MODE_DUPLICATE_CHECKABLE` のいずれかで決まる
- `Empty` / `Null`
  - 制約: 要素として使用できない
- Excel エラー値
  - 制約: `CVErr` は要素として扱える
- オブジェクト比較
  - 制約: `ObjectKeyMode` が明示されない従来推論では、`IDuplicateCheckable`、`IEquatable`、参照一致の順で要素能力が推論される。`ObjectSet` の重複判定では `IDuplicateCheckable` を優先できる
- 配列要素
  - 制約: 型付き値キーで比較される
- `ObjectDictionary.Item(Key)`
  - 制約: getter は存在しない key を暗黙追加せずエラーにする。setter は追加または更新として扱う
- `ObjectDictionary.CompareMode`
  - 制約: 文字列 key の比較モードは最初の追加前だけ変更できる。`RemoveAll` 後も value の要素型契約と `CompareMode` は維持される
- 標準列挙
  - 制約: `ObjectList` と `ObjectSet` の `For Each` は要素を列挙する。`ObjectDictionary` の `For Each` は key を列挙する。列挙開始時点のスナップショットを列挙するが、列挙中に元のコレクションを更新または削除した場合、実行中のループの継続は保証しない。`ObjectDictionary` は `GetEnumerator` を公開しないため、key は `Keys`、value は `Items` または `ConvertToArray` を使い、更新や削除は明示 API で行う

### ファイル/テキスト API の制約

- 環境変数
  - 制約: `FileSystemService.GetAbsolutePath` は URL ではないローカルパスに限り、絶対パス判定前に OS の環境変数展開を適用する。未定義変数や `%` を含む文字列の扱いは OS の `ExpandEnvironmentStrings` の結果に従う
- 相対パス
  - 制約: `FileSystemService` は `WbSrv.GetThisWorkbookDirectoryPath` 基準で絶対化する
- ドライブ相対パス
  - 制約: `C:foo` のような形式はカレントディレクトリ依存のためエラーにする
- 一時フォルダー
  - 制約: `GetTemporaryDirectoryPath` は OS ユーザー一時フォルダーを末尾 `\` なしで返す。`CreateTemporaryDirectory` は同フォルダー直下へ `Scripting.FileSystemObject.GetTempName` ベースのディレクトリを作成し、削除は呼び出し側が行う
- 破壊的操作
  - 制約: `Move*`、`Copy*`、`Remove*`、`CreateTemporaryDirectory`、`CreateBackupFile(Move:=True)` は実ファイルシステムを変更する
- テキスト文字コード
  - 制約: `TextFileEntity` に文字コード指定 API はない
- 書き込み
  - 制約: `OpenFile(AsWrite:=True, AsAppend:=False, Force:=False)` で既存ファイルがある場合はエラー

### IPv4 の制約

- 値表現
  - 制約: IPv4 は符号付き `Long` で保持する。最上位ビットが立つアドレスは負値になる
- マスク長
  - 制約: `ConvertFromMaskLength` は 0-32 のみ許可する
- マスク値
  - 制約: `ConvertToMaskLength` は連続した 1 のマスクのみ許可する
- ネットワーク狭窄
  - 制約: `NarrowNetwork` は実装上 `/30` から `/32` をエラーにする。仕様意図は未確認

## 依存関係

### 外部

| 依存先 | 利用内容 |
| --- | --- |
| Excel Object Model | ブック、シート、範囲、書式、図形、アプリケーション状態 |
| VBIDE / VBA Extensibility | `UnitTestMain`、`RemoveVBComponents`、実行用一時モジュール |
| `Scripting.FileSystemObject` | ファイル/ディレクトリ操作 |
| `Scripting.Dictionary` | `ObjectSet`、`ObjectDictionary`、キー管理、テストダブル記録 |
| `WScript.Shell` | ローカルパスの環境変数展開 |
| `VBScript.RegExp` | パス/IPv4/テスト検出などの正規表現 |
| `Forms.TextBox.1` | クリップボード操作 |
| Excel `WorksheetFunction.XLookup` | `WorksheetService.XLookup` |

### 内部

以下は、標準的にインポートするモジュールとして扱う。

- `Lib_Common`
- `Lib_CommonConstructor`
- `Fx_Common`
- `ApplicationScreenUpdateManager`
- `CommonRunStateManager`
- `DebugInformation`
- `ArrayObject`
- `ObjectList`
- `ObjectSet`
- `ObjectDictionary`
- `Counter`
- `CounterSet`
- `Enumerator`
- `IEnumerator`
- `IElementTypeProvider`
- `IEquatable`
- `IStringable`
- `IDuplicateCheckable`
- `IComparable`
- `IWorkbookService`
- `WorkbookService`
- `IWorksheetService`
- `WorksheetService`
- `WorksheetRangeBounds`
- `WorksheetRangeBoundsEnumerator`
- `WorksheetVirtualTable`
- `WorksheetVirtualTableEnumerator`

その他、必要に応じてインポートするモジュールの依存は以下のとおりである。

| 領域 | 利用元モジュール | 依存先モジュール |
| --- | --- | --- |
| 入力画面 | `Lib_InputSheet` | `IUserInputSheet`、`UserInputSheet` |
| 入力画面 | `UserInputSheet` | `IUserInputSheet` |
| コレクション・列挙 | `CounterSet` | `Counter` |
| ファイル操作 | `Lib_FileSystem` | `IFileSystemService`、`FileSystemService` |
| ファイル操作 | `FileSystemService` | `IFileSystemService` |
| ファイル操作 | `Lib_TextFile` | `ITextFileService`、`TextFileService` |
| ファイル操作 | `TextFileService` | `ITextFileService`、`ITextFileEntity`、`TextFileEntity` |
| ファイル操作 | `TextFileEntity` | `ITextFileEntity` |
| ファイル操作 | `ITextFileService` | `ITextFileEntity` |

ユニット テストを行う場合は、以下を標準的にインポートするモジュールとする。

- `Lib_UnitTest`
- `UnitTestAssert`
- `TestDoubleBehaviorStore`
- `TestDoubleCallRecord`
- `TestDoubleVariantKeyBuilder`

その他、必要に応じてインポートするモジュールは以下のとおりである。

| 領域 | 利用元モジュール | 依存先モジュール |
| --- | --- | --- |
| テスト支援 | `FileSystemServiceTestDouble` | `IFileSystemService` |
| テスト支援 | `TextFileServiceTestDouble` | `ITextFileService`、`ITextFileEntity`、`TextFileEntity`、`TextFileEntityTestDouble` |
| テスト支援 | `TextFileEntityTestDouble` | `ITextFileEntity` |
| テスト支援 | `UserInputSheetTestDouble` | `IUserInputSheet` |

## 内部実装の要約

- インターフェイスは VBA の `Implements` 用クラスとして定義され、直接生成されるとエラーになる。
- 主要サービスと入力シートは具象クラスとテストダブルの両方が同じインターフェイスを実装する。
- 比較・重複・文字列化は `IComparable`、`IEquatable`、`IDuplicateCheckable`、`IStringable` によって拡張される。
- 要素型契約は `TypeName` または `IElementTypeProvider.ElementTypeKey` によって決まり、キー生成は `GetTypedValueKey` / `TestDoubleVariantKeyBuilder` に集約され、特殊値、配列、オブジェクト参照、インターフェイス実装を区別する。
- Excel 操作 API は、操作前後のアクティブブック・アプリケーション状態をできるだけ復元するよう実装されている。
