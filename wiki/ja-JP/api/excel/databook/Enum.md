---
title: "Enum"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/databook/Enum.html
updated: 1789700916000
hash: DgIfYgjA
---

# Enum

列挙型は `Application` から公開されます。

## XlAboveBelow

値が平均より上か下かを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlAboveAverage | 0 | 平均より上 |
| xlBelowAverage | 1 | 平均より下 |
| xlEqualAboveAverage | 2 | 平均と等しい |

## XlAutoFillType

コピー元の範囲に基づいて、貼り付け先の範囲をどのように塗りつぶすかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlFillDefault | 0 | 貼り付け先を塗りつぶすために使用する値と書式を決定します |
| xlFillCopy | 1 | コピー元から貼り付け先へ値と書式をコピーします。必要に応じて繰り返します |
| xlFillSeries | 2 | コピー元から貼り付け先へ、数列として値を拡張します（例:「1, 2」→「3, 4, 5」）。書式はコピー元から必要に応じて繰り返してコピーします |
| xlFillFormats | 3 | 書式のみをコピー元から貼り付け先へコピーします。必要に応じて繰り返します |
| xlFillValues | 4 | 値のみをコピー元から貼り付け先へコピーします。必要に応じて繰り返します |
| xlFillDays | 5 | コピー元から貼り付け先へ曜日名を拡張します。書式はコピー元から必要に応じて繰り返してコピーします |
| xlFillWeekdays | 6 | コピー元から貼り付け先へ平日の曜日名を拡張します。書式はコピー元から必要に応じて繰り返してコピーします |
| xlFillMonths | 7 | コピー元から貼り付け先へ月名を拡張します。書式はコピー元から必要に応じて繰り返してコピーします |
| xlFillYears | 8 | コピー元から貼り付け先へ年を拡張します。書式はコピー元から必要に応じて繰り返してコピーします |
| xlLinearTrend | 9 | 加算的トレンドとみなして数値を拡張します（例:「1, 2」→「3, 4, 5」）。書式はコピー元から必要に応じて繰り返してコピーします |
| xlGrowthTrend | 10 | 乗算的トレンドとみなして数値を拡張します（例:「1, 2」→「4, 8, 16」）。書式はコピー元から必要に応じて繰り返してコピーします |

## XlBordersIndex

取得する罫線を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlDiagonalDown | 5 | 各セルの左上から右下への罫線 |
| xlDiagonalUp | 6 | 各セルの左下から右上への罫線 |
| xlEdgeLeft | 7 | 範囲の左端の罫線 |
| xlEdgeTop | 8 | 範囲の上端の罫線 |
| xlEdgeBottom | 9 | 範囲の下端の罫線 |
| xlEdgeRight | 10 | 範囲の右端の罫線 |
| xlInsideVertical | 11 | 範囲内の縦罫線（外枠を除く） |
| xlInsideHorizontal | 12 | 範囲内の横罫線（外枠を除く） |
| xlOutside | 13 | 範囲の上・下・左・右 |
| xlInside | 14 | 内側の領域 |

## XlBorderWeight

範囲周りの罫線の太さを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlMedium | -4138 | 中 |
| xlHairline | 1 | 最細線（最も細い罫線） |
| xlThin | 2 | 細線 |
| xlThick | 4 | 太線（最も太い罫線） |

## XlCalcModeType

反復計算モード。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| manual | `manual` | 手動 |
| automatic | `automatic` | 自動 |

## XlChartType

グラフの種類を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xl3DArea | -4098 | 3D 面グラフ |
| xl3DAreaStacked | 78 | 3D 積み上げ面グラフ |
| xl3DAreaStacked100 | 79 | 100% 積み上げ面グラフ |
| xl3DBarClustered | 60 | 3D 集合横棒グラフ |
| xl3DBarStacked | 61 | 3D 積み上げ横棒グラフ |
| xl3DBarStacked100 | 62 | 3D 100% 積み上げ横棒グラフ |
| xl3DColumn | -4100 | 3D 縦棒グラフ |
| xl3DColumnClustered | 54 | 3D 集合縦棒グラフ |
| xl3DColumnStacked | 55 | 3D 積み上げ縦棒グラフ |
| xl3DColumnStacked100 | 56 | 3D 100% 積み上げ縦棒グラフ |
| xl3DLine | -4101 | 3D 折れ線グラフ |
| xl3DPie | -4102 | 3D 円グラフ |
| xl3DPieExploded | 70 | 切り離し 3D 円グラフ |
| xlArea | 1 | 面グラフ |
| xlAreaStacked | 76 | 積み上げ面グラフ |
| xlAreaStacked100 | 77 | 100% 積み上げ面グラフ |
| xlBarClustered | 57 | 集合横棒グラフ |
| xlBarOfPie | 71 | 円の横棒グラフ |
| xlBarStacked | 58 | 積み上げ横棒グラフ |
| xlBarStacked100 | 59 | 100% 積み上げ横棒グラフ |
| xlBubble | 15 | バブルチャート |
| xlBubble3DEffect | 87 | 3D バブルチャート |
| xlColumnClustered | 51 | 集合縦棒グラフ |
| xlColumnStacked | 52 | 積み上げ縦棒グラフ |
| xlColumnStacked100 | 53 | 100% 積み上げ縦棒グラフ |
| xlConeBarClustered | 102 | 集合円錐横棒グラフ |
| xlConeBarStacked | 103 | 積み上げ円錐横棒グラフ |
| xlConeBarStacked100 | 104 | 100% 積み上げ円錐横棒グラフ |
| xlConeCol | 105 | 3D 円錐縦棒グラフ |
| xlConeColClustered | 99 | 集合円錐縦棒グラフ |
| xlConeColStacked | 100 | 積み上げ円錐縦棒グラフ |
| xlConeColStacked100 | 101 | 100% 積み上げ円錐縦棒グラフ |
| xlCylinderBarClustered | 95 | 集合円柱横棒グラフ |
| xlCylinderBarStacked | 96 | 積み上げ円柱横棒グラフ |
| xlCylinderBarStacked100 | 97 | 100% 積み上げ円柱横棒グラフ |
| xlCylinderCol | 98 | 3D 円柱縦棒グラフ |
| xlCylinderColClustered | 92 | 集合円柱縦棒グラフ |
| xlCylinderColStacked | 93 | 積み上げ円柱縦棒グラフ |
| xlCylinderColStacked100 | 94 | 100% 積み上げ円柱縦棒グラフ |
| xlDoughnut | -4120 | ドーナツグラフ |
| xlDoughnutExploded | 80 | 切り離しドーナツグラフ |
| xlLine | 4 | 折れ線グラフ |
| xlLineMarkers | 65 | データマーカー付き折れ線グラフ |
| xlLineMarkersStacked | 66 | データマーカー付き積み上げ折れ線グラフ |
| xlLineMarkersStacked100 | 67 | データマーカー付き 100% 積み上げ折れ線グラフ |
| xlLineStacked | 63 | 積み上げ折れ線グラフ |
| xlLineStacked100 | 64 | 100% 積み上げ折れ線グラフ |
| xlPie | 5 | 円グラフ |
| xlPieExploded | 69 | 切り離し円グラフ |
| xlPieOfPie | 68 | 円の円グラフ |
| xlPyramidBarClustered | 109 | 集合ピラミッド横棒グラフ |
| xlPyramidBarStacked | 110 | 積み上げピラミッド横棒グラフ |
| xlPyramidBarStacked100 | 111 | 100% 積み上げピラミッド横棒グラフ |
| xlPyramidCol | 112 | 3D ピラミッド縦棒グラフ |
| xlPyramidColClustered | 106 | 集合ピラミッド縦棒グラフ |
| xlPyramidColStacked | 107 | 積み上げピラミッド縦棒グラフ |
| xlPyramidColStacked100 | 108 | 100% 積み上げピラミッド縦棒グラフ |
| xlRadar | -4151 | レーダーチャート |
| xlRadarFilled | 82 | 塗りつぶしレーダーチャート |
| xlRadarMarkers | 81 | データマーカー付きレーダーチャート |
| xlStockHLC | 88 | 高値・安値・終値の株価グラフ |
| xlStockOHLC | 89 | 始値・高値・安値・終値の株価グラフ |
| xlStockVHLC | 90 | 出来高・高値・安値・終値の株価グラフ |
| xlStockVOHLC | 91 | 出来高・始値・高値・安値・終値の株価グラフ |
| xlSurface | 83 | 3D サーフェスグラフ |
| xlSurfaceTopView | 85 | サーフェスグラフ（上面表示） |
| xlSurfaceTopViewWireframe | 86 | サーフェスグラフ上面（ワイヤーフレーム） |
| xlSurfaceWireframe | 84 | 3D サーフェスグラフ（ワイヤーフレーム） |
| xlXYScatter | -4169 | 散布図 |
| xlXYScatterLines | 74 | 直線付き散布図 |
| xlXYScatterLinesNoMarkers | 75 | 直線付き散布図（マーカーなし） |
| xlXYScatterSmooth | 72 | 平滑線付き散布図 |
| xlXYScatterSmoothNoMarkers | 73 | 平滑線付き散布図（マーカーなし） |

## XlContainsOperator

関数で使用する演算子を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlContains | 0 | 指定値を含む |
| xlDoesNotContain | 1 | 指定値を含まない |
| xlBeginsWith | 2 | 指定値で始まる |
| xlEndsWith | 3 | 指定値で終わる |

## XlDeleteShiftDirection

削除されたセルを埋めるためにセルがどう移動するかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlShiftToLeft | -4159 | 左へシフト |
| xlShiftUp | -4162 | 上へシフト |

## XlDirection

移動する方向を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlDown | -4121 | 下 |
| xlToLeft | -4159 | 左 |
| xlToRight | -4161 | 右 |
| xlUp | -4162 | 上 |

## XlDVAlertStyle

入力規則のメッセージボックスに使用するアイコンを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlValidAlertStop | 1 | 停止アイコン |
| xlValidAlertWarning | 2 | 警告アイコン |
| xlValidAlertInformation | 3 | 情報アイコン |

## XlDVType

値に適用する入力規則の種類を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlValidateWholeNumber | 1 | 整数 |
| xlValidateDecimal | 2 | 小数 |
| xlValidateList | 3 | 値は指定リストに存在する必要あり |
| xlValidateDate | 4 | 日付 |
| xlValidateTime | 5 | 時刻 |
| xlValidateTextLength | 6 | 文字数 |
| xlValidateCustom | 7 | カスタム数式による入力規則 |

## XlExportImgFormatType

画像のエクスポート形式。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlImgTypePNG | 0 | .png をエクスポート |
| xlImgTypeJPG | 1 | .jpg をエクスポート |
| xlImgTypeBMP | 2 | .bmp をエクスポート |
| xlImgTypeTIF | 2 | .tif をエクスポート |

## XlFixedFormatType

ファイル形式の種類を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlTypePDF | 0 | PDF（.pdf） |
| xlTypeXPS | 1 | XPS（.xps） |
| xlTypeIMG | 2 | IMG（.png、.jpg、.bmp、.tif） |

## XlFormatConditionOperator

数式とセル値を比較する演算子、または xlBetween と xlNotBetween で 2 つの数式を比較する演算子を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlBetween | 1 | 範囲内。2 つの数式が指定されている場合のみ有効 |
| xlNotBetween | 2 | 範囲外。2 つの数式が指定されている場合のみ有効 |
| xlEqual | 3 | 等しい |
| xlNotEqual | 4 | 等しくない |
| xlGreater | 5 | より大きい |
| xlLess | 6 | より小さい |
| xlGreaterEqual | 7 | 以上 |
| xlLessEqual | 8 | 以下 |

## XlFormatConditionType

条件付き書式がセル値に基づくか式に基づくかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlCellValue | 1 | セル値 |
| xlExpression | 2 | 式 |
| xlColorScale | 3 | カラースケール |
| xlTop10 | 5 | 上位 10 項目 |
| xlUniqueValues | 8 | 一意の値 |
| xlTextString | 9 | 文字列 |
| xlBlanksCondition | 10 | 空白の条件 |
| xlTimePeriod | 11 | 期間 |
| xlAboveAverageCondition | 12 | 平均より上の条件 |
| xlNoBlanksCondition | 13 | 空白でない条件 |
| xlErrorsCondition | 16 | エラーの条件 |
| xlNoErrorsCondition | 17 | エラーでない条件 |

## XlHAlign

水平方向の配置を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlHAlignRight | -4152 | 右寄せ |
| xlHAlignLeft | -4131 | 左寄せ |
| xlHAlignJustify | -4130 | 両端揃え |
| xlHAlignDistributed | -4117 | 分散揃え |
| xlHAlignCenter | -4108 | 中央 |
| xlHAlignGeneral | 1 | 標準（データ型に応じて） |
| xlHAlignFill | 5 | 詰め |
| xlHAlignCenterAcrossSelection | 7 | 選択範囲内で中央 |

## XlInsertFormatOrigin

セルを挿入するときに書式をどこからコピーするかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlFormatFromLeftOrAbove | 0 | 左または上から書式をコピー |
| xlFormatFromRightOrBelow | 1 | 右または下から書式をコピー |

## XlInsertShiftDirection

挿入時にセルがどうシフトするかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlShiftDown | -4121 | 下へシフト |
| xlShiftToRight | -4161 | 右へシフト |

## XlLineStyle

罫線の線種を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlLineStyleNone | -4142 | なし |
| xlDouble | -4119 | 二重線 |
| xlDot | -4118 | 点線 |
| xlDash | -4115 | 破線 |
| xlContinuous | 1 | 実線 |
| xlDashDot | 4 | 一点鎖線 |
| xlDashDotDot | 5 | 二点鎖線 |
| xlSlantDashDot | 13 | 斜め一点鎖線 |

## XlPasteSpecialOperation

XlPaste の特殊操作。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlPasteSpecialOperationAdd | 0x1 | コピーしたデータを貼り付け先の値に加算 |
| xlPasteSpecialOperationDivide | 0x4 | コピーしたデータを貼り付け先の値で除算 |
| xlPasteSpecialOperationMultiply | 0x3 | コピーしたデータに貼り付け先の値を乗算 |
| xlPasteSpecialOperationNone | 0x0 | 貼り付け中は計算しない |
| xlPasteSpecialOperationSubtract | 0x2 | コピーしたデータから貼り付け先の値を減算 |

## XlPasteType

貼り付けの種類。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlPasteFormulas | 0x2 | 数式 |
| xlPasteAllExceptBorders | 0x4 | 罫線以外すべて |
| xlPasteColumnWidths | 0x5 | 列幅 |
| xlPasteValues | 0x3 | 値 |
| xlPasteValuesAndNumberFormats | 0x7 | 値と表示形式 |
| xlPasteFormats | 0x8 | 書式 |
| xlPastePasteAll | 0x0 | すべて貼り付け |
| xlPasteComments | 0x9 | コメント |
| xlPasteValidation | 0xa | 入力規則 |

## XlReferenceStyle

参照形式を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlR1C1 | -4150 | xlR1C1 を使用して R1C1 形式の参照を返す |
| xlA1 | 1 | 既定。xlA1 を使用して A1 形式の参照を返す |

## XlRowCol

データ系列の値が行と列のどちらに並ぶかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlRows | 1 | データ系列は列方向 |
| xlColumns | 2 | データ系列は行方向 |

## XlSheetType

ワークシートの種類を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlWorksheet | -4167 | ワークシート |
| xlDialogSheet | -4116 | ダイアログシート |
| xlChart | -4109 | グラフ |
| xlExcel4MacroSheet | 3 | Excel 4.0 マクロシート |
| xlExcel4IntlMacroSheet | 4 | Excel 4.0 国際マクロシート |

## XlTimePeriods

期間を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlToday | 0 | 今日 |
| xlYesterday | 1 | 昨日 |
| xlLast7Days | 2 | 過去 7 日間 |
| xlThisWeek | 3 | 今週 |
| xlLastWeek | 4 | 先週 |
| xlLastMonth | 5 | 先月 |
| xlTomorrow | 6 | 明日 |
| xlNextWeek | 7 | 来週 |
| xlNextMonth | 8 | 来月 |
| xlThisMonth | 9 | 今月 |

## XlUnderlineStyle

フォントに適用する下線のスタイルを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlUnderlineStyleDouble | -4119 | 二重下線 |
| xlUnderlineStyleDoubleAccounting | 5 | 会計用二重下線（細い下線 2 本） |
| xlUnderlineStyleNone | -4142 | 下線なし |
| xlUnderlineStyleSingle | 2 | 単下線 |
| xlUnderlineStyleSingleAccounting | 4 | 未サポート |

## XlVAlign

垂直方向の配置を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlVAlignTop | -4160 | 上 |
| xlVAlignJustify | -4130 | 両端揃え（縦） |
| xlVAlignDistributed | -4117 | 分散揃え（縦） |
| xlVAlignCenter | -4108 | 中央 |
| xlVAlignBottom | -4107 | 下 |

## XlXLMMacroType

名前がワークシート内のマクロを参照するか、そもそもマクロを参照するかを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlFunction | 1 | ユーザー定義関数 |
| xlCommand | 2 | ユーザー定義コマンド |
| xlNotXLM | 3 | マクロではない |

## XlAutoFilterOperator

現在のワークシートのフィルター種類、または 2 つのフィルター条件を結合する演算子を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlAnd | 1 | Criteria1 と Criteria2 の論理 AND |
| xlOr | 2 | Criteria1 と Criteria2 の論理 OR |
| xlTop10Items | 3 | 上位項目（件数は Criteria1） |
| xlBottom10Items | 4 | 下位項目（件数は Criteria1） |
| xlTop10Percent | 5 | 上位パーセント（割合は Criteria1） |
| xlBottom10Percent | 6 | 下位パーセント（割合は Criteria1） |
| xlFilterValues | 7 | フィルター値 |
| xlFilterCellColor | 8 | セルの塗りつぶし色 |
| xlFilterFontColor | 9 | フォントの色 |
| xlFilterIcon | 10 | フィルターアイコン |
| xlFilterDynamic | 11 | 動的フィルター |

## XlYesNoGuess

先頭行に見出しが含まれるかを指定します。ピボットテーブルの並べ替えでは使用できません。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlGuess | 0 | Excel が見出しを持つか推測する |
| xlYes | 1 | 既定。範囲全体を並べ替える |
| xlNo | 2 | 範囲全体を並べ替えない |

## XlSortOrientation

並べ替えの向きを指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlSortColumns | 1 | 列で並べ替え |
| xlSortRows | 2 | 行で並べ替え（既定） |

## XlSortMethod

並べ替え方法を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlPinYin | 1 | 中国語文字を拼音で並べ替え（既定） |
| xlStroke | 2 | 各文字の画数で並べ替え |

## XlSortDataOption

文字列の並べ替え方法を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlSortNormal | 0 | 数値と文字列を別々に並べ替え（既定） |
| xlSortTextAsNumbers | 1 | 文字列を数値データとして並べ替え |

## XlSortOn

並べ替えの基準を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| SortOnCellColor | 1 | セルの塗りつぶし色 |
| SortOnFontColor | 2 | フォントの色 |
| SortOnIcon | 3 | アイコン |
| SortOnValues | 0 | 値 |

## XlSortOrder

フィールドまたは範囲の並べ替え順を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlAscending | 1 | 既定。昇順 |
| xlDescending | 2 | 降順 |

## XlTextParsingType

テキストファイルをクエリ表に取り込むときの列データ形式を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlDelimited | 1 | 既定。区切りファイル |
| xlFixedWidth | 2 | 固定幅列 |

## XlTextQualifier

テキストの修飾子を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlTextQualifierDoubleQuote | 1 | 二重引用符（"） |
| xlTextQualifierNone | -4142 | 修飾子なし |
| xlTextQualifierSingleQuote | 2 | 単一引用符（'） |

## XlColumnDataType

列の解析方法を指定します。

| メンバー | 値 | 説明 |
| --- | --- | --- |
| xlDMYFormat | 4 | DMY 日付 |
| xlDYMFormat | 7 | DYM 日付 |
| xlEMDFormat | 10 | EMD 日付 |
| xlGeneralFormat | 1 | 標準 |
| xlMDYFormat | 3 | MDY 日付 |
| xlMYDFormat | 6 | MYD 日付 |
| xlSkipColumn | 9 | 列をスキップ |
| xlTextFormat | 2 | 文字列 |
| xlYDMFormat | 8 | YDM 日付 |
| xlYMDFormat | 5 | YMD 日付 |
