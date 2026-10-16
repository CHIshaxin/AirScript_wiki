---
title: "Enum"
source: https://airsheet.wps.cn/docs/ja-JP/api/excel/workbook/Enum.html
updated: 1789700916000
hash: hE69Ii-G
---

# Enum

Application に格納される列挙型

## XlAboveBelow

値が平均より上か下かを表します。

| フィールド | 値 | 意味 |
|---|---|---|
| xlAboveAverage | 0 | 平均より大きい |
| xlBelowAverage | 1 | 平均より小さい |
| xlEqualAboveAverage | 2 | 平均と等しい |

## XlAutoFillType

コピー元の範囲から、コピー先の範囲をどのように埋めるかを指定します。

| フィールド | 値 | 意味 |
|---|---|---|
| xlFillDefault | 0 | コピー先を埋めるために使用する値と書式を決定する |
| xlFillCopy | 1 | コピー元からコピー先へ値と書式をコピーする。必要に応じて繰り返す |
| xlFillSeries | 2 | コピー元の値を数列として拡張する（例: 「1, 2」から「3, 4, 5」）。コピー元から書式をコピーし、必要に応じて繰り返す |
| xlFillFormats | 3 | コピー元からコピー先へ書式のみをコピーする。必要に応じて繰り返す |
| xlFillValues | 4 | コピー元からコピー先へ値のみをコピーする。必要に応じて繰り返す |
| xlFillDays | 5 | コピー元からコピー先へ曜日名を拡張する。必要に応じて書式をコピーする |
| xlFillWeekdays | 6 | コピー元からコピー先へ稼働日の曜日名を拡張する。必要に応じて書式をコピーする |
| xlFillMonths | 7 | コピー元からコピー先へ月名を拡張する。必要に応じて書式をコピーする |
| xlFillYears | 8 | コピー元からコピー先へ年を拡張する。必要に応じて書式をコピーする |
| xlLinearTrend | 9 | 加算を仮定して数値を拡張する（例: 「1, 2」から「3, 4, 5」）。必要に応じて書式をコピーする |
| xlGrowthTrend | 10 | 乗算を仮定して数値を拡張する（例: 「1, 2」から「4, 8, 16」）。必要に応じて書式をコピーする |

## XlBordersIndex

取得する罫線のインデックスです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlDiagonalDown | 5 | 各セルの左上から右下への対角線 |
| xlDiagonalUp | 6 | 各セルの左下から右上への対角線 |
| xlEdgeLeft | 7 | 範囲の左端 |
| xlEdgeTop | 8 | 範囲の上端 |
| xlEdgeBottom | 9 | 範囲の下端 |
| xlEdgeRight | 10 | 範囲の右端 |
| xlInsideVertical | 11 | セル間の縦方向の内側罫線（外枠を除く） |
| xlInsideHorizontal | 12 | セル間の横方向の内側罫線（外枠を除く） |
| xlOutside | 13 | 四方の外枠 |
| xlInside | 14 | 内側領域 |

## XlBorderWeight

範囲周りの罫線の太さです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlMedium | -4138 | 中線 |
| xlHairline | 1 | ヘアライン（最細） |
| xlThin | 2 | 細線 |
| xlThick | 4 | 太線（最太） |

## XlCalcModeType

計算モードです。

| フィールド | 値 | 意味 |
|---|---|---|
| manual | `manual` | 手動 |
| automatic | `automatic` | 自動 |

## XlChartType

グラフの種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xl3DArea | -4098 | 3-D 面 |
| xl3DAreaStacked | 78 | 3-D 積み上げ面 |
| xl3DAreaStacked100 | 79 | 3-D 100% 積み上げ面 |
| xl3DBarClustered | 60 | 3-D 群集横棒 |
| xl3DBarStacked | 61 | 3-D 積み上げ横棒 |
| xl3DBarStacked100 | 62 | 3-D 100% 積み上げ横棒 |
| xl3DColumn | -4100 | 3-D 縦棒 |
| xl3DColumnClustered | 54 | 3-D 群集縦棒 |
| xl3DColumnStacked | 55 | 3-D 積み上げ縦棒 |
| xl3DColumnStacked100 | 56 | 3-D 100% 積み上げ縦棒 |
| xl3DLine | -4101 | 3-D 折れ線 |
| xl3DPie | -4102 | 3-D 円 |
| xl3DPieExploded | 70 | 分離した 3-D 円 |
| xlArea | 1 | 面 |
| xlAreaStacked | 76 | 積み上げ面 |
| xlAreaStacked100 | 77 | 100% 積み上げ面 |
| xlBarClustered | 57 | 群集横棒 |
| xlBarOfPie | 71 | 円の横棒 |
| xlBarStacked | 58 | 積み上げ横棒 |
| xlBarStacked100 | 59 | 100% 積み上げ横棒 |
| xlBubble | 15 | バブル |
| xlBubble3DEffect | 87 | 3-D バブル |
| xlColumnClustered | 51 | 群集縦棒 |
| xlColumnStacked | 52 | 積み上げ縦棒 |
| xlColumnStacked100 | 53 | 100% 積み上げ縦棒 |
| xlConeBarClustered | 102 | 群集円すい横棒 |
| xlConeBarStacked | 103 | 積み上げ円すい横棒 |
| xlConeBarStacked100 | 104 | 100% 積み上げ円すい横棒 |
| xlConeCol | 105 | 3-D 円すい縦棒 |
| xlConeColClustered | 99 | 群集円すい縦棒 |
| xlConeColStacked | 100 | 積み上げ円すい縦棒 |
| xlConeColStacked100 | 101 | 100% 積み上げ円すい縦棒 |
| xlCylinderBarClustered | 95 | 群集円柱横棒 |
| xlCylinderBarStacked | 96 | 積み上げ円柱横棒 |
| xlCylinderBarStacked100 | 97 | 100% 積み上げ円柱横棒 |
| xlCylinderCol | 98 | 3-D 円柱縦棒 |
| xlCylinderColClustered | 92 | 群集円柱縦棒 |
| xlCylinderColStacked | 93 | 積み上げ円柱縦棒 |
| xlCylinderColStacked100 | 94 | 100% 積み上げ円柱縦棒 |
| xlDoughnut | -4120 | ドーナツ |
| xlDoughnutExploded | 80 | 分離したドーナツ |
| xlLine | 4 | 折れ線 |
| xlLineMarkers | 65 | データ点付き折れ線 |
| xlLineMarkersStacked | 66 | 積み上げデータ点付き折れ線 |
| xlLineMarkersStacked100 | 67 | 100% 積み上げデータ点付き折れ線 |
| xlLineStacked | 63 | 積み上げ折れ線 |
| xlLineStacked100 | 64 | 100% 積み上げ折れ線 |
| xlPie | 5 | 円 |
| xlPieExploded | 69 | 分離した円 |
| xlPieOfPie | 68 | 円の円 |
| xlPyramidBarClustered | 109 | 群集ピラミッド横棒 |
| xlPyramidBarStacked | 110 | 積み上げピラミッド横棒 |
| xlPyramidBarStacked100 | 111 | 100% 積み上げピラミッド横棒 |
| xlPyramidCol | 112 | 3-D ピラミッド縦棒 |
| xlPyramidColClustered | 106 | 群集ピラミッド縦棒 |
| xlPyramidColStacked | 107 | 積み上げピラミッド縦棒 |
| xlPyramidColStacked100 | 108 | 100% 積み上げピラミッド縦棒 |
| xlRadar | -4151 | レーダー |
| xlRadarFilled | 82 | 塗りつぶしレーダー |
| xlRadarMarkers | 81 | データ点付きレーダー |
| xlStockHLC | 88 | 高値・安値・終値の株価 |
| xlStockOHLC | 89 | 始値・高値・安値・終値の株価 |
| xlStockVHLC | 90 | 出来高・高値・安値・終値の株価 |
| xlStockVOHLC | 91 | 出来高・始値・高値・安値・終値の株価 |
| xlSurface | 83 | 3-D サーフェス |
| xlSurfaceTopView | 85 | サーフェス上面図 |
| xlSurfaceTopViewWireframe | 86 | サーフェス上面ワイヤーフレーム |
| xlSurfaceWireframe | 84 | 3-D サーフェス ワイヤーフレーム |
| xlXYScatter | -4169 | 散布図 |
| xlXYScatterLines | 74 | 線付き散布図 |
| xlXYScatterLinesNoMarkers | 75 | 線付き散布図（データ点なし） |
| xlXYScatterSmooth | 72 | 平滑線付き散布図 |
| xlXYScatterSmoothNoMarkers | 73 | 平滑線付き散布図（データ点なし） |

## XlContainsOperator

「含む」系の関数で使用する演算子です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlContains | 0 | 値を含む |
| xlDoesNotContain | 1 | 値を含まない |
| xlBeginsWith | 2 | 値で始まる |
| xlEndsWith | 3 | 値で終わる |

## XlDeleteShiftDirection

削除時にセルがどちらにずれるかです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlShiftToLeft | -4159 | セルを左へずらす |
| xlShiftUp | -4162 | セルを上へずらす |

## XlDirection

移動方向です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlDown | -4121 | 下 |
| xlToLeft | -4159 | 左 |
| xlToRight | -4161 | 右 |
| xlUp | -4162 | 上 |

## XlDVAlertStyle

入力規則のアラートで使用するアイコンです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlValidAlertStop | 1 | 停止アイコン |
| xlValidAlertWarning | 2 | 警告アイコン |
| xlValidAlertInformation | 3 | 情報アイコン |

## XlDVType

値に対して行う入力規則の種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlValidateWholeNumber | 1 | 整数 |
| xlValidateDecimal | 2 | 小数 |
| xlValidateList | 3 | 値はリストに存在する必要がある |
| xlValidateDate | 4 | 日付 |
| xlValidateTime | 5 | 時刻 |
| xlValidateTextLength | 6 | 文字列の長さ |
| xlValidateCustom | 7 | カスタム数式による入力規則 |

## XlExportImgFormatType

エクスポートする画像の形式です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlImgTypePNG | 0 | .png をエクスポート |
| xlImgTypeJPG | 1 | .jpg をエクスポート |
| xlImgTypeBMP | 2 | .bmp をエクスポート |
| xlImgTypeTIF | 2 | .tif をエクスポート |

## XlFixedFormatType

固定エクスポート形式の種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlTypePDF | 0 | PDF（.pdf） |
| xlTypeXPS | 1 | XPS（.xps） |
| xlTypeIMG | 2 | IMG（.png, .jpg, .bmp, .tif） |

## XlFormatConditionOperator

数式とセル値を比較する演算子。xlBetween および xlNotBetween を含みます。

| フィールド | 値 | 意味 |
|---|---|---|
| xlBetween | 1 | 範囲内。2 つの数式が必要 |
| xlNotBetween | 2 | 範囲外。2 つの数式が必要 |
| xlEqual | 3 | 等しい |
| xlNotEqual | 4 | 等しくない |
| xlGreater | 5 | より大きい |
| xlLess | 6 | より小さい |
| xlGreaterEqual | 7 | 以上 |
| xlLessEqual | 8 | 以下 |

## XlFormatConditionType

条件付き書式がセル値ベースか式ベースかを表します。

| フィールド | 値 | 意味 |
|---|---|---|
| xlCellValue | 1 | セル値 |
| xlExpression | 2 | 式 |
| xlColorScale | 3 | カラースケール |
| xlTop10 | 5 | 上位または下位の値 |
| xlUniqueValues | 8 | 一意の値 |
| xlTextString | 9 | 文字列 |
| xlBlanksCondition | 10 | 空白の条件 |
| xlTimePeriod | 11 | 期間 |
| xlAboveAverageCondition | 12 | 平均より上の条件 |
| xlNoBlanksCondition | 13 | 空白でない条件 |
| xlErrorsCondition | 16 | エラーの条件 |
| xlNoErrorsCondition | 17 | エラーでない条件 |

## XlHAlign

水平方向の配置です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlHAlignRight | -4152 | 右寄せ |
| xlHAlignLeft | -4131 | 左寄せ |
| xlHAlignJustify | -4130 | 両端揃え |
| xlHAlignDistributed | -4117 | 分散揃え |
| xlHAlignCenter | -4108 | 中央揃え |
| xlHAlignGeneral | 1 | 標準（データ型に応じて） |
| xlHAlignFill | 5 | 詰め |
| xlHAlignCenterAcrossSelection | 7 | 選択範囲にわたる中央揃え |

## XlInsertFormatOrigin

挿入したセルの書式をどこからコピーするかです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlFormatFromLeftOrAbove | 0 | 左または上から書式をコピー |
| xlFormatFromRightOrBelow | 1 | 右または下から書式をコピー |

## XlInsertShiftDirection

挿入時にセルがずれる方向です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlShiftDown | -4121 | セルを下へずらす |
| xlShiftToRight | -4161 | セルを右へずらす |

## XlLineStyle

罫線の線種です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlLineStyleNone | -4142 | 線なし |
| xlDouble | -4119 | 二重線 |
| xlDot | -4118 | 点線 |
| xlDash | -4115 | 破線 |
| xlContinuous | 1 | 実線 |
| xlDashDot | 4 | 一点鎖線 |
| xlDashDotDot | 5 | 二点鎖線 |
| xlSlantDashDot | 13 | 斜め一点鎖線 |

## XlPasteSpecialOperation

形式を選択して貼り付けるときの演算です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlPasteSpecialOperationAdd | 0x1 | 貼り付けデータをコピー先に加算 |
| xlPasteSpecialOperationDivide | 0x4 | 貼り付けデータをコピー先で除算 |
| xlPasteSpecialOperationMultiply | 0x3 | 貼り付けデータをコピー先で乗算 |
| xlPasteSpecialOperationNone | 0x0 | 貼り付け中に演算しない |
| xlPasteSpecialOperationSubtract | 0x2 | 貼り付けデータからコピー先を減算 |

## XlPasteType

貼り付けの種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlPasteAll | 0x1 | すべて貼り付け |
| xlPasteFormulas | 0x2 | 数式を貼り付け |
| xlPasteAllExceptBorders | 0x4 | 罫線以外をすべて貼り付け |
| xlPasteColumnWidths | 0x5 | 列幅を貼り付け |
| xlPasteValues | 0x3 | 値を貼り付け |
| xlPasteValuesAndNumberFormats | 0x7 | 値と表示形式を貼り付け |
| xlPasteFormats | 0x8 | 書式を貼り付け |
| xlPasteComments | 0x9 | コメントを貼り付け |
| xlPasteValidation | 0xa | 入力規則を貼り付け |

## XlReferenceStyle

参照スタイルです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlR1C1 | -4150 | R1C1 スタイルの参照に xlR1C1 を使用 |
| xlA1 | 1 | 既定。A1 スタイルの参照に xlA1 を使用 |

## XlRowCol

系列の値が行で並ぶか列で並ぶかです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlRows | 1 | データ系列が行方向 |
| xlColumns | 2 | データ系列が列方向 |

## XlSheetType

ワークシートの種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlWorksheet | -4167 | ワークシート |
| xlDialogSheet | -4116 | ダイアログシート |
| xlChart | -4109 | グラフ |
| xlExcel4MacroSheet | 3 | Excel 4.0 マクロシート |
| xlExcel4IntlMacroSheet | 4 | Excel 4.0 国際マクロシート |

## XlTimePeriods

期間フィルターです。

| フィールド | 値 | 意味 |
|---|---|---|
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

フォントの下線スタイルです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlUnderlineStyleDouble | -4119 | 二重下線 |
| xlUnderlineStyleDoubleAccounting | 5 | 会計用二重下線 |
| xlUnderlineStyleNone | -4142 | 下線なし |
| xlUnderlineStyleSingle | 2 | 単線下線 |
| xlUnderlineStyleSingleAccounting | 4 | 未対応 |

## XlVAlign

垂直方向の配置です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlVAlignTop | -4160 | 上揃え |
| xlVAlignJustify | -4130 | 両端揃え（縦） |
| xlVAlignDistributed | -4117 | 分散揃え（縦） |
| xlVAlignCenter | -4108 | 中央揃え |
| xlVAlignBottom | -4107 | 下揃え |

## XlXLMMacroType

シート上の名前が参照するマクロの種類です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlFunction | 1 | ユーザー定義関数 |
| xlCommand | 2 | ユーザー定義コマンド |
| xlNotXLM | 3 | マクロではない |

## XlAutoFilterOperator

AutoFilter の演算子、または 2 つの条件の論理の組み合わせです。

| フィールド | 値 | 意味 |
|---|---|---|
| xlAnd | 1 | Criteria1 と Criteria2 の論理積 |
| xlOr | 2 | Criteria1 と Criteria2 の論理和 |
| xlTop10Items | 3 | 上位 N 件（N は Criteria1） |
| xlBottom10Items | 4 | 下位 N 件（N は Criteria1） |
| xlTop10Percent | 5 | 上位の割合（割合は Criteria1） |
| xlBottom10Percent | 6 | 下位の割合（割合は Criteria1） |
| xlFilterValues | 7 | 値によるフィルター |
| xlFilterCellColor | 8 | セルの色 |
| xlFilterFontColor | 9 | フォントの色 |
| xlFilterIcon | 10 | アイコンによるフィルター |
| xlFilterDynamic | 11 | 動的フィルター |

## XlYesNoGuess

先頭行に見出しが含まれるか。ピボットテーブルの並べ替えでは使用しません。

| フィールド | 値 | 意味 |
|---|---|---|
| xlGuess | 0 | Excel に見出しを検出させる |
| xlYes | 1 | 既定。範囲全体を並べ替える |
| xlNo | 2 | 範囲全体を見出しとして並べ替えない |

## XlSortOrientation

並べ替えの向き（行と列）です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlSortColumns | 1 | 列で並べ替え |
| xlSortRows | 2 | 行で並べ替え（既定） |

## XlSortMethod

東アジア語の並べ替え方法です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlPinYin | 1 | 東アジア語テキストの読み順（既定） |
| xlStroke | 2 | 東アジア語テキストの画数順 |

## XlSortDataOption

テキストの並べ替え方法です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlSortNormal | 0 | 数値とテキストを別々に並べ替える（既定） |
| xlSortTextAsNumbers | 1 | テキストを数値として並べ替える |

## XlSortOn

並べ替えに使用するプロパティです。

| フィールド | 値 | 意味 |
|---|---|---|
| SortOnCellColor | 1 | セルの色 |
| SortOnFontColor | 2 | フォントの色 |
| SortOnIcon | 3 | アイコン |
| SortOnValues | 0 | 値 |

## XlSortOrder

フィールドまたは範囲の並べ替え順序です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlAscending | 1 | 既定。フィールドの昇順 |
| xlDescending | 2 | 降順 |

## XlTextParsingType

区切りテキストを取り込むときの列の形式です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlDelimited | 1 | 既定。区切り文字付きテキストファイル |
| xlFixedWidth | 2 | 固定幅の列 |

## XlTextQualifier

解析時のテキスト修飾子です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlTextQualifierDoubleQuote | 1 | 二重引用符（"） |
| xlTextQualifierNone | -4142 | 修飾子なし |
| xlTextQualifierSingleQuote | 2 | 単一引用符（'） |

## XlColumnDataType

テキストを分割するときの列のデータ型です。

| フィールド | 値 | 意味 |
|---|---|---|
| xlDMYFormat | 4 | DMY 日付形式 |
| xlDYMFormat | 7 | DYM 日付形式 |
| xlEMDFormat | 10 | EMD 日付形式 |
| xlGeneralFormat | 1 | 標準 |
| xlMDYFormat | 3 | MDY 日付形式 |
| xlMYDFormat | 6 | MYD 日付形式 |
| xlSkipColumn | 9 | 列をインポートしない |
| xlTextFormat | 2 | テキスト |
| xlYDMFormat | 8 | YDM 日付形式 |
| xlYMDFormat | 5 | YMD 日付形式 |
