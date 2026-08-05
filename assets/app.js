
  (function () {
    "use strict";

    var SUPPORTED_EXTENSIONS = {
      doc: { family: "word", container: "cfb" },
      docx: { family: "word", container: "zip" },
      docm: { family: "word", container: "zip" },
      xls: { family: "excel", container: "cfb" },
      xlsx: { family: "excel", container: "zip" },
      xlsm: { family: "excel", container: "zip" },
      xltx: { family: "excel", container: "zip" },
      xltm: { family: "excel", container: "zip" },
      ppt: { family: "powerpoint", container: "cfb" },
      pptx: { family: "powerpoint", container: "zip" },
      pptm: { family: "powerpoint", container: "zip" },
      potx: { family: "powerpoint", container: "zip" },
      potm: { family: "powerpoint", container: "zip" }
    };

    var OOXML_MAIN_PARTS = {
      word: "word/document.xml",
      excel: "xl/workbook.xml",
      powerpoint: "ppt/presentation.xml"
    };

    var CFB_SIGNATURE = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    var ZIP_LOCAL_SIGNATURE = 0x04034B50;
    var ZIP_CENTRAL_SIGNATURE = 0x02014B50;
    var ZIP_EOCD_SIGNATURE = 0x06054B50;
    var MAX_ZIP_ENTRIES = 100000;
    var MAX_XML_BYTES = 64 * 1024 * 1024;
    var MAX_OUTPUT_FILE_NAME = 64;
    var MAX_SOURCE_OUTPUT_STEM = 40;
    var MAX_MEDIA_SOURCE_PREFIX = 32;
    var MAX_OUTPUT_PARENT_PATH = 72;
    var MAX_OUTPUT_ROOT_PATH = 96;
    var MAX_OUTPUT_RELATIVE_PATH = 180;
    var MAX_OUTPUT_PARENT_SEGMENT = 32;
    var MIN_MEDIA_FILE_STEM = 8;
    var MAX_EXCEL_TABLE_CELLS = 1000000;
    var MAX_INPUT_FILE_BYTES = 256 * 1024 * 1024;
    var MAX_TOTAL_INPUT_BYTES = 512 * 1024 * 1024;
    var MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES = 128 * 1024 * 1024;
    var MAX_ZIP_TOTAL_UNCOMPRESSED_BYTES = 512 * 1024 * 1024;
    var MAX_COMPRESSION_RATIO = 200;
    var MAX_TOTAL_OUTPUT_BYTES = 512 * 1024 * 1024;
    var MAX_MEDIA_FILES = 10000;
    var MAX_EMBEDDING_FILES = 10000;
    var MAX_TEXT_OUTPUT_BYTES = 64 * 1024 * 1024;
    var MAX_CACHED_INPUT_BYTES = 32 * 1024 * 1024;
    var MAX_TOTAL_CACHED_INPUT_BYTES = 128 * 1024 * 1024;
    var MAX_ESTIMATED_WORKING_BYTES = 768 * 1024 * 1024;
    var MAX_LOG_LINES = 1000;
    var OUTPUT_NAME_CHARACTER_REPLACEMENTS = [
      ["髙", "高"]
    ];

    var TRANSLATIONS = {
      ja: {
        "meta.description": "Word・Excel・PowerPointファイルから画像、テキスト、埋め込みデータをブラウザ内で抽出し、ZIPにまとめるローカル処理ツールです。",
        "language.label": "表示言語",
        "hero.subtitle": "Word・Excel・PowerPointからメディアやテキストを一括抽出。ブラウザ内だけで抽出して1個のZIPにまとめます。",
        "hero.localBadge": "ブラウザ内で完結・外部送信なし",
        "input.heading": "入力ファイル",
        "button.addFiles": "ファイルを追加",
        "button.addFolder": "フォルダを追加",
        "button.removeSelected": "選択した項目を削除",
        "button.clearFiles": "クリア",
        "button.cancelProcessing": "処理を中止",
        "button.exportZip": "全件をZIPで書き出す",
        "button.clearLog": "ログを消去",
        "drop.heading": "ファイルまたはフォルダをここへドロップ",
        "drop.anywhere": "ページ内のどこにドロップしても追加できます",
        "drop.formats": "対応形式は .doc / .docx / .docm / .xls / .xlsx / .xlsm / .xltx / .xltm / .ppt / .pptx / .pptm / .potx / .potm",
        "drop.ariaLabel": "Officeファイルを追加。フォルダはドロップまたはフォルダを追加ボタンから選択できます",
        "option.includeSubfolders": "フォルダ追加時にサブフォルダも読み込む",
        "option.includeText": "テキストも抽出する",
        "option.includeEmbeddings": "埋め込みデータも抽出する",
        "warning.browserCompatibility": "このブラウザではOOXML形式のZIP展開機能を利用できません。ChromeまたはEdgeの現行安定版を使用してください。旧 .doc / .xls / .ppt は検査できます。",
        "table.selectAll": "削除対象をすべて選択",
        "table.selectItem": "{name}を削除対象として選択",
        "table.selectItemTitle": "削除対象として選択",
        "table.fileName": "ファイル名",
        "table.relativePath": "相対パス",
        "table.extension": "拡張子",
        "table.size": "サイズ",
        "table.status": "判定状態",
        "table.empty": "Officeファイルが追加されていません。",
        "table.selectionHelp": "一覧のチェックは削除対象の選択にのみ使用します。ZIP書き出しでは、一覧内の処理可能な全ファイルが対象になります。",
        "progress.label": "処理進捗",
        "progress.waiting": "待機中",
        "progress.checkingInternalFormat": "内部形式を確認しています",
        "progress.extractingMedia": "メディアを抽出しています",
        "progress.extractingText": "テキストを抽出しています",
        "progress.extractingEmbeddings": "埋め込みデータを抽出しています",
        "progress.addingToZip": "ZIPへ追加しています",
        "progress.readingFile": "{name}：ファイルを読み込んでいます",
        "progress.fileStage": "{name}：{stage}",
        "progress.fileFinished": "{name}：処理を終了しました",
        "progress.completed": "完了しました",
        "progress.generatingZip": "ZIPを生成しています",
        "progress.cancelled": "処理を中止しました",
        "progress.error": "エラーで終了しました",
        "log.heading": "処理ログ",
        "log.level.info": "情報",
        "log.level.warning": "警告",
        "log.level.error": "エラー",
        "log.level.success": "完了",
        "log.addPrompt": "ファイルを追加してください。",
        "log.cleared": "ログを消去しました。",
        "log.zeroByteSkipped": "0バイトのため追加しませんでした: {name}",
        "log.fileTooLarge": "入力ファイルのサイズ上限を超えています: {name}",
        "log.totalTooLarge": "入力ファイルの合計サイズが安全上限を超えるため追加しませんでした: {name}",
        "log.filesAdded.one": "{count}件のファイルを追加しました。",
        "log.filesAdded.other": "{count}件のファイルを追加しました。",
        "log.subfoldersSkipped.one": "サブフォルダ内の{count}件はオプションがオフのため追加しませんでした。",
        "log.subfoldersSkipped.other": "サブフォルダ内の{count}件はオプションがオフのため追加しませんでした。",
        "log.noFilesAdded": "追加できるファイルはありませんでした。",
        "log.noProcessableFiles": "処理対象ファイルがありません。",
        "log.fileWarning": "{name}：{warning}",
        "log.noExtractableContent": "抽出可能なメディア、テキスト、埋め込みデータが見つかりません: {name}",
        "log.fileError": "{message}: {name}",
        "log.noZipResults": "ZIPへ保存できる抽出結果がありませんでした。",
        "log.zipGenerated": "{name} を生成しました（抽出成功{success}件、結果なし{empty}件、失敗{failed}件）。",
        "log.cancelled": "利用者の操作により処理を中止しました。",
        "log.folderReadFailed.one": "{count}件のフォルダをドラッグ＆ドロップで読み取れませんでした。「フォルダを追加」を使用してください。",
        "log.folderReadFailed.other": "{count}件のフォルダをドラッグ＆ドロップで読み取れませんでした。「フォルダを追加」を使用してください。",
        "log.dropItemsFailed.one": "{count}件のドロップ項目を読み取れませんでした。「ファイルを追加」を使用してください。",
        "log.dropItemsFailed.other": "{count}件のドロップ項目を読み取れませんでした。「ファイルを追加」を使用してください。",
        "log.removed.one": "{count}件を一覧から削除しました。",
        "log.removed.other": "{count}件を一覧から削除しました。",
        "log.listCleared.one": "{count}件の入力一覧をクリアしました。",
        "log.listCleared.other": "{count}件の入力一覧をクリアしました。",
        "log.cancelRequested": "処理の中止を要求しました。現在の安全な区切りで停止します。",
        "log.dropReadFailed": "ドロップした項目を読み取れませんでした: {message}",
        "log.browserUnsupported": "このブラウザはDeflate展開APIに対応していません。ChromeまたはEdgeの現行安定版を使用してください。",
        "summary.result": "{name}（{media}、{characters}、{embeddings}）",
        "summary.media.one": "メディア{count}件",
        "summary.media.other": "メディア{count}件",
        "summary.characters.one": "テキスト{count}文字",
        "summary.characters.other": "テキスト{count}文字",
        "summary.embeddings.one": "埋め込み{count}件",
        "summary.embeddings.other": "埋め込み{count}件",
        "status.ready": "処理対象",
        "status.checking": "確認中",
        "status.unsupported": "未対応形式",
        "status.formatMismatch": "内部形式不一致",
        "status.encrypted": "暗号化の可能性",
        "status.browserUnsupported": "ブラウザ非対応",
        "status.processing": "処理中",
        "status.completed": "抽出完了",
        "status.completedWithWarnings": "警告あり",
        "status.noResults": "結果なし",
        "status.error": "解析エラー",
        "status.processingFailed": "処理失敗",
        "status.cancelled": "中止",
        "status.detail.unsupported": "未対応の拡張子です。",
        "status.detail.formatMismatchOoxml": "拡張子とOOXML内部形式が一致しません。",
        "status.detail.formatMismatchCfb": "拡張子とCFB内部形式が一致しません。",
        "status.detail.missingSignature": "ZIPまたはCFBシグネチャがありません。",
        "status.detail.ooxml": "OOXML {family}",
        "status.detail.cfb": "CFB {family}",
        "status.detail.encryptedOoxml": "暗号化OOXMLパッケージです。",
        "status.detail.encryptedLegacy": "Officeバイナリ形式に暗号化情報があります。",
        "status.detail.browserUnsupported": "このブラウザではOOXML形式のZIP展開機能を利用できません。",
        "status.detail.checking": "内部形式を確認しています。",
        "status.detail.processing": "抽出処理を実行しています。",
        "status.detail.noResults": "解析は完了しましたが、抽出対象がありませんでした。",
        "status.detail.completed": "抽出処理が完了しました。",
        "status.detail.completedWithWarnings": "一部の抽出処理に警告があります。",
        "status.detail.cancelled": "利用者の操作により処理を中止しました。",
        "message.value": "{message}",
        "warning.textOmitted": "抽出テキストが安全上限を超えたため、テキスト出力を省略しました。",
        "warning.noText": "有効な文書テキストが見つかりませんでした。",
        "warning.media": "一部のメディアを抽出できませんでした。",
        "warning.text": "一部のテキストを抽出できませんでした。",
        "warning.embedded": "一部の埋め込みデータを抽出できませんでした。",
        "warning.word": "一部のWord文書構造を完全に復元できませんでした。",
        "warning.excel": "一部のExcel文書構造を完全に復元できませんでした。",
        "warning.powerpoint": "一部のPowerPoint文書構造を完全に復元できませんでした。",
        "warning.generic": "一部の文書内容を完全に処理できませんでした。",
        "warning.subject.wordChart": "Wordグラフ",
        "warning.subject.altChunk": "altChunk",
        "warning.subject.footnote": "脚注",
        "warning.subject.endnote": "文末脚注",
        "warning.subject.legacyFootnote": "旧Word脚注",
        "warning.subject.legacyEndnote": "旧Word文末脚注",
        "warning.subject.mainTextbox": "本文テキストボックス",
        "warning.subject.headerFooterTextbox": "ヘッダー／フッター内テキストボックス",
        "warning.subject.mainShapeAnchor": "本文図形アンカー文字",
        "warning.subject.headerFooterShapeAnchor": "ヘッダー／フッター図形アンカー文字",
        "warning.subject.media": "メディア",
        "warning.subject.text": "テキスト",
        "warning.subject.embedding": "埋め込みデータ",
        "warning.document.word": "Word",
        "warning.document.legacyWord": "旧Word",
        "warning.value.empty": "（空）",
        "warning.value.noContentType": "Content Typeなし",
        "warning.section.unplacedMainTextbox": "配置を復元できないテキストボックス",
        "warning.section.unplacedHeaderFooterTextbox": "配置を復元できないヘッダー／フッター内テキストボックス",
        "warning.referenceMissingId": "{subject}参照にrelationship IDがないため読み飛ばしました。",
        "warning.referenceRelationshipMissing": "{subject}参照 {id} のrelationshipが見つかりませんでした。",
        "warning.externalReferenceSkipped": "外部{subject}参照は安全のため読み飛ばしました: {id}",
        "warning.referenceKindMismatch": "{subject}参照 {id} のrelationship種類が一致しませんでした。",
        "warning.partMissing": "{subject}パーツが見つかりませんでした: {name}",
        "warning.unsupportedAltChunk": "未対応のaltChunk形式を読み飛ばしました: {type}",
        "warning.partParseFailed": "{subject}パーツを解析できなかったため読み飛ばしました: {name}",
        "warning.annotationDefinitionMissingId": "{subject}定義にIDがないため、その定義を読み飛ばしました。",
        "warning.annotationDuplicateId": "{subject}ID {id} が重複しているため、最初の定義を使用しました。",
        "warning.annotationBodyMissing": "{subject}参照ID {id} に対応する本文が見つかりませんでした。",
        "warning.wordNoPageBreak": "{document}文書に利用可能なページ区切り情報がないため、文書全体を1つの論理ページとして脚注を配置しました。",
        "warning.unreferencedAnnotation.one": "未参照の{subject}が{count}件あります。",
        "warning.unreferencedAnnotation.other": "未参照の{subject}が{count}件あります。",
        "warning.legacyAnnotationEmptyStory": "旧Wordの{subject}本文がない状態で参照PLCを解析できませんでした。",
        "warning.legacyAnnotationCountMismatch": "旧Wordの{subject}参照数と本文数が一致しないため、対応できる範囲だけを配置します。",
        "warning.legacyAnnotationPositionFallback": "旧Wordの{subject}参照位置を復元できないため、独立セクションへ出力します。理由: {message}",
        "warning.legacyTextboxPlcFallback": "旧Wordの{subject}本文PLCを解析できないため、「{section}」セクションへ出力します。理由: {message}",
        "warning.legacyTextboxInvalidRange.one": "旧Wordの{subject}に不正なテキスト範囲が{count}件あるため、取得可能な本文を「{section}」セクションへ出力します。",
        "warning.legacyTextboxInvalidRange.other": "旧Wordの{subject}に不正なテキスト範囲が{count}件あるため、取得可能な本文を「{section}」セクションへ出力します。",
        "warning.legacyTextboxMissingParagraphEnd.one": "旧Wordの{subject}に段落終端文字がない定義が{count}件あるため、取得可能な本文を「{section}」セクションへ出力します。",
        "warning.legacyTextboxMissingParagraphEnd.other": "旧Wordの{subject}に段落終端文字がない定義が{count}件あるため、取得可能な本文を「{section}」セクションへ出力します。",
        "warning.legacyTextboxMetadataInvalid.one": "旧Wordの{subject}に不正な{field}が{count}件あるため、「{section}」セクションへ出力します。",
        "warning.legacyTextboxMetadataInvalid.other": "旧Wordの{subject}に不正な{field}が{count}件あるため、「{section}」セクションへ出力します。",
        "warning.legacyTextboxUnsupportedFlags.one": "旧Wordの{subject}に未対応のFTXBXSフラグが{count}件あります。",
        "warning.legacyTextboxUnsupportedFlags.other": "旧Wordの{subject}に未対応のFTXBXSフラグが{count}件あります。",
        "warning.legacyTextboxReuseManagement": "旧Wordの{subject}の再利用用管理情報に不整合を検出しました。\n再利用用定義は本文へ出力せず、通常テキストボックスの抽出を継続しました。",
        "warning.legacyTextboxTextRangeMismatch": "旧Wordの{subject}本文数とPLC範囲が一致しないため、対応不能部分を「{section}」セクションへ出力します。",
        "warning.legacyTextboxAnchorPlcFallback": "旧Wordの{subject}図形アンカーPLCを解析できないため、「{section}」セクションへ出力します。理由: {message}",
        "warning.legacyTextboxInvalidAnchorMarker.one": "旧Wordの{subject}が不正なものが{count}件あるため、対応するテキストボックスを「{section}」セクションへ出力しました。",
        "warning.legacyTextboxInvalidAnchorMarker.other": "旧Wordの{subject}が不正なものが{count}件あるため、対応するテキストボックスを「{section}」セクションへ出力しました。",
        "warning.legacyTextboxDuplicateAnchor.one": "旧Wordの{subject}図形アンカーに同じshape IDが{count}件重複しているため、最初のアンカーを使用します。",
        "warning.legacyTextboxDuplicateAnchor.other": "旧Wordの{subject}図形アンカーに同じshape IDが{count}件重複しているため、最初のアンカーを使用します。",
        "warning.legacyTextboxMissingAnchor.one": "旧Wordの{subject}を図形アンカーへ対応付けられないものが{count}件あるため、「{section}」セクションへ出力します。",
        "warning.legacyTextboxMissingAnchor.other": "旧Wordの{subject}を図形アンカーへ対応付けられないものが{count}件あるため、「{section}」セクションへ出力します。",
        "warning.legacyTextboxMissingCoordinate.one": "旧Wordの{subject}図形座標を取得できないアンカーが{count}件あります。CPとshape IDで配置を継続します。",
        "warning.legacyTextboxMissingCoordinate.other": "旧Wordの{subject}図形座標を取得できないアンカーが{count}件あります。CPとshape IDで配置を継続します。",
        "warning.legacyTextboxUnplaced.one": "旧Wordの{subject}{count}件を図形アンカーへ配置できなかったため、「{section}」セクションへ出力しました。",
        "warning.legacyTextboxUnplaced.other": "旧Wordの{subject}{count}件を図形アンカーへ配置できなかったため、「{section}」セクションへ出力しました。",
        "warning.legacyTextboxUnplacedUnknown": "旧Wordの{subject}の配置を復元できないため、「{section}」セクションへ出力しました。",
        "warning.legacyPapxRangeMismatch.one": "旧Wordの段落書式ページにBTE範囲との不整合が{count}件あったため、範囲内の段落だけを使用しました。",
        "warning.legacyPapxRangeMismatch.other": "旧Wordの段落書式ページにBTE範囲との不整合が{count}件あったため、範囲内の段落だけを使用しました。",
        "warning.legacyPapxFallback": "旧Wordの段落前改ページ情報を解析できないため、手動改ページとセクション区切りだけを使用しました。",
        "warning.legacyPapxFallbackWithReason": "旧Wordの段落前改ページ情報を解析できないため、手動改ページとセクション区切りだけを使用しました。理由: {message}",
        "warning.legacyPapxPartial": "旧Wordの段落前改ページ情報の一部を解析できないため、取得できたページ区切り情報だけを使用しました。",
        "warning.legacyAnnotationIndexMissing": "旧Wordの{subject}参照インデックス {id} に対応する本文が見つかりませんでした。",
        "warning.legacySectionBoundaryApproximation": "旧Wordのセクション境界に終端文字がないため、取得できたセクション種別からページ位置を近似しました。",
        "warning.legacyReferenceCharacter": "旧Wordの自動注釈参照位置に参照制御文字がないため、元の本文文字を保持しました。",
        "warning.legacySectionFallback": "旧Wordのセクション種別を解析できないため、本文中の改ページ文字だけを使用します。理由: {message}",
        "warning.legacyRemainder": "WordのFIB文字数範囲外にテキストがあり、末尾へ追加しました。",
        "warning.excelNonWorksheetExcluded": "Worksheet以外のsheetをテキスト抽出から除外しました: {name}",
        "warning.excelSstOutOfRange": "LabelSstが範囲外のSST indexを参照しています: {name}",
        "warning.excelShapeTextFailed": "Excel図形内テキストを読み取れませんでした（{name}）: {message}",
        "warning.powerPointCurrentUserMissing": "Current User streamがないため、最新編集状態の特定を省略しました。",
        "warning.powerPointHeaderTokenUnknown": "CurrentUserAtomのheader tokenが既知の値ではありません。",
        "warning.powerPointSlideOrderFallback": "最新Persist Directoryからスライド順を復元できないため、現行Document内の保存順を使用しました。",
        "warning.powerPointLegacyCodePage": "旧PowerPointの1バイト文字列は日本語コードページとして解析しました。他言語の文書では文字化けする場合があります。",
        "warning.officeArtMetafileHeader": "OfficeArtメタファイルヘッダーを解析できませんでした。",
        "warning.officeArtMetafileDecompression": "圧縮OfficeArtメタファイルを展開できませんでした。",
        "warning.officeArtMetafileSize": "OfficeArtメタファイルの保存サイズが一致しませんでした。",
        "warning.officeArtBlipLocation": "OfficeArt BLIPのメディア本体位置を特定できませんでした。",
        "warning.powerPointPicturesMissing": "PowerPoint Pictures streamがありません。",
        "warning.oleNativeFallback": "Ole10Nativeを元ファイルへ分離できないため、binaryで保存します。",
        "warning.oleRecordInstanceUnsupported": "ExOleObjStgのrecord instanceが未対応です。",
        "warning.oleExpandedSizeInvalid": "ExOleObjStgの展開サイズが不正なため読み飛ばしました。",
        "warning.oleStorageRestoreFailed": "PowerPointの圧縮OLE storageを復元できませんでした。",
        "warning.packageEntryExtractionFailed": "{subject}「{name}」を抽出できませんでした。\n理由: {message}",
        "warning.categoryExtractionFailed": "{subject}を抽出できませんでした: {message}",
        "error.cancelled": "処理を中止しました。",
        "error.emptyFile": "ファイルは空です。",
        "error.encrypted": "ファイルは暗号化またはパスワード保護されています。",
        "error.unsupported": "このファイル形式には対応していません。",
        "error.format": "拡張子と内部形式が一致しないか、ファイル構造が不正です。",
        "error.inputSize": "入力ファイルのサイズが安全上限の範囲外です。",
        "error.memory": "メモリの安全上限を超える可能性があります。入力ファイル数またはファイルサイズを減らして再実行してください。",
        "error.output": "ZIPの安全上限内で出力を作成できませんでした。",
        "error.zip": "ZIP構造が不正または未対応です。",
        "error.word": "Word文書構造が不正または未対応です。",
        "error.excel": "Excel文書構造が不正または未対応です。",
        "error.powerpoint": "PowerPoint文書構造が不正または未対応です。",
        "error.xml": "XMLパーツが大きすぎるか、安全でないか、不正です。",
        "error.folder": "安全上限内でフォルダを読み取れませんでした。",
        "error.general": "ファイルを処理できませんでした。",
        "error.range": "ブラウザで処理できるメモリ量を超えた可能性があります。入力ファイル数またはファイルサイズを減らして再実行してください。",
        "error.unexpected": "予期しない処理エラーが発生しました。",
        "error.unexpectedWithMessage": "予期しないエラーが発生しました: {message}",
        "note.local.heading": "外部送信は行いません",
        "note.local.body": "Officeファイルはサーバーへ送信されず、現在のブラウザ内だけで処理されます。",
        "note.caution.heading": "ご利用前の注意",
        "note.caution.body": "本ツールは元のOfficeファイルを書き換えませんが、すべてのファイルでの正常動作や抽出結果の完全性を保証するものではありません。重要なファイルはバックアップを保管したうえで、利用者自身の判断で使用してください。",
        "note.legacy.heading": "旧Office形式について",
        "note.legacy.body": "Office 97～2003形式にも対応しています。新しいOffice形式とは内部構造が異なるため、ファイルの内容や作成環境によっては、取得できる画像やテキストに違いが生じる場合があります。",
        "note.text.heading": "テキスト抽出の範囲",
        "note.text.body": "主要な本文、セル、図形内テキスト、表、発表者ノートを抽出します。配置情報をもとに見た目に近い順序へ整えますが、複雑なレイアウトでは画面上の順序と異なる場合があります。",
        "note.embedded.heading": "埋め込みデータについて",
        "note.embedded.body": "OLEや添付ファイルの内部データを抽出します。元のファイル形式へ復元できず、.binのまま出力される場合があります。抽出物は実行しないでください。",
        "note.cancel.heading": "処理中止について",
        "note.cancel.body": "大きなファイルや複雑な旧Officeファイルでは、中止操作が反映されるまで時間がかかる場合があります。",
        "footer.ariaLabel": "サイト情報",
        "footer.copyright": "© 2026 ラクジタワークス（RakuJita Works）",
        "footer.usageLink": "使い方・注意事項",
        "footer.termsLink": "利用条件",
        "footer.trademark": "Microsoft、Microsoft 365、Office、Word、ExcelおよびPowerPointは、Microsoftグループの企業の商標です。",
        "footer.disclaimer": "本ツールはMicrosoftによって提供、承認、後援されているものではありません。"
      },
      en: {
        "meta.description": "A local browser-based tool that extracts images, text, and embedded data from Word, Excel, and PowerPoint files and packages them into a ZIP file.",
        "language.label": "Display language",
        "hero.subtitle": "Extract media and text from Word, Excel, and PowerPoint files in bulk. Everything runs in your browser and is packaged into a single ZIP file.",
        "hero.localBadge": "Processed locally in your browser · No uploads",
        "input.heading": "Input files",
        "button.addFiles": "Add Files",
        "button.addFolder": "Add Folder",
        "button.removeSelected": "Remove Selected",
        "button.clearFiles": "Clear",
        "button.cancelProcessing": "Cancel Processing",
        "button.exportZip": "Export All as ZIP",
        "button.clearLog": "Clear Log",
        "drop.heading": "Drop files or folders here",
        "drop.anywhere": "Drop them anywhere on this page to add them",
        "drop.formats": "Supported formats: .doc / .docx / .docm / .xls / .xlsx / .xlsm / .xltx / .xltm / .ppt / .pptx / .pptm / .potx / .potm",
        "drop.ariaLabel": "Add Office files. Drop folders here or use the Add Folder button.",
        "option.includeSubfolders": "Include subfolders when adding a folder",
        "option.includeText": "Extract text",
        "option.includeEmbeddings": "Extract embedded data",
        "warning.browserCompatibility": "This browser cannot decompress ZIP-based OOXML files. Use the current stable version of Chrome or Edge. Legacy .doc, .xls, and .ppt files can still be inspected.",
        "table.selectAll": "Select all items for removal",
        "table.selectItem": "Select {name} for removal",
        "table.selectItemTitle": "Select for removal",
        "table.fileName": "File name",
        "table.relativePath": "Relative path",
        "table.extension": "Extension",
        "table.size": "Size",
        "table.status": "Status",
        "table.empty": "No Office files have been added.",
        "table.selectionHelp": "The checkboxes are only used to select items for removal. When exporting a ZIP, all processable files in the list are included.",
        "progress.label": "Processing progress",
        "progress.waiting": "Waiting",
        "progress.checkingInternalFormat": "Checking the internal format",
        "progress.extractingMedia": "Extracting media",
        "progress.extractingText": "Extracting text",
        "progress.extractingEmbeddings": "Extracting embedded data",
        "progress.addingToZip": "Adding files to the ZIP",
        "progress.readingFile": "{name}: Reading the file",
        "progress.fileStage": "{name}: {stage}",
        "progress.fileFinished": "{name}: Processing finished",
        "progress.completed": "Completed",
        "progress.generatingZip": "Generating the ZIP",
        "progress.cancelled": "Processing cancelled",
        "progress.error": "Processing ended with an error",
        "log.heading": "Processing Log",
        "log.level.info": "Info",
        "log.level.warning": "Warning",
        "log.level.error": "Error",
        "log.level.success": "Success",
        "log.addPrompt": "Add files to begin.",
        "log.cleared": "The log has been cleared.",
        "log.zeroByteSkipped": "The file was not added because it is empty: {name}",
        "log.fileTooLarge": "The file exceeds the input size limit: {name}",
        "log.totalTooLarge": "The file was not added because the total input size would exceed the safety limit: {name}",
        "log.filesAdded.one": "{count} file was added.",
        "log.filesAdded.other": "{count} files were added.",
        "log.subfoldersSkipped.one": "{count} file in a subfolder was skipped because the option is off.",
        "log.subfoldersSkipped.other": "{count} files in subfolders were skipped because the option is off.",
        "log.noFilesAdded": "No files could be added.",
        "log.noProcessableFiles": "There are no files ready for processing.",
        "log.fileWarning": "{name}: {warning}",
        "log.noExtractableContent": "No extractable media, text, or embedded data was found: {name}",
        "log.fileError": "{message}: {name}",
        "log.noZipResults": "There are no extracted results to save in a ZIP file.",
        "log.zipGenerated": "Generated {name} ({success} successful, {empty} with no results, {failed} failed).",
        "log.cancelled": "Processing was cancelled by the user.",
        "log.folderReadFailed.one": "{count} folder could not be read by drag and drop. Use Add Folder instead.",
        "log.folderReadFailed.other": "{count} folders could not be read by drag and drop. Use Add Folder instead.",
        "log.dropItemsFailed.one": "{count} dropped item could not be read. Use Add Files instead.",
        "log.dropItemsFailed.other": "{count} dropped items could not be read. Use Add Files instead.",
        "log.removed.one": "{count} item was removed from the list.",
        "log.removed.other": "{count} items were removed from the list.",
        "log.listCleared.one": "{count} item was cleared from the input list.",
        "log.listCleared.other": "{count} items were cleared from the input list.",
        "log.cancelRequested": "Cancellation was requested. Processing will stop at the next safe boundary.",
        "log.dropReadFailed": "The dropped items could not be read: {message}",
        "log.browserUnsupported": "This browser does not support the Deflate decompression API. Use the current stable version of Chrome or Edge.",
        "summary.result": "{name} ({media}, {characters}, {embeddings})",
        "summary.media.one": "{count} media item",
        "summary.media.other": "{count} media items",
        "summary.characters.one": "{count} character of text",
        "summary.characters.other": "{count} characters of text",
        "summary.embeddings.one": "{count} embedded item",
        "summary.embeddings.other": "{count} embedded items",
        "status.ready": "Ready",
        "status.checking": "Checking",
        "status.unsupported": "Unsupported format",
        "status.formatMismatch": "Format mismatch",
        "status.encrypted": "Possibly encrypted",
        "status.browserUnsupported": "Unsupported browser",
        "status.processing": "Processing",
        "status.completed": "Completed",
        "status.completedWithWarnings": "Completed with warnings",
        "status.noResults": "No results",
        "status.error": "Error",
        "status.processingFailed": "Error",
        "status.cancelled": "Cancelled",
        "status.detail.unsupported": "This file extension is not supported.",
        "status.detail.formatMismatchOoxml": "The file extension does not match the internal OOXML format.",
        "status.detail.formatMismatchCfb": "The file extension does not match the internal CFB format.",
        "status.detail.missingSignature": "The file does not have a ZIP or CFB signature.",
        "status.detail.ooxml": "OOXML {family}",
        "status.detail.cfb": "CFB {family}",
        "status.detail.encryptedOoxml": "This is an encrypted OOXML package.",
        "status.detail.encryptedLegacy": "The legacy Office binary contains encryption information.",
        "status.detail.browserUnsupported": "This browser cannot decompress ZIP-based OOXML files.",
        "status.detail.checking": "Checking the internal format.",
        "status.detail.processing": "Extracting content from the file.",
        "status.detail.noResults": "Analysis completed, but there was no content to extract.",
        "status.detail.completed": "Extraction completed.",
        "status.detail.completedWithWarnings": "Some extraction steps produced warnings.",
        "status.detail.cancelled": "Processing was cancelled by the user.",
        "message.value": "{message}",
        "warning.textOmitted": "Text output was omitted because the extracted text exceeded the safety limit.",
        "warning.noText": "No usable document text was found.",
        "warning.media": "Some media could not be extracted.",
        "warning.text": "Some text could not be extracted.",
        "warning.embedded": "Some embedded data could not be extracted.",
        "warning.word": "Some Word document structures could not be fully reconstructed.",
        "warning.excel": "Some Excel document structures could not be fully reconstructed.",
        "warning.powerpoint": "Some PowerPoint document structures could not be fully reconstructed.",
        "warning.generic": "Some document content could not be fully processed.",
        "warning.subject.wordChart": "Word chart",
        "warning.subject.altChunk": "altChunk",
        "warning.subject.footnote": "footnote",
        "warning.subject.endnote": "endnote",
        "warning.subject.legacyFootnote": "legacy Word footnote",
        "warning.subject.legacyEndnote": "legacy Word endnote",
        "warning.subject.mainTextbox": "body text box",
        "warning.subject.headerFooterTextbox": "header/footer text box",
        "warning.subject.mainShapeAnchor": "body shape-anchor character",
        "warning.subject.headerFooterShapeAnchor": "header/footer shape-anchor character",
        "warning.subject.media": "media",
        "warning.subject.text": "text",
        "warning.subject.embedding": "embedded data",
        "warning.document.word": "Word",
        "warning.document.legacyWord": "legacy Word",
        "warning.value.empty": "(empty)",
        "warning.value.noContentType": "no Content Type",
        "warning.section.unplacedMainTextbox": "Text Boxes Whose Position Could Not Be Reconstructed",
        "warning.section.unplacedHeaderFooterTextbox": "Header/Footer Text Boxes Whose Position Could Not Be Reconstructed",
        "warning.referenceMissingId": "The {subject} reference has no relationship ID and was skipped.",
        "warning.referenceRelationshipMissing": "The relationship for {subject} reference {id} was not found.",
        "warning.externalReferenceSkipped": "External {subject} reference {id} was skipped for safety.",
        "warning.referenceKindMismatch": "The relationship type for {subject} reference {id} did not match the expected type.",
        "warning.partMissing": "The {subject} part was not found: {name}",
        "warning.unsupportedAltChunk": "An unsupported altChunk format was skipped: {type}",
        "warning.partParseFailed": "The {subject} part could not be parsed and was skipped: {name}",
        "warning.annotationDefinitionMissingId": "A {subject} definition had no ID and was skipped.",
        "warning.annotationDuplicateId": "The {subject} ID {id} is duplicated; the first definition was used.",
        "warning.annotationBodyMissing": "No body text was found for {subject} reference ID {id}.",
        "warning.wordNoPageBreak": "The {document} document has no usable page-break information, so its footnotes were placed as if the entire document were one logical page.",
        "warning.unreferencedAnnotation.one": "There is {count} unreferenced {subject} entry.",
        "warning.unreferencedAnnotation.other": "There are {count} unreferenced {subject} entries.",
        "warning.legacyAnnotationEmptyStory": "The legacy Word {subject} reference PLC could not be parsed because no annotation body text was present.",
        "warning.legacyAnnotationCountMismatch": "The legacy Word {subject} reference count did not match the body count; only matching entries were placed.",
        "warning.legacyAnnotationPositionFallback": "The positions of legacy Word {subject} references could not be reconstructed, so they were placed in a separate section. Reason: {message}",
        "warning.legacyTextboxPlcFallback": "The body PLC for the legacy Word {subject} could not be parsed, so its text will be placed in the \"{section}\" section. Reason: {message}",
        "warning.legacyTextboxInvalidRange.one": "The legacy Word {subject} contains {count} invalid text range, so recoverable text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxInvalidRange.other": "The legacy Word {subject} contains {count} invalid text ranges, so recoverable text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMissingParagraphEnd.one": "The legacy Word {subject} contains {count} definition without a paragraph terminator, so recoverable text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMissingParagraphEnd.other": "The legacy Word {subject} contains {count} definitions without paragraph terminators, so recoverable text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMetadataInvalid.one": "The legacy Word {subject} contains {count} invalid {field} value, so affected text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMetadataInvalid.other": "The legacy Word {subject} contains {count} invalid {field} values, so affected text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxUnsupportedFlags.one": "The legacy Word {subject} contains {count} unsupported FTXBXS flag value.",
        "warning.legacyTextboxUnsupportedFlags.other": "The legacy Word {subject} contains {count} unsupported FTXBXS flag values.",
        "warning.legacyTextboxReuseManagement": "Inconsistent reuse-management information was found for the legacy Word {subject}.\nReusable definitions were not written to the document body, and extraction of ordinary text boxes continued.",
        "warning.legacyTextboxTextRangeMismatch": "The legacy Word {subject} body count did not match its PLC range, so unmatched text will be placed in the \"{section}\" section.",
        "warning.legacyTextboxAnchorPlcFallback": "The shape-anchor PLC for the legacy Word {subject} could not be parsed, so its text will be placed in the \"{section}\" section. Reason: {message}",
        "warning.legacyTextboxInvalidAnchorMarker.one": "The legacy Word document contains {count} invalid {subject}; the corresponding text box was placed in the \"{section}\" section.",
        "warning.legacyTextboxInvalidAnchorMarker.other": "The legacy Word document contains {count} invalid {subject} values; the corresponding text boxes were placed in the \"{section}\" section.",
        "warning.legacyTextboxDuplicateAnchor.one": "The shape anchors for the legacy Word {subject} contain {count} duplicate shape ID; the first anchor was used.",
        "warning.legacyTextboxDuplicateAnchor.other": "The shape anchors for the legacy Word {subject} contain {count} duplicate shape IDs; the first anchor for each ID was used.",
        "warning.legacyTextboxMissingAnchor.one": "{count} legacy Word {subject} could not be matched to a shape anchor and will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMissingAnchor.other": "{count} legacy Word {subject} entries could not be matched to shape anchors and will be placed in the \"{section}\" section.",
        "warning.legacyTextboxMissingCoordinate.one": "Coordinates could not be obtained for {count} legacy Word {subject} shape anchor. Placement continued using CP and shape ID.",
        "warning.legacyTextboxMissingCoordinate.other": "Coordinates could not be obtained for {count} legacy Word {subject} shape anchors. Placement continued using CP and shape ID.",
        "warning.legacyTextboxUnplaced.one": "{count} legacy Word {subject} could not be placed at its shape anchor and was placed in the \"{section}\" section.",
        "warning.legacyTextboxUnplaced.other": "{count} legacy Word {subject} entries could not be placed at their shape anchors and were placed in the \"{section}\" section.",
        "warning.legacyTextboxUnplacedUnknown": "The position of a legacy Word {subject} could not be reconstructed, so it was placed in the \"{section}\" section.",
        "warning.legacyPapxRangeMismatch.one": "The legacy Word paragraph-format page contained {count} inconsistency with its BTE range, so only in-range paragraphs were used.",
        "warning.legacyPapxRangeMismatch.other": "The legacy Word paragraph-format pages contained {count} inconsistencies with their BTE ranges, so only in-range paragraphs were used.",
        "warning.legacyPapxFallback": "Legacy Word page-break-before information could not be parsed, so only manual page breaks and section breaks were used.",
        "warning.legacyPapxFallbackWithReason": "Legacy Word page-break-before information could not be parsed, so only manual page breaks and section breaks were used. Reason: {message}",
        "warning.legacyPapxPartial": "Some legacy Word page-break-before information could not be parsed, so only the available page-break information was used.",
        "warning.legacyAnnotationIndexMissing": "No body text was found for legacy Word {subject} reference index {id}.",
        "warning.legacySectionBoundaryApproximation": "A legacy Word section boundary had no terminating character, so the page position was approximated from the available section type.",
        "warning.legacyReferenceCharacter": "A legacy Word automatic annotation reference had no reference control character, so the original body character was preserved.",
        "warning.legacySectionFallback": "Legacy Word section types could not be parsed, so only page-break characters in the document body were used. Reason: {message}",
        "warning.legacyRemainder": "Text was found beyond the Word FIB character-count range and was appended to the end.",
        "warning.excelNonWorksheetExcluded": "A non-worksheet sheet was excluded from text extraction: {name}",
        "warning.excelSstOutOfRange": "A LabelSst record referenced an out-of-range SST index: {name}",
        "warning.excelShapeTextFailed": "Text inside an Excel shape could not be read ({name}): {message}",
        "warning.powerPointCurrentUserMissing": "The latest editing state could not be determined because the Current User stream was missing.",
        "warning.powerPointHeaderTokenUnknown": "The CurrentUserAtom header token was not a known value.",
        "warning.powerPointSlideOrderFallback": "The slide order could not be reconstructed from the latest Persist Directory, so the saved order in the current Document was used.",
        "warning.powerPointLegacyCodePage": "Single-byte strings in the legacy PowerPoint file were decoded using the Japanese code page. Text may be garbled in documents created for other languages.",
        "warning.officeArtMetafileHeader": "The OfficeArt metafile header could not be parsed.",
        "warning.officeArtMetafileDecompression": "The compressed OfficeArt metafile could not be decompressed.",
        "warning.officeArtMetafileSize": "The stored OfficeArt metafile size did not match the expected size.",
        "warning.officeArtBlipLocation": "The media payload position in the OfficeArt BLIP could not be determined.",
        "warning.powerPointPicturesMissing": "The PowerPoint Pictures stream was missing.",
        "warning.oleNativeFallback": "The Ole10Native payload could not be separated into its original file and was saved as binary data.",
        "warning.oleRecordInstanceUnsupported": "The ExOleObjStg record instance is unsupported.",
        "warning.oleExpandedSizeInvalid": "The ExOleObjStg expanded size was invalid, so the record was skipped.",
        "warning.oleStorageRestoreFailed": "The compressed PowerPoint OLE storage could not be reconstructed.",
        "warning.packageEntryExtractionFailed": "The {subject} \"{name}\" could not be extracted.\nReason: {message}",
        "warning.categoryExtractionFailed": "The {subject} could not be extracted: {message}",
        "error.cancelled": "Processing was cancelled.",
        "error.emptyFile": "The file is empty.",
        "error.encrypted": "The file is encrypted or password-protected.",
        "error.unsupported": "This file format is not supported.",
        "error.format": "The file extension and internal format do not match, or the file structure is invalid.",
        "error.inputSize": "The input file size is outside the safety limit.",
        "error.memory": "The operation would exceed a memory safety limit. Reduce the number or size of the input files and try again.",
        "error.output": "The output could not be created within the ZIP safety limits.",
        "error.zip": "The ZIP structure is invalid or unsupported.",
        "error.word": "The Word document structure is invalid or unsupported.",
        "error.excel": "The Excel document structure is invalid or unsupported.",
        "error.powerpoint": "The PowerPoint document structure is invalid or unsupported.",
        "error.xml": "An XML part is too large, unsafe, or invalid.",
        "error.folder": "The folder could not be read within the safety limits.",
        "error.general": "The file could not be processed.",
        "error.range": "The operation may exceed the memory available to this browser. Reduce the number or size of the input files and try again.",
        "error.unexpected": "An unexpected processing error occurred.",
        "error.unexpectedWithMessage": "An unexpected error occurred: {message}",
        "note.local.heading": "No files are uploaded",
        "note.local.body": "Office files are not sent to a server. All processing takes place in the current browser.",
        "note.caution.heading": "Before using this tool",
        "note.caution.body": "This tool does not modify the original Office files. However, correct operation and complete extraction cannot be guaranteed for every file. Keep backups of important files and use the tool at your own discretion.",
        "note.legacy.heading": "Legacy Office formats",
        "note.legacy.body": "Office 97–2003 formats are also supported. Because their internal structure differs from newer Office formats, the images and text that can be extracted may vary depending on the file contents and the environment in which the file was created.",
        "note.text.heading": "Text extraction coverage",
        "note.text.body": "The tool extracts primary document text, cells, text inside shapes, tables, and presenter notes. It uses available position information to approximate the visual order, but complex layouts may produce a different order from what appears on screen.",
        "note.embedded.heading": "Embedded data",
        "note.embedded.body": "The tool extracts internal data from OLE objects and attached files. Some data cannot be restored to its original file format and may be exported as a .bin file. Do not execute extracted files.",
        "note.cancel.heading": "Cancelling processing",
        "note.cancel.body": "For large files or complex legacy Office files, it may take some time for a cancellation request to take effect.",
        "footer.ariaLabel": "Site information",
        "footer.copyright": "© 2026 RakuJita Works",
        "footer.usageLink": "Usage & Notes (Japanese)",
        "footer.termsLink": "Terms of Use (Japanese)",
        "footer.trademark": "Microsoft, Microsoft 365, Office, Word, Excel, and PowerPoint are trademarks of the Microsoft group of companies.",
        "footer.disclaimer": "This tool is not provided, approved, or sponsored by Microsoft."
      }
    };
    var missingTranslationWarnings = Object.create(null);

    function detectInitialLanguage() {
      var language = "";
      try {
        if (navigator.languages && navigator.languages.length > 0) {
          language = navigator.languages[0];
        } else {
          language = navigator.language || navigator.userLanguage || "";
        }
      } catch (error) {
        language = "";
      }
      language = String(language).toLowerCase();
      return language === "ja" || language.indexOf("ja-") === 0 ? "ja" : "en";
    }

    var currentLanguage = detectInitialLanguage();

    function localeForLanguage(language) {
      return language === "ja" ? "ja-JP" : "en-US";
    }

    function translationValue(key, params) {
      return { translationKey: key, translationParams: params || {} };
    }

    function pluralValue(baseKey, count) {
      return { pluralBaseKey: baseKey, count: count };
    }

    function warningValue(key, params, technicalMessage, pluralCount) {
      var rawTechnicalMessage = "";
      var technicalCode = "";
      var technicalMessageKey = "";
      var technicalName = "";
      if (technicalMessage instanceof AppError) {
        rawTechnicalMessage = technicalMessage.rawMessage;
        technicalCode = technicalMessage.code;
        technicalMessageKey = technicalMessage.messageKey;
        technicalName = technicalMessage.name;
      } else if (technicalMessage instanceof Error) {
        rawTechnicalMessage = technicalMessage.message || String(technicalMessage);
        technicalName = technicalMessage.name || "Error";
      } else {
        rawTechnicalMessage = String(technicalMessage || "");
      }
      var warning = {
        key: key,
        params: params || {}
      };
      if (rawTechnicalMessage) {
        warning.technicalMessage = rawTechnicalMessage;
        if (technicalCode) {
          warning.technicalCode = technicalCode;
        }
        if (technicalMessageKey) {
          warning.technicalMessageKey = technicalMessageKey;
        }
        if (technicalName) {
          warning.technicalName = technicalName;
        }
      }
      if (typeof pluralCount === "number") {
        warning.pluralCount = pluralCount;
      }
      return warning;
    }

    function runtimeWarningValue(warning) {
      return { runtimeWarning: warning };
    }

    function numberValue(value) {
      return { localizedNumber: Number(value || 0) };
    }

    function translationParameter(value, language) {
      if (value && value.translationKey) {
        return t(value.translationKey, value.translationParams, language);
      }
      if (value && value.pluralBaseKey) {
        return t(pluralKey(value.pluralBaseKey, value.count, language), {
          count: formatNumber(value.count, language)
        }, language);
      }
      if (value && Object.prototype.hasOwnProperty.call(value, "runtimeWarning")) {
        return localizeRuntimeWarning(value.runtimeWarning, language);
      }
      if (value && Object.prototype.hasOwnProperty.call(value, "localizedNumber")) {
        return formatNumber(value.localizedNumber, language);
      }
      if (value instanceof AppError || value instanceof Error) {
        return friendlyError(value, language);
      }
      return String(value === null || typeof value === "undefined" ? "" : value);
    }

    function t(key, params, language) {
      var selectedLanguage = language === "ja" || language === "en" ?
        language : currentLanguage;
      var selected = TRANSLATIONS[selectedLanguage] || TRANSLATIONS.en;
      var value = selected[key];
      if (typeof value !== "string") {
        value = TRANSLATIONS.ja[key];
        if (!missingTranslationWarnings[selectedLanguage + ":" + key]) {
          missingTranslationWarnings[selectedLanguage + ":" + key] = true;
          if (typeof console !== "undefined" && console.warn) {
            console.warn("Missing translation: " + selectedLanguage + ":" + key);
          }
        }
      }
      if (typeof value !== "string") {
        value = key;
      }
      var output = value;
      var replacements = params || {};
      var name;
      for (name in replacements) {
        if (Object.prototype.hasOwnProperty.call(replacements, name)) {
          output = output.split("{" + name + "}").join(
            translationParameter(replacements[name], selectedLanguage)
          );
        }
      }
      var unresolved = output.match(/\{[A-Za-z][A-Za-z0-9_]*\}/g) || [];
      var unresolvedIndex;
      for (unresolvedIndex = 0;
        unresolvedIndex < unresolved.length;
        unresolvedIndex += 1) {
        var warningId = selectedLanguage + ":" + key + ":" +
          unresolved[unresolvedIndex];
        if (!missingTranslationWarnings[warningId]) {
          missingTranslationWarnings[warningId] = true;
          if (typeof console !== "undefined" && console.warn) {
            console.warn(
              "Missing translation parameter: " + warningId
            );
          }
        }
      }
      return output;
    }

    function pluralKey(baseKey, count, language) {
      return language === "en" && Number(count) === 1 ?
        baseKey + ".one" : baseKey + ".other";
    }

    function formatNumber(value, language) {
      return Number(value || 0).toLocaleString(localeForLanguage(language || currentLanguage));
    }

    function errorKeyForCode(code) {
      var value = String(code || "");
      if (value === "CANCELLED") { return "error.cancelled"; }
      if (value === "EMPTY_FILE") { return "error.emptyFile"; }
      if (value === "ENCRYPTED" || value === "ZIP_ENCRYPTED") { return "error.encrypted"; }
      if (value === "UNSUPPORTED") { return "error.unsupported"; }
      if (value === "FORMAT_MISMATCH" || value === "FORMAT_SIGNATURE") { return "error.format"; }
      if (value === "INPUT_SIZE") { return "error.inputSize"; }
      if (value === "MEMORY" || value === "ESTIMATED_MEMORY") { return "error.memory"; }
      if (value.indexOf("OUTPUT_") === 0 || value === "NAME_COLLISION") { return "error.output"; }
      if (value.indexOf("ZIP") === 0 || value === "DEFLATE" || value.indexOf("DEFLATE_") === 0) { return "error.zip"; }
      if (value.indexOf("WORD_") === 0) { return "error.word"; }
      if (value.indexOf("BIFF_") === 0 || value.indexOf("EXCEL_") === 0) { return "error.excel"; }
      if (value.indexOf("PPT_") === 0) { return "error.powerpoint"; }
      if (value.indexOf("XML_") === 0) { return "error.xml"; }
      if (value.indexOf("FOLDER_") === 0) { return "error.folder"; }
      return "error.general";
    }

    function AppError(code, message) {
      this.name = "AppError";
      this.code = code || "GENERAL";
      this.rawMessage = message || "処理中にエラーが発生しました。";
      this.messageKey = errorKeyForCode(this.code);
      this.messageParams = { message: this.rawMessage };
      this.message = this.rawMessage;
      if (Error.captureStackTrace) {
        Error.captureStackTrace(this, AppError);
      }
    }
    AppError.prototype = Object.create(Error.prototype);
    AppError.prototype.constructor = AppError;

    function fail(code, message) {
      throw new AppError(code, message);
    }

    function ensure(condition, code, message) {
      if (!condition) {
        fail(code, message);
      }
    }

    function checkCancelled() {
      ensure(
        !appState.cancelRequested,
        "CANCELLED",
        "処理が中止されました。"
      );
    }

    function throwIfCancelled(error) {
      if (error instanceof AppError && error.code === "CANCELLED") {
        throw error;
      }
    }

    function isProcessingSafetyError(error) {
      return !!(
        (error instanceof AppError &&
          (error.code === "MEDIA_LIMIT" ||
           error.code === "EMBEDDING_LIMIT" ||
           error.code === "OUTPUT_SIZE" ||
           error.code === "OUTPUT_ZIP64" ||
           error.code === "MEMORY" ||
           error.code === "ESTIMATED_MEMORY" ||
           error.code === "ZIP_ENTRY_TOO_LARGE" ||
           error.code === "ZIP_TOTAL_TOO_LARGE" ||
           error.code === "ZIP_RATIO" ||
           error.code === "ZIP_ENTRIES" ||
           error.code === "ZIP64" ||
           error.code === "EXCEL_TEXT_RANGE" ||
           error.code === "NAME_COLLISION")) ||
        (error && error.name === "RangeError")
      );
    }

    function throwIfCategoryLimitOrCancelled(error) {
      throwIfCancelled(error);
      if (isProcessingSafetyError(error)) {
        throw error;
      }
    }

    function isRangeValid(length, offset, size) {
      return Number.isFinite(offset) &&
        Number.isFinite(size) &&
        offset >= 0 &&
        size >= 0 &&
        offset <= length &&
        size <= length - offset;
    }

    function requireRange(bytes, offset, size, label) {
      ensure(
        isRangeValid(bytes.length, offset, size),
        "BOUNDS",
        (label || "バイナリデータ") + "がファイル範囲外を参照しています。"
      );
    }

    function viewOf(bytes) {
      return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    }

    function u16(bytes, offset) {
      requireRange(bytes, offset, 2, "16ビット値");
      return viewOf(bytes).getUint16(offset, true);
    }

    function u32(bytes, offset) {
      requireRange(bytes, offset, 4, "32ビット値");
      return viewOf(bytes).getUint32(offset, true);
    }

    function i32(bytes, offset) {
      requireRange(bytes, offset, 4, "32ビット値");
      return viewOf(bytes).getInt32(offset, true);
    }

    function f64(bytes, offset) {
      requireRange(bytes, offset, 8, "浮動小数点値");
      return viewOf(bytes).getFloat64(offset, true);
    }

    function sliceBytes(bytes, offset, size, label) {
      requireRange(bytes, offset, size, label);
      return bytes.subarray(offset, offset + size);
    }

    function concatBytes(parts, expectedLength) {
      var total = 0;
      var index;
      for (index = 0; index < parts.length; index += 1) {
        total += parts[index].length;
        ensure(total <= 0x7FFFFFFF, "MEMORY", "連結するデータが大きすぎます。");
      }
      if (typeof expectedLength === "number") {
        total = Math.min(total, expectedLength);
      }
      var result = new Uint8Array(total);
      var position = 0;
      for (index = 0; index < parts.length && position < total; index += 1) {
        var count = Math.min(parts[index].length, total - position);
        result.set(parts[index].subarray(0, count), position);
        position += count;
      }
      return result;
    }

    function startsWithBytes(bytes, signature) {
      if (bytes.length < signature.length) {
        return false;
      }
      var index;
      for (index = 0; index < signature.length; index += 1) {
        if (bytes[index] !== signature[index]) {
          return false;
        }
      }
      return true;
    }

    function bytesEqual(left, right) {
      if (left.length !== right.length) {
        return false;
      }
      var index;
      for (index = 0; index < left.length; index += 1) {
        if (left[index] !== right[index]) {
          return false;
        }
      }
      return true;
    }

    function makeDecoder(label, fatal) {
      try {
        return new TextDecoder(label, { fatal: !!fatal });
      } catch (error) {
        return new TextDecoder("windows-1252", { fatal: false });
      }
    }

    function decodeUtf8(bytes, fatal) {
      return makeDecoder("utf-8", fatal).decode(bytes);
    }

    function decodeUtf16Le(bytes) {
      var evenLength = bytes.length - (bytes.length % 2);
      return makeDecoder("utf-16le", false).decode(bytes.subarray(0, evenLength));
    }

    function decodeSingleByte(bytes, codePage) {
      var label = "windows-1252";
      if (codePage === 932 || codePage === 943 || codePage === 128) {
        label = "shift_jis";
      } else if (codePage === 936) {
        label = "gbk";
      } else if (codePage === 949) {
        label = "euc-kr";
      } else if (codePage === 950) {
        label = "big5";
      } else if (codePage === 65001) {
        label = "utf-8";
      }
      return makeDecoder(label, false).decode(bytes);
    }

    function decodeCompressedUnicode(bytes) {
      var chunks = [];
      var blockSize = 8192;
      var offset;
      for (offset = 0; offset < bytes.length; offset += blockSize) {
        var end = Math.min(bytes.length, offset + blockSize);
        var chars = [];
        var index;
        for (index = offset; index < end; index += 1) {
          chars.push(bytes[index]);
        }
        chunks.push(String.fromCharCode.apply(null, chars));
      }
      return chunks.join("");
    }

    function encodeUtf8(text) {
      return new TextEncoder().encode(text);
    }

    function repairSurrogates(text) {
      var result = "";
      var index;
      for (index = 0; index < text.length; index += 1) {
        var code = text.charCodeAt(index);
        if (code >= 0xD800 && code <= 0xDBFF) {
          if (index + 1 < text.length) {
            var next = text.charCodeAt(index + 1);
            if (next >= 0xDC00 && next <= 0xDFFF) {
              result += text.charAt(index) + text.charAt(index + 1);
              index += 1;
              continue;
            }
          }
          result += "\uFFFD";
        } else if (code >= 0xDC00 && code <= 0xDFFF) {
          result += "\uFFFD";
        } else {
          result += text.charAt(index);
        }
      }
      return result;
    }

    function normalizeOutputNameCharacters(value) {
      var normalized = repairSurrogates(String(value || ""));
      if (typeof String.prototype.normalize === "function") {
        try {
          normalized = normalized.normalize("NFKC");
        } catch (error) {
          // Explicit replacements and filename safety still run without NFKC.
        }
      }
      var replacementIndex;
      for (replacementIndex = 0;
        replacementIndex < OUTPUT_NAME_CHARACTER_REPLACEMENTS.length;
        replacementIndex += 1) {
        var replacement = OUTPUT_NAME_CHARACTER_REPLACEMENTS[replacementIndex];
        normalized = normalized.split(replacement[0]).join(replacement[1]);
      }
      return normalized.replace(/[ \u3000]+/g, "_");
    }

    function normalizeText(text) {
      var repaired = repairSurrogates(String(text || ""))
        .replace(/\u0000/g, "")
        .replace(/\r\n?/g, "\n")
        .replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "");
      var lines = repaired.split("\n");
      var normalized = [];
      var blankCount = 0;
      var index;
      for (index = 0; index < lines.length; index += 1) {
        var line = lines[index].replace(/[ \u3000]+$/g, "");
        if (line.length === 0) {
          blankCount += 1;
          if (blankCount <= 2) {
            normalized.push("");
          }
        } else {
          blankCount = 0;
          normalized.push(line);
        }
      }
      while (normalized.length && normalized[0] === "") {
        normalized.shift();
      }
      while (normalized.length && normalized[normalized.length - 1] === "") {
        normalized.pop();
      }
      return normalized.join("\r\n");
    }

    function hasMeaningfulText(text) {
      return String(text || "").replace(/[\s\uFEFF]/g, "").length > 0;
    }

    function normalizeTextForDeduplication(text) {
      var repaired = repairSurrogates(String(text || ""))
        .replace(/\r\n?/g, "\n")
        .replace(/\u0000/g, "")
        .replace(/[\u0001-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "");
      var sourceLines = repaired.split("\n");
      var lines = [];
      var previousBlank = true;
      var index;
      for (index = 0; index < sourceLines.length; index += 1) {
        var line = sourceLines[index].replace(/[ \t\u3000]+$/g, "");
        var blank = line.length === 0;
        if (!blank || !previousBlank) {
          lines.push(line);
        }
        previousBlank = blank;
      }
      while (lines.length && lines[0] === "") {
        lines.shift();
      }
      while (lines.length && lines[lines.length - 1] === "") {
        lines.pop();
      }
      return lines.join("\n");
    }

    function deduplicateTextBlocks(blocks) {
      var output = [];
      var seen = Object.create(null);
      var index;
      for (index = 0; index < blocks.length; index += 1) {
        var normalized = normalizeTextForDeduplication(blocks[index]);
        if (!hasMeaningfulText(normalized)) {
          continue;
        }
        var key = normalized;
        if (!seen[key]) {
          seen[key] = true;
          output.push(normalized);
        }
      }
      return output;
    }

    function textToOutputBytes(text) {
      var normalized = normalizeText(text);
      if (!hasMeaningfulText(normalized)) {
        return null;
      }
      var bytes = encodeUtf8(normalized);
      if (bytes.length > MAX_TEXT_OUTPUT_BYTES) {
        return {
          omitted: true,
          warning: warningValue("warning.textOmitted")
        };
      }
      return {
        text: normalized,
        bytes: bytes,
        characterCount: normalized.length
      };
    }

    function parseXml(bytes, partName) {
      ensure(bytes.length <= MAX_XML_BYTES, "XML_SIZE", "XMLパーツが大きすぎます: " + partName);
      var source = decodeUtf8(bytes, false);
      ensure(!/<!\s*(?:DOCTYPE|ENTITY)\b/i.test(source),
        "XML_UNSAFE_DECLARATION",
        "安全でないXML宣言が含まれています: " + partName);
      var documentNode = new DOMParser().parseFromString(source, "application/xml");
      var errors = documentNode.getElementsByTagName("parsererror");
      ensure(errors.length === 0, "XML_PARSE", "XMLを解析できません: " + partName);
      return documentNode;
    }

    function elementsByLocalName(root, name) {
      var all = root.getElementsByTagName("*");
      var result = [];
      var index;
      for (index = 0; index < all.length; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        if (all[index].localName === name) {
          result.push(all[index]);
        }
      }
      return result;
    }

    function firstElementByLocalName(root, name) {
      var elements = elementsByLocalName(root, name);
      return elements.length ? elements[0] : null;
    }

    function attributeByLocalName(element, name) {
      if (!element || !element.attributes) {
        return null;
      }
      var index;
      for (index = 0; index < element.attributes.length; index += 1) {
        if (element.attributes[index].localName === name) {
          return element.attributes[index].value;
        }
      }
      return null;
    }

    function relationshipIdAttribute(element) {
      if (!element || !element.attributes) {
        return null;
      }
      var fallback = null;
      var index;
      for (index = 0; index < element.attributes.length; index += 1) {
        var attribute = element.attributes[index];
        if (attribute.localName !== "id") {
          continue;
        }
        if (String(attribute.namespaceURI || "").toLowerCase().indexOf(
          "relationships"
        ) >= 0) {
          return attribute.value;
        }
        if (attribute.prefix === "r") {
          return attribute.value;
        }
        fallback = attribute.value;
      }
      return fallback;
    }

    function extensionOf(name) {
      var match = /\.([^.\\\/]+)$/.exec(String(name || ""));
      return match ? match[1].toLowerCase() : "";
    }

    function removeExtension(name) {
      return String(name || "").replace(/\.[^.]*$/, "");
    }

    function leafName(path) {
      var normalized = String(path || "").replace(/\\/g, "/");
      var index = normalized.lastIndexOf("/");
      return index >= 0 ? normalized.slice(index + 1) : normalized;
    }

    function parentPath(path) {
      var normalized = String(path || "").replace(/\\/g, "/");
      var index = normalized.lastIndexOf("/");
      return index >= 0 ? normalized.slice(0, index) : "";
    }

    function truncateUnicode(text, maximum) {
      var value = String(text || "");
      if (value.length <= maximum) {
        return value;
      }
      var result = value.slice(0, maximum);
      var last = result.charCodeAt(result.length - 1);
      if (last >= 0xD800 && last <= 0xDBFF) {
        result = result.slice(0, -1);
      }
      return result;
    }

    function zeroPad(value, width) {
      var text = String(value);
      while (text.length < width) {
        text = "0" + text;
      }
      return text;
    }

    function normalizedOutputLimit(maximum, fallback) {
      var value = Number(maximum);
      if (!Number.isFinite(value) || value < 1) {
        value = fallback;
      }
      return Math.max(1, Math.floor(value));
    }

    function sanitizeStem(value, fallback, maximum) {
      var limit = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
      var name = repairSurrogates(String(value || ""))
        .replace(/[\u0000-\u001F\u007F\u202A-\u202E\u2066-\u2069<>:"/\\|?*]/g, "_")
        .replace(/[. ]+$/g, "")
        .replace(/^[ ]+/g, "");
      if (!name) {
        name = fallback || "unnamed";
      }
      if (/^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)/i.test(name)) {
        name = "_" + name;
      }
      return truncateUnicode(name, limit);
    }

    function sanitizeFileName(value, fallback, maximum) {
      var limit = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
      var raw = repairSurrogates(leafName(value));
      var extension = "";
      var stem = raw;
      var dot = raw.lastIndexOf(".");
      if (dot > 0 && dot < raw.length - 1) {
        extension = raw.slice(dot + 1)
          .replace(/[\u0000-\u001F\u007F\u202A-\u202E\u2066-\u2069<>:"/\\|?*]/g, "_")
          .replace(/[. ]+$/g, "");
        stem = raw.slice(0, dot);
      }
      ensure(extension.length <= Math.max(0, limit - 2), "OUTPUT_PATH",
        "出力ファイルの拡張子を保持したまま短縮できませんでした。");
      stem = sanitizeStem(stem, fallback || "file", limit);
      var available = limit - (extension ? extension.length + 1 : 0);
      stem = truncateUnicode(stem, Math.max(1, available));
      return extension ? stem + "." + extension : stem;
    }

    function buildSourceOutputStem(fileName) {
      return sanitizeStem(
        normalizeOutputNameCharacters(removeExtension(fileName)),
        "document",
        MAX_SOURCE_OUTPUT_STEM
      );
    }

    function mediaStemParts(mediaName) {
      var stem = mediaName;
      var extension = "";
      var dot = mediaName.lastIndexOf(".");
      if (dot > 0) {
        stem = mediaName.slice(0, dot);
        extension = mediaName.slice(dot);
      }
      var sequence = "";
      var sequenceMatch = /(_[0-9]{4,})$/.exec(stem);
      if (sequenceMatch && sequenceMatch.index > 0) {
        sequence = sequenceMatch[1];
        stem = stem.slice(0, sequenceMatch.index);
      }
      return {
        stem: stem || "media",
        sequence: sequence,
        extension: extension
      };
    }

    function buildPrefixedMediaFileName(
      sourceStem,
      originalMediaName,
      collisionSuffix,
      maximum
    ) {
      var limit = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
      var sourcePrefix = truncateUnicode(
        sanitizeStem(sourceStem, "document", MAX_SOURCE_OUTPUT_STEM),
        MAX_MEDIA_SOURCE_PREFIX
      );
      var mediaName = sanitizeFileName(
        normalizeOutputNameCharacters(leafName(originalMediaName)),
        "media",
        MAX_OUTPUT_FILE_NAME
      );
      var parts = mediaStemParts(mediaName);
      var serial = parts.sequence + String(collisionSuffix || "");
      var extension = parts.extension;
      var maximumExtensionLength = Math.max(
        0,
        limit - serial.length - 3
      );
      ensure(extension.length <= maximumExtensionLength, "OUTPUT_PATH",
        "メディアファイルの拡張子を保持したまま短縮できませんでした。");
      var minimumMediaLength = Math.min(
        MIN_MEDIA_FILE_STEM,
        parts.stem.length
      );
      var prefixLimit = limit - 1 - minimumMediaLength -
        serial.length - extension.length;
      sourcePrefix = truncateUnicode(sourcePrefix, Math.max(1, prefixLimit));
      var available = limit - sourcePrefix.length - 1 -
        serial.length - extension.length;
      if (available < minimumMediaLength && sourcePrefix.length > 1) {
        sourcePrefix = truncateUnicode(
          sourcePrefix,
          Math.max(1, sourcePrefix.length - (minimumMediaLength - available))
        );
        available = limit - sourcePrefix.length - 1 -
          serial.length - extension.length;
      }
      var mediaStem = truncateUnicode(parts.stem, Math.max(1, available));
      if (mediaStem.length < minimumMediaLength && sourcePrefix.length > 1) {
        sourcePrefix = truncateUnicode(
          sourcePrefix,
          Math.max(1, sourcePrefix.length -
            (minimumMediaLength - mediaStem.length))
        );
        available = limit - sourcePrefix.length - 1 -
          serial.length - extension.length;
        mediaStem = truncateUnicode(parts.stem, Math.max(1, available));
      }
      var output = sourcePrefix + "_" + mediaStem + serial + extension;
      ensure(output.length <= limit, "OUTPUT_PATH",
        "メディアファイル名を安全な長さへ短縮できませんでした。");
      return output;
    }

    function buildTextOutputFileName(sourceStem, maximum) {
      var limit = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
      var suffix = "_text.txt";
      var safeStem = sanitizeStem(
        sourceStem,
        "document",
        MAX_SOURCE_OUTPUT_STEM
      );
      return truncateUnicode(
        safeStem,
        Math.max(1, limit - suffix.length)
      ) + suffix;
    }

    function sanitizePathSegment(segment) {
      return sanitizeStem(segment, "folder", MAX_OUTPUT_PARENT_SEGMENT);
    }

    function relativeParentSegments(path) {
      var raw = String(path || "").replace(/\\/g, "/");
      if (/^[A-Za-z]:\//.test(raw) || /^\/\//.test(raw) || raw.charAt(0) === "/") {
        return [];
      }
      var pieces = raw.split("/");
      if (pieces.length) {
        pieces.pop();
      }
      var output = [];
      var index;
      for (index = 0; index < pieces.length; index += 1) {
        var piece = pieces[index];
        if (!piece || piece === "." || piece === ".." || /^[A-Za-z]:$/.test(piece)) {
          continue;
        }
        output.push(piece);
      }
      return output;
    }

    function relativeParentPathKey(path) {
      return relativeParentSegments(path).join("/");
    }

    function parentPathIdentifierValue(rawParentPath, salt) {
      var source = "parent:" + String(rawParentPath || "") +
        (salt ? ":" + String(salt) : "");
      var identifier = crc32(encodeUtf8(source)).toString(16);
      while (identifier.length < 8) {
        identifier = "0" + identifier;
      }
      return identifier;
    }

    function parentPathIdentifier(rawParentPath, salt) {
      return "_" + parentPathIdentifierValue(rawParentPath, salt);
    }

    function parentOmissionSegment(rawParentPath, salt) {
      return "__" + parentPathIdentifierValue(
        "omitted:" + String(rawParentPath || ""),
        salt
      );
    }

    function fitParentTreeSegment(node, salt) {
      var maximum = normalizedOutputLimit(
        node.maximumLength,
        MAX_OUTPUT_PARENT_SEGMENT
      );
      if (!node.forceIdentifier) {
        return truncateUnicode(node.safeName, maximum);
      }
      var identifier = parentPathIdentifier(node.fullPath, salt);
      ensure(maximum > identifier.length, "OUTPUT_PATH",
        "入力元の親フォルダセグメントを安全な長さへ短縮できませんでした。");
      return truncateUnicode(
        node.safeName,
        maximum - identifier.length
      ) + identifier;
    }

    function assignParentTreeChildNames(parentNode) {
      var children = parentNode.children.slice();
      children.sort(function (left, right) {
        return left.fullPath < right.fullPath ? -1 :
          (left.fullPath > right.fullPath ? 1 : 0);
      });
      var safeCounts = Object.create(null);
      var index;
      for (index = 0; index < children.length; index += 1) {
        var safeKey = children[index].safeName.toLowerCase();
        safeCounts[safeKey] = (safeCounts[safeKey] || 0) + 1;
      }
      for (index = 0; index < children.length; index += 1) {
        var child = children[index];
        child.forceIdentifier = safeCounts[child.safeName.toLowerCase()] > 1 ||
          child.maximumLength < child.safeName.length;
      }

      var used = Object.create(null);
      for (index = 0; index < children.length; index += 1) {
        child = children[index];
        var salt = 0;
        var candidate = fitParentTreeSegment(child, salt);
        var candidateKey = candidate.toLowerCase();
        while (used[candidateKey] && salt < 1000000) {
          child.forceIdentifier = true;
          salt += 1;
          candidate = fitParentTreeSegment(child, salt);
          candidateKey = candidate.toLowerCase();
        }
        ensure(salt < 1000000, "NAME_COLLISION",
          "入力元の親フォルダセグメント名を一意にできませんでした。");
        used[candidateKey] = child.fullPath;
        child.outputName = candidate;
      }
      for (index = 0; index < children.length; index += 1) {
        assignParentTreeChildNames(children[index]);
      }
    }

    function parentTreeOutputPath(nodes) {
      var output = [];
      var index;
      for (index = 0; index < nodes.length; index += 1) {
        output.push(nodes[index].outputName);
      }
      return output.join("/");
    }

    function parentTreeOmissionStart(nodes) {
      var omissionLength = 10;
      var start = nodes.length - 1;
      var suffixLength = nodes[start].outputName.length;
      while (start > 2) {
        var candidateSuffixLength = nodes[start - 1].outputName.length + 1 +
          suffixLength;
        var candidateLength = nodes[0].outputName.length + 1 +
          omissionLength + 1 + candidateSuffixLength;
        if (candidateLength > MAX_OUTPUT_PARENT_PATH) {
          break;
        }
        start -= 1;
        suffixLength = candidateSuffixLength;
      }
      return start;
    }

    function fitFixedParentPath(record, omissionValues) {
      var safePath = parentTreeOutputPath(record.nodes);
      if (safePath.length <= MAX_OUTPUT_PARENT_PATH) {
        return safePath;
      }
      ensure(record.nodes.length >= 3, "OUTPUT_PATH",
        "入力元の親フォルダを安全な長さへ短縮できませんでした。");
      var start = parentTreeOmissionStart(record.nodes);
      var omissionKey = record.nodes[start - 1].fullPath;
      var omission = omissionValues[omissionKey];
      ensure(omission, "OUTPUT_PATH",
        "入力元の親フォルダの固定省略名がありません。");
      var output = [record.nodes[0].outputName, omission];
      var index;
      for (index = start; index < record.nodes.length; index += 1) {
        output.push(record.nodes[index].outputName);
      }
      var outputPath = output.join("/");
      ensure(outputPath.length <= MAX_OUTPUT_PARENT_PATH, "OUTPUT_PATH",
        "入力元の親フォルダを安全な長さへ短縮できませんでした。");
      return outputPath;
    }

    function ParentPathMap(items) {
      this.values = Object.create(null);
      var root = {
        fullPath: "",
        children: [],
        childrenByOriginal: Object.create(null)
      };
      var recordsByKey = Object.create(null);
      var records = [];
      var index;
      for (index = 0; index < items.length; index += 1) {
        var key = relativeParentPathKey(items[index].path);
        if (Object.prototype.hasOwnProperty.call(recordsByKey, key)) {
          continue;
        }
        var rawSegments = relativeParentSegments(items[index].path);
        var nodes = [];
        var parentNode = root;
        var segmentIndex;
        for (segmentIndex = 0;
          segmentIndex < rawSegments.length;
          segmentIndex += 1) {
          var originalName = rawSegments[segmentIndex];
          var node = parentNode.childrenByOriginal[originalName];
          if (!node) {
            var fullPath = parentNode.fullPath ?
              parentNode.fullPath + "/" + originalName : originalName;
            node = {
              originalName: originalName,
              safeName: sanitizePathSegment(originalName),
              parent: parentNode,
              fullPath: fullPath,
              outputName: "",
              maximumLength: MAX_OUTPUT_PARENT_SEGMENT,
              forceIdentifier: false,
              children: [],
              childrenByOriginal: Object.create(null)
            };
            parentNode.childrenByOriginal[originalName] = node;
            parentNode.children.push(node);
          }
          nodes.push(node);
          parentNode = node;
        }
        var record = {
          key: key,
          nodes: nodes
        };
        recordsByKey[key] = record;
        records.push(record);
      }

      records.sort(function (left, right) {
        return left.key < right.key ? -1 : (left.key > right.key ? 1 : 0);
      });
      assignParentTreeChildNames(root);

      for (index = 0; index < records.length; index += 1) {
        var current = records[index];
        var currentPath = parentTreeOutputPath(current.nodes);
        if (currentPath.length <= MAX_OUTPUT_PARENT_PATH ||
          current.nodes.length < 3) {
          continue;
        }
        var rootLength = current.nodes[0].outputName.length;
        var lastNode = current.nodes[current.nodes.length - 1];
        var maximumLastLength = MAX_OUTPUT_PARENT_PATH - rootLength - 12;
        ensure(maximumLastLength > 9, "OUTPUT_PATH",
          "入力元の最内側フォルダを固定長へ短縮できませんでした。");
        if (lastNode.maximumLength > maximumLastLength) {
          lastNode.maximumLength = maximumLastLength;
        }
      }
      assignParentTreeChildNames(root);

      var omissionRequests = [];
      var omissionRequestKeys = Object.create(null);
      for (index = 0; index < records.length; index += 1) {
        current = records[index];
        if (parentTreeOutputPath(current.nodes).length <=
          MAX_OUTPUT_PARENT_PATH) {
          continue;
        }
        var start = parentTreeOmissionStart(current.nodes);
        var omittedNode = current.nodes[start - 1];
        if (!omissionRequestKeys[omittedNode.fullPath]) {
          omissionRequestKeys[omittedNode.fullPath] = true;
          omissionRequests.push({
            key: omittedNode.fullPath,
            rootNode: current.nodes[0]
          });
        }
      }
      omissionRequests.sort(function (left, right) {
        return left.key < right.key ? -1 : (left.key > right.key ? 1 : 0);
      });
      var omissionValues = Object.create(null);
      var omissionUsedByRoot = Object.create(null);
      for (index = 0; index < omissionRequests.length; index += 1) {
        var request = omissionRequests[index];
        var rootKey = request.rootNode.fullPath;
        var rootUsed = omissionUsedByRoot[rootKey];
        if (!rootUsed) {
          rootUsed = Object.create(null);
          var childIndex;
          for (childIndex = 0;
            childIndex < request.rootNode.children.length;
            childIndex += 1) {
            rootUsed[request.rootNode.children[childIndex].outputName.toLowerCase()] =
              true;
          }
          omissionUsedByRoot[rootKey] = rootUsed;
        }
        var omissionSalt = 0;
        var omissionCandidate = parentOmissionSegment(request.key, omissionSalt);
        while (rootUsed[omissionCandidate.toLowerCase()] &&
          omissionSalt < 1000000) {
          omissionSalt += 1;
          omissionCandidate = parentOmissionSegment(request.key, omissionSalt);
        }
        ensure(omissionSalt < 1000000, "NAME_COLLISION",
          "入力元の親フォルダの固定省略名を一意にできませんでした。");
        rootUsed[omissionCandidate.toLowerCase()] = true;
        omissionValues[request.key] = omissionCandidate;
      }

      for (index = 0; index < records.length; index += 1) {
        this.values[records[index].key] = fitFixedParentPath(
          records[index],
          omissionValues
        );
      }
    }

    ParentPathMap.prototype.get = function (path) {
      var key = relativeParentPathKey(path);
      if (Object.prototype.hasOwnProperty.call(this.values, key)) {
        return this.values[key];
      }
      fail("OUTPUT_PATH",
        "エクスポート開始時に収集されていない入力親フォルダです。");
    };

    function NameAllocator(maximum) {
      this.used = Object.create(null);
      this.maximum = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
    }

    NameAllocator.prototype.claim = function (requested) {
      var candidate = String(requested || "");
      var key = candidate.toLowerCase();
      if (this.used[key]) {
        return false;
      }
      this.used[key] = true;
      return true;
    };

    NameAllocator.prototype.allocate = function (requested, isFile, maximum) {
      var limit = normalizedOutputLimit(maximum, this.maximum);
      var safe = isFile ?
        sanitizeFileName(requested, "file", limit) :
        sanitizeStem(requested, "folder", limit);
      if (this.claim(safe)) {
        return safe;
      }
      var stem = safe;
      var extension = "";
      if (isFile) {
        var dot = safe.lastIndexOf(".");
        if (dot > 0) {
          stem = safe.slice(0, dot);
          extension = safe.slice(dot);
        }
      }
      var index = 1;
      while (index < 1000000) {
        var suffix = "_" + zeroPad(index, 4);
        var candidateExtension = extension;
        ensure(
          candidateExtension.length <=
            Math.max(0, limit - suffix.length - 1),
          "OUTPUT_PATH",
          "出力ファイルの拡張子と連番を保持したまま短縮できませんでした。"
        );
        var available = limit - suffix.length - candidateExtension.length;
        var candidate = truncateUnicode(stem, Math.max(1, available)) +
          suffix + candidateExtension;
        if (this.claim(candidate)) {
          return candidate;
        }
        index += 1;
      }
      fail("NAME_COLLISION", "出力ファイル名の連番を確保できませんでした。");
    };

    function allocatePrefixedMediaFileName(
      allocator,
      sourceStem,
      originalMediaName,
      maximum
    ) {
      var limit = normalizedOutputLimit(maximum, MAX_OUTPUT_FILE_NAME);
      var index = 0;
      while (index < 1000000) {
        var suffix = index ? "_" + zeroPad(index, 4) : "";
        var candidate = buildPrefixedMediaFileName(
          sourceStem,
          originalMediaName,
          suffix,
          limit
        );
        if (allocator.claim(candidate)) {
          return candidate;
        }
        index += 1;
      }
      fail("NAME_COLLISION", "出力ファイル名の連番を確保できませんでした。");
    }

    function outputRootValue(parentPath, sourceStem, suffix) {
      var folderName = sourceStem + suffix;
      return (parentPath ? parentPath + "/" : "") +
        folderName;
    }

    function fitOutputRootPath(parentPath, sourceStem, suffix) {
      var serial = String(suffix || "");
      var fixedParent = normalizePackagePath(parentPath);
      ensure(fixedParent.length <= MAX_OUTPUT_PARENT_PATH, "OUTPUT_PATH",
        "入力元の親フォルダが固定長の上限を超えています。");
      var safeStem = sanitizeStem(
        sourceStem,
        "document",
        MAX_SOURCE_OUTPUT_STEM
      );
      var availableStemLength = MAX_OUTPUT_ROOT_PATH -
        (fixedParent ? fixedParent.length + 1 : 0) - serial.length;
      ensure(availableStemLength >= 1, "OUTPUT_PATH",
        "出力ルートへOfficeファイル別フォルダを追加できませんでした。");
      safeStem = truncateUnicode(
        safeStem,
        Math.max(
          1,
          Math.min(
            MAX_SOURCE_OUTPUT_STEM - serial.length,
            availableStemLength
          )
        )
      );
      var result = outputRootValue(fixedParent, safeStem, serial);
      ensure(result.length <= MAX_OUTPUT_ROOT_PATH, "OUTPUT_PATH",
        "出力ルートを安全な長さへ短縮できませんでした。");
      return result;
    }

    function OutputRootAllocator() {
      this.used = Object.create(null);
    }

    OutputRootAllocator.prototype.allocate = function (requested) {
      var normalized = normalizePackagePath(requested);
      var pieces = normalized.split("/");
      var stem = pieces.pop() || "document";
      var fixedParent = pieces.join("/");
      normalized = fitOutputRootPath(fixedParent, stem, "");
      var key = normalized.toLowerCase();
      if (!this.used[key]) {
        this.used[key] = true;
        return normalized;
      }
      var index = 1;
      while (index < 1000000) {
        var suffix = "_" + zeroPad(index, 4);
        var candidate = fitOutputRootPath(fixedParent, stem, suffix);
        key = candidate.toLowerCase();
        if (!this.used[key]) {
          this.used[key] = true;
          return candidate;
        }
        index += 1;
      }
      fail("NAME_COLLISION", "出力フォルダ名の連番を確保できませんでした。");
    };

    function buildOutputRootPath(item, allocator, sourceStem, parentPathMap) {
      var parent = parentPathMap.get(item.path);
      var stem = sourceStem || buildSourceOutputStem(item.file.name);
      return allocator.allocate(fitOutputRootPath(parent, stem, ""));
    }

    function outputRelativeFileNameLimit(rootName, category) {
      var categoryName = String(category || "");
      var fixedLength = String(rootName || "").length + 1 +
        (categoryName ? categoryName.length + 1 : 0);
      return Math.max(
        1,
        Math.min(MAX_OUTPUT_FILE_NAME, MAX_OUTPUT_RELATIVE_PATH - fixedLength)
      );
    }

    function buildOutputRelativePath(rootName, category, fileName) {
      var categoryName = String(category || "");
      ensure(String(rootName || "").length <= MAX_OUTPUT_ROOT_PATH, "OUTPUT_PATH",
        "出力ルートを安全な長さへ短縮できませんでした。");
      var path = String(rootName || "") + "/" +
        (categoryName ? categoryName + "/" : "") + String(fileName || "");
      var normalized = normalizePackagePath(path);
      ensure(leafName(normalized).length <= MAX_OUTPUT_FILE_NAME, "OUTPUT_PATH",
        "出力ファイル名を安全な長さへ短縮できませんでした。");
      ensure(normalized.length <= MAX_OUTPUT_RELATIVE_PATH, "OUTPUT_PATH",
        "ZIP内の相対パスを安全な長さへ短縮できませんでした。");
      return normalized;
    }

    function normalizePackagePath(path) {
      var pieces = String(path || "").replace(/\\/g, "/").split("/");
      var output = [];
      var index;
      for (index = 0; index < pieces.length; index += 1) {
        if (!pieces[index] || pieces[index] === ".") {
          continue;
        }
        if (pieces[index] === "..") {
          ensure(output.length > 0, "PATH", "パッケージ外を参照するパスです。");
          output.pop();
        } else {
          output.push(pieces[index]);
        }
      }
      return output.join("/");
    }

    function resolvePartTarget(sourcePart, target) {
      var targetText = String(target || "").replace(/\\/g, "/");
      if (targetText.charAt(0) === "/") {
        return normalizePackagePath(targetText.slice(1));
      }
      var base = parentPath(sourcePart);
      return normalizePackagePath((base ? base + "/" : "") + targetText);
    }

    function relationshipPartName(part) {
      var directory = parentPath(part);
      var file = leafName(part);
      return (directory ? directory + "/" : "") + "_rels/" + file + ".rels";
    }

    function delayTurn() {
      return new Promise(function (resolve) {
        setTimeout(resolve, 0);
      });
    }

    async function cooperativeYield(index, interval) {
      if (index > 0 && index % interval === 0) {
        checkCancelled();
        await delayTurn();
        checkCancelled();
      }
    }

    function formatBytes(size) {
      var value = Number(size) || 0;
      var units = ["B", "KB", "MB", "GB"];
      var unit = 0;
      while (value >= 1024 && unit < units.length - 1) {
        value /= 1024;
        unit += 1;
      }
      var digits = unit === 0 ? 0 : (value >= 10 ? 1 : 2);
      return value.toLocaleString(localeForLanguage(currentLanguage), {
        minimumFractionDigits: digits,
        maximumFractionDigits: digits
      }) + " " + units[unit];
    }

    var CRC_TABLE = (function () {
      var table = new Uint32Array(256);
      var index;
      for (index = 0; index < 256; index += 1) {
        var value = index;
        var bit;
        for (bit = 0; bit < 8; bit += 1) {
          value = (value & 1) ? (0xEDB88320 ^ (value >>> 1)) : (value >>> 1);
        }
        table[index] = value >>> 0;
      }
      return table;
    }());

    function crc32(bytes) {
      var value = 0xFFFFFFFF;
      var index;
      for (index = 0; index < bytes.length; index += 1) {
        value = CRC_TABLE[(value ^ bytes[index]) & 0xFF] ^ (value >>> 8);
      }
      return (value ^ 0xFFFFFFFF) >>> 0;
    }

    async function decompressBytes(bytes, format, expectedSize) {
      ensure(typeof DecompressionStream === "function", "DEFLATE_UNAVAILABLE",
        "このブラウザはZIPのDeflate展開に対応していません。");
      if (typeof expectedSize === "number") {
        ensure(expectedSize <= MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES, "ZIP_ENTRY_TOO_LARGE",
          "このファイルは展開後サイズが大きすぎるため処理できません。");
      }
      var reader = null;
      try {
        var decompressedStream = new Blob([bytes])
          .stream()
          .pipeThrough(new DecompressionStream(format));
        reader = decompressedStream.getReader();
        var parts = [];
        var total = 0;
        while (true) {
          checkCancelled();
          var chunk = await reader.read();
          if (chunk.done) {
            break;
          }
          var value = chunk.value instanceof Uint8Array ?
            chunk.value :
            new Uint8Array(chunk.value);
          total += value.length;
          ensure(Number.isSafeInteger(total), "ZIP_INVALID_SIZE",
            "展開後サイズを安全に計算できません。");
          ensure(total <= MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES, "ZIP_ENTRY_TOO_LARGE",
            "このファイルは展開後サイズが大きすぎるため処理できません。");
          if (typeof expectedSize === "number") {
            ensure(total <= expectedSize, "DEFLATE_SIZE",
              "展開後のデータサイズがZIP情報を超えています。");
          }
          parts.push(value);
        }
        var result = concatBytes(parts, total);
        if (typeof expectedSize === "number") {
          ensure(result.length === expectedSize, "DEFLATE_SIZE",
            "展開後のデータサイズがZIP情報と一致しません。");
        }
        return result;
      } catch (error) {
        if (error instanceof AppError) {
          throw error;
        }
        fail("DEFLATE", "圧縮データを展開できませんでした。");
      } finally {
        if (reader) {
          try {
            await reader.cancel();
          } catch (ignoreCancelError) {
            // The stream may already be closed after a successful read.
          }
          try {
            reader.releaseLock();
          } catch (ignoreReleaseError) {
            // Some browsers release the lock when cancellation finishes.
          }
        }
      }
    }

    function findEocd(bytes) {
      var minimum = Math.max(0, bytes.length - 65557);
      var offset;
      for (offset = bytes.length - 22; offset >= minimum; offset -= 1) {
        if (u32(bytes, offset) === ZIP_EOCD_SIGNATURE) {
          var commentLength = u16(bytes, offset + 20);
          if (offset + 22 + commentLength === bytes.length) {
            return offset;
          }
        }
      }
      return -1;
    }

    function parseZipExtraFields(extraBytes) {
      var fields = [];
      var offset = 0;
      while (offset < extraBytes.length) {
        ensure(extraBytes.length - offset >= 4, "ZIP_EXTRA",
          "ZIP Extra Fieldのheaderが不完全です。");
        var id = u16(extraBytes, offset);
        var size = u16(extraBytes, offset + 2);
        offset += 4;
        requireRange(extraBytes, offset, size, "ZIP Extra Field");
        fields.push({
          id: id,
          data: extraBytes.subarray(offset, offset + size)
        });
        offset += size;
      }
      return fields;
    }

    function decodeCp437(bytes) {
      var extended =
        "ÇüéâäàåçêëèïîìÄÅ" +
        "ÉæÆôöòûùÿÖÜ¢£¥₧ƒ" +
        "áíóúñÑªº¿⌐¬½¼¡«»" +
        "░▒▓│┤╡╢╖╕╣║╗╝╜╛┐" +
        "└┴┬├─┼╞╟╚╔╩╦╠═╬╧" +
        "╨╤╥╙╘╒╓╫╪┘┌█▄▌▐▀" +
        "αßΓπΣσµτΦΘΩδ∞φε∩" +
        "≡±≥≤⌠⌡÷≈°∙·√ⁿ²■ ";
      var output = "";
      var index;
      for (index = 0; index < bytes.length; index += 1) {
        output += bytes[index] < 0x80 ?
          String.fromCharCode(bytes[index]) :
          extended.charAt(bytes[index] - 0x80);
      }
      return output;
    }

    function decodeZipEntryName(nameBytes, flags, extraBytes) {
      if ((flags & 0x0800) !== 0) {
        return decodeUtf8(nameBytes, false);
      }
      var ascii = true;
      var index;
      for (index = 0; index < nameBytes.length; index += 1) {
        if (nameBytes[index] >= 0x80) {
          ascii = false;
          break;
        }
      }
      var fields = parseZipExtraFields(extraBytes);
      for (index = 0; index < fields.length; index += 1) {
        if (fields[index].id !== 0x7075 || fields[index].data.length < 5) {
          continue;
        }
        var data = fields[index].data;
        if (data[0] === 1 && u32(data, 1) === crc32(nameBytes)) {
          return decodeUtf8(data.subarray(5), false);
        }
      }
      if (ascii) {
        return decodeCompressedUnicode(nameBytes);
      }
      return decodeCp437(nameBytes);
    }

    function ZipArchive(bytes) {
      this.bytes = bytes;
      this.entries = [];
      this.byName = Object.create(null);
      this.totalCompressedSize = 0;
      this.totalUncompressedSize = 0;
      this.parse();
    }

    ZipArchive.prototype.parse = function () {
      var eocd = findEocd(this.bytes);
      ensure(eocd >= 0, "ZIP_EOCD", "ZIPの終端情報が見つかりません。");
      var diskNumber = u16(this.bytes, eocd + 4);
      var centralDisk = u16(this.bytes, eocd + 6);
      var diskEntries = u16(this.bytes, eocd + 8);
      var totalEntries = u16(this.bytes, eocd + 10);
      var centralSize = u32(this.bytes, eocd + 12);
      var centralOffset = u32(this.bytes, eocd + 16);
      ensure(diskNumber === 0 && centralDisk === 0 && diskEntries === totalEntries,
        "ZIP_SPLIT", "分割ZIPには対応していません。");
      ensure(totalEntries !== 0xFFFF && centralSize !== 0xFFFFFFFF &&
        centralOffset !== 0xFFFFFFFF, "ZIP64", "ZIP64には対応していません。");
      ensure(totalEntries <= MAX_ZIP_ENTRIES, "ZIP_ENTRIES",
        "ZIP内のファイル数が安全上限を超えています。");
      requireRange(this.bytes, centralOffset, centralSize, "ZIP Central Directory");
      ensure(centralOffset + centralSize <= eocd, "ZIP_CENTRAL",
        "ZIP Central Directoryの位置が不正です。");

      var position = centralOffset;
      var index;
      for (index = 0; index < totalEntries; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        requireRange(this.bytes, position, 46, "ZIP Central Directory entry");
        ensure(u32(this.bytes, position) === ZIP_CENTRAL_SIGNATURE,
          "ZIP_CENTRAL", "ZIP Central Directoryが壊れています。");
        var flags = u16(this.bytes, position + 8);
        var method = u16(this.bytes, position + 10);
        var checksum = u32(this.bytes, position + 16);
        var compressedSize = u32(this.bytes, position + 20);
        var uncompressedSize = u32(this.bytes, position + 24);
        var nameLength = u16(this.bytes, position + 28);
        var extraLength = u16(this.bytes, position + 30);
        var commentLength = u16(this.bytes, position + 32);
        var localOffset = u32(this.bytes, position + 42);
        var entryLength = 46 + nameLength + extraLength + commentLength;
        requireRange(this.bytes, position, entryLength, "ZIP Central Directory entry");
        ensure(compressedSize !== 0xFFFFFFFF && uncompressedSize !== 0xFFFFFFFF &&
          localOffset !== 0xFFFFFFFF, "ZIP64", "ZIP64エントリーには対応していません。");
        ensure((flags & 0x0001) === 0, "ZIP_ENCRYPTED",
          "暗号化されたZIPエントリーには対応していません。");
        ensure(method === 0 || method === 8, "ZIP_METHOD",
          "未対応のZIP圧縮方式が使われています。");
        var nameBytes = sliceBytes(this.bytes, position + 46, nameLength, "ZIP entry name");
        var extraBytes = sliceBytes(
          this.bytes,
          position + 46 + nameLength,
          extraLength,
          "ZIP entry extra field"
        );
        ensure(uncompressedSize <= MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES,
          "ZIP_ENTRY_TOO_LARGE",
          "このファイルは展開後サイズが大きすぎるため処理できません。");
        ensure(!(compressedSize === 0 && uncompressedSize > 0), "ZIP_INVALID_SIZE",
          "ZIPエントリーの圧縮サイズ情報が不正です。");
        if (compressedSize > 0 && uncompressedSize >= 1024 * 1024) {
          ensure(uncompressedSize / compressedSize <= MAX_COMPRESSION_RATIO,
            "ZIP_RATIO", "このファイルは圧縮率が安全上限を超えています。");
        }
        var nextCompressedTotal = this.totalCompressedSize + compressedSize;
        var nextUncompressedTotal = this.totalUncompressedSize + uncompressedSize;
        ensure(Number.isSafeInteger(nextCompressedTotal) &&
          Number.isSafeInteger(nextUncompressedTotal), "ZIP_INVALID_SIZE",
        "ZIP全体のサイズを安全に計算できません。");
        ensure(nextUncompressedTotal <= MAX_ZIP_TOTAL_UNCOMPRESSED_BYTES,
          "ZIP_TOTAL_TOO_LARGE",
          "このファイルはZIP全体の展開後サイズが大きすぎるため処理できません。");
        this.totalCompressedSize = nextCompressedTotal;
        this.totalUncompressedSize = nextUncompressedTotal;
        var decodedName = decodeZipEntryName(nameBytes, flags, extraBytes);
        var name = normalizePackagePath(decodedName);
        ensure(name.length > 0, "ZIP_PATH", "空のZIPエントリー名があります。");
        var key = name.toLowerCase();
        ensure(!this.byName[key], "ZIP_DUPLICATE", "ZIP内に同名エントリーがあります: " + name);
        var entry = {
          name: name,
          flags: flags,
          method: method,
          crc: checksum,
          compressedSize: compressedSize,
          uncompressedSize: uncompressedSize,
          localOffset: localOffset,
          directory: /\/$/.test(decodedName)
        };
        this.entries.push(entry);
        this.byName[key] = entry;
        position += entryLength;
      }
      ensure(position === centralOffset + centralSize, "ZIP_CENTRAL",
        "ZIP Central Directoryのサイズが一致しません。");
    };

    ZipArchive.prototype.has = function (name) {
      return !!this.byName[normalizePackagePath(name).toLowerCase()];
    };

    ZipArchive.prototype.get = function (name) {
      return this.byName[normalizePackagePath(name).toLowerCase()] || null;
    };

    ZipArchive.prototype.listPrefix = function (prefix) {
      var normalized = normalizePackagePath(prefix).toLowerCase();
      if (normalized && normalized.charAt(normalized.length - 1) !== "/") {
        normalized += "/";
      }
      var output = [];
      var index;
      for (index = 0; index < this.entries.length; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        var entry = this.entries[index];
        if (!entry.directory &&
          entry.name.toLowerCase().indexOf(normalized) === 0 &&
          entry.name.length > normalized.length) {
          output.push(entry);
        }
      }
      return output;
    };

    ZipArchive.prototype.extract = async function (entryOrName) {
      var entry = typeof entryOrName === "string" ? this.get(entryOrName) : entryOrName;
      ensure(entry, "ZIP_ENTRY", "ZIPエントリーが見つかりません。");
      checkCancelled();
      ensure(entry.uncompressedSize <= MAX_ZIP_ENTRY_UNCOMPRESSED_BYTES,
        "ZIP_ENTRY_TOO_LARGE",
        "このファイルは展開後サイズが大きすぎるため処理できません。");
      requireRange(this.bytes, entry.localOffset, 30, "ZIP local header");
      ensure(u32(this.bytes, entry.localOffset) === ZIP_LOCAL_SIGNATURE,
        "ZIP_LOCAL", "ZIP local headerが壊れています: " + entry.name);
      var localFlags = u16(this.bytes, entry.localOffset + 6);
      var localMethod = u16(this.bytes, entry.localOffset + 8);
      ensure((localFlags & 0x0001) === 0, "ZIP_ENCRYPTED",
        "暗号化されたZIPエントリーです: " + entry.name);
      ensure(localMethod === entry.method, "ZIP_METHOD",
        "ZIPの圧縮方式情報が一致しません: " + entry.name);
      var nameLength = u16(this.bytes, entry.localOffset + 26);
      var extraLength = u16(this.bytes, entry.localOffset + 28);
      var dataOffset = entry.localOffset + 30 + nameLength + extraLength;
      var compressed = sliceBytes(
        this.bytes,
        dataOffset,
        entry.compressedSize,
        "ZIP compressed data"
      );
      var output;
      if (entry.method === 0) {
        ensure(entry.compressedSize === entry.uncompressedSize, "ZIP_SIZE",
          "Storeエントリーのサイズが一致しません: " + entry.name);
        output = compressed;
      } else {
        output = await decompressBytes(compressed, "deflate-raw", entry.uncompressedSize);
      }
      ensure(crc32(output) === entry.crc, "ZIP_CRC",
        "ZIPエントリーのCRCが一致しません: " + entry.name);
      return output;
    };

    ZipArchive.prototype.extractXml = async function (name) {
      var entry = this.get(name);
      ensure(entry, "OOXML_PART", "必要なXMLパーツがありません: " + name);
      return parseXml(await this.extract(entry), name);
    };

    function writeU16(bytes, offset, value) {
      viewOf(bytes).setUint16(offset, value & 0xFFFF, true);
    }

    function writeU32(bytes, offset, value) {
      viewOf(bytes).setUint32(offset, value >>> 0, true);
    }

    function dosDateTime(date) {
      var year = Math.max(1980, Math.min(2107, date.getFullYear()));
      return {
        time: ((date.getHours() & 31) << 11) |
          ((date.getMinutes() & 63) << 5) |
          (Math.floor(date.getSeconds() / 2) & 31),
        date: (((year - 1980) & 127) << 9) |
          (((date.getMonth() + 1) & 15) << 5) |
          (date.getDate() & 31)
      };
    }

    function ZipBuilder() {
      this.entries = [];
      this.paths = Object.create(null);
      this.totalBytes = 0;
    }

    ZipBuilder.prototype.add = function (path, data) {
      var normalized = normalizePackagePath(path);
      ensure(normalized && normalized.charAt(normalized.length - 1) !== "/",
        "OUTPUT_PATH", "出力ZIPにはファイルだけを追加できます。");
      ensure(leafName(normalized).length <= MAX_OUTPUT_FILE_NAME,
        "OUTPUT_PATH", "出力ファイル名が安全な長さを超えています。");
      ensure(normalized.length <= MAX_OUTPUT_RELATIVE_PATH,
        "OUTPUT_PATH", "ZIP内の相対パスが安全な長さを超えています。");
      var key = normalized.toLowerCase();
      ensure(!this.paths[key], "OUTPUT_DUPLICATE", "出力ZIPの名前が重複しています: " + normalized);
      ensure(this.entries.length < 0xFFFF, "OUTPUT_ZIP64",
        "出力ファイル数がZIP32の上限を超えています。");
      var bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
      var nextTotal = this.totalBytes + bytes.length;
      ensure(Number.isSafeInteger(nextTotal), "OUTPUT_SIZE",
        "出力サイズを安全に計算できません。");
      ensure(nextTotal <= MAX_TOTAL_OUTPUT_BYTES, "OUTPUT_SIZE",
        "抽出結果の合計サイズが安全上限を超えています。");
      this.totalBytes = nextTotal;
      this.paths[key] = true;
      this.entries.push({
        path: normalized,
        nameBytes: encodeUtf8(normalized),
        data: bytes,
        crc: crc32(bytes)
      });
    };

    ZipBuilder.prototype.toBlob = function () {
      checkCancelled();
      ensure(Number.isSafeInteger(this.totalBytes) &&
        this.totalBytes <= MAX_TOTAL_OUTPUT_BYTES, "OUTPUT_SIZE",
      "抽出結果の合計サイズが安全上限を超えています。");
      ensure(this.entries.length <= 0xFFFF, "OUTPUT_ZIP64",
        "出力ファイル数がZIP32の上限を超えています。");
      var now = dosDateTime(new Date());
      var parts = [];
      var centrals = [];
      var offset = 0;
      var centralSize = 0;
      var index;
      for (index = 0; index < this.entries.length; index += 1) {
        checkCancelled();
        var entry = this.entries[index];
        ensure(entry.data.length <= 0xFFFFFFFF, "OUTPUT_ZIP64",
          "出力ファイルがZIP32のサイズ上限を超えています。");
        ensure(offset <= 0xFFFFFFFF, "OUTPUT_ZIP64", "出力ZIPがZIP32の上限を超えています。");
        var local = new Uint8Array(30 + entry.nameBytes.length);
        writeU32(local, 0, ZIP_LOCAL_SIGNATURE);
        writeU16(local, 4, 20);
        writeU16(local, 6, 0x0800);
        writeU16(local, 8, 0);
        writeU16(local, 10, now.time);
        writeU16(local, 12, now.date);
        writeU32(local, 14, entry.crc);
        writeU32(local, 18, entry.data.length);
        writeU32(local, 22, entry.data.length);
        writeU16(local, 26, entry.nameBytes.length);
        writeU16(local, 28, 0);
        local.set(entry.nameBytes, 30);
        parts.push(local, entry.data);

        var central = new Uint8Array(46 + entry.nameBytes.length);
        writeU32(central, 0, ZIP_CENTRAL_SIGNATURE);
        writeU16(central, 4, 20);
        writeU16(central, 6, 20);
        writeU16(central, 8, 0x0800);
        writeU16(central, 10, 0);
        writeU16(central, 12, now.time);
        writeU16(central, 14, now.date);
        writeU32(central, 16, entry.crc);
        writeU32(central, 20, entry.data.length);
        writeU32(central, 24, entry.data.length);
        writeU16(central, 28, entry.nameBytes.length);
        writeU16(central, 30, 0);
        writeU16(central, 32, 0);
        writeU16(central, 34, 0);
        writeU16(central, 36, 0);
        writeU32(central, 38, 0);
        writeU32(central, 42, offset);
        central.set(entry.nameBytes, 46);
        centrals.push(central);
        centralSize += central.length;
        offset += local.length + entry.data.length;
      }
      ensure(offset <= 0xFFFFFFFF && centralSize <= 0xFFFFFFFF,
        "OUTPUT_ZIP64", "出力ZIPがZIP32の上限を超えています。");
      for (index = 0; index < centrals.length; index += 1) {
        parts.push(centrals[index]);
      }
      var eocd = new Uint8Array(22);
      writeU32(eocd, 0, ZIP_EOCD_SIGNATURE);
      writeU16(eocd, 4, 0);
      writeU16(eocd, 6, 0);
      writeU16(eocd, 8, this.entries.length);
      writeU16(eocd, 10, this.entries.length);
      writeU32(eocd, 12, centralSize);
      writeU32(eocd, 16, offset);
      writeU16(eocd, 20, 0);
      parts.push(eocd);
      var blob = new Blob(parts, { type: "application/zip" });
      ensure(this.totalBytes <= MAX_TOTAL_OUTPUT_BYTES, "OUTPUT_SIZE",
        "抽出結果の合計サイズが安全上限を超えています。");
      ensure(blob.size <= 0xFFFFFFFF, "OUTPUT_ZIP64",
        "出力ZIPがZIP32の上限を超えています。");
      return blob;
    };

    function timestampForFileName(date) {
      function two(value) {
        var text = String(value);
        return text.length < 2 ? "0" + text : text;
      }
      return String(date.getFullYear()) +
        two(date.getMonth() + 1) +
        two(date.getDate()) + "_" +
        two(date.getHours()) +
        two(date.getMinutes()) +
        two(date.getSeconds());
    }

    function downloadBlob(blob, fileName) {
      var url = URL.createObjectURL(blob);
      var anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = fileName;
      anchor.style.display = "none";
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      setTimeout(function () {
        URL.revokeObjectURL(url);
      }, 2000);
    }

    var CFB_FREESECT = 0xFFFFFFFF;
    var CFB_ENDOFCHAIN = 0xFFFFFFFE;
    var CFB_FATSECT = 0xFFFFFFFD;
    var CFB_DIFSECT = 0xFFFFFFFC;
    var CFB_MAXREGSECT = 0xFFFFFFFA;

    function CompoundFile(bytes) {
      this.bytes = bytes;
      this.majorVersion = 0;
      this.sectorSize = 0;
      this.miniSectorSize = 0;
      this.sectorCount = 0;
      this.miniStreamCutoff = 4096;
      this.fat = [];
      this.miniFat = [];
      this.directory = [];
      this.root = null;
      this.miniStream = new Uint8Array(0);
      this.parse();
    }

    CompoundFile.prototype.sector = function (sectorId) {
      ensure(
        Number.isInteger(sectorId) &&
        sectorId >= 0 &&
        sectorId < this.sectorCount,
        "CFB_SECTOR",
        "CFBセクター番号が範囲外です。"
      );
      var offset = (sectorId + 1) * this.sectorSize;
      return sliceBytes(this.bytes, offset, this.sectorSize, "CFB sector");
    };

    CompoundFile.prototype.chainIds = function (startSector, allocationTable, limit, label) {
      if (startSector === CFB_ENDOFCHAIN || startSector === CFB_FREESECT) {
        return [];
      }
      var maximum = typeof limit === "number" ?
        Math.min(limit, allocationTable.length + 1) :
        allocationTable.length + 1;
      var ids = [];
      var seen = Object.create(null);
      var sectorId = startSector >>> 0;
      while (sectorId !== CFB_ENDOFCHAIN) {
        ensure(sectorId < CFB_MAXREGSECT, "CFB_CHAIN",
          (label || "CFB") + "のchainに特殊sectorが含まれています。");
        ensure(sectorId < allocationTable.length, "CFB_CHAIN",
          (label || "CFB") + "のchainがallocation table範囲外です。");
        ensure(!seen[sectorId], "CFB_CHAIN",
          (label || "CFB") + "のchainが循環しています。");
        ensure(ids.length < maximum, "CFB_CHAIN",
          (label || "CFB") + "のchainが安全上限を超えています。");
        seen[sectorId] = true;
        ids.push(sectorId);
        sectorId = allocationTable[sectorId] >>> 0;
      }
      return ids;
    };

    CompoundFile.prototype.regularChainBytes = function (startSector, size, label) {
      if (size === 0) {
        return new Uint8Array(0);
      }
      var needed = Math.ceil(size / this.sectorSize);
      var ids = this.chainIds(startSector, this.fat, needed + 1, label);
      ensure(ids.length >= needed, "CFB_STREAM",
        (label || "CFB stream") + "のchainが途中で終了しています。");
      var parts = [];
      var index;
      for (index = 0; index < needed; index += 1) {
        parts.push(this.sector(ids[index]));
      }
      return concatBytes(parts, size);
    };

    CompoundFile.prototype.readDirectoryTree = function () {
      var self = this;
      var assigned = Object.create(null);

      function validDirectoryId(id) {
        return id !== CFB_FREESECT && id < self.directory.length;
      }

      function walkSiblings(id, parent, parentPath, active, depth) {
        if (id === CFB_FREESECT) {
          return;
        }
        ensure(validDirectoryId(id), "CFB_DIRECTORY", "CFB Directory IDが範囲外です。");
        ensure(depth <= self.directory.length, "CFB_DIRECTORY",
          "CFB Directory treeが深すぎます。");
        ensure(!active[id], "CFB_DIRECTORY", "CFB Directory treeが循環しています。");
        active[id] = true;
        var entry = self.directory[id];
        if (validDirectoryId(entry.left)) {
          walkSiblings(entry.left, parent, parentPath, active, depth + 1);
        } else {
          ensure(entry.left === CFB_FREESECT, "CFB_DIRECTORY",
            "CFB Directoryのleft IDが不正です。");
        }
        ensure(!assigned[id], "CFB_DIRECTORY",
          "CFB Directory entryが複数のstorageに属しています。");
        assigned[id] = true;
        entry.parent = parent;
        entry.path = parentPath ? parentPath + "/" + entry.name : entry.name;
        if ((entry.type === 1 || entry.type === 5) && entry.child !== CFB_FREESECT) {
          walkSiblings(entry.child, entry.id, entry.path, Object.create(null), depth + 1);
        }
        if (validDirectoryId(entry.right)) {
          walkSiblings(entry.right, parent, parentPath, active, depth + 1);
        } else {
          ensure(entry.right === CFB_FREESECT, "CFB_DIRECTORY",
            "CFB Directoryのright IDが不正です。");
        }
        delete active[id];
      }

      this.root.path = "";
      this.root.parent = null;
      assigned[this.root.id] = true;
      if (this.root.child !== CFB_FREESECT) {
        walkSiblings(this.root.child, this.root.id, "", Object.create(null), 0);
      }
      this.directory.forEach(function (entry) {
        if (entry.type !== 0 && !assigned[entry.id]) {
          entry.path = entry.name;
          entry.parent = null;
        }
      });
    };

    CompoundFile.prototype.parse = function () {
      ensure(startsWithBytes(this.bytes, CFB_SIGNATURE), "CFB_SIGNATURE",
        "CFBシグネチャが一致しません。");
      requireRange(this.bytes, 0, 512, "CFB header");
      this.majorVersion = u16(this.bytes, 26);
      ensure(this.majorVersion === 3 || this.majorVersion === 4,
        "CFB_VERSION", "未対応のCFBバージョンです。");
      ensure(u16(this.bytes, 28) === 0xFFFE, "CFB_BYTE_ORDER",
        "CFBのbyte orderが不正です。");
      var sectorShift = u16(this.bytes, 30);
      var miniSectorShift = u16(this.bytes, 32);
      ensure((this.majorVersion === 3 && sectorShift === 9) ||
        (this.majorVersion === 4 && sectorShift === 12),
      "CFB_SECTOR_SIZE", "CFBのsector sizeが不正です。");
      ensure(miniSectorShift === 6, "CFB_MINI_SECTOR",
        "CFBのmini sector sizeが不正です。");
      this.sectorSize = Math.pow(2, sectorShift);
      this.miniSectorSize = Math.pow(2, miniSectorShift);
      ensure(this.bytes.length >= this.sectorSize, "CFB_HEADER",
        "CFBヘッダーsectorが不足しています。");
      this.sectorCount = Math.floor(this.bytes.length / this.sectorSize) - 1;
      ensure(this.sectorCount >= 2, "CFB_SIZE",
        "CFBファイルはheader、FAT、Directoryの最小3 sectorを満たしていません。");
      ensure(this.bytes.length % this.sectorSize === 0, "CFB_SIZE",
        "CFBファイルサイズがsector境界と一致しません。");

      var numberOfDirectorySectors = u32(this.bytes, 40);
      if (this.majorVersion === 3) {
        ensure(numberOfDirectorySectors === 0, "CFB_HEADER",
          "CFB v3のDirectory sector countが不正です。");
      }
      var numberOfFatSectors = u32(this.bytes, 44);
      var firstDirectorySector = u32(this.bytes, 48);
      this.miniStreamCutoff = u32(this.bytes, 56);
      var firstMiniFatSector = u32(this.bytes, 60);
      var numberOfMiniFatSectors = u32(this.bytes, 64);
      var firstDifatSector = u32(this.bytes, 68);
      var numberOfDifatSectors = u32(this.bytes, 72);
      ensure(this.miniStreamCutoff === 4096,
        "CFB_HEADER", "CFB mini stream cutoffが4096ではありません。");
      ensure(numberOfFatSectors <= this.sectorCount, "CFB_FAT",
        "CFB FAT sector数が不正です。");
      ensure(numberOfDifatSectors <= this.sectorCount, "CFB_DIFAT",
        "CFB DIFAT sector数が不正です。");
      ensure(numberOfMiniFatSectors <= this.sectorCount, "CFB_MINIFAT",
        "CFB MiniFAT sector数が不正です。");

      var difat = [];
      var headerIndex;
      for (headerIndex = 0; headerIndex < 109; headerIndex += 1) {
        var fatSector = u32(this.bytes, 76 + headerIndex * 4);
        if (fatSector !== CFB_FREESECT) {
          difat.push(fatSector);
        }
      }
      var difatSector = firstDifatSector;
      var difatSeen = Object.create(null);
      var difatIndex;
      for (difatIndex = 0; difatIndex < numberOfDifatSectors; difatIndex += 1) {
        ensure(difatSector < CFB_MAXREGSECT && difatSector < this.sectorCount,
          "CFB_DIFAT", "CFB DIFAT sectorが範囲外です。");
        ensure(!difatSeen[difatSector], "CFB_DIFAT", "CFB DIFAT chainが循環しています。");
        difatSeen[difatSector] = true;
        var difatBytes = this.sector(difatSector);
        var entriesPerDifat = this.sectorSize / 4 - 1;
        var entryIndex;
        for (entryIndex = 0; entryIndex < entriesPerDifat; entryIndex += 1) {
          fatSector = u32(difatBytes, entryIndex * 4);
          if (fatSector !== CFB_FREESECT) {
            difat.push(fatSector);
          }
        }
        difatSector = u32(difatBytes, entriesPerDifat * 4);
      }
      if (numberOfDifatSectors === 0) {
        ensure(firstDifatSector === CFB_ENDOFCHAIN || firstDifatSector === CFB_FREESECT,
          "CFB_DIFAT", "CFB DIFAT headerが矛盾しています。");
      } else {
        ensure(difatSector === CFB_ENDOFCHAIN, "CFB_DIFAT",
          "CFB DIFAT chainの終端が不正です。");
      }
      ensure(difat.length >= numberOfFatSectors, "CFB_FAT",
        "CFB FAT sector一覧が不足しています。");

      var fatValues = [];
      var fatSeen = Object.create(null);
      var fatIndex;
      for (fatIndex = 0; fatIndex < numberOfFatSectors; fatIndex += 1) {
        fatSector = difat[fatIndex];
        ensure(fatSector < CFB_MAXREGSECT && fatSector < this.sectorCount,
          "CFB_FAT", "CFB FAT sectorが範囲外です。");
        ensure(!fatSeen[fatSector], "CFB_FAT", "CFB FAT sectorが重複しています。");
        fatSeen[fatSector] = true;
        var fatBytes = this.sector(fatSector);
        var fatEntry;
        for (fatEntry = 0; fatEntry < this.sectorSize / 4; fatEntry += 1) {
          fatValues.push(u32(fatBytes, fatEntry * 4));
        }
      }
      ensure(fatValues.length >= this.sectorCount, "CFB_FAT",
        "CFB FATがファイルsectorを網羅していません。");
      this.fat = fatValues;
      Object.keys(fatSeen).forEach(function (sectorKey) {
        var sectorId = Number(sectorKey);
        ensure(fatValues[sectorId] === CFB_FATSECT, "CFB_FAT",
          "FAT sectorがFATSECTとしてマークされていません。");
      });
      Object.keys(difatSeen).forEach(function (sectorKey) {
        var sectorId = Number(sectorKey);
        ensure(fatValues[sectorId] === CFB_DIFSECT, "CFB_DIFAT",
          "DIFAT sectorがDIFSECTとしてマークされていません。");
      });

      var directoryIds = this.chainIds(
        firstDirectorySector,
        this.fat,
        this.sectorCount + 1,
        "CFB Directory"
      );
      if (this.majorVersion === 4 && numberOfDirectorySectors !== 0) {
        ensure(directoryIds.length === numberOfDirectorySectors, "CFB_DIRECTORY",
          "CFB Directory sector数が一致しません。");
      }
      ensure(directoryIds.length > 0, "CFB_DIRECTORY",
        "CFB Directory streamがありません。");
      var directoryParts = [];
      var directorySectorIndex;
      for (directorySectorIndex = 0;
        directorySectorIndex < directoryIds.length;
        directorySectorIndex += 1) {
        directoryParts.push(this.sector(directoryIds[directorySectorIndex]));
      }
      var directoryBytes = concatBytes(directoryParts);
      var directoryCount = Math.floor(directoryBytes.length / 128);
      ensure(directoryCount > 0, "CFB_DIRECTORY", "CFB Directory entryがありません。");
      var directoryIndex;
      for (directoryIndex = 0; directoryIndex < directoryCount; directoryIndex += 1) {
        var offset = directoryIndex * 128;
        var nameLength = u16(directoryBytes, offset + 64);
        var type = directoryBytes[offset + 66];
        ensure(type >= 0 && type <= 5 && type !== 3 && type !== 4,
          "CFB_DIRECTORY", "CFB Directory entry typeが不正です。");
        var name = "";
        if (type !== 0) {
          ensure(nameLength >= 2 && nameLength <= 64 && nameLength % 2 === 0,
            "CFB_DIRECTORY", "CFB Directory entry name lengthが不正です。");
          name = decodeUtf16Le(sliceBytes(directoryBytes, offset, nameLength - 2,
            "CFB Directory name"));
        }
        var sizeLow = u32(directoryBytes, offset + 120);
        var sizeHigh = this.majorVersion === 4 ? u32(directoryBytes, offset + 124) : 0;
        var size = sizeLow + sizeHigh * 4294967296;
        ensure(Number.isSafeInteger(size), "CFB_STREAM_SIZE",
          "CFB stream sizeがJavaScriptの安全範囲を超えています。");
        this.directory.push({
          id: directoryIndex,
          name: name,
          type: type,
          color: directoryBytes[offset + 67],
          left: u32(directoryBytes, offset + 68),
          right: u32(directoryBytes, offset + 72),
          child: u32(directoryBytes, offset + 76),
          startSector: u32(directoryBytes, offset + 116),
          size: size,
          path: "",
          parent: null
        });
      }
      this.root = this.directory.find(function (entry) {
        return entry.type === 5;
      }) || null;
      ensure(this.root && this.root.id === 0, "CFB_ROOT",
        "CFB Root Entryが見つかりません。");
      this.readDirectoryTree();

      if (numberOfMiniFatSectors > 0) {
        var miniFatIds = this.chainIds(
          firstMiniFatSector,
          this.fat,
          numberOfMiniFatSectors + 1,
          "CFB MiniFAT"
        );
        ensure(miniFatIds.length === numberOfMiniFatSectors, "CFB_MINIFAT",
          "CFB MiniFAT sector数が一致しません。");
        var miniIndex;
        for (miniIndex = 0; miniIndex < miniFatIds.length; miniIndex += 1) {
          var miniFatBytes = this.sector(miniFatIds[miniIndex]);
          var miniEntry;
          for (miniEntry = 0; miniEntry < this.sectorSize / 4; miniEntry += 1) {
            this.miniFat.push(u32(miniFatBytes, miniEntry * 4));
          }
        }
      } else {
        ensure(firstMiniFatSector === CFB_ENDOFCHAIN ||
          firstMiniFatSector === CFB_FREESECT, "CFB_MINIFAT",
        "CFB MiniFAT headerが矛盾しています。");
      }

      if (this.root.size > 0) {
        this.miniStream = this.regularChainBytes(
          this.root.startSector,
          this.root.size,
          "CFB Root mini stream"
        );
      }
    };

    CompoundFile.prototype.getStream = function (entry) {
      ensure(entry && entry.type === 2, "CFB_STREAM", "CFB stream entryではありません。");
      if (entry.size === 0) {
        return new Uint8Array(0);
      }
      ensure(entry.size <= this.bytes.length, "CFB_STREAM_SIZE",
        "CFB stream sizeがファイルサイズを超えています: " + entry.name);
      if (entry.size < this.miniStreamCutoff) {
        ensure(this.miniFat.length > 0 && this.miniStream.length > 0,
          "CFB_MINISTREAM", "CFB mini streamがありません: " + entry.name);
        var needed = Math.ceil(entry.size / this.miniSectorSize);
        var ids = this.chainIds(
          entry.startSector,
          this.miniFat,
          needed + 1,
          "CFB mini stream " + entry.name
        );
        ensure(ids.length >= needed, "CFB_MINISTREAM",
          "CFB mini stream chainが途中で終了しています: " + entry.name);
        var parts = [];
        var index;
        for (index = 0; index < needed; index += 1) {
          var offset = ids[index] * this.miniSectorSize;
          requireRange(this.miniStream, offset, this.miniSectorSize, "CFB mini sector");
          parts.push(this.miniStream.subarray(offset, offset + this.miniSectorSize));
        }
        return concatBytes(parts, entry.size);
      }
      return this.regularChainBytes(entry.startSector, entry.size, "CFB stream " + entry.name);
    };

    CompoundFile.prototype.entriesByName = function (name, type) {
      var lowered = String(name || "").toLowerCase();
      return this.directory.filter(function (entry) {
        return entry.type !== 0 &&
          (typeof type !== "number" || entry.type === type) &&
          entry.name.toLowerCase() === lowered;
      });
    };

    CompoundFile.prototype.rootStreamsByName = function (name) {
      var lowered = String(name || "").toLowerCase();
      var rootId = this.root.id;
      // Main-document detection must not inspect Office streams inside embedded storages.
      return this.directory.filter(function (entry) {
        return entry.type === 2 &&
          entry.parent === rootId &&
          entry.name.toLowerCase() === lowered;
      });
    };

    CompoundFile.prototype.firstRootStream = function (name) {
      var entries = this.rootStreamsByName(name);
      return entries.length ? entries[0] : null;
    };

    CompoundFile.prototype.hasRootStream = function (name) {
      return !!this.firstRootStream(name);
    };

    CompoundFile.prototype.rootStreamBytes = function (name) {
      var entry = this.firstRootStream(name);
      ensure(entry, "CFB_STREAM", "CFB streamが見つかりません: " + name);
      return this.getStream(entry);
    };

    CompoundFile.prototype.firstStream = function (name) {
      var entries = this.entriesByName(name, 2);
      return entries.length ? entries[0] : null;
    };

    CompoundFile.prototype.hasStream = function (name) {
      return !!this.firstStream(name);
    };

    CompoundFile.prototype.streamBytes = function (name) {
      var entry = this.firstStream(name);
      ensure(entry, "CFB_STREAM", "CFB streamが見つかりません: " + name);
      return this.getStream(entry);
    };

    CompoundFile.prototype.streams = function () {
      return this.directory.filter(function (entry) {
        return entry.type === 2;
      });
    };

    CompoundFile.prototype.isEncryptedPackage = function () {
      return this.hasRootStream("EncryptedPackage") ||
        this.hasRootStream("EncryptionInfo");
    };

    function detectCompoundFamily(cfb) {
      var hasWord = cfb.hasRootStream("WordDocument") &&
        (cfb.hasRootStream("0Table") || cfb.hasRootStream("1Table"));
      var hasExcel = cfb.hasRootStream("Workbook") || cfb.hasRootStream("Book");
      var hasPowerPoint = cfb.hasRootStream("PowerPoint Document");
      var count = (hasWord ? 1 : 0) + (hasExcel ? 1 : 0) + (hasPowerPoint ? 1 : 0);
      ensure(count === 1, "CFB_FORMAT",
        count === 0 ?
          "Office形式固有の主要streamが見つかりません。" :
          "複数のOffice主要streamがあり、内部形式を一意に判定できません。");
      return hasWord ? "word" : (hasExcel ? "excel" : "powerpoint");
    }

    function relationshipKind(type) {
      var text = String(type || "");
      var slash = text.lastIndexOf("/");
      return (slash >= 0 ? text.slice(slash + 1) : text).toLowerCase();
    }

    async function readRelationships(zip, sourcePart) {
      var relsPath = sourcePart ? relationshipPartName(sourcePart) : "_rels/.rels";
      var entry = zip.get(relsPath);
      if (!entry) {
        return [];
      }
      var documentNode = parseXml(await zip.extract(entry), relsPath);
      return elementsByLocalName(documentNode, "Relationship").map(function (element) {
        var target = element.getAttribute("Target") || "";
        var mode = element.getAttribute("TargetMode") || "";
        return {
          id: element.getAttribute("Id") || "",
          type: element.getAttribute("Type") || "",
          kind: relationshipKind(element.getAttribute("Type") || ""),
          target: target,
          external: mode.toLowerCase() === "external",
          part: mode.toLowerCase() === "external" ?
            "" :
            resolvePartTarget(sourcePart || "", target)
        };
      });
    }

    function relationshipMap(relationships) {
      var result = Object.create(null);
      relationships.forEach(function (relationship) {
        if (relationship.id && !relationship.external) {
          result[relationship.id] = relationship;
        }
      });
      return result;
    }

    function placementNumber(value) {
      var number = Number(value);
      return Number.isFinite(number) ? number : null;
    }

    function createPlacementBlock(type, text, sourcePart, sourceOrder) {
      return {
        type: type || "text",
        text: String(text == null ? "" : text).replace(/\r\n?|\n/g, "\n"),
        events: null,
        sourcePart: sourcePart || "",
        sourceOrder: Number.isFinite(sourceOrder) ? sourceOrder : 0,
        sourceId: "",
        x: null,
        y: null,
        width: null,
        height: null,
        row: null,
        column: null,
        rowOffset: null,
        columnOffset: null,
        anchorOrder: null,
        pageHint: null,
        positionKnown: false,
        keepEmpty: false
      };
    }

    function renderPlacementBlocks(blocks, separator) {
      var lines = [];
      var index;
      var blockSeparator = separator == null ? "\n" : String(separator);
      for (index = 0; index < blocks.length; index += 1) {
        var block = blocks[index];
        if (block.keepEmpty || hasMeaningfulText(block.text)) {
          lines.push(block.text);
        }
      }
      return lines.join(blockSeparator);
    }

    async function readOoxmlContentTypes(zip) {
      ensure(zip.has("[Content_Types].xml"), "OOXML_CONTENT_TYPES",
        "[Content_Types].xmlがありません。");
      var contentTypes = await zip.extractXml("[Content_Types].xml");
      var defaults = elementsByLocalName(contentTypes, "Default");
      var defaultExtensions = Object.create(null);
      defaults.forEach(function (element) {
        var extension = String(element.getAttribute("Extension") || "")
          .replace(/^\./, "")
          .toLowerCase();
        var contentType = String(element.getAttribute("ContentType") || "")
          .replace(/^\s+|\s+$/g, "")
          .toLowerCase();
        if (extension && contentType) {
          defaultExtensions[extension] = contentType;
        }
      });
      var overrides = elementsByLocalName(contentTypes, "Override");
      var overrideParts = Object.create(null);
      overrides.forEach(function (element) {
        var partName = normalizePackagePath(
          String(element.getAttribute("PartName") || "").replace(/^\//, "")
        );
        var contentType = String(element.getAttribute("ContentType") || "")
          .replace(/^\s+|\s+$/g, "")
          .toLowerCase();
        if (partName && contentType) {
          overrideParts[partName.toLowerCase()] = contentType;
        }
      });
      return {
        defaults: defaultExtensions,
        overrides: overrideParts
      };
    }

    function ooxmlContentTypeForPart(contentTypes, partName) {
      var normalized = normalizePackagePath(partName).toLowerCase();
      return contentTypes.overrides[normalized] ||
        contentTypes.defaults[extensionOf(normalized)] || "";
    }

    async function detectOoxmlFamily(zip) {
      ensure(zip.has("_rels/.rels"), "OOXML_RELATIONSHIPS",
        "パッケージrelationshipsがありません。");
      var contentTypes = await readOoxmlContentTypes(zip);
      var rootRelationships = await readRelationships(zip, "");
      var officeRelationships = rootRelationships.filter(function (relationship) {
        return relationship.kind === "officedocument" && !relationship.external;
      });
      ensure(officeRelationships.length === 1, "OOXML_OFFICE_DOCUMENT",
        "Office文書のmain relationshipを一意に特定できません。");
      var mainPart = officeRelationships[0].part.toLowerCase();
      var family = "";
      Object.keys(OOXML_MAIN_PARTS).forEach(function (candidate) {
        if (OOXML_MAIN_PARTS[candidate].toLowerCase() === mainPart) {
          family = candidate;
        }
      });
      ensure(family, "OOXML_MAIN_PART",
        "未対応のOffice main partです: " + officeRelationships[0].part);
      ensure(zip.has(OOXML_MAIN_PARTS[family]), "OOXML_MAIN_PART",
        "Office main partがZIP内にありません: " + OOXML_MAIN_PARTS[family]);
      var mainContentType = ooxmlContentTypeForPart(contentTypes, mainPart);
      ensure(mainContentType, "OOXML_CONTENT_TYPES",
        "main partのContent Type定義がありません。");
      return family;
    }

    async function extractZipPrefix(zip, prefix, category) {
      var entries = zip.listPrefix(prefix);
      var maximum = category === "media" ? MAX_MEDIA_FILES : MAX_EMBEDDING_FILES;
      ensure(entries.length <= maximum,
        category === "media" ? "MEDIA_LIMIT" : "EMBEDDING_LIMIT",
        category === "media" ?
          "メディア数が安全上限を超えたため、メディアの抽出を中止しました。" :
          "埋め込み数が安全上限を超えたため、埋め込みの抽出を中止しました。");
      var allocator = new NameAllocator();
      var output = {
        files: [],
        warnings: []
      };
      var index;
      for (index = 0; index < entries.length; index += 1) {
        checkCancelled();
        var entry = entries[index];
        try {
          var data = await zip.extract(entry);
          output.files.push({
            name: allocator.allocate(leafName(entry.name), true),
            data: data
          });
        } catch (error) {
          throwIfCategoryLimitOrCancelled(error);
          // A damaged package member must not discard earlier successful members.
          output.warnings.push(warningValue(
            "warning.packageEntryExtractionFailed",
            {
              subject: translationValue(
                category === "media" ?
                  "warning.subject.media" : "warning.subject.embedding"
              ),
              name: entry.name
            },
            error
          ));
        }
        await cooperativeYield(index + 1, 10);
      }
      return output;
    }

    function renderWordRuby(node) {
      var baseText = "";
      var rubyText = "";
      var child;
      for (child = node.firstElementChild;
        child;
        child = child.nextElementSibling) {
        if (child.localName === "rubyBase") {
          baseText += renderWordElement(child);
        } else if (child.localName === "rt") {
          rubyText += renderWordElement(child);
        }
      }
      if (baseText && rubyText) {
        return "〓" + baseText + "《" + rubyText + "》";
      }
      return baseText || rubyText;
    }

    function findWordFieldClosingParenthesis(value, openIndex) {
      if (openIndex < 0 || value.charAt(openIndex) !== "(") {
        return -1;
      }
      var depth = 0;
      var index;
      for (index = openIndex; index < value.length; index += 1) {
        if (value.charAt(index) === "(") {
          depth += 1;
        } else if (value.charAt(index) === ")") {
          depth -= 1;
          if (depth === 0) {
            return index;
          }
          if (depth < 0) {
            return -1;
          }
        }
      }
      return -1;
    }

    function splitWordFieldArguments(value) {
      var depth = 0;
      var separator = -1;
      var index;
      for (index = 0; index < value.length; index += 1) {
        if (value.charAt(index) === "(") {
          depth += 1;
        } else if (value.charAt(index) === ")") {
          if (depth === 0) {
            return null;
          }
          depth -= 1;
        } else if (value.charAt(index) === "," &&
          depth === 0 && separator < 0) {
          separator = index;
        }
      }
      if (depth !== 0 || separator < 0) {
        return null;
      }
      return [
        value.substring(0, separator),
        value.substring(separator + 1)
      ];
    }

    function parseWordEqRubyInstruction(instruction) {
      var value = String(instruction || "");
      if (!/^\s*EQ(?:\s|\\)/i.test(value)) {
        return "";
      }
      var overlayMatch = /\\o\s*\\ad\s*\(/i.exec(value);
      if (!overlayMatch) {
        return "";
      }
      var overlayOpen = overlayMatch.index + overlayMatch[0].length - 1;
      var overlayClose = findWordFieldClosingParenthesis(value, overlayOpen);
      if (overlayClose < 0) {
        return "";
      }
      var argumentsValue = value.substring(overlayOpen + 1, overlayClose);
      var argumentsList = splitWordFieldArguments(argumentsValue);
      if (!argumentsList) {
        return "";
      }
      var rubyMatch = /\\s\s*\\up\s*-?\d+\s*\(/i.exec(argumentsList[0]);
      if (!rubyMatch) {
        return "";
      }
      var rubyOpen = rubyMatch.index + rubyMatch[0].length - 1;
      var rubyClose = findWordFieldClosingParenthesis(argumentsList[0], rubyOpen);
      if (rubyClose < 0) {
        return "";
      }
      if (argumentsList[0].substring(0, rubyMatch.index).replace(/\s/g, "") ||
        argumentsList[0].substring(rubyClose + 1).replace(/\s/g, "")) {
        return "";
      }
      var rubyText = argumentsList[0].substring(rubyOpen + 1, rubyClose)
        .replace(/^\s+|\s+$/g, "");
      var baseText = argumentsList[1].replace(/^\s+|\s+$/g, "");
      if (!baseText || !rubyText) {
        return "";
      }
      return "〓" + baseText + "《" + rubyText + "》";
    }

    function addWordParagraphTextToken(tokens, value) {
      if (!value) {
        return;
      }
      var last = tokens.length ? tokens[tokens.length - 1] : null;
      if (last && last.kind === "text") {
        last.value += value;
      } else {
        tokens.push({ kind: "text", value: value });
      }
    }

    function wordTextIsWhitespaceOnly(value) {
      return /^[\t \u3000]*$/.test(String(value || ""));
    }

    function appendWordInlineShapeLineBreak(token) {
      if (token && !/\n$/.test(token.value)) {
        token.value += "\n";
        if (token.events) {
          addWordRenderEvent(token.events, "lineBreak", "");
        }
      }
    }

    function normalizeWordInlineShapeTokenBoundaries(tokens) {
      var separated = [];
      var index;
      for (index = 0; index < tokens.length; index += 1) {
        var token = tokens[index];
        if (token.kind !== "inlineShape") {
          separated.push(token);
          continue;
        }
        var previous = separated.length ? separated[separated.length - 1] : null;
        if (previous && previous.kind === "text" &&
          wordTextIsWhitespaceOnly(previous.value) &&
          separated.length > 1 &&
          separated[separated.length - 2].kind === "inlineShape") {
          separated.pop();
          previous = separated[separated.length - 1];
        }
        if (previous && previous.kind === "inlineShape") {
          appendWordInlineShapeLineBreak(previous);
        }
        separated.push({
          kind: "inlineShape",
          value: token.value,
          events: token.events || null
        });
      }

      var normalized = [];
      for (index = 0; index < separated.length; index += 1) {
        var separatedToken = separated[index];
        if (separatedToken.kind === "text") {
          addWordParagraphTextToken(normalized, separatedToken.value);
        } else {
          normalized.push(separatedToken);
        }
      }
      return normalized;
    }

    function wordNodeIsExcluded(node, excludedNodes) {
      var index;
      for (index = 0; excludedNodes && index < excludedNodes.length; index += 1) {
        if (excludedNodes[index] === node) {
          return true;
        }
      }
      return false;
    }

    function collectWordParagraphTokens(node, tokens, paragraphRoot, excludedNodes) {
      if (!node || node.nodeType !== 1) {
        return;
      }
      if (wordNodeIsExcluded(node, excludedNodes)) {
        return;
      }
      var name = node.localName;
      if (name === "chart" || name === "altChunk") {
        tokens.push({
          kind: name === "chart" ? "chartReference" : "altChunkReference",
          value: String(relationshipIdAttribute(node) || "")
        });
        return;
      }
      if (wordNodeIsInlineShape(node)) {
        var inlineEvents = renderWordElementEvents(node, true);
        trimWordRenderEventLineBreaks(inlineEvents);
        var inlineText = wordRenderEventsToText(inlineEvents, true);
        if (hasMeaningfulText(inlineText) ||
          wordRenderEventsHaveReferences(inlineEvents)) {
          tokens.push({
            kind: "inlineShape",
            value: inlineText,
            events: inlineEvents
          });
        }
        return;
      }
      if (name === "p" && node !== paragraphRoot) {
        addWordParagraphTextToken(tokens, renderWordParagraph(node, excludedNodes));
        return;
      }
      if (name === "t") {
        addWordParagraphTextToken(tokens, node.textContent || "");
        return;
      }
      if (name === "ruby") {
        addWordParagraphTextToken(tokens, renderWordRuby(node));
        return;
      }
      if (name === "fldChar") {
        tokens.push({
          kind: "field",
          value: String(attributeByLocalName(node, "fldCharType") || "")
            .toLowerCase()
        });
        return;
      }
      if (name === "instrText") {
        tokens.push({
          kind: "instruction",
          value: node.textContent || ""
        });
        return;
      }
      if (name === "del" || name === "delText" || name === "moveFrom") {
        return;
      }
      if (name === "footnoteRef" || name === "endnoteRef") {
        return;
      }
      if (name === "footnoteReference" || name === "endnoteReference") {
        tokens.push({
          kind: name,
          value: String(attributeByLocalName(node, "id") || "")
        });
        return;
      }
      if (name === "AlternateContent") {
        var preferred = null;
        var alternateChild;
        for (alternateChild = node.firstElementChild;
          alternateChild;
          alternateChild = alternateChild.nextElementSibling) {
          if (alternateChild.localName === "Choice") {
            preferred = alternateChild;
            break;
          }
          if (!preferred && alternateChild.localName === "Fallback") {
            preferred = alternateChild;
          }
        }
        if (preferred) {
          collectWordParagraphTokens(preferred, tokens, paragraphRoot, excludedNodes);
        }
        return;
      }
      if (name === "tab" || name === "ptab") {
        addWordParagraphTextToken(tokens, "\t");
        return;
      }
      if (name === "br") {
        var breakType = String(attributeByLocalName(node, "type") || "")
          .toLowerCase();
        tokens.push({
          kind: breakType === "page" ? "pageBreak" : "lineBreak",
          value: breakType || "line"
        });
        return;
      }
      if (name === "lastRenderedPageBreak") {
        tokens.push({ kind: "pageBreak", value: "lastRendered" });
        return;
      }
      if (name === "cr") {
        tokens.push({ kind: "lineBreak", value: "line" });
        return;
      }
      if (name === "noBreakHyphen") {
        addWordParagraphTextToken(tokens, "\u2011");
        return;
      }
      if (name === "softHyphen") {
        addWordParagraphTextToken(tokens, "\u00AD");
        return;
      }
      if (name === "proofErr" || name === "bookmarkStart" || name === "bookmarkEnd") {
        return;
      }
      var child;
      for (child = node.firstChild; child; child = child.nextSibling) {
        collectWordParagraphTokens(child, tokens, paragraphRoot, excludedNodes);
      }
    }

    function addWordRenderEvent(events, kind, value) {
      var eventKind = kind || "text";
      var eventValue = String(value == null ? "" : value);
      if (eventKind === "text" && !eventValue) {
        return;
      }
      var last = events.length ? events[events.length - 1] : null;
      if (eventKind === "text" && last && last.kind === "text") {
        last.value += eventValue;
        return;
      }
      events.push({ kind: eventKind, value: eventValue });
    }

    function appendWordRenderEvents(target, source) {
      var index;
      for (index = 0; index < source.length; index += 1) {
        addWordRenderEvent(target, source[index].kind, source[index].value);
      }
    }

    function appendWordFieldEvents(fields, output, events) {
      if (!events.length) {
        return;
      }
      if (!fields.length) {
        appendWordRenderEvents(output, events);
        return;
      }
      var field = fields[fields.length - 1];
      if (field.phase === "result") {
        appendWordRenderEvents(field.result, events);
      } else {
        appendWordRenderEvents(field.fallback, events);
      }
    }

    function wordRenderEventsToText(events, pageBreaksAsLines) {
      var output = "";
      var index;
      for (index = 0; index < events.length; index += 1) {
        var event = events[index];
        if (event.kind === "text") {
          output += event.value;
        } else if (event.kind === "lineBreak" ||
          (pageBreaksAsLines && event.kind === "pageBreak")) {
          output += "\n";
        }
      }
      return output;
    }

    function wordRenderEventsHaveReferences(events) {
      var index;
      for (index = 0; index < events.length; index += 1) {
        if (events[index].kind === "footnoteReference" ||
          events[index].kind === "endnoteReference" ||
          events[index].kind === "chartReference" ||
          events[index].kind === "altChunkReference") {
          return true;
        }
      }
      return false;
    }

    function wordRenderEventsHavePageBreaks(events) {
      var index;
      for (index = 0; index < events.length; index += 1) {
        if (events[index].kind === "pageBreak") {
          return true;
        }
      }
      return false;
    }

    function trimWordRenderEventLineBreaks(events) {
      while (events.length) {
        var last = events[events.length - 1];
        if (last.kind === "lineBreak") {
          events.pop();
          continue;
        }
        if (last.kind === "text" && /\n$/.test(last.value)) {
          last.value = last.value.replace(/\n+$/g, "");
          if (!last.value) {
            events.pop();
          }
          continue;
        }
        break;
      }
      return events;
    }

    function trimWordParagraphBlockEvents(events) {
      var pageBreak = null;
      if (events.length && events[events.length - 1].kind === "pageBreak") {
        pageBreak = events.pop();
      }
      trimWordRenderEventLineBreaks(events);
      if (pageBreak) {
        events.push(pageBreak);
      }
      return events;
    }

    function wordOnOffEnabled(node) {
      if (!node) {
        return false;
      }
      var value = String(attributeByLocalName(node, "val") || "").toLowerCase();
      return value !== "0" && value !== "false" &&
        value !== "off" && value !== "no";
    }

    function wordParagraphProperties(node) {
      var child;
      for (child = node ? node.firstElementChild : null;
        child;
        child = child.nextElementSibling) {
        if (child.localName === "pPr") {
          return child;
        }
      }
      return null;
    }

    function wordParagraphPageBreakBefore(node) {
      var properties = wordParagraphProperties(node);
      return wordOnOffEnabled(firstChildByLocalName(properties, "pageBreakBefore"));
    }

    function wordParagraphSectionBreakKind(node) {
      var properties = wordParagraphProperties(node);
      var section = firstChildByLocalName(properties, "sectPr");
      if (!section) {
        return "";
      }
      var type = firstChildByLocalName(section, "type");
      return String(type ? attributeByLocalName(type, "val") || "nextPage" : "nextPage")
        .toLowerCase();
    }

    function wordSectionKindCreatesPage(kind) {
      return kind === "nextpage" || kind === "oddpage" || kind === "evenpage";
    }

    function renderWordParagraphEvents(node, excludedNodes) {
      var tokens = [];
      collectWordParagraphTokens(node, tokens, node, excludedNodes);
      tokens = normalizeWordInlineShapeTokenBoundaries(tokens);
      var fields = [];
      var output = [];
      if (wordParagraphPageBreakBefore(node)) {
        addWordRenderEvent(output, "pageBreak", "pageBreakBefore");
      }
      var index;
      for (index = 0; index < tokens.length; index += 1) {
        var token = tokens[index];
        if (token.kind === "text") {
          appendWordFieldEvents(fields, output, [
            { kind: "text", value: token.value }
          ]);
          continue;
        }
        if (token.kind === "inlineShape") {
          appendWordFieldEvents(
            fields,
            output,
            token.events || [{ kind: "text", value: token.value }]
          );
          continue;
        }
        if (token.kind === "instruction") {
          if (fields.length && fields[fields.length - 1].phase === "instruction") {
            fields[fields.length - 1].instruction += token.value;
          }
          continue;
        }
        if (token.value === "begin") {
          fields.push({
            instruction: "",
            fallback: [],
            result: [],
            phase: "instruction"
          });
        } else if (token.value === "separate" && fields.length) {
          fields[fields.length - 1].phase = "result";
        } else if (token.value === "end" && fields.length) {
          var completed = fields.pop();
          var ruby = parseWordEqRubyInstruction(completed.instruction);
          var rendered = ruby ?
            [{ kind: "text", value: ruby }] :
            (completed.result.length ? completed.result : completed.fallback);
          appendWordFieldEvents(fields, output, rendered);
        } else if (token.kind === "lineBreak" ||
          token.kind === "pageBreak" ||
          token.kind === "footnoteReference" ||
          token.kind === "endnoteReference" ||
          token.kind === "chartReference" ||
          token.kind === "altChunkReference") {
          appendWordFieldEvents(fields, output, [
            { kind: token.kind, value: token.value }
          ]);
        }
      }
      while (fields.length) {
        var incomplete = fields.pop();
        appendWordFieldEvents(
          fields,
          output,
          incomplete.result.length ? incomplete.result : incomplete.fallback
        );
      }
      trimWordRenderEventLineBreaks(output);
      addWordRenderEvent(output, "lineBreak", "");
      var sectionKind = wordParagraphSectionBreakKind(node);
      if (wordSectionKindCreatesPage(sectionKind)) {
        addWordRenderEvent(output, "pageBreak", "section:" + sectionKind);
      }
      return output;
    }

    function renderWordParagraph(node, excludedNodes) {
      return wordRenderEventsToText(
        renderWordParagraphEvents(node, excludedNodes),
        true
      ).replace(/\n+$/g, "") + "\n";
    }

    function renderWordElement(node) {
      if (!node) {
        return "";
      }
      if (node.nodeType !== 1) {
        return "";
      }
      var name = node.localName;
      if (name === "t") {
        return node.textContent || "";
      }
      if (name === "ruby") {
        return renderWordRuby(node);
      }
      if (name === "del" || name === "delText" || name === "moveFrom") {
        return "";
      }
      if (name === "footnoteRef" || name === "endnoteRef" ||
        name === "footnoteReference" || name === "endnoteReference") {
        return "";
      }
      if (name === "AlternateContent") {
        var preferred = null;
        var alternateChild;
        for (alternateChild = node.firstElementChild;
          alternateChild;
          alternateChild = alternateChild.nextElementSibling) {
          if (alternateChild.localName === "Choice") {
            preferred = alternateChild;
            break;
          }
          if (!preferred && alternateChild.localName === "Fallback") {
            preferred = alternateChild;
          }
        }
        return preferred ? renderWordElement(preferred) : "";
      }
      if (name === "tab" || name === "ptab") {
        return "\t";
      }
      if (name === "br" || name === "cr" || name === "lastRenderedPageBreak") {
        return "\n";
      }
      if (name === "noBreakHyphen") {
        return "\u2011";
      }
      if (name === "softHyphen") {
        return "\u00AD";
      }
      if (name === "instrText" || name === "fldChar" ||
        name === "proofErr" || name === "bookmarkStart" || name === "bookmarkEnd") {
        return "";
      }
      if (name === "p") {
        return renderWordParagraph(node);
      }
      var value = "";
      var child;
      for (child = node.firstChild; child; child = child.nextSibling) {
        value += renderWordElement(child);
      }
      if (name === "tc") {
        return value.replace(/[\t\n]+$/g, "") + "\t";
      }
      if (name === "tr") {
        return value.replace(/[\t\n]+$/g, "") + "\n";
      }
      return value;
    }

    function trimWordContainerEventEnd(events) {
      var index = events.length - 1;
      while (index >= 0) {
        var event = events[index];
        if (event.kind === "footnoteReference" ||
          event.kind === "endnoteReference") {
          index -= 1;
          continue;
        }
        if (event.kind === "lineBreak") {
          events.splice(index, 1);
          index -= 1;
          continue;
        }
        if (event.kind === "text") {
          event.value = event.value.replace(/[\t\n]+$/g, "");
          if (!event.value) {
            events.splice(index, 1);
            index -= 1;
            continue;
          }
        }
        break;
      }
      return events;
    }

    function renderWordElementEvents(node, pageBreaksAsLines) {
      var events = [];
      if (!node || node.nodeType !== 1) {
        return events;
      }
      var name = node.localName;
      if (name === "shape" && wordVmlShapeHasTextContent(node)) {
        var wordArtText = extractWordVmlTextPathText(node);
        if (hasMeaningfulText(wordArtText)) {
          addWordRenderEvent(events, "text", wordArtText);
          return events;
        }
      }
      if (name === "textpath") {
        addWordRenderEvent(events, "text", extractWordVmlTextPathText(node));
        return events;
      }
      if (name === "chart" || name === "altChunk") {
        addWordRenderEvent(
          events,
          name === "chart" ? "chartReference" : "altChunkReference",
          String(relationshipIdAttribute(node) || "")
        );
        return events;
      }
      if (name === "t") {
        addWordRenderEvent(events, "text", node.textContent || "");
        return events;
      }
      if (name === "ruby") {
        addWordRenderEvent(events, "text", renderWordRuby(node));
        return events;
      }
      if (name === "del" || name === "delText" || name === "moveFrom" ||
        name === "footnoteRef" || name === "endnoteRef") {
        return events;
      }
      if (name === "footnoteReference" || name === "endnoteReference") {
        addWordRenderEvent(
          events,
          name,
          String(attributeByLocalName(node, "id") || "")
        );
        return events;
      }
      if (name === "AlternateContent") {
        var preferred = preferredAlternateContentChild(node);
        return preferred ?
          renderWordElementEvents(preferred, pageBreaksAsLines) :
          events;
      }
      if (name === "tab" || name === "ptab") {
        addWordRenderEvent(events, "text", "\t");
        return events;
      }
      if (name === "br") {
        var breakType = String(attributeByLocalName(node, "type") || "")
          .toLowerCase();
        addWordRenderEvent(
          events,
          breakType === "page" && !pageBreaksAsLines ? "pageBreak" : "lineBreak",
          breakType || "line"
        );
        return events;
      }
      if (name === "lastRenderedPageBreak") {
        addWordRenderEvent(
          events,
          pageBreaksAsLines ? "lineBreak" : "pageBreak",
          "lastRendered"
        );
        return events;
      }
      if (name === "cr") {
        addWordRenderEvent(events, "lineBreak", "line");
        return events;
      }
      if (name === "noBreakHyphen") {
        addWordRenderEvent(events, "text", "\u2011");
        return events;
      }
      if (name === "softHyphen") {
        addWordRenderEvent(events, "text", "\u00AD");
        return events;
      }
      if (name === "instrText" || name === "fldChar" ||
        name === "proofErr" || name === "bookmarkStart" ||
        name === "bookmarkEnd") {
        return events;
      }
      if (name === "p") {
        events = renderWordParagraphEvents(node);
        if (pageBreaksAsLines) {
          var paragraphIndex;
          for (paragraphIndex = 0;
            paragraphIndex < events.length;
            paragraphIndex += 1) {
            if (events[paragraphIndex].kind === "pageBreak") {
              events[paragraphIndex].kind = "lineBreak";
            }
          }
        }
        return events;
      }
      var child;
      for (child = node.firstChild; child; child = child.nextSibling) {
        appendWordRenderEvents(
          events,
          renderWordElementEvents(child, pageBreaksAsLines)
        );
      }
      if (name === "tc") {
        trimWordContainerEventEnd(events);
        addWordRenderEvent(events, "text", "\t");
      } else if (name === "tr") {
        trimWordContainerEventEnd(events);
        addWordRenderEvent(events, "lineBreak", "row");
      }
      return events;
    }

    function preferredAlternateContentChild(node) {
      var preferred = null;
      var child;
      for (child = node ? node.firstElementChild : null;
        child;
        child = child.nextElementSibling) {
        if (child.localName === "Choice") {
          return child;
        }
        if (!preferred && child.localName === "Fallback") {
          preferred = child;
        }
      }
      return preferred;
    }

    function parseCssLength(value) {
      var match = /^\s*(-?(?:\d+(?:\.\d*)?|\.\d+))\s*(pt|px|in|cm|mm|pc)?\s*$/i
        .exec(String(value == null ? "" : value));
      if (!match) {
        return null;
      }
      var number = Number(match[1]);
      var unit = (match[2] || "pt").toLowerCase();
      if (unit === "px") {
        number *= 0.75;
      } else if (unit === "in") {
        number *= 72;
      } else if (unit === "cm") {
        number *= 72 / 2.54;
      } else if (unit === "mm") {
        number *= 72 / 25.4;
      } else if (unit === "pc") {
        number *= 12;
      }
      return Number.isFinite(number) ? number : null;
    }

    function cssStyleMap(value) {
      var result = Object.create(null);
      String(value || "").split(";").forEach(function (declaration) {
        var colon = declaration.indexOf(":");
        if (colon < 0) {
          return;
        }
        var name = declaration.slice(0, colon)
          .replace(/^\s+|\s+$/g, "")
          .toLowerCase();
        if (name) {
          result[name] = declaration.slice(colon + 1).replace(/^\s+|\s+$/g, "");
        }
      });
      return result;
    }

    function wordNodeIsVmlShape(node) {
      return !!node && node.localName === "shape" &&
        String(node.namespaceURI || "").toLowerCase().indexOf("vml") >= 0;
    }

    function wordVmlTextPathElements(node) {
      var candidates = [];
      if (node && node.localName === "textpath") {
        candidates.push(node);
      } else if (node) {
        candidates = elementsByLocalName(node, "textpath");
      }
      return candidates.filter(function (candidate) {
        return String(candidate.namespaceURI || "").toLowerCase()
          .indexOf("vml") >= 0;
      });
    }

    function extractWordVmlTextPathText(node) {
      var textPaths = wordVmlTextPathElements(node);
      var output = [];
      var seen = Object.create(null);
      var index;
      for (index = 0; index < textPaths.length; index += 1) {
        var textPath = textPaths[index];
        var value = textPath.hasAttribute("string") ?
          textPath.getAttribute("string") : textPath.textContent;
        value = repairSurrogates(String(value || ""));
        if (hasMeaningfulText(value) && !seen[value]) {
          seen[value] = true;
          output.push(value);
        }
      }
      return output.join("\n");
    }

    function wordVmlShapeHasTextContent(node) {
      return wordNodeIsVmlShape(node) &&
        (elementsByLocalName(node, "textbox").length > 0 ||
          wordVmlTextPathElements(node).length > 0);
    }

    function wordVmlShapeIsFloating(node) {
      if (!wordVmlShapeHasTextContent(node)) {
        return false;
      }
      var style = cssStyleMap(node.getAttribute("style") || "");
      var position = String(style.position || "").toLowerCase();
      return position === "absolute" || position === "relative" ||
        Object.prototype.hasOwnProperty.call(style, "left") ||
        Object.prototype.hasOwnProperty.call(style, "top") ||
        Object.prototype.hasOwnProperty.call(style, "margin-left") ||
        Object.prototype.hasOwnProperty.call(style, "margin-top");
    }

    function wordVmlShapeIsInline(node) {
      return wordVmlShapeHasTextContent(node) &&
        !wordVmlShapeIsFloating(node);
    }

    function wordNodeIsDrawingInline(node) {
      return !!node && node.localName === "inline" &&
        String(node.namespaceURI || "").toLowerCase()
          .indexOf("wordprocessingdrawing") >= 0;
    }

    function wordNodeIsInlineShape(node) {
      var drawingContainer = !!node && node.localName === "drawing" &&
        !elementsByLocalName(node, "anchor").length &&
        elementsByLocalName(node, "txBody").length > 0;
      return wordNodeIsDrawingInline(node) ||
        drawingContainer ||
        wordVmlShapeIsInline(node);
    }

    function wordNodeIsDrawingAnchor(node) {
      return !!node && node.localName === "anchor" &&
        String(node.namespaceURI || "").toLowerCase()
          .indexOf("wordprocessingdrawing") >= 0;
    }

    function collectWordFloatingShapes(paragraph) {
      var result = [];
      function walk(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "AlternateContent") {
          var preferred = preferredAlternateContentChild(node);
          if (preferred) {
            walk(preferred);
          }
          return;
        }
        if (wordNodeIsDrawingAnchor(node) || wordVmlShapeIsFloating(node)) {
          result.push(node);
          return;
        }
        var child;
        for (child = node.firstElementChild; child; child = child.nextElementSibling) {
          walk(child);
        }
      }
      walk(paragraph);
      return result;
    }

    function wordPositionComponent(anchor, name) {
      var position = firstElementByLocalName(anchor, name);
      if (!position) {
        return { value: null, relativeFrom: "" };
      }
      var offset = firstElementByLocalName(position, "posOffset");
      var value = offset ? placementNumber(offset.textContent) : null;
      if (value === null) {
        var align = firstElementByLocalName(position, "align");
        var alignValue = align ? String(align.textContent || "").toLowerCase() : "";
        var ranks = {
          left: 0,
          top: 0,
          inside: 0,
          center: 0.5,
          right: 1,
          bottom: 1,
          outside: 1
        };
        value = Object.prototype.hasOwnProperty.call(ranks, alignValue) ?
          ranks[alignValue] : null;
      }
      return {
        value: value,
        relativeFrom: position.getAttribute("relativeFrom") || ""
      };
    }

    function applyWordShapePosition(block, shape) {
      if (wordNodeIsDrawingAnchor(shape)) {
        var horizontal = wordPositionComponent(shape, "positionH");
        var vertical = wordPositionComponent(shape, "positionV");
        block.x = horizontal.value;
        block.y = vertical.value;
        block.relativeFromH = horizontal.relativeFrom;
        block.relativeFromV = vertical.relativeFrom;
        var extent = firstElementByLocalName(shape, "extent");
        if (extent) {
          block.width = placementNumber(extent.getAttribute("cx"));
          block.height = placementNumber(extent.getAttribute("cy"));
        }
        block.positionKnown = block.x !== null || block.y !== null;
        return;
      }
      var style = cssStyleMap(shape.getAttribute("style") || "");
      block.x = parseCssLength(style.left);
      if (block.x === null) {
        block.x = parseCssLength(style["margin-left"]);
      }
      block.y = parseCssLength(style.top);
      if (block.y === null) {
        block.y = parseCssLength(style["margin-top"]);
      }
      block.width = parseCssLength(style.width);
      block.height = parseCssLength(style.height);
      block.positionKnown = block.x !== null || block.y !== null;
    }

    function sortWordAnchorBlocks(blocks) {
      return blocks.sort(function (left, right) {
        if (left.positionKnown !== right.positionKnown) {
          return left.positionKnown ? -1 : 1;
        }
        if (left.positionKnown) {
          var leftY = left.y === null ? Infinity : left.y;
          var rightY = right.y === null ? Infinity : right.y;
          if (leftY !== rightY) {
            return leftY - rightY;
          }
          var leftX = left.x === null ? Infinity : left.x;
          var rightX = right.x === null ? Infinity : right.x;
          if (leftX !== rightX) {
            return leftX - rightX;
          }
        }
        return left.sourceOrder - right.sourceOrder;
      });
    }

    function collectWordPlacementBlocks(documentNode, sourcePart) {
      var blocks = [];
      var order = 0;
      var anchorOrder = 0;

      function addShape(shape, currentAnchorOrder) {
        var events = renderWordElementEvents(shape, true);
        trimWordRenderEventLineBreaks(events);
        var text = wordRenderEventsToText(events, true);
        if (!hasMeaningfulText(text) && !wordRenderEventsHaveReferences(events)) {
          return null;
        }
        var block = createPlacementBlock("shape", text, sourcePart, order);
        block.events = events;
        order += 1;
        block.anchorOrder = currentAnchorOrder;
        block.sourceId = sourcePart + "#shape-" + block.sourceOrder;
        applyWordShapePosition(block, shape);
        return block;
      }

      function walk(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "AlternateContent") {
          var preferred = preferredAlternateContentChild(node);
          if (preferred) {
            walk(preferred);
          }
          return;
        }
        if (node.localName === "altChunk" || node.localName === "chart") {
          var referenceKind = node.localName === "chart" ?
            "chartReference" : "altChunkReference";
          var referenceBlock = createPlacementBlock(
            referenceKind,
            "",
            sourcePart,
            order
          );
          referenceBlock.events = [{
            kind: referenceKind,
            value: String(relationshipIdAttribute(node) || "")
          }];
          referenceBlock.sourceId = sourcePart + "#reference-" + order;
          blocks.push(referenceBlock);
          order += 1;
          return;
        }
        if (node.localName === "p") {
          var currentAnchorOrder = anchorOrder;
          anchorOrder += 1;
          var shapes = collectWordFloatingShapes(node);
          var paragraphEvents = renderWordParagraphEvents(node, shapes);
          trimWordParagraphBlockEvents(paragraphEvents);
          var paragraphText = wordRenderEventsToText(paragraphEvents, true);
          if (hasMeaningfulText(paragraphText) || !shapes.length ||
            wordRenderEventsHaveReferences(paragraphEvents) ||
            wordRenderEventsHavePageBreaks(paragraphEvents)) {
            var paragraphBlock = createPlacementBlock(
              "paragraph",
              paragraphText,
              sourcePart,
              order
            );
            paragraphBlock.events = paragraphEvents;
            order += 1;
            paragraphBlock.anchorOrder = currentAnchorOrder;
            paragraphBlock.keepEmpty = !hasMeaningfulText(paragraphText);
            blocks.push(paragraphBlock);
          }
          var shapeBlocks = [];
          var shapeIndex;
          for (shapeIndex = 0; shapeIndex < shapes.length; shapeIndex += 1) {
            var shapeBlock = addShape(shapes[shapeIndex], currentAnchorOrder);
            if (shapeBlock) {
              shapeBlocks.push(shapeBlock);
            }
          }
          Array.prototype.push.apply(blocks, sortWordAnchorBlocks(shapeBlocks));
          return;
        }
        if (node.localName === "tbl") {
          var tableEvents = renderWordElementEvents(node, false);
          trimWordRenderEventLineBreaks(tableEvents);
          var tableText = wordRenderEventsToText(tableEvents, true);
          if (hasMeaningfulText(tableText) ||
            wordRenderEventsHaveReferences(tableEvents)) {
            var tableBlock = createPlacementBlock(
              "table",
              tableText,
              sourcePart,
              order
            );
            tableBlock.events = tableEvents;
            blocks.push(tableBlock);
            order += 1;
          }
          return;
        }
        if (wordNodeIsDrawingAnchor(node) || wordVmlShapeIsFloating(node)) {
          var detachedShape = addShape(node, anchorOrder);
          if (detachedShape) {
            blocks.push(detachedShape);
          }
          return;
        }
        var child;
        for (child = node.firstElementChild; child; child = child.nextElementSibling) {
          walk(child);
        }
      }

      if (documentNode && documentNode.documentElement) {
        walk(documentNode.documentElement);
      }
      return blocks;
    }

    function addWordExtractionWarning(warnings, seen, key, message) {
      var warningKey = String(key || message);
      if (!seen[warningKey]) {
        seen[warningKey] = true;
        warnings.push(message);
      }
    }

    function addUniqueWordReferenceText(output, seen, value) {
      var lines = repairSurrogates(String(value || ""))
        .replace(/\r\n?/g, "\n")
        .split("\n");
      var index;
      for (index = 0; index < lines.length; index += 1) {
        var text = lines[index].replace(/^[ \t]+|[ \t]+$/g, "");
        if (hasMeaningfulText(text) && !seen[text]) {
          seen[text] = true;
          output.push(text);
        }
      }
    }

    function wordChartDrawingText(root) {
      var output = [];
      var paragraphs = elementsByLocalName(root, "p");
      var paragraphIndex;
      for (paragraphIndex = 0;
        paragraphIndex < paragraphs.length;
        paragraphIndex += 1) {
        var paragraph = paragraphs[paragraphIndex];
        if (String(paragraph.namespaceURI || "").toLowerCase()
          .indexOf("drawingml") < 0) {
          continue;
        }
        var textNodes = elementsByLocalName(paragraph, "t");
        var paragraphText = "";
        var textIndex;
        for (textIndex = 0; textIndex < textNodes.length; textIndex += 1) {
          if (String(textNodes[textIndex].namespaceURI || "").toLowerCase()
            .indexOf("drawingml") >= 0) {
            paragraphText += textNodes[textIndex].textContent || "";
          }
        }
        if (hasMeaningfulText(paragraphText)) {
          output.push(paragraphText);
        }
      }
      if (!output.length) {
        var fallbackNodes = elementsByLocalName(root, "t");
        var fallbackText = "";
        var fallbackIndex;
        for (fallbackIndex = 0;
          fallbackIndex < fallbackNodes.length;
          fallbackIndex += 1) {
          if (String(fallbackNodes[fallbackIndex].namespaceURI || "")
            .toLowerCase().indexOf("drawingml") >= 0) {
            fallbackText += fallbackNodes[fallbackIndex].textContent || "";
          }
        }
        if (hasMeaningfulText(fallbackText)) {
          output.push(fallbackText);
        }
      }
      return output.join("\n");
    }

    function appendWordChartValues(container, output, seen) {
      if (!container) {
        return;
      }
      var values = elementsByLocalName(container, "v");
      var index;
      for (index = 0; index < values.length; index += 1) {
        addUniqueWordReferenceText(output, seen, values[index].textContent || "");
      }
    }

    function extractWordChartText(documentNode) {
      var output = [];
      var seen = Object.create(null);
      var title = firstElementByLocalName(documentNode, "title");
      if (title) {
        var titleText = wordChartDrawingText(title);
        if (hasMeaningfulText(titleText)) {
          addUniqueWordReferenceText(output, seen, titleText);
        } else {
          appendWordChartValues(title, output, seen);
        }
      }

      var series = elementsByLocalName(documentNode, "ser");
      var index;
      for (index = 0; index < series.length; index += 1) {
        appendWordChartValues(
          firstChildByLocalName(series[index], "cat"),
          output,
          seen
        );
      }
      for (index = 0; index < series.length; index += 1) {
        appendWordChartValues(
          firstChildByLocalName(series[index], "tx"),
          output,
          seen
        );
      }

      var dataLabels = elementsByLocalName(documentNode, "dLbl");
      for (index = 0; index < dataLabels.length; index += 1) {
        var labelText = firstChildByLocalName(dataLabels[index], "tx");
        if (labelText) {
          var drawingLabel = wordChartDrawingText(labelText);
          if (hasMeaningfulText(drawingLabel)) {
            addUniqueWordReferenceText(output, seen, drawingLabel);
          } else {
            appendWordChartValues(labelText, output, seen);
          }
        }
      }
      return output.join("\n");
    }

    function decodeUtf16Be(bytes) {
      var evenLength = bytes.length - (bytes.length % 2);
      var swapped = new Uint8Array(evenLength);
      var index;
      for (index = 0; index < evenLength; index += 2) {
        swapped[index] = bytes[index + 1];
        swapped[index + 1] = bytes[index];
      }
      return decodeUtf16Le(swapped);
    }

    function decodeWordAltChunkBytes(bytes) {
      var text;
      if (bytes.length >= 3 && bytes[0] === 0xEF &&
        bytes[1] === 0xBB && bytes[2] === 0xBF) {
        text = decodeUtf8(bytes.subarray(3), false);
      } else if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
        text = decodeUtf16Le(bytes.subarray(2));
      } else if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
        text = decodeUtf16Be(bytes.subarray(2));
      } else {
        text = decodeUtf8(bytes, false);
      }
      return repairSurrogates(String(text || "")).replace(/^\uFEFF/, "");
    }

    function wordAltChunkHtmlElementIsExcluded(element) {
      var excluded = {
        script: true,
        style: true,
        noscript: true,
        template: true,
        object: true,
        iframe: true,
        canvas: true,
        svg: true
      };
      var name = String(element.localName || element.tagName || "").toLowerCase();
      if (excluded[name] || element.hasAttribute("hidden") ||
        String(element.getAttribute("aria-hidden") || "").toLowerCase() ===
          "true") {
        return true;
      }
      var style = String(element.getAttribute("style") || "").toLowerCase();
      return /(?:^|;)\s*display\s*:\s*none(?:\s*!important)?\s*(?:;|$)/
        .test(style) ||
        /(?:^|;)\s*visibility\s*:\s*hidden(?:\s*!important)?\s*(?:;|$)/
          .test(style);
    }

    function extractWordAltChunkHtmlText(source) {
      var documentNode = new DOMParser().parseFromString(
        String(source || ""),
        "text/html"
      );
      var root = documentNode.body || documentNode.documentElement;
      if (!root) {
        return "";
      }
      var blocks = {
        address: true,
        article: true,
        aside: true,
        blockquote: true,
        dd: true,
        div: true,
        dl: true,
        dt: true,
        fieldset: true,
        figcaption: true,
        figure: true,
        footer: true,
        form: true,
        h1: true,
        h2: true,
        h3: true,
        h4: true,
        h5: true,
        h6: true,
        header: true,
        hr: true,
        li: true,
        main: true,
        nav: true,
        ol: true,
        p: true,
        pre: true,
        section: true,
        table: true,
        tbody: true,
        tfoot: true,
        th: true,
        thead: true,
        tr: true,
        ul: true
      };
      var tokens = [];

      function addBoundary() {
        if (tokens.length && tokens[tokens.length - 1] !== "\n") {
          tokens.push("\n");
        }
      }

      function walk(node) {
        if (!node) {
          return;
        }
        if (node.nodeType === 3) {
          var value = String(node.nodeValue || "")
            .replace(/\u00A0/g, " ")
            .replace(/[\t\r\n\f ]+/g, " ");
          if (value) {
            tokens.push(value);
          }
          return;
        }
        if (node.nodeType !== 1 || wordAltChunkHtmlElementIsExcluded(node)) {
          return;
        }
        var name = String(node.localName || node.tagName || "").toLowerCase();
        if (name === "br") {
          addBoundary();
          return;
        }
        var isBlock = !!blocks[name];
        if (isBlock) {
          addBoundary();
        }
        var child;
        for (child = node.firstChild; child; child = child.nextSibling) {
          walk(child);
        }
        if (name === "td") {
          tokens.push("\t");
        } else if (isBlock) {
          addBoundary();
        }
      }

      walk(root);
      return tokens.join("")
        .replace(/[ \t]+\n/g, "\n")
        .replace(/\n[ \t]+/g, "\n")
        .replace(/\n{3,}/g, "\n\n")
        .replace(/^[ \t\n]+|[ \t\n]+$/g, "");
    }

    function wordRelationshipById(relationships, id) {
      var index;
      for (index = 0; index < relationships.length; index += 1) {
        if (relationships[index].id === id) {
          return relationships[index];
        }
      }
      return null;
    }

    function throwIfWordReferenceErrorMustAbort(error) {
      throwIfCancelled(error);
      if (isProcessingSafetyError(error) ||
        (error instanceof AppError && error.code === "XML_SIZE")) {
        throw error;
      }
    }

    async function resolveWordReferenceText(
      zip,
      event,
      sourcePart,
      relationships,
      contentTypes,
      warnings,
      warningSeen,
      cache
    ) {
      var id = String(event.value || "");
      var label = translationValue(
        event.kind === "chartReference" ?
          "warning.subject.wordChart" : "warning.subject.altChunk"
      );
      if (!id) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "missing-reference-id:" + event.kind + ":" + sourcePart,
          warningValue("warning.referenceMissingId", { subject: label })
        );
        return "";
      }
      var relationship = wordRelationshipById(relationships, id);
      if (!relationship) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "missing-relationship:" + event.kind + ":" + sourcePart + ":" + id,
          warningValue("warning.referenceRelationshipMissing", {
            subject: label,
            id: id
          })
        );
        return "";
      }
      if (relationship.external) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "external-relationship:" + event.kind + ":" + sourcePart + ":" + id,
          warningValue("warning.externalReferenceSkipped", {
            subject: label,
            id: id
          })
        );
        return "";
      }
      var expectedKind = event.kind === "chartReference" ? "chart" : "afchunk";
      if (relationship.kind !== expectedKind) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "unexpected-relationship:" + event.kind + ":" + sourcePart + ":" + id,
          warningValue("warning.referenceKindMismatch", {
            subject: label,
            id: id
          })
        );
        return "";
      }
      if (!zip.has(relationship.part)) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "missing-reference-part:" + event.kind + ":" + relationship.part,
          warningValue("warning.partMissing", {
            subject: label,
            name: relationship.part
          })
        );
        return "";
      }
      var cacheKey = event.kind + ":" + relationship.part.toLowerCase();
      if (Object.prototype.hasOwnProperty.call(cache, cacheKey)) {
        return cache[cacheKey];
      }
      try {
        var text = "";
        if (event.kind === "chartReference") {
          text = extractWordChartText(await zip.extractXml(relationship.part));
        } else {
          var contentType = ooxmlContentTypeForPart(
            contentTypes,
            relationship.part
          ).split(";", 1)[0].replace(/^\s+|\s+$/g, "");
          if (contentType === "text/html" ||
            contentType === "application/xhtml+xml") {
            text = extractWordAltChunkHtmlText(
              decodeWordAltChunkBytes(await zip.extract(relationship.part))
            );
          } else if (contentType === "text/plain") {
            text = decodeWordAltChunkBytes(await zip.extract(relationship.part));
          } else {
            addWordExtractionWarning(
              warnings,
              warningSeen,
              "unsupported-altchunk:" + relationship.part + ":" + contentType,
              warningValue("warning.unsupportedAltChunk", {
                type: contentType ||
                  translationValue("warning.value.noContentType")
              })
            );
          }
        }
        text = repairSurrogates(String(text || "")).replace(/\r\n?/g, "\n");
        cache[cacheKey] = text;
        return text;
      } catch (error) {
        throwIfWordReferenceErrorMustAbort(error);
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "invalid-reference-part:" + event.kind + ":" + relationship.part,
          warningValue("warning.partParseFailed", {
            subject: label,
            name: relationship.part
          })
        );
        cache[cacheKey] = "";
        return "";
      }
    }

    async function resolveWordPlacementBlockReferences(
      zip,
      blocks,
      sourcePart,
      relationships,
      contentTypes,
      warnings,
      warningSeen,
      cache
    ) {
      var referenceCache = cache || Object.create(null);
      var blockIndex;
      for (blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
        var block = blocks[blockIndex];
        var events = block.events || [{ kind: "text", value: block.text }];
        var resolvedEvents = [];
        var eventIndex;
        for (eventIndex = 0; eventIndex < events.length; eventIndex += 1) {
          var event = events[eventIndex];
          if (event.kind === "chartReference" ||
            event.kind === "altChunkReference") {
            var referenceText = await resolveWordReferenceText(
              zip,
              event,
              sourcePart,
              relationships,
              contentTypes,
              warnings,
              warningSeen,
              referenceCache
            );
            if (hasMeaningfulText(referenceText)) {
              addWordRenderEvent(resolvedEvents, "text", referenceText);
            }
          } else {
            addWordRenderEvent(resolvedEvents, event.kind, event.value);
          }
        }
        block.events = resolvedEvents;
        block.text = wordRenderEventsToText(resolvedEvents, true);
        await cooperativeYield(blockIndex + 1, 50);
      }
      return blocks;
    }

    function wordAnnotationDefinitionIsSpecial(id, type) {
      var normalizedType = String(type || "").toLowerCase();
      if (normalizedType && normalizedType !== "normal") {
        return true;
      }
      return /^-\d+$/.test(String(id || ""));
    }

    async function readWordAnnotationDefinitions(
      zip,
      relationships,
      relationshipKindName,
      elementName,
      label,
      warnings,
      warningSeen,
      contentTypes,
      referenceCache
    ) {
      var result = {
        definitions: [],
        byId: Object.create(null)
      };
      var matching = relationships.filter(function (relationship) {
        return !relationship.external &&
          relationship.kind === relationshipKindName;
      });
      var partIndex;
      for (partIndex = 0; partIndex < matching.length; partIndex += 1) {
        checkCancelled();
        var relationship = matching[partIndex];
        if (!zip.has(relationship.part)) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "missing-part:" + relationshipKindName + ":" + relationship.part,
            warningValue("warning.partMissing", {
              subject: translationValue(
                relationshipKindName === "footnotes" ?
                  "warning.subject.footnote" : "warning.subject.endnote"
              ),
              name: relationship.part
            })
          );
          continue;
        }
        var documentNode = await zip.extractXml(relationship.part);
        var annotationRelationships = await readRelationships(
          zip,
          relationship.part
        );
        var elements = elementsByLocalName(documentNode, elementName);
        ensure(elements.length <= 1000000, "WORD_ANNOTATION_LIMIT",
          label + "定義数が安全上限を超えています。");
        var elementIndex;
        for (elementIndex = 0; elementIndex < elements.length; elementIndex += 1) {
          if (elementIndex > 0 && elementIndex % 1000 === 0) {
            checkCancelled();
            await cooperativeYield(elementIndex, 1000);
          }
          var element = elements[elementIndex];
          var id = String(attributeByLocalName(element, "id") || "");
          var type = String(attributeByLocalName(element, "type") || "");
          if (!id) {
            addWordExtractionWarning(
              warnings,
              warningSeen,
              "invalid-id:" + relationshipKindName,
              warningValue("warning.annotationDefinitionMissingId", {
                subject: translationValue(
                  relationshipKindName === "footnotes" ?
                    "warning.subject.footnote" : "warning.subject.endnote"
                )
              })
            );
            continue;
          }
          if (wordAnnotationDefinitionIsSpecial(id, type)) {
            continue;
          }
          if (Object.prototype.hasOwnProperty.call(result.byId, id)) {
            addWordExtractionWarning(
              warnings,
              warningSeen,
              "duplicate-id:" + relationshipKindName + ":" + id,
              warningValue("warning.annotationDuplicateId", {
                subject: translationValue(
                  relationshipKindName === "footnotes" ?
                    "warning.subject.footnote" : "warning.subject.endnote"
                ),
                id: id
              })
            );
            continue;
          }
          var events = renderWordElementEvents(element, true);
          trimWordRenderEventLineBreaks(events);
          var definitionBlock = createPlacementBlock(
            elementName,
            wordRenderEventsToText(events, true),
            relationship.part,
            elementIndex
          );
          definitionBlock.events = events;
          await resolveWordPlacementBlockReferences(
            zip,
            [definitionBlock],
            relationship.part,
            annotationRelationships,
            contentTypes,
            warnings,
            warningSeen,
            referenceCache
          );
          var definitionText = definitionBlock.text;
          if (!hasMeaningfulText(definitionText)) {
            continue;
          }
          var definition = {
            id: id,
            text: definitionText,
            order: result.definitions.length,
            referenced: false
          };
          result.byId[id] = definition;
          result.definitions.push(definition);
        }
        await cooperativeYield(partIndex + 1, 1);
      }
      return result;
    }

    function createWordLogicalPage() {
      return {
        blocks: [],
        footnotes: []
      };
    }

    function wordLogicalPageHasContent(page) {
      if (page.footnotes.length) {
        return true;
      }
      var index;
      for (index = 0; index < page.blocks.length; index += 1) {
        if (page.blocks[index].keepEmpty ||
          hasMeaningfulText(page.blocks[index].text)) {
          return true;
        }
      }
      return false;
    }

    function addWordLogicalPageBlock(page, sourceBlock, value, keepEmpty) {
      if (!keepEmpty && !hasMeaningfulText(value)) {
        return;
      }
      var block = createPlacementBlock(
        sourceBlock.type,
        value,
        sourceBlock.sourcePart,
        sourceBlock.sourceOrder
      );
      block.keepEmpty = !!keepEmpty;
      block.anchorOrder = sourceBlock.anchorOrder;
      page.blocks.push(block);
    }

    async function renderOoxmlWordPages(
      blocks,
      footnoteDefinitions,
      endnoteDefinitions,
      warnings,
      warningSeen
    ) {
      var pages = [];
      var currentPage = createWordLogicalPage();
      var seenFootnotes = Object.create(null);
      var seenEndnotes = Object.create(null);
      var referencedEndnotes = [];
      var pageBreakSeen = false;

      function finishPage() {
        if (wordLogicalPageHasContent(currentPage)) {
          pages.push(currentPage);
        }
        currentPage = createWordLogicalPage();
      }

      function registerReference(event) {
        var id = String(event.value || "");
        var definitions = event.kind === "footnoteReference" ?
          footnoteDefinitions :
          endnoteDefinitions;
        var seen = event.kind === "footnoteReference" ?
          seenFootnotes :
          seenEndnotes;
        var label = translationValue(
          event.kind === "footnoteReference" ?
            "warning.subject.footnote" : "warning.subject.endnote"
        );
        if (!id || !Object.prototype.hasOwnProperty.call(definitions.byId, id)) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "missing-definition:" + event.kind + ":" + id,
            warningValue("warning.annotationBodyMissing", {
              subject: label,
              id: id || translationValue("warning.value.empty")
            })
          );
          return;
        }
        if (seen[id]) {
          return;
        }
        seen[id] = true;
        var definition = definitions.byId[id];
        definition.referenced = true;
        if (event.kind === "footnoteReference") {
          currentPage.footnotes.push(definition);
        } else {
          referencedEndnotes.push(definition);
        }
      }

      var blockIndex;
      for (blockIndex = 0; blockIndex < blocks.length; blockIndex += 1) {
        if (blockIndex > 0 && blockIndex % 500 === 0) {
          checkCancelled();
          await cooperativeYield(blockIndex, 500);
        }
        var sourceBlock = blocks[blockIndex];
        var events = sourceBlock.events || [
          { kind: "text", value: sourceBlock.text }
        ];
        var fragment = "";
        var fragmentAdded = false;
        var eventIndex;
        for (eventIndex = 0; eventIndex < events.length; eventIndex += 1) {
          var event = events[eventIndex];
          if (event.kind === "text") {
            fragment += event.value;
          } else if (event.kind === "lineBreak") {
            fragment += "\n";
          } else if (event.kind === "pageBreak") {
            if (hasMeaningfulText(fragment)) {
              addWordLogicalPageBlock(
                currentPage,
                sourceBlock,
                fragment,
                false
              );
              fragmentAdded = true;
            }
            fragment = "";
            finishPage();
            pageBreakSeen = true;
          } else if (event.kind === "footnoteReference" ||
            event.kind === "endnoteReference") {
            registerReference(event);
          }
        }
        if (hasMeaningfulText(fragment) ||
          (sourceBlock.keepEmpty && !fragmentAdded)) {
          addWordLogicalPageBlock(
            currentPage,
            sourceBlock,
            fragment,
            sourceBlock.keepEmpty && !hasMeaningfulText(fragment)
          );
        }
      }
      if (wordLogicalPageHasContent(currentPage) || !pages.length) {
        pages.push(currentPage);
      }

      if (footnoteDefinitions.definitions.length && !pageBreakSeen) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "no-page-break:ooxml",
          warningValue("warning.wordNoPageBreak", {
            document: translationValue("warning.document.word")
          })
        );
      }

      var renderedPages = [];
      var pageIndex;
      for (pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
        var page = pages[pageIndex];
        var pageText = renderPlacementBlocks(page.blocks);
        if (page.footnotes.length) {
          var footnoteTexts = page.footnotes.map(function (definition) {
            return definition.text;
          });
          var footnoteSection = "----- 脚注 -----";
          if (footnoteTexts.length) {
            footnoteSection += "\n" + footnoteTexts.join("\n\n");
          }
          pageText += (hasMeaningfulText(pageText) ? "\n\n" : "") +
            footnoteSection;
        }
        if (hasMeaningfulText(pageText)) {
          renderedPages.push(pageText);
        }
      }
      return {
        text: renderedPages.join("\n\n"),
        referencedEndnotes: referencedEndnotes
      };
    }

    async function extractOoxmlWordText(zip, warnings) {
      var outputWarnings = warnings || [];
      var warningSeen = Object.create(null);
      var mainPart = "word/document.xml";
      var mainDocument = await zip.extractXml(mainPart);
      var relationships = await readRelationships(zip, mainPart);
      var contentTypes = await readOoxmlContentTypes(zip);
      var referenceCache = Object.create(null);
      var footnoteDefinitions = await readWordAnnotationDefinitions(
        zip,
        relationships,
        "footnotes",
        "footnote",
        "脚注",
        outputWarnings,
        warningSeen,
        contentTypes,
        referenceCache
      );
      var endnoteDefinitions = await readWordAnnotationDefinitions(
        zip,
        relationships,
        "endnotes",
        "endnote",
        "文末脚注",
        outputWarnings,
        warningSeen,
        contentTypes,
        referenceCache
      );
      var mainBlocks = collectWordPlacementBlocks(mainDocument, mainPart);
      await resolveWordPlacementBlockReferences(
        zip,
        mainBlocks,
        mainPart,
        relationships,
        contentTypes,
        outputWarnings,
        warningSeen,
        referenceCache
      );
      var mainResult = await renderOoxmlWordPages(
        mainBlocks,
        footnoteDefinitions,
        endnoteDefinitions,
        outputWarnings,
        warningSeen
      );
      var sections = [];
      if (hasMeaningfulText(mainResult.text)) {
        sections.push(mainResult.text);
      }
      var configured = [
        { kinds: ["header"], label: "ヘッダー" },
        { kinds: ["footer"], label: "フッター" },
        { kinds: ["comments"], label: "コメント" }
      ];
      var groupIndex;
      for (groupIndex = 0; groupIndex < configured.length; groupIndex += 1) {
        var group = configured[groupIndex];
        var groupParts = relationships.filter(function (relationship) {
          return !relationship.external &&
            group.kinds.indexOf(relationship.kind) >= 0 &&
            zip.has(relationship.part);
        });
        var texts = [];
        var partIndex;
        for (partIndex = 0; partIndex < groupParts.length; partIndex += 1) {
          checkCancelled();
          var groupPart = groupParts[partIndex].part;
          var partDocument = await zip.extractXml(groupPart);
          var partRelationships = await readRelationships(zip, groupPart);
          var partBlocks = collectWordPlacementBlocks(partDocument, groupPart);
          await resolveWordPlacementBlockReferences(
            zip,
            partBlocks,
            groupPart,
            partRelationships,
            contentTypes,
            outputWarnings,
            warningSeen,
            referenceCache
          );
          var partText = renderPlacementBlocks(partBlocks);
          if (hasMeaningfulText(partText)) {
            texts.push(partText);
          }
          await cooperativeYield(partIndex + 1, 4);
        }
        if (texts.length) {
          sections.push("===== " + group.label + " =====\n" + texts.join("\n"));
        }
      }

      var unreferencedFootnotes = footnoteDefinitions.definitions.filter(
        function (definition) {
          return !definition.referenced && hasMeaningfulText(definition.text);
        }
      );
      if (unreferencedFootnotes.length) {
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "unreferenced:footnotes",
          warningValue(
            "warning.unreferencedAnnotation",
            { subject: translationValue("warning.subject.footnote") },
            "",
            unreferencedFootnotes.length
          )
        );
        sections.push(
          "===== 未参照の脚注 =====\n" +
          unreferencedFootnotes.map(function (definition) {
            return definition.text;
          }).join("\n\n")
        );
      }

      var endnotes = mainResult.referencedEndnotes.filter(function (definition) {
        return hasMeaningfulText(definition.text);
      });
      var unreferencedEndnoteCount = 0;
      endnoteDefinitions.definitions.forEach(function (definition) {
        if (!definition.referenced && hasMeaningfulText(definition.text)) {
          unreferencedEndnoteCount += 1;
          endnotes.push(definition);
        }
      });
      if (unreferencedEndnoteCount) {
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "unreferenced:endnotes",
          warningValue(
            "warning.unreferencedAnnotation",
            { subject: translationValue("warning.subject.endnote") },
            "",
            unreferencedEndnoteCount
          )
        );
      }
      if (endnotes.length) {
        sections.push(
          "===== 文末脚注 =====\n" +
          endnotes.map(function (definition) {
            return definition.text;
          }).join("\n\n")
        );
      }
      return sections.join("\n\n");
    }

    function childElementsByLocalName(element, name) {
      var result = [];
      var child;
      for (child = element ? element.firstElementChild : null;
        child;
        child = child.nextElementSibling) {
        if (child.localName === name) {
          result.push(child);
        }
      }
      return result;
    }

    function firstChildByLocalName(element, name) {
      var children = childElementsByLocalName(element, name);
      return children.length ? children[0] : null;
    }

    function collectTextNodesExcluding(root, excludedAncestor) {
      var texts = elementsByLocalName(root, "t");
      var result = "";
      texts.forEach(function (element) {
        var current = element.parentElement;
        var excluded = false;
        while (current && current !== root) {
          if (current.localName === excludedAncestor) {
            excluded = true;
            break;
          }
          current = current.parentElement;
        }
        if (!excluded) {
          result += element.textContent || "";
        }
      });
      return result;
    }

    function parseOoxmlSharedStrings(documentNode) {
      return elementsByLocalName(documentNode, "si").map(function (item) {
        return collectTextNodesExcluding(item, "rPh");
      });
    }

    var BUILTIN_DATE_FORMATS = {
      14: true, 15: true, 16: true, 17: true, 18: true, 19: true, 20: true,
      21: true, 22: true, 27: true, 28: true, 29: true, 30: true, 31: true,
      32: true, 33: true, 34: true, 35: true, 36: true, 45: true, 46: true,
      47: true, 50: true, 51: true, 52: true, 53: true, 54: true, 55: true,
      56: true, 57: true, 58: true
    };

    function looksLikeDateFormat(formatCode) {
      var format = String(formatCode || "")
        .replace(/"[^"]*"/g, "")
        .replace(/\\./g, "")
        .replace(/\[[^\]]*\]/g, "")
        .replace(/_.| \*./g, "")
        .toLowerCase();
      return /[ymdhis]/.test(format);
    }

    function parseOoxmlStyles(documentNode) {
      var custom = Object.create(null);
      elementsByLocalName(documentNode, "numFmt").forEach(function (element) {
        var id = Number(element.getAttribute("numFmtId"));
        if (Number.isInteger(id)) {
          custom[id] = element.getAttribute("formatCode") || "";
        }
      });
      var cellXfs = firstElementByLocalName(documentNode, "cellXfs");
      var styles = [];
      if (cellXfs) {
        childElementsByLocalName(cellXfs, "xf").forEach(function (element) {
          var id = Number(element.getAttribute("numFmtId") || 0);
          styles.push({
            numFmtId: id,
            date: !!BUILTIN_DATE_FORMATS[id] || looksLikeDateFormat(custom[id])
          });
        });
      }
      return styles;
    }

    function excelSerialToText(value, date1904) {
      if (!Number.isFinite(value) || value < 0 || value > 2958465) {
        return null;
      }
      var days = value;
      if (!date1904 && days >= 60) {
        days -= 1;
      }
      var epoch = date1904 ?
        Date.UTC(1904, 0, 1) :
        Date.UTC(1899, 11, 31);
      var milliseconds = epoch + Math.round(days * 86400000);
      var date = new Date(milliseconds);
      if (!Number.isFinite(date.getTime())) {
        return null;
      }
      function two(number) {
        return zeroPad(number, 2);
      }
      var output = date.getUTCFullYear() + "-" +
        two(date.getUTCMonth() + 1) + "-" +
        two(date.getUTCDate());
      var fraction = Math.abs(days - Math.floor(days));
      if (fraction > 0.0000001) {
        output += " " + two(date.getUTCHours()) + ":" +
          two(date.getUTCMinutes()) + ":" +
          two(date.getUTCSeconds());
      }
      return output;
    }

    function visibleCellText(value) {
      return String(value == null ? "" : value)
        .replace(/\\/g, "\\\\")
        .replace(/\r\n?|\n/g, "\\n")
        .replace(/\t/g, "\\t");
    }

    function normalizeExcelNumericText(value) {
      var text = String(value == null ? "" : value);
      var number = Number(text);
      if (!Number.isFinite(number)) {
        return text;
      }
      if (number === 0) {
        return "0";
      }
      return Number(number.toPrecision(15)).toString();
    }

    function excelCellPosition(address) {
      var match = /^([A-Z]+)([1-9][0-9]*)$/i.exec(String(address || ""));
      if (!match) {
        return null;
      }
      var letters = match[1].toUpperCase();
      var column = 0;
      var index;
      for (index = 0; index < letters.length; index += 1) {
        column = column * 26 + letters.charCodeAt(index) - 64;
      }
      var row = Number(match[2]) - 1;
      column -= 1;
      if (!Number.isInteger(row) || row < 0 || row > 1048575 ||
        column < 0 || column > 16383) {
        return null;
      }
      return { row: row, column: column };
    }

    function excelTableRowBlocks(cells, sourcePart) {
      var values = Object.create(null);
      var minRow = Infinity;
      var maxRow = -1;
      var minColumn = Infinity;
      var maxColumn = -1;
      var index;
      for (index = 0; index < cells.length; index += 1) {
        var cell = cells[index];
        var text = visibleCellText(cell.value);
        if (text === "") {
          continue;
        }
        values[cell.row + ":" + cell.column] = text;
        minRow = Math.min(minRow, cell.row);
        maxRow = Math.max(maxRow, cell.row);
        minColumn = Math.min(minColumn, cell.column);
        maxColumn = Math.max(maxColumn, cell.column);
      }
      if (maxRow < minRow || maxColumn < minColumn) {
        return [];
      }
      var rowCount = maxRow - minRow + 1;
      var columnCount = maxColumn - minColumn + 1;
      ensure(
        rowCount <= MAX_EXCEL_TABLE_CELLS &&
        columnCount <= MAX_EXCEL_TABLE_CELLS &&
        rowCount <= Math.floor(MAX_EXCEL_TABLE_CELLS / columnCount),
        "EXCEL_TEXT_RANGE",
        "Excelのテキスト出力範囲が大きすぎます。"
      );
      var blocks = [];
      var generatedCells = 0;
      var row;
      for (row = minRow; row <= maxRow; row += 1) {
        checkCancelled();
        var fields = [];
        var column;
        for (column = minColumn; column <= maxColumn; column += 1) {
          var key = row + ":" + column;
          fields.push(Object.prototype.hasOwnProperty.call(values, key) ?
            values[key] : "");
          generatedCells += 1;
          if (generatedCells % 4096 === 0) {
            checkCancelled();
          }
        }
        var rowText = fields.join("\t");
        if (!hasMeaningfulText(rowText)) {
          rowText = "";
        }
        var block = createPlacementBlock(
          "cell-row",
          rowText,
          sourcePart || "",
          blocks.length
        );
        block.row = row;
        block.column = minColumn;
        block.positionKnown = true;
        block.keepEmpty = rowText === "";
        blocks.push(block);
      }
      return blocks;
    }

    function findWorkbookRelationship(relationships, kind) {
      var index;
      for (index = 0; index < relationships.length; index += 1) {
        if (!relationships[index].external && relationships[index].kind === kind) {
          return relationships[index];
        }
      }
      return null;
    }

    function vmlTextBoxText(textBox) {
      var output = "";
      function walk(node) {
        if (!node) {
          return;
        }
        if (node.nodeType === 3) {
          var value = node.nodeValue || "";
          if (/[^\t\r\n ]/.test(value)) {
            output += value;
          }
          return;
        }
        if (node.nodeType !== 1) {
          return;
        }
        var block = node.localName === "div" || node.localName === "p";
        if (block && output && output.charAt(output.length - 1) !== "\n") {
          output += "\n";
        }
        if (node.localName === "br") {
          output += "\n";
          return;
        }
        var child;
        for (child = node.firstChild; child; child = child.nextSibling) {
          walk(child);
        }
        if (block && output && output.charAt(output.length - 1) !== "\n") {
          output += "\n";
        }
      }
      walk(textBox);
      return output.replace(/\r\n?|\n/g, "\n").replace(/^\n+|\n+$/g, "");
    }

    function excelSheetMetrics(worksheet) {
      var sheetFormat = firstElementByLocalName(worksheet, "sheetFormatPr");
      var defaultRowHeight = sheetFormat ?
        placementNumber(sheetFormat.getAttribute("defaultRowHeight")) : null;
      var defaultColumnWidth = sheetFormat ?
        placementNumber(
          sheetFormat.getAttribute("defaultColWidth") ||
          sheetFormat.getAttribute("baseColWidth")
        ) : null;
      var metrics = {
        defaultRowHeightEmu: (defaultRowHeight || 15) * 12700,
        defaultColumnWidthEmu: (defaultColumnWidth || 8.43) * 7 * 9525,
        rowHeights: Object.create(null)
      };
      elementsByLocalName(worksheet, "row").forEach(function (row) {
        var rowNumber = placementNumber(row.getAttribute("r"));
        var height = placementNumber(row.getAttribute("ht"));
        if (rowNumber !== null && rowNumber >= 1 && height !== null && height > 0) {
          metrics.rowHeights[rowNumber - 1] = height * 12700;
        }
      });
      return metrics;
    }

    function excelRowHeightEmu(metrics, row) {
      return Object.prototype.hasOwnProperty.call(metrics.rowHeights, row) ?
        metrics.rowHeights[row] :
        metrics.defaultRowHeightEmu;
    }

    function applyExcelAbsolutePosition(block, metrics) {
      if (block.y !== null && block.row === null) {
        block.row = Math.max(0, Math.floor(block.y / metrics.defaultRowHeightEmu));
        block.rowOffset = block.y % metrics.defaultRowHeightEmu;
      }
      if (block.x !== null && block.column === null) {
        block.column = Math.max(0, Math.floor(block.x / metrics.defaultColumnWidthEmu));
        block.columnOffset = block.x % metrics.defaultColumnWidthEmu;
      }
      block.positionKnown = block.row !== null || block.y !== null;
    }

    function parseExcelDrawingAnchor(block, anchor, metrics) {
      var from = firstChildByLocalName(anchor, "from");
      if (from) {
        var row = firstChildByLocalName(from, "row");
        var column = firstChildByLocalName(from, "col");
        var rowOffset = firstChildByLocalName(from, "rowOff");
        var columnOffset = firstChildByLocalName(from, "colOff");
        block.row = row ? placementNumber(row.textContent) : null;
        block.column = column ? placementNumber(column.textContent) : null;
        block.rowOffset = rowOffset ? placementNumber(rowOffset.textContent) : null;
        block.columnOffset = columnOffset ?
          placementNumber(columnOffset.textContent) : null;
      } else if (anchor.localName === "absoluteAnchor") {
        var position = firstChildByLocalName(anchor, "pos");
        if (position) {
          block.x = placementNumber(position.getAttribute("x"));
          block.y = placementNumber(position.getAttribute("y"));
        }
      }
      var extent = firstChildByLocalName(anchor, "ext");
      if (extent) {
        block.width = placementNumber(extent.getAttribute("cx"));
        block.height = placementNumber(extent.getAttribute("cy"));
      }
      applyExcelAbsolutePosition(block, metrics);
    }

    function collectExcelDrawingAnchors(documentNode) {
      var result = [];
      function walk(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "AlternateContent") {
          var preferred = preferredAlternateContentChild(node);
          if (preferred) {
            walk(preferred);
          }
          return;
        }
        if (node.localName === "twoCellAnchor" ||
          node.localName === "oneCellAnchor" ||
          node.localName === "absoluteAnchor") {
          result.push(node);
          return;
        }
        var child;
        for (child = node.firstElementChild; child; child = child.nextElementSibling) {
          walk(child);
        }
      }
      if (documentNode && documentNode.documentElement) {
        walk(documentNode.documentElement);
      }
      return result;
    }

    function extractVmlDrawingBlocks(documentNode, sourcePart, startOrder, metrics) {
      var blocks = [];
      var textBoxes = elementsByLocalName(documentNode, "textbox");
      var index;
      for (index = 0; index < textBoxes.length; index += 1) {
        var text = vmlTextBoxText(textBoxes[index]);
        if (!hasMeaningfulText(text)) {
          continue;
        }
        var shape = textBoxes[index].parentElement;
        while (shape && shape.localName !== "shape") {
          shape = shape.parentElement;
        }
        var block = createPlacementBlock(
          "shape",
          text,
          sourcePart,
          startOrder + blocks.length
        );
        block.sourceId = sourcePart + "#vml-" + index;
        var anchor = shape ?
          firstElementByLocalName(shape, "Anchor") ||
          firstElementByLocalName(shape, "anchor") :
          null;
        if (anchor) {
          var parts = String(anchor.textContent || "").split(",");
          if (parts.length >= 4) {
            block.column = placementNumber(parts[0]);
            block.columnOffset = placementNumber(parts[1]);
            block.row = placementNumber(parts[2]);
            block.rowOffset = placementNumber(parts[3]);
            block.vmlOffsetUnits = true;
          }
        }
        if (shape) {
          var style = cssStyleMap(shape.getAttribute("style") || "");
          var left = parseCssLength(style.left);
          var top = parseCssLength(style.top);
          if (left === null) {
            left = parseCssLength(style["margin-left"]);
          }
          if (top === null) {
            top = parseCssLength(style["margin-top"]);
          }
          block.x = left === null ? null : left * 12700;
          block.y = top === null ? null : top * 12700;
          block.width = parseCssLength(style.width);
          block.height = parseCssLength(style.height);
        }
        applyExcelAbsolutePosition(block, metrics);
        blocks.push(block);
      }
      return blocks;
    }

    async function extractOoxmlExcelShapeBlocks(
      zip,
      worksheet,
      worksheetPart,
      metrics
    ) {
      var relationships = await readRelationships(zip, worksheetPart);
      var relMap = relationshipMap(relationships);
      var references = worksheet.getElementsByTagName("*");
      var seenParts = Object.create(null);
      var blocks = [];
      var index;
      for (index = 0; index < references.length; index += 1) {
        await cooperativeYield(index + 1, 100);
        var reference = references[index];
        if (reference.localName !== "drawing" &&
          reference.localName !== "legacyDrawing") {
          continue;
        }
        var relationship = relMap[relationshipIdAttribute(reference)];
        if (!relationship || relationship.external ||
          (relationship.kind !== "drawing" && relationship.kind !== "vmldrawing") ||
          !zip.has(relationship.part) || seenParts[relationship.part]) {
          continue;
        }
        seenParts[relationship.part] = true;
        var drawing = await zip.extractXml(relationship.part);
        if (relationship.kind === "vmldrawing") {
          Array.prototype.push.apply(
            blocks,
            extractVmlDrawingBlocks(
              drawing,
              relationship.part,
              blocks.length,
              metrics
            )
          );
          continue;
        }
        var anchors = collectExcelDrawingAnchors(drawing);
        var anchorIndex;
        for (anchorIndex = 0; anchorIndex < anchors.length; anchorIndex += 1) {
          var text = drawingObjectText(anchors[anchorIndex], false);
          if (!hasMeaningfulText(text)) {
            continue;
          }
          var block = createPlacementBlock(
            "shape",
            text,
            relationship.part,
            blocks.length
          );
          block.sourceId = relationship.part + "#anchor-" + anchorIndex;
          parseExcelDrawingAnchor(block, anchors[anchorIndex], metrics);
          blocks.push(block);
        }
      }
      return blocks;
    }

    function excelShapeRowPhase(block, metrics) {
      if (block.type === "cell-row") {
        return 0;
      }
      if (block.row === null) {
        return 2;
      }
      if (block.rowOffset === null) {
        return 1;
      }
      var half = block.vmlOffsetUnits ?
        128 :
        excelRowHeightEmu(metrics, block.row) / 2;
      return block.rowOffset < half ? -1 : 1;
    }

    function sortExcelPlacementBlocks(blocks, metrics) {
      return blocks.sort(function (left, right) {
        var leftRow = left.row === null ? Infinity : left.row;
        var rightRow = right.row === null ? Infinity : right.row;
        if (leftRow !== rightRow) {
          return leftRow - rightRow;
        }
        var leftPhase = excelShapeRowPhase(left, metrics);
        var rightPhase = excelShapeRowPhase(right, metrics);
        if (leftPhase !== rightPhase) {
          return leftPhase - rightPhase;
        }
        var leftColumn = left.column === null ? Infinity : left.column;
        var rightColumn = right.column === null ? Infinity : right.column;
        if (leftColumn !== rightColumn) {
          return leftColumn - rightColumn;
        }
        var leftOffset = left.columnOffset === null ? Infinity : left.columnOffset;
        var rightOffset = right.columnOffset === null ? Infinity : right.columnOffset;
        if (leftOffset !== rightOffset) {
          return leftOffset - rightOffset;
        }
        return left.sourceOrder - right.sourceOrder;
      });
    }

    async function extractOoxmlExcelText(zip) {
      var workbookPart = "xl/workbook.xml";
      var workbook = await zip.extractXml(workbookPart);
      var relationships = await readRelationships(zip, workbookPart);
      var relMap = relationshipMap(relationships);
      var sharedStrings = [];
      var sharedRel = findWorkbookRelationship(relationships, "sharedstrings");
      if (sharedRel && zip.has(sharedRel.part)) {
        sharedStrings = parseOoxmlSharedStrings(await zip.extractXml(sharedRel.part));
      }
      var styles = [];
      var stylesRel = findWorkbookRelationship(relationships, "styles");
      if (stylesRel && zip.has(stylesRel.part)) {
        styles = parseOoxmlStyles(await zip.extractXml(stylesRel.part));
      }
      var workbookPr = firstElementByLocalName(workbook, "workbookPr");
      var date1904Value = workbookPr ? workbookPr.getAttribute("date1904") : "";
      var date1904 = date1904Value === "1" || date1904Value === "true";
      var sheetElements = elementsByLocalName(workbook, "sheet");
      var output = [];
      var sheetIndex;
      for (sheetIndex = 0; sheetIndex < sheetElements.length; sheetIndex += 1) {
        checkCancelled();
        var sheetElement = sheetElements[sheetIndex];
        var relationId = relationshipIdAttribute(sheetElement);
        var relationship = relMap[relationId];
        if (!relationship || relationship.external || !zip.has(relationship.part)) {
          continue;
        }
        var sheetName = sheetElement.getAttribute("name") || ("Sheet" + (sheetIndex + 1));
        var state = (sheetElement.getAttribute("state") || "visible").toLowerCase();
        var heading = "===== Sheet: " + sheetName +
          (state === "visible" ? "" : "（非表示）") + " =====";
        var tableCells = [];
        var worksheet = await zip.extractXml(relationship.part);
        var cells = elementsByLocalName(worksheet, "c");
        var cellIndex;
        for (cellIndex = 0; cellIndex < cells.length; cellIndex += 1) {
          await cooperativeYield(cellIndex + 1, 500);
          var cell = cells[cellIndex];
          var address = cell.getAttribute("r") || "";
          var position = excelCellPosition(address);
          if (!position) {
            continue;
          }
          var type = cell.getAttribute("t") || "n";
          var styleIndex = Number(cell.getAttribute("s") || -1);
          var formulaElement = firstChildByLocalName(cell, "f");
          var valueElement = firstChildByLocalName(cell, "v");
          var formula = formulaElement ? (formulaElement.textContent || "") : "";
          if (formula && formula.charAt(0) !== "=") {
            formula = "=" + formula;
          }
          var value = "";
          if (type === "inlineStr") {
            var inline = firstChildByLocalName(cell, "is");
            value = inline ? collectTextNodesExcluding(inline, "rPh") : "";
          } else if (type === "s") {
            var sharedIndex = valueElement ? Number(valueElement.textContent) : -1;
            value = Number.isInteger(sharedIndex) && sharedIndex >= 0 &&
              sharedIndex < sharedStrings.length ? sharedStrings[sharedIndex] : "";
          } else if (type === "str") {
            value = valueElement ? valueElement.textContent || "" : "";
          } else if (type === "b") {
            value = valueElement && valueElement.textContent === "1" ? "TRUE" : "FALSE";
          } else if (type === "e") {
            value = valueElement ? valueElement.textContent || "" : "";
          } else if (valueElement) {
            value = valueElement.textContent || "";
            var numeric = Number(value);
            if (Number.isFinite(numeric)) {
              var convertedDate = false;
              if (styleIndex >= 0 &&
                styles[styleIndex] && styles[styleIndex].date) {
                var dateText = excelSerialToText(numeric, date1904);
                if (dateText !== null) {
                  value = dateText;
                  convertedDate = true;
                }
              }
              if (!convertedDate) {
                value = normalizeExcelNumericText(value);
              }
            }
          }
          if (value !== "" || formula !== "") {
            tableCells.push({
              row: position.row,
              column: position.column,
              value: value !== "" ? value : formula
            });
          }
        }
        var metrics = excelSheetMetrics(worksheet);
        var placementBlocks = excelTableRowBlocks(tableCells, relationship.part);
        var shapeBlocks = await extractOoxmlExcelShapeBlocks(
          zip,
          worksheet,
          relationship.part,
          metrics
        );
        Array.prototype.push.apply(placementBlocks, shapeBlocks);
        sortExcelPlacementBlocks(placementBlocks, metrics);
        var lines = [heading];
        if (placementBlocks.length) {
          lines.push(renderPlacementBlocks(placementBlocks));
        }
        output.push(lines.join("\n"));
        await cooperativeYield(sheetIndex + 1, 1);
      }
      return output.join("\n\n");
    }

    function drawingParagraphText(paragraph) {
      var output = "";
      function walk(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "t") {
          output += node.textContent || "";
          return;
        }
        if (node.localName === "br") {
          output += "\n";
          return;
        }
        if (node.localName === "tab") {
          output += "\t";
          return;
        }
        var child;
        for (child = node.firstElementChild; child; child = child.nextElementSibling) {
          walk(child);
        }
      }
      walk(paragraph);
      return output;
    }

    function paragraphIsNotesMetadata(paragraph) {
      var current = paragraph;
      while (current && current.localName !== "sp") {
        current = current.parentElement;
      }
      if (!current) {
        return false;
      }
      var placeholders = elementsByLocalName(current, "ph");
      if (!placeholders.length) {
        return false;
      }
      var type = (placeholders[0].getAttribute("type") || "").toLowerCase();
      return type === "dt" || type === "ftr" || type === "sldnum" || type === "hdr";
    }

    function drawingTableText(table) {
      var lines = [];
      childElementsByLocalName(table, "tr").forEach(function (row) {
        var cells = childElementsByLocalName(row, "tc");
        var values = cells.map(function (cell) {
          return elementsByLocalName(cell, "p")
            .map(drawingParagraphText)
            .filter(hasMeaningfulText)
            .join("\n");
        });
        if (values.some(hasMeaningfulText)) {
          lines.push(values.join("\t"));
        }
      });
      return lines.join("\n");
    }

    function drawingObjectText(root, notesMode) {
      var blocks = [];
      function walk(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "AlternateContent") {
          var preferred = preferredAlternateContentChild(node);
          if (preferred) {
            walk(preferred);
          }
          return;
        }
        if (node.localName === "tbl") {
          var tableText = drawingTableText(node);
          if (hasMeaningfulText(tableText)) {
            blocks.push(tableText);
          }
          return;
        }
        if (node.localName === "p") {
          if (!notesMode || !paragraphIsNotesMetadata(node)) {
            var paragraphText = drawingParagraphText(node);
            if (hasMeaningfulText(paragraphText)) {
              blocks.push(paragraphText);
            }
          }
          return;
        }
        var child;
        for (child = node.firstElementChild; child; child = child.nextElementSibling) {
          walk(child);
        }
      }
      walk(root);
      return blocks.join("\n");
    }

    function powerPointTreeObjects(documentNode) {
      var tree = firstElementByLocalName(documentNode, "spTree");
      var result = [];
      function append(node) {
        if (!node || node.nodeType !== 1) {
          return;
        }
        if (node.localName === "AlternateContent") {
          var preferred = preferredAlternateContentChild(node);
          var alternateChild;
          for (alternateChild = preferred ? preferred.firstElementChild : null;
            alternateChild;
            alternateChild = alternateChild.nextElementSibling) {
            append(alternateChild);
          }
          return;
        }
        if (node.localName === "sp" ||
          node.localName === "graphicFrame" ||
          node.localName === "grpSp" ||
          node.localName === "cxnSp") {
          result.push(node);
        }
      }
      var child;
      for (child = tree ? tree.firstElementChild : null;
        child;
        child = child.nextElementSibling) {
        append(child);
      }
      return result;
    }

    function powerPointPlaceholderIdentity(object) {
      var placeholders = elementsByLocalName(object, "ph");
      if (!placeholders.length) {
        return null;
      }
      return {
        index: placeholders[0].getAttribute("idx") || "",
        type: (placeholders[0].getAttribute("type") || "").toLowerCase()
      };
    }

    function powerPointTransformElement(object) {
      var child;
      for (child = object ? object.firstElementChild : null;
        child;
        child = child.nextElementSibling) {
        if (child.localName === "xfrm") {
          return child;
        }
        if (child.localName === "spPr" || child.localName === "grpSpPr") {
          return firstChildByLocalName(child, "xfrm");
        }
      }
      return null;
    }

    function parsePowerPointTransform(object) {
      var transform = powerPointTransformElement(object);
      if (!transform) {
        return null;
      }
      var offset = firstChildByLocalName(transform, "off");
      var extent = firstChildByLocalName(transform, "ext");
      var result = {
        x: offset ? placementNumber(offset.getAttribute("x")) : null,
        y: offset ? placementNumber(offset.getAttribute("y")) : null,
        width: extent ? placementNumber(extent.getAttribute("cx")) : null,
        height: extent ? placementNumber(extent.getAttribute("cy")) : null
      };
      return result.x !== null || result.y !== null ? result : null;
    }

    function powerPointPlaceholderMatches(candidate, identity) {
      var candidateIdentity = powerPointPlaceholderIdentity(candidate);
      if (!candidateIdentity) {
        return false;
      }
      if (identity.index && candidateIdentity.index) {
        return identity.index === candidateIdentity.index;
      }
      return !!identity.type && candidateIdentity.type === identity.type;
    }

    function resolvePowerPointPlaceholderPosition(object, inheritedObjectGroups) {
      var identity = powerPointPlaceholderIdentity(object);
      if (!identity) {
        return null;
      }
      var groupIndex;
      for (groupIndex = 0;
        inheritedObjectGroups && groupIndex < inheritedObjectGroups.length;
        groupIndex += 1) {
        var candidates = inheritedObjectGroups[groupIndex];
        var candidateIndex;
        for (candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
          if (!powerPointPlaceholderMatches(candidates[candidateIndex], identity)) {
            continue;
          }
          var transform = parsePowerPointTransform(candidates[candidateIndex]);
          if (transform) {
            return transform;
          }
        }
      }
      return null;
    }

    function collectPowerPointBlocks(
      documentNode,
      notesMode,
      inheritedDocuments,
      sourcePart
    ) {
      var objects = powerPointTreeObjects(documentNode);
      var inheritedObjectGroups = [];
      var inheritedIndex;
      for (inheritedIndex = 0;
        inheritedDocuments && inheritedIndex < inheritedDocuments.length;
        inheritedIndex += 1) {
        inheritedObjectGroups.push(
          powerPointTreeObjects(inheritedDocuments[inheritedIndex])
        );
      }
      var blocks = [];
      var index;
      for (index = 0; index < objects.length; index += 1) {
        var text = drawingObjectText(objects[index], notesMode);
        if (!hasMeaningfulText(text)) {
          continue;
        }
        var block = createPlacementBlock(
          objects[index].localName === "graphicFrame" &&
            elementsByLocalName(objects[index], "tbl").length ?
            "table" :
            objects[index].localName,
          text,
          sourcePart || "",
          index
        );
        block.sourceId = (sourcePart || "") + "#object-" + index;
        var transform = parsePowerPointTransform(objects[index]) ||
          resolvePowerPointPlaceholderPosition(objects[index], inheritedObjectGroups);
        if (transform) {
          block.x = transform.x;
          block.y = transform.y;
          block.width = transform.width;
          block.height = transform.height;
          block.positionKnown = block.x !== null && block.y !== null;
        }
        blocks.push(block);
      }
      return blocks;
    }

    function powerPointBlocksShareRow(block, group) {
      var blockHeight = block.height !== null && block.height > 0 ? block.height : 0;
      var groupHeight = group.bottom > group.top ? group.bottom - group.top : 0;
      if (blockHeight > 0 && groupHeight > 0) {
        var overlap = Math.min(block.y + blockHeight, group.bottom) -
          Math.max(block.y, group.top);
        if (overlap >= Math.min(blockHeight, groupHeight) * 0.5) {
          return true;
        }
      }
      var smallestHeight = blockHeight > 0 && groupHeight > 0 ?
        Math.min(blockHeight, groupHeight) :
        0;
      var tolerance = Math.max(127000, smallestHeight * 0.25);
      return Math.abs(block.y - group.referenceY) <= tolerance;
    }

    function sortPowerPointBlocks(blocks) {
      var known = blocks.filter(function (block) {
        return block.positionKnown;
      }).sort(function (left, right) {
        if (left.y !== right.y) {
          return left.y - right.y;
        }
        return left.sourceOrder - right.sourceOrder;
      });
      var unknown = blocks.filter(function (block) {
        return !block.positionKnown;
      }).sort(function (left, right) {
        return left.sourceOrder - right.sourceOrder;
      });
      var groups = [];
      known.forEach(function (block) {
        var group = groups.length ? groups[groups.length - 1] : null;
        if (!group || !powerPointBlocksShareRow(block, group)) {
          group = {
            referenceY: block.y,
            top: block.y,
            bottom: block.y + (block.height && block.height > 0 ? block.height : 0),
            blocks: []
          };
          groups.push(group);
        }
        group.blocks.push(block);
        group.top = Math.min(group.top, block.y);
        group.bottom = Math.max(
          group.bottom,
          block.y + (block.height && block.height > 0 ? block.height : 0)
        );
      });
      var result = [];
      groups.forEach(function (group) {
        group.blocks.sort(function (left, right) {
          if (left.x !== right.x) {
            return left.x - right.x;
          }
          return left.sourceOrder - right.sourceOrder;
        });
        Array.prototype.push.apply(result, group.blocks);
      });
      Array.prototype.push.apply(result, unknown);
      return result;
    }

    function extractDrawingText(
      documentNode,
      notesMode,
      inheritedDocuments,
      sourcePart
    ) {
      return renderPlacementBlocks(sortPowerPointBlocks(
        collectPowerPointBlocks(
          documentNode,
          notesMode,
          inheritedDocuments || [],
          sourcePart || ""
        )
      ), "\n\n");
    }

    async function extractOoxmlPowerPointText(zip) {
      var presentationPart = "ppt/presentation.xml";
      var presentation = await zip.extractXml(presentationPart);
      var presentationRelationships = await readRelationships(zip, presentationPart);
      var relationMap = relationshipMap(presentationRelationships);
      var slideIds = elementsByLocalName(presentation, "sldId");
      var sections = [];
      var documentCache = Object.create(null);
      var relationshipsCache = Object.create(null);

      async function cachedDocument(part) {
        if (!documentCache[part]) {
          documentCache[part] = await zip.extractXml(part);
        }
        return documentCache[part];
      }

      async function cachedRelationships(part) {
        if (!relationshipsCache[part]) {
          relationshipsCache[part] = await readRelationships(zip, part);
        }
        return relationshipsCache[part];
      }

      var slideIndex;
      for (slideIndex = 0; slideIndex < slideIds.length; slideIndex += 1) {
        checkCancelled();
        var id = relationshipIdAttribute(slideIds[slideIndex]);
        var relationship = relationMap[id];
        if (!relationship || relationship.external || !zip.has(relationship.part)) {
          continue;
        }
        var slideDocument = await cachedDocument(relationship.part);
        var root = slideDocument.documentElement;
        var show = root ? root.getAttribute("show") : "";
        var idShow = slideIds[slideIndex].getAttribute("show");
        var hidden = show === "0" || show === "false" ||
          idShow === "0" || idShow === "false";
        var slideRelationships = await cachedRelationships(relationship.part);
        var inheritedDocuments = [];
        var layoutRelationship = slideRelationships.find(function (candidate) {
          return candidate.kind === "slidelayout" &&
            !candidate.external &&
            zip.has(candidate.part);
        });
        if (layoutRelationship) {
          var layoutDocument = await cachedDocument(layoutRelationship.part);
          inheritedDocuments.push(layoutDocument);
          var layoutRelationships = await cachedRelationships(layoutRelationship.part);
          var masterRelationship = layoutRelationships.find(function (candidate) {
            return candidate.kind === "slidemaster" &&
              !candidate.external &&
              zip.has(candidate.part);
          });
          if (masterRelationship) {
            inheritedDocuments.push(await cachedDocument(masterRelationship.part));
          }
        }
        var section = "===== スライド " + (slideIndex + 1) +
          (hidden ? "（非表示）" : "") + " =====\n";
        section += extractDrawingText(
          slideDocument,
          false,
          inheritedDocuments,
          relationship.part
        );
        var notesRelationship = slideRelationships.find(function (candidate) {
          return candidate.kind === "notesslide" &&
            !candidate.external &&
            zip.has(candidate.part);
        });
        if (notesRelationship) {
          var notesDocument = await cachedDocument(notesRelationship.part);
          var notes = extractDrawingText(
            notesDocument,
            true,
            [],
            notesRelationship.part
          );
          if (hasMeaningfulText(notes)) {
            section += "\n\n----- 発表者ノート -----\n" + notes;
          }
        }
        sections.push(section);
        await cooperativeYield(slideIndex + 1, 1);
      }
      return sections.join("\n\n");
    }

    async function extractOoxmlText(zip, family, warnings) {
      if (family === "word") {
        return extractOoxmlWordText(zip, warnings);
      }
      if (family === "excel") {
        return extractOoxmlExcelText(zip);
      }
      return extractOoxmlPowerPointText(zip);
    }

    function ooxmlPrefix(family, category) {
      var root = family === "word" ? "word" :
        (family === "excel" ? "xl" : "ppt");
      return root + "/" + category;
    }

    function wordFibInfo(wordBytes) {
      requireRange(wordBytes, 0, 32, "Word FIB base");
      ensure(u16(wordBytes, 0) === 0xA5EC, "WORD_FIB",
        "WordDocument streamのFIB識別子が一致しません。");
      var flags = u16(wordBytes, 10);
      var nFib = u16(wordBytes, 2);
      ensure(nFib >= 0x00C1, "WORD_VERSION",
        "Word 97より前のBinary File Formatには対応していません。");
      var position = 32;
      var csw = u16(wordBytes, position);
      position += 2;
      ensure(csw >= 0x000E, "WORD_FIB",
        "Word FIBのfibRgWカウントが不足しています。");
      requireRange(wordBytes, position, csw * 2, "Word FIB fibRgW");
      position += csw * 2;
      var cslw = u16(wordBytes, position);
      position += 2;
      ensure(cslw >= 0x0016, "WORD_FIB",
        "Word FIBのfibRgLwカウントが不足しています。");
      var fibRgLwOffset = position;
      requireRange(wordBytes, position, cslw * 4, "Word FIB fibRgLw");
      position += cslw * 4;
      var cbRgFcLcb = u16(wordBytes, position);
      position += 2;
      requireRange(
        wordBytes,
        position,
        cbRgFcLcb * 8,
        "Word FIB fibRgFcLcb"
      );

      function readFcLcbPair(index, name) {
        if (index >= cbRgFcLcb) {
          return {
            name: name,
            fc: 0,
            lcb: 0,
            present: false
          };
        }
        var pairOffset = position + index * 8;
        return {
          name: name,
          fc: u32(wordBytes, pairOffset),
          lcb: u32(wordBytes, pairOffset + 4),
          present: true
        };
      }

      var counts = {
        main: u32(wordBytes, fibRgLwOffset + 3 * 4),
        footnote: u32(wordBytes, fibRgLwOffset + 4 * 4),
        header: u32(wordBytes, fibRgLwOffset + 5 * 4),
        macro: u32(wordBytes, fibRgLwOffset + 6 * 4),
        annotation: u32(wordBytes, fibRgLwOffset + 7 * 4),
        endnote: u32(wordBytes, fibRgLwOffset + 8 * 4),
        textbox: u32(wordBytes, fibRgLwOffset + 9 * 4),
        headerTextbox: u32(wordBytes, fibRgLwOffset + 10 * 4)
      };
      var fcLcb = {
        plcffndRef: readFcLcbPair(2, "PlcffndRef"),
        plcffndTxt: readFcLcbPair(3, "PlcffndTxt"),
        plcfSed: readFcLcbPair(6, "PlcfSed"),
        plcfBtePapx: readFcLcbPair(13, "PlcfBtePapx"),
        clx: readFcLcbPair(33, "Clx"),
        plcSpaMom: readFcLcbPair(40, "PlcSpaMom"),
        plcSpaHdr: readFcLcbPair(41, "PlcSpaHdr"),
        plcfendRef: readFcLcbPair(46, "PlcfendRef"),
        plcfendTxt: readFcLcbPair(47, "PlcfendTxt"),
        plcftxbxTxt: readFcLcbPair(56, "PlcftxbxTxt"),
        plcfHdrtxbxTxt: readFcLcbPair(58, "PlcfHdrtxbxTxt")
      };
      ensure(fcLcb.clx.present, "WORD_FIB",
        "Word FIBにClxのfc/lcbペアがありません。");
      return {
        nFib: nFib,
        lid: u16(wordBytes, 6),
        encrypted: (flags & 0x0100) !== 0 || (flags & 0x8000) !== 0,
        tableName: (flags & 0x0200) !== 0 ? "1Table" : "0Table",
        fcClx: fcLcb.clx.fc,
        lcbClx: fcLcb.clx.lcb,
        csw: csw,
        cslw: cslw,
        cbRgFcLcb: cbRgFcLcb,
        fcLcb: fcLcb,
        counts: counts
      };
    }

    function decodeWordCompressedUnicode(bytes) {
      var mapped = {
        0x82: 0x201A,
        0x83: 0x0192,
        0x84: 0x201E,
        0x85: 0x2026,
        0x86: 0x2020,
        0x87: 0x2021,
        0x88: 0x02C6,
        0x89: 0x2030,
        0x8A: 0x0160,
        0x8B: 0x2039,
        0x8C: 0x0152,
        0x91: 0x2018,
        0x92: 0x2019,
        0x93: 0x201C,
        0x94: 0x201D,
        0x95: 0x2022,
        0x96: 0x2013,
        0x97: 0x2014,
        0x98: 0x02DC,
        0x99: 0x2122,
        0x9A: 0x0161,
        0x9B: 0x203A,
        0x9C: 0x0153,
        0x9F: 0x0178
      };
      var output = [];
      var index;
      for (index = 0; index < bytes.length; index += 1) {
        output.push(String.fromCharCode(mapped[bytes[index]] || bytes[index]));
      }
      return output.join("");
    }

    function parseWordPieceTable(wordBytes, tableBytes, fib) {
      ensure(fib.lcbClx > 0, "WORD_CLX", "Word Piece TableのClxがありません。");
      requireRange(tableBytes, fib.fcClx, fib.lcbClx, "Word Clx");
      var position = fib.fcClx;
      var end = fib.fcClx + fib.lcbClx;
      var plcBytes = null;
      var prcGrpprls = [];
      while (position < end) {
        var marker = tableBytes[position];
        position += 1;
        if (marker === 0x01) {
          requireRange(tableBytes, position, 2, "Word Prc");
          var grpprlLength = u16(tableBytes, position);
          position += 2;
          ensure(grpprlLength <= 0x3FA2,
            "WORD_PCD", "Word PrcDataのgrpprl長が不正です。");
          requireRange(tableBytes, position, grpprlLength, "Word grpprl");
          ensure(prcGrpprls.length < 0x8000, "WORD_PCD",
            "Word RgPrcの定義数が安全上限を超えています。");
          prcGrpprls.push(
            tableBytes.subarray(position, position + grpprlLength)
          );
          position += grpprlLength;
        } else if (marker === 0x02) {
          requireRange(tableBytes, position, 4, "Word Pcdt");
          var plcLength = u32(tableBytes, position);
          position += 4;
          requireRange(tableBytes, position, plcLength, "Word PlcPcd");
          plcBytes = tableBytes.subarray(position, position + plcLength);
          position += plcLength;
          break;
        } else {
          fail("WORD_CLX", "Word Clxに未対応のrecord markerがあります。");
        }
      }
      ensure(plcBytes && plcBytes.length >= 4 && (plcBytes.length - 4) % 12 === 0,
        "WORD_PIECE_TABLE", "Word PlcPcdのサイズが不正です。");
      var pieceCount = (plcBytes.length - 4) / 12;
      ensure(pieceCount > 0 && pieceCount <= 1000000, "WORD_PIECE_TABLE",
        "Word Piece数が不正です。");
      var cpArraySize = (pieceCount + 1) * 4;
      var pieces = [];
      var index;
      for (index = 0; index < pieceCount; index += 1) {
        var cpStart = u32(plcBytes, index * 4);
        var cpEnd = u32(plcBytes, (index + 1) * 4);
        ensure(cpEnd >= cpStart, "WORD_PIECE_TABLE",
          "Word PieceのCP範囲が逆転しています。");
        var characterCount = cpEnd - cpStart;
        var pcdOffset = cpArraySize + index * 8;
        requireRange(plcBytes, pcdOffset, 8, "Word PCD");
        var prm = u16(plcBytes, pcdOffset + 6);
        var prmGrpprl = null;
        if ((prm & 0x0001) !== 0) {
          var prmIndex = prm >>> 1;
          ensure(prmIndex < prcGrpprls.length, "WORD_PCD",
            "Word Prm1が範囲外のPrcを参照しています。");
          prmGrpprl = prcGrpprls[prmIndex];
        }
        var encodedFc = u32(plcBytes, pcdOffset + 2);
        var compressed = (encodedFc & 0x40000000) !== 0;
        var fileOffset = encodedFc & 0x3FFFFFFF;
        if (compressed) {
          ensure(fileOffset % 2 === 0, "WORD_PCD",
            "Word圧縮Pieceのfile offsetが不正です。");
          fileOffset = fileOffset / 2;
        }
        var byteLength = compressed ? characterCount : characterCount * 2;
        requireRange(wordBytes, fileOffset, byteLength, "Word Piece text");
        var pieceBytes = wordBytes.subarray(fileOffset, fileOffset + byteLength);
        var text = compressed ?
          decodeWordCompressedUnicode(pieceBytes) :
          decodeUtf16Le(pieceBytes);
        if (text.length > characterCount) {
          text = text.slice(0, characterCount);
        }
        pieces.push({
          cpStart: cpStart,
          cpEnd: cpEnd,
          fileOffset: fileOffset,
          byteLength: byteLength,
          bytesPerCharacter: compressed ? 1 : 2,
          compressed: compressed,
          prm: prm,
          prmGrpprl: prmGrpprl,
          text: text
        });
      }
      return pieces;
    }

    function wordPiecesRange(pieces, start, count) {
      var end = start + count;
      var output = "";
      var index;
      for (index = 0; index < pieces.length; index += 1) {
        var piece = pieces[index];
        var overlapStart = Math.max(start, piece.cpStart);
        var overlapEnd = Math.min(end, piece.cpEnd);
        if (overlapEnd > overlapStart) {
          var localStart = overlapStart - piece.cpStart;
          var localEnd = overlapEnd - piece.cpStart;
          output += piece.text.substring(localStart, localEnd);
        }
      }
      return output;
    }

    function normalizeWordBinaryControls(text) {
      var output = "";
      var fieldStack = [];
      var index;
      for (index = 0; index < text.length; index += 1) {
        var code = text.charCodeAt(index);
        if (code === 0x13) {
          fieldStack.push(false);
          continue;
        }
        if (code === 0x14 && fieldStack.length) {
          fieldStack[fieldStack.length - 1] = true;
          continue;
        }
        if (code === 0x15 && fieldStack.length) {
          fieldStack.pop();
          continue;
        }
        if (fieldStack.length && !fieldStack[fieldStack.length - 1]) {
          continue;
        }
        if (code === 0x0D || code === 0x0B || code === 0x0C) {
          output += "\n";
        } else if (code === 0x09 || code === 0x07) {
          output += "\t";
        } else if (code >= 0x20 || code === 0x0A) {
          output += text.charAt(index);
        }
      }
      return output
        .replace(/\t+\n/g, "\n")
        .replace(/\n\t+/g, "\n");
    }

    function requireWordTablePair(tableBytes, pair) {
      ensure(pair && pair.present && pair.lcb > 0, "WORD_ANNOTATION_PLC",
        (pair && pair.name ? pair.name : "Word PLC") + "がありません。");
      requireRange(tableBytes, pair.fc, pair.lcb, pair.name);
      return tableBytes.subarray(pair.fc, pair.fc + pair.lcb);
    }

    function parseWordAnnotationReferencePlc(
      tableBytes,
      pair,
      mainCharacterCount
    ) {
      var bytes = requireWordTablePair(tableBytes, pair);
      ensure(bytes.length >= 4 && (bytes.length - 4) % 6 === 0,
        "WORD_ANNOTATION_PLC", pair.name + "のサイズが不正です。");
      var referenceCount = (bytes.length - 4) / 6;
      ensure(referenceCount <= 1000000, "WORD_ANNOTATION_PLC",
        pair.name + "の参照数が安全上限を超えています。");
      var cpBytes = (referenceCount + 1) * 4;
      var references = [];
      var previousCp = -1;
      var index;
      for (index = 0; index <= referenceCount; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        var cp = u32(bytes, index * 4);
        ensure(cp > previousCp, "WORD_ANNOTATION_PLC",
          pair.name + "のCPが昇順ではないか、重複しています。");
        if (index < referenceCount) {
          ensure(cp < mainCharacterCount, "WORD_ANNOTATION_PLC",
            pair.name + "の参照CPが本文範囲外です。");
          references.push({
            index: index,
            cp: cp,
            automatic: u16(bytes, cpBytes + index * 2) !== 0
          });
        }
        previousCp = cp;
      }
      return references;
    }

    function parseWordAnnotationTextPlc(
      tableBytes,
      pair,
      storyCharacterCount
    ) {
      var bytes = requireWordTablePair(tableBytes, pair);
      ensure(bytes.length >= 8 && bytes.length % 4 === 0,
        "WORD_ANNOTATION_PLC", pair.name + "のサイズが不正です。");
      var cpCount = bytes.length / 4;
      var definitionCount = cpCount - 2;
      ensure(definitionCount >= 0 && definitionCount <= 1000000,
        "WORD_ANNOTATION_PLC", pair.name + "の定義数が不正です。");
      var cps = [];
      var previousCp = -1;
      var index;
      for (index = 0; index < cpCount - 1; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        var cp = u32(bytes, index * 4);
        ensure(cp > previousCp, "WORD_ANNOTATION_PLC",
          pair.name + "のCPが昇順ではないか、重複しています。");
        ensure(cp < storyCharacterCount, "WORD_ANNOTATION_PLC",
          pair.name + "のCPが注釈ストーリー範囲外です。");
        cps.push(cp);
        previousCp = cp;
      }
      ensure(
        cps.length && cps[cps.length - 1] === storyCharacterCount - 1,
        "WORD_ANNOTATION_PLC",
        pair.name + "の本文範囲終端CPがFIB文字数と一致しません。"
      );
      var ranges = [];
      for (index = 0; index < definitionCount; index += 1) {
        ranges.push({
          index: index,
          start: cps[index],
          end: cps[index + 1]
        });
      }
      return ranges;
    }

    function legacyWordAnnotationDefinitions(pieces, baseCp, ranges) {
      var definitions = [];
      var index;
      for (index = 0; index < ranges.length; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        var range = ranges[index];
        var raw = wordPiecesRange(
          pieces,
          baseCp + range.start,
          range.end - range.start
        );
        definitions.push({
          id: String(range.index),
          index: range.index,
          text: normalizeWordBinaryControls(raw).replace(/\n+$/g, ""),
          order: index,
          referenced: false
        });
      }
      return definitions;
    }

    function throwIfLegacyWordAnnotationSafetyError(error) {
      throwIfCategoryLimitOrCancelled(error);
      if (error instanceof AppError &&
        (error.code === "BOUNDS" ||
         error.code === "WORD_CLX" ||
         error.code === "WORD_PIECE_TABLE" ||
         error.code === "WORD_PCD" ||
         error.code === "WORD_TEXTBOX_SAFETY" ||
         error.code === "WORD_TEXTBOX_LIMIT" ||
         error.code === "WORD_PAPX_SAFETY" ||
         error.code === "WORD_PAPX_LIMIT")) {
        throw error;
      }
    }

    function readLegacyWordAnnotationSet(
      kind,
      fib,
      tableBytes,
      pieces,
      story,
      warnings,
      warningSeen
    ) {
      var isFootnote = kind === "footnote";
      var label = translationValue(
        isFootnote ?
          "warning.subject.footnote" : "warning.subject.endnote"
      );
      var referencePair = fib.fcLcb[
        isFootnote ? "plcffndRef" : "plcfendRef"
      ];
      var textPair = fib.fcLcb[
        isFootnote ? "plcffndTxt" : "plcfendTxt"
      ];
      var empty = {
        usable: true,
        references: [],
        definitions: [],
        fallbackText: ""
      };
      if (!story.count) {
        if (referencePair && referencePair.present && referencePair.lcb > 0) {
          try {
            empty.references = parseWordAnnotationReferencePlc(
              tableBytes,
              referencePair,
              fib.counts.main
            );
          } catch (error) {
            throwIfLegacyWordAnnotationSafetyError(error);
            addWordExtractionWarning(
              warnings,
              warningSeen,
              "legacy-empty-story-reference:" + kind,
              warningValue("warning.legacyAnnotationEmptyStory", {
                subject: label
              })
            );
          }
        }
        return empty;
      }
      try {
        var references = parseWordAnnotationReferencePlc(
          tableBytes,
          referencePair,
          fib.counts.main
        );
        var ranges = parseWordAnnotationTextPlc(
          tableBytes,
          textPair,
          story.count
        );
        if (references.length !== ranges.length) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "legacy-count-mismatch:" + kind,
            warningValue("warning.legacyAnnotationCountMismatch", {
              subject: label
            })
          );
        }
        return {
          usable: true,
          references: references,
          definitions: legacyWordAnnotationDefinitions(
            pieces,
            story.baseCp,
            ranges
          ),
          fallbackText: ""
        };
      } catch (error) {
        throwIfLegacyWordAnnotationSafetyError(error);
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-plc-fallback:" + kind,
          warningValue(
            "warning.legacyAnnotationPositionFallback",
            { subject: label },
            error
          )
        );
        return {
          usable: false,
          references: [],
          definitions: [],
          fallbackText: renderLegacyWordStory(story.raw)
        };
      }
    }

    function emptyLegacyWordTextboxSet() {
      return {
        usable: true,
        definitions: [],
        byAnchorCp: Object.create(null),
        fallbackText: ""
      };
    }

    function requireLegacyWordTextboxTablePair(tableBytes, pair) {
      ensure(pair && pair.present && pair.lcb > 0, "WORD_TEXTBOX_PLC",
        (pair && pair.name ? pair.name : "Word textbox PLC") + "がありません。");
      requireRange(tableBytes, pair.fc, pair.lcb, pair.name);
      return tableBytes.subarray(pair.fc, pair.fc + pair.lcb);
    }

    async function validateLegacyWordTextboxReuseChain(
      definitions,
      validationCounts
    ) {
      var definitionCount = definitions.length;
      if (!definitionCount) {
        return;
      }
      var visited = Object.create(null);
      var currentIndex = definitionCount - 1;
      var traversalCount = 0;
      while (currentIndex !== -1) {
        if (traversalCount >= definitionCount) {
          validationCounts.reuseCycle += 1;
          break;
        }
        if (!Number.isSafeInteger(currentIndex) ||
          currentIndex < 0 ||
          currentIndex >= definitionCount) {
          validationCounts.invalidReuseIndex += 1;
          break;
        }
        var currentKey = String(currentIndex);
        if (visited[currentKey]) {
          validationCounts.reuseCycle += 1;
          break;
        }
        visited[currentKey] = true;
        traversalCount += 1;

        var definition = definitions[currentIndex];
        if (!definition.reusable) {
          validationCounts.reusePointsToNormal += 1;
          break;
        }
        var nextReuseIndex = definition.nextReuseIndex;
        var reusableCount = definition.reusableCount;
        if (!Number.isSafeInteger(reusableCount) ||
          reusableCount < 0 ||
          reusableCount > definitionCount) {
          validationCounts.invalidReuseCount += 1;
        }
        if (nextReuseIndex === currentIndex) {
          validationCounts.reuseCycle += 1;
          break;
        }
        if (nextReuseIndex === -1) {
          if (reusableCount !== 0) {
            validationCounts.invalidReuseCount += 1;
          }
          break;
        }
        if (!Number.isSafeInteger(nextReuseIndex) ||
          nextReuseIndex < 0 ||
          nextReuseIndex >= definitionCount) {
          validationCounts.invalidReuseIndex += 1;
          break;
        }
        var nextDefinition = definitions[nextReuseIndex];
        if (!nextDefinition.reusable) {
          validationCounts.reusePointsToNormal += 1;
          break;
        }
        if (reusableCount === 0 ||
          nextDefinition.reusableCount !== reusableCount - 1) {
          validationCounts.invalidReuseCount += 1;
        }
        currentIndex = nextReuseIndex;
        await cooperativeYield(traversalCount, 1000);
      }

      var index;
      for (index = 0; index < definitionCount; index += 1) {
        if (definitions[index].reusable && !visited[String(index)]) {
          validationCounts.invalidReuseCount += 1;
        }
        await cooperativeYield(index + 1, 1000);
      }
    }

    async function parseLegacyWordTextboxTextPlc(
      tableBytes,
      pair,
      story,
      source
    ) {
      var bytes = requireLegacyWordTextboxTablePair(tableBytes, pair);
      ensure(bytes.length >= 4 && (bytes.length - 4) % 26 === 0,
        "WORD_TEXTBOX_SAFETY", pair.name + "のサイズが不正です。");
      var definitionCount = (bytes.length - 4) / 26;
      ensure(definitionCount <= 1000000, "WORD_TEXTBOX_LIMIT",
        pair.name + "の定義数が安全上限を超えています。");
      var cpBytes = (definitionCount + 1) * 4;
      ensure(
        Number.isSafeInteger(cpBytes) &&
          cpBytes <= bytes.length &&
          story &&
          typeof story.raw === "string" &&
          story.raw.length >= story.count,
        "WORD_TEXTBOX_SAFETY",
        pair.name + "のCP配列またはテキストストーリー範囲が不正です。"
      );
      var cps = [];
      var previousCp = -1;
      var index;
      for (index = 0; index <= definitionCount; index += 1) {
        var cp = u32(bytes, index * 4);
        ensure(cp > previousCp, "WORD_TEXTBOX_PLC",
          pair.name + "のCPが昇順ではないか、重複しています。");
        ensure(cp <= story.count, "WORD_TEXTBOX_SAFETY",
          pair.name + "のCPがテキストボックスストーリー範囲外です。");
        cps.push(cp);
        previousCp = cp;
        await cooperativeYield(index + 1, 1000);
      }

      var definitions = [];
      var validationCounts = {
        invalidRange: 0,
        missingParagraphEnd: 0,
        invalidLid: 0,
        invalidChainCount: 0,
        invalidEditCount: 0,
        invalidUndoId: 0,
        invalidReusableUndoId: 0,
        unsupportedFlags: 0,
        invalidReusable: 0,
        invalidFinal: 0,
        invalidReuseIndex: 0,
        invalidReuseCount: 0,
        reuseCycle: 0,
        reusePointsToNormal: 0
      };
      for (index = 0; index < definitionCount; index += 1) {
        var dataOffset = cpBytes + index * 22;
        requireRange(bytes, dataOffset, 22, pair.name + " FTXBXS");
        var unionFirstValue = i32(bytes, dataOffset);
        var unionSecondValue = i32(bytes, dataOffset + 4);
        var reusableFlag = u16(bytes, dataOffset + 8);
        var storedReusable = (reusableFlag & 0x0001) !== 0;
        var unsupportedReusableFlags = reusableFlag & 0xFFFE;
        var isFinal = index === definitionCount - 1;
        var reusable = isFinal || storedReusable;
        var cTxbx = reusable ? null : unionFirstValue;
        var cTxbxEdit = reusable ? null : unionSecondValue;
        var nextReuseIndex = reusable ? unionFirstValue : null;
        var reusableCount = reusable ? unionSecondValue : null;
        var itxbxsDest = u32(bytes, dataOffset + 10);
        var lid = u32(bytes, dataOffset + 14);
        var txidUndo = u32(bytes, dataOffset + 18);
        var start = cps[index];
        var end = cps[index + 1];
        var rangeLength = end - start;
        var rawText = story.raw.substring(start, end);
        var text = "";
        var fallbackReason = null;
        if (unsupportedReusableFlags) {
          validationCounts.unsupportedFlags += 1;
        }
        if (!reusable) {
          text = normalizeWordBinaryControls(rawText).replace(/\n+$/g, "");
          if (rangeLength <= 1) {
            fallbackReason = "invalid-range";
            validationCounts.invalidRange += 1;
          } else if (rawText.charCodeAt(rawText.length - 1) !== 0x0D) {
            fallbackReason = "missing-paragraph-end";
            validationCounts.missingParagraphEnd += 1;
          } else if (!lid) {
            fallbackReason = "invalid-lid";
            validationCounts.invalidLid += 1;
          } else if (cTxbx < 1) {
            fallbackReason = "invalid-chain-count";
            validationCounts.invalidChainCount += 1;
          } else if (cTxbxEdit !== 0) {
            fallbackReason = "invalid-edit-count";
            validationCounts.invalidEditCount += 1;
          } else if (txidUndo !== 0) {
            fallbackReason = "invalid-undo-id";
            validationCounts.invalidUndoId += 1;
          } else if (unsupportedReusableFlags) {
            fallbackReason = "unsupported-flags";
          }
        } else {
          var reusableInvalid =
            (storedReusable && rangeLength !== 1) ||
            lid !== 0;
          if (reusableInvalid) {
            if (isFinal) {
              validationCounts.invalidFinal += 1;
            } else {
              validationCounts.invalidReusable += 1;
            }
          }
          if (txidUndo !== 0) {
            validationCounts.invalidReusableUndoId += 1;
          }
        }
        definitions.push({
          index: index,
          start: start,
          end: end,
          rangeLength: rangeLength,
          cTxbx: cTxbx,
          cTxbxEdit: cTxbxEdit,
          nextReuseIndex: nextReuseIndex,
          reusableCount: reusableCount,
          isFinal: isFinal,
          storedReusable: storedReusable,
          itxbxsDest: itxbxsDest,
          lid: lid,
          txidUndo: txidUndo,
          reusable: reusable,
          chainShapeCount: reusable ? 0 : cTxbx,
          text: text,
          placed: false,
          fallbackReason: fallbackReason,
          source: source,
          anchor: null,
          anchorCp: null
        });
        await cooperativeYield(index + 1, 1000);
      }
      await validateLegacyWordTextboxReuseChain(
        definitions,
        validationCounts
      );
      return {
        definitions: definitions,
        firstCp: cps.length ? cps[0] : 0,
        finalCp: cps.length ? cps[cps.length - 1] : 0,
        validationCounts: validationCounts
      };
    }

    async function parseLegacyWordShapeAnchorPlc(
      tableBytes,
      pair,
      story
    ) {
      var bytes = requireLegacyWordTextboxTablePair(tableBytes, pair);
      ensure(bytes.length >= 4 && (bytes.length - 4) % 30 === 0,
        "WORD_TEXTBOX_SAFETY", pair.name + "のサイズが不正です。");
      var anchorCount = (bytes.length - 4) / 30;
      ensure(anchorCount <= 1000000, "WORD_TEXTBOX_LIMIT",
        pair.name + "のアンカー数が安全上限を超えています。");
      var cpBytes = (anchorCount + 1) * 4;
      ensure(
        Number.isSafeInteger(cpBytes) &&
          cpBytes <= bytes.length &&
          story &&
          typeof story.raw === "string" &&
          story.raw.length >= story.count,
        "WORD_TEXTBOX_SAFETY",
        pair.name + "のCP配列またはアンカー先ストーリー範囲が不正です。"
      );
      var cps = [];
      var previousCp = -1;
      var index;
      for (index = 0; index <= anchorCount; index += 1) {
        var cp = u32(bytes, index * 4);
        ensure(cp > previousCp, "WORD_TEXTBOX_PLC",
          pair.name + "のCPが昇順ではないか、重複しています。");
        if (index < anchorCount) {
          ensure(cp < story.count, "WORD_TEXTBOX_SAFETY",
            pair.name + "のアンカーCPがストーリー範囲外です。");
        }
        cps.push(cp);
        previousCp = cp;
        await cooperativeYield(index + 1, 1000);
      }

      var anchors = [];
      for (index = 0; index < anchorCount; index += 1) {
        var dataOffset = cpBytes + index * 26;
        requireRange(bytes, dataOffset, 26, pair.name + " SPA");
        var left = i32(bytes, dataOffset + 4);
        var top = i32(bytes, dataOffset + 8);
        var right = i32(bytes, dataOffset + 12);
        var bottom = i32(bytes, dataOffset + 16);
        var coordinateAvailable = right >= left && bottom >= top;
        var markerCode = story.raw.charCodeAt(cps[index]);
        anchors.push({
          cp: cps[index],
          spid: u32(bytes, dataOffset),
          order: index,
          markerValid: markerCode === 0x08,
          markerCode: markerCode,
          x: coordinateAvailable ? left : null,
          y: coordinateAvailable ? top : null,
          width: coordinateAvailable ? right - left : null,
          height: coordinateAvailable ? bottom - top : null
        });
        await cooperativeYield(index + 1, 1000);
      }
      return anchors;
    }

    function compareLegacyWordTextboxDefinitions(left, right) {
      var leftHasCoordinates = left.anchor &&
        typeof left.anchor.y === "number" &&
        typeof left.anchor.x === "number";
      var rightHasCoordinates = right.anchor &&
        typeof right.anchor.y === "number" &&
        typeof right.anchor.x === "number";
      if (leftHasCoordinates !== rightHasCoordinates) {
        return leftHasCoordinates ? -1 : 1;
      }
      if (leftHasCoordinates) {
        if (left.anchor.y !== right.anchor.y) {
          return left.anchor.y - right.anchor.y;
        }
        if (left.anchor.x !== right.anchor.x) {
          return left.anchor.x - right.anchor.x;
        }
      }
      if (left.anchor.order !== right.anchor.order) {
        return left.anchor.order - right.anchor.order;
      }
      return left.index - right.index;
    }

    async function readLegacyWordTextboxSet(
      source,
      fib,
      tableBytes,
      story,
      anchorStory,
      warnings,
      warningSeen
    ) {
      var isHeader = source === "header";
      var sourceLabel = translationValue(
        isHeader ?
          "warning.subject.headerFooterTextbox" :
          "warning.subject.mainTextbox"
      );
      var sectionLabel = translationValue(
        isHeader ?
          "warning.section.unplacedHeaderFooterTextbox" :
          "warning.section.unplacedMainTextbox"
      );
      if (!story.count) {
        return emptyLegacyWordTextboxSet();
      }

      var textPair = fib.fcLcb[
        isHeader ? "plcfHdrtxbxTxt" : "plcftxbxTxt"
      ];
      var anchorPair = fib.fcLcb[
        isHeader ? "plcSpaHdr" : "plcSpaMom"
      ];
      var parsedText;
      try {
        parsedText = await parseLegacyWordTextboxTextPlc(
          tableBytes,
          textPair,
          story,
          source
        );
      } catch (error) {
        throwIfLegacyWordAnnotationSafetyError(error);
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-text-fallback:" + source,
          warningValue(
            "warning.legacyTextboxPlcFallback",
            { subject: sourceLabel, section: sectionLabel },
            error
          )
        );
        return {
          usable: false,
          definitions: [],
          byAnchorCp: Object.create(null),
          fallbackText: renderLegacyWordStory(story.raw)
        };
      }

      var validationCounts = parsedText.validationCounts;
      if (validationCounts.invalidRange) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-range:" + source,
          warningValue(
            "warning.legacyTextboxInvalidRange",
            { subject: sourceLabel, section: sectionLabel },
            "",
            validationCounts.invalidRange
          )
        );
      }
      if (validationCounts.missingParagraphEnd) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-missing-paragraph-end:" + source,
          warningValue(
            "warning.legacyTextboxMissingParagraphEnd",
            { subject: sourceLabel, section: sectionLabel },
            "",
            validationCounts.missingParagraphEnd
          )
        );
      }
      if (validationCounts.invalidLid) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-lid:" + source,
          warningValue(
            "warning.legacyTextboxMetadataInvalid",
            { subject: sourceLabel, section: sectionLabel, field: "lid" },
            "",
            validationCounts.invalidLid
          )
        );
      }
      if (validationCounts.invalidChainCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-chain:" + source,
          warningValue(
            "warning.legacyTextboxMetadataInvalid",
            { subject: sourceLabel, section: sectionLabel, field: "cTxbx" },
            "",
            validationCounts.invalidChainCount
          )
        );
      }
      if (validationCounts.invalidEditCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-edit:" + source,
          warningValue(
            "warning.legacyTextboxMetadataInvalid",
            { subject: sourceLabel, section: sectionLabel, field: "cTxbxEdit" },
            "",
            validationCounts.invalidEditCount
          )
        );
      }
      if (validationCounts.invalidUndoId) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-undo:" + source,
          warningValue(
            "warning.legacyTextboxMetadataInvalid",
            { subject: sourceLabel, section: sectionLabel, field: "txidUndo" },
            "",
            validationCounts.invalidUndoId
          )
        );
      }
      if (validationCounts.unsupportedFlags) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-unsupported-flags:" + source,
          warningValue(
            "warning.legacyTextboxUnsupportedFlags",
            { subject: sourceLabel },
            "",
            validationCounts.unsupportedFlags
          )
        );
      }
      var hasInvalidReuseManagement = !!(
        validationCounts.invalidReusable ||
        validationCounts.invalidFinal ||
        validationCounts.invalidReusableUndoId ||
        validationCounts.invalidReuseIndex ||
        validationCounts.invalidReuseCount ||
        validationCounts.reuseCycle ||
        validationCounts.reusePointsToNormal
      );
      if (hasInvalidReuseManagement) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-reuse-management:" + source,
          warningValue("warning.legacyTextboxReuseManagement", {
            subject: sourceLabel
          })
        );
      }

      var fallbackTextParts = [];
      if (parsedText.firstCp > 0) {
        var prefix = renderLegacyWordStory(
          story.raw.substring(0, parsedText.firstCp)
        );
        if (hasMeaningfulText(prefix)) {
          fallbackTextParts.push(prefix);
        }
      }
      if (parsedText.finalCp < story.count) {
        var suffix = renderLegacyWordStory(
          story.raw.substring(parsedText.finalCp)
        );
        if (hasMeaningfulText(suffix)) {
          fallbackTextParts.push(suffix);
        }
      }
      if (fallbackTextParts.length) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-text-range:" + source,
          warningValue("warning.legacyTextboxTextRangeMismatch", {
            subject: sourceLabel,
            section: sectionLabel
          })
        );
      }

      var anchors;
      try {
        anchors = await parseLegacyWordShapeAnchorPlc(
          tableBytes,
          anchorPair,
          anchorStory
        );
      } catch (error) {
        throwIfLegacyWordAnnotationSafetyError(error);
        var fallbackIndex;
        for (fallbackIndex = 0;
          fallbackIndex < parsedText.definitions.length;
          fallbackIndex += 1) {
          var fallbackDefinition = parsedText.definitions[fallbackIndex];
          if (!fallbackDefinition.reusable &&
            !fallbackDefinition.fallbackReason) {
            fallbackDefinition.fallbackReason = "missing-anchor";
          }
          await cooperativeYield(fallbackIndex + 1, 1000);
        }
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-anchor-fallback:" + source,
          warningValue(
            "warning.legacyTextboxAnchorPlcFallback",
            { subject: sourceLabel, section: sectionLabel },
            error
          )
        );
        return {
          usable: false,
          definitions: parsedText.definitions,
          byAnchorCp: Object.create(null),
          fallbackText: fallbackTextParts.join("\n\n")
        };
      }

      var anchorsByShapeId = Object.create(null);
      var invalidAnchorsByShapeId = Object.create(null);
      var invalidMarkerCpSeen = Object.create(null);
      var invalidMarkerCount = 0;
      var duplicateAnchorCount = 0;
      var anchorIndex;
      for (anchorIndex = 0; anchorIndex < anchors.length; anchorIndex += 1) {
        var anchor = anchors[anchorIndex];
        if (!anchor.markerValid) {
          var invalidCpKey = String(anchor.cp);
          if (!invalidMarkerCpSeen[invalidCpKey]) {
            invalidMarkerCpSeen[invalidCpKey] = true;
            invalidMarkerCount += 1;
          }
          if (anchor.spid &&
            !invalidAnchorsByShapeId[String(anchor.spid)]) {
            invalidAnchorsByShapeId[String(anchor.spid)] = anchor;
          }
          await cooperativeYield(anchorIndex + 1, 1000);
          continue;
        }
        if (!anchor.spid) {
          await cooperativeYield(anchorIndex + 1, 1000);
          continue;
        }
        var anchorKey = String(anchor.spid);
        if (anchorsByShapeId[anchorKey]) {
          duplicateAnchorCount += 1;
        } else {
          anchorsByShapeId[anchorKey] = anchor;
        }
        await cooperativeYield(anchorIndex + 1, 1000);
      }
      if (invalidMarkerCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-invalid-anchor-marker:" + source,
          warningValue(
            "warning.legacyTextboxInvalidAnchorMarker",
            {
              subject: translationValue(
                isHeader ?
                  "warning.subject.headerFooterShapeAnchor" :
                  "warning.subject.mainShapeAnchor"
              ),
              section: sectionLabel
            },
            "",
            invalidMarkerCount
          )
        );
      }
      if (duplicateAnchorCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-duplicate-anchor:" + source,
          warningValue(
            "warning.legacyTextboxDuplicateAnchor",
            { subject: sourceLabel },
            "",
            duplicateAnchorCount
          )
        );
      }

      var byAnchorCp = Object.create(null);
      var missingAnchorCount = 0;
      var missingCoordinateCount = 0;
      var missingCoordinateSeen = Object.create(null);
      var definitionIndex;
      for (definitionIndex = 0;
        definitionIndex < parsedText.definitions.length;
        definitionIndex += 1) {
        var definition = parsedText.definitions[definitionIndex];
        if (definition.reusable ||
          definition.fallbackReason ||
          !hasMeaningfulText(definition.text)) {
          await cooperativeYield(definitionIndex + 1, 1000);
          continue;
        }
        var matchedAnchor = anchorsByShapeId[String(definition.lid)];
        if (!matchedAnchor) {
          if (invalidAnchorsByShapeId[String(definition.lid)]) {
            definition.fallbackReason = "invalid-anchor-marker";
          } else {
            definition.fallbackReason = "missing-anchor";
            missingAnchorCount += 1;
          }
          await cooperativeYield(definitionIndex + 1, 1000);
          continue;
        }
        definition.anchor = matchedAnchor;
        definition.anchorCp = matchedAnchor.cp;
        if ((matchedAnchor.x === null || matchedAnchor.y === null) &&
          !missingCoordinateSeen[String(matchedAnchor.spid)]) {
          missingCoordinateSeen[String(matchedAnchor.spid)] = true;
          missingCoordinateCount += 1;
        }
        var cpKey = String(matchedAnchor.cp);
        if (!byAnchorCp[cpKey]) {
          byAnchorCp[cpKey] = [];
        }
        byAnchorCp[cpKey].push(definition);
        await cooperativeYield(definitionIndex + 1, 1000);
      }
      if (missingAnchorCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-missing-anchor:" + source,
          warningValue(
            "warning.legacyTextboxMissingAnchor",
            { subject: sourceLabel, section: sectionLabel },
            "",
            missingAnchorCount
          )
        );
      }
      if (missingCoordinateCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-textbox-coordinate:" + source,
          warningValue(
            "warning.legacyTextboxMissingCoordinate",
            { subject: sourceLabel },
            "",
            missingCoordinateCount
          )
        );
      }

      var cpKeys = Object.keys(byAnchorCp);
      var cpKeyIndex;
      for (cpKeyIndex = 0; cpKeyIndex < cpKeys.length; cpKeyIndex += 1) {
        byAnchorCp[cpKeys[cpKeyIndex]].sort(
          compareLegacyWordTextboxDefinitions
        );
        await cooperativeYield(cpKeyIndex + 1, 500);
      }
      return {
        usable: true,
        definitions: parsedText.definitions,
        byAnchorCp: byAnchorCp,
        fallbackText: fallbackTextParts.join("\n\n")
      };
    }

    function collectLegacyWordTextboxesAtCp(textboxSet, cp, target) {
      if (!textboxSet) {
        return;
      }
      var definitions = textboxSet.byAnchorCp[String(cp)];
      if (!definitions) {
        return;
      }
      var index;
      for (index = 0; index < definitions.length; index += 1) {
        if (!definitions[index].placed) {
          target.push(definitions[index]);
        }
      }
    }

    function appendLegacyWordParagraphTextboxes(raw, definitions) {
      if (!definitions.length) {
        return raw;
      }
      definitions.sort(compareLegacyWordTextboxDefinitions);
      var output = raw;
      var index;
      for (index = 0; index < definitions.length; index += 1) {
        var definition = definitions[index];
        if (definition.placed || !hasMeaningfulText(definition.text)) {
          continue;
        }
        if (output.length &&
          !/[\r\n\u000B\u000C]$/.test(output)) {
          output += "\r";
        }
        output += definition.text + "\r";
        definition.placed = true;
      }
      return output;
    }

    async function renderLegacyWordHeaderStory(raw, textboxSet) {
      var combined = "";
      var paragraphTextboxes = [];
      var cp;
      for (cp = 0; cp < raw.length; cp += 1) {
        collectLegacyWordTextboxesAtCp(
          textboxSet,
          cp,
          paragraphTextboxes
        );
        var code = raw.charCodeAt(cp);
        if (code === 0x0C) {
          combined = appendLegacyWordParagraphTextboxes(
            combined,
            paragraphTextboxes
          );
          paragraphTextboxes = [];
          combined += raw.charAt(cp);
          await cooperativeYield(cp + 1, 1000);
          continue;
        }
        combined += raw.charAt(cp);
        if (code === 0x0D) {
          combined = appendLegacyWordParagraphTextboxes(
            combined,
            paragraphTextboxes
          );
          paragraphTextboxes = [];
        }
        await cooperativeYield(cp + 1, 1000);
      }
      combined = appendLegacyWordParagraphTextboxes(
        combined,
        paragraphTextboxes
      );
      return renderLegacyWordStory(combined);
    }

    async function collectLegacyWordTextboxFallback(textboxSet) {
      var parts = [];
      var count = 0;
      if (hasMeaningfulText(textboxSet.fallbackText)) {
        parts.push(textboxSet.fallbackText);
      }
      var index;
      for (index = 0; index < textboxSet.definitions.length; index += 1) {
        var definition = textboxSet.definitions[index];
        if (!definition.reusable &&
          !definition.placed &&
          hasMeaningfulText(definition.text)) {
          parts.push(definition.text);
          count += 1;
        }
        await cooperativeYield(index + 1, 1000);
      }
      return {
        text: parts.join("\n\n"),
        count: count
      };
    }

    function wordSprmOperandLength(
      bytes,
      operandOffset,
      end,
      sprm,
      errorCode,
      label
    ) {
      var code = errorCode || "WORD_SECTION_PLC";
      var context = label || "Word section sprm";
      var spra = (sprm >>> 13) & 0x07;
      var length;
      if (spra === 0 || spra === 1) {
        length = 1;
      } else if (spra === 2 || spra === 4 || spra === 5) {
        length = 2;
      } else if (spra === 3) {
        length = 4;
      } else if (spra === 7) {
        length = 3;
      } else {
        ensure(operandOffset < end, code,
          context + "の可変長operandがありません。");
        if (sprm === 0xD608) {
          ensure(operandOffset + 2 <= end, code,
            context + "のsprmTDefTable長がありません。");
          var tableLength = u16(bytes, operandOffset);
          ensure(tableLength >= 1, code,
            context + "のsprmTDefTable長が不正です。");
          length = 1 + tableLength;
        } else if (sprm === 0xC615 &&
          bytes[operandOffset] === 0xFF) {
          ensure(operandOffset + 2 <= end, code,
            context + "のsprmPChgTabs削除件数がありません。");
          var deleteTabCount = bytes[operandOffset + 1];
          ensure(deleteTabCount <= 64, code,
            context + "のsprmPChgTabs削除件数が不正です。");
          var addTabCountOffset = operandOffset + 2 +
            deleteTabCount * 4;
          ensure(addTabCountOffset < end, code,
            context + "のsprmPChgTabs追加件数がありません。");
          var addTabCount = bytes[addTabCountOffset];
          ensure(addTabCount <= 64, code,
            context + "のsprmPChgTabs追加件数が不正です。");
          length = 3 + deleteTabCount * 4 + addTabCount * 3;
        } else {
          length = 1 + bytes[operandOffset];
        }
      }
      ensure(
        Number.isSafeInteger(operandOffset + length) &&
          operandOffset + length <= end,
        code,
        context + "のoperandが範囲外です。"
      );
      return length;
    }

    function legacyWordPageBreakBeforeFromGrpprl(
      bytes,
      start,
      end,
      initialValue,
      errorCode,
      label
    ) {
      var pageBreakBefore = initialValue;
      var position = start;
      while (position < end) {
        ensure(position + 2 <= end, errorCode,
          label + "のSPRMが途中で終了しています。");
        var sprm = u16(bytes, position);
        position += 2;
        var operandLength = wordSprmOperandLength(
          bytes,
          position,
          end,
          sprm,
          errorCode,
          label
        );
        if (sprm === 0x2407) {
          pageBreakBefore = bytes[position] !== 0;
        }
        position += operandLength;
      }
      return pageBreakBefore;
    }

    function legacyWordPcdPageBreakBefore(piece, initialValue) {
      if (!piece) {
        return initialValue;
      }
      var prm = piece.prm;
      if ((prm & 0x0001) === 0) {
        var isprm = (prm >>> 1) & 0x007F;
        if (isprm === 0x09) {
          return ((prm >>> 8) & 0x00FF) !== 0;
        }
        return initialValue;
      }
      return legacyWordPageBreakBeforeFromGrpprl(
        piece.prmGrpprl,
        0,
        piece.prmGrpprl.length,
        initialValue,
        "WORD_PAPX",
        "Word PCD Prm1 grpprl"
      );
    }

    function buildLegacyWordFilePieceIndex(pieces) {
      var entries = [];
      var index;
      for (index = 0; index < pieces.length; index += 1) {
        var piece = pieces[index];
        entries.push({
          piece: piece,
          order: index,
          start: piece.fileOffset,
          end: piece.fileOffset + piece.byteLength
        });
      }
      entries.sort(function (left, right) {
        if (left.start !== right.start) {
          return left.start - right.start;
        }
        return left.order - right.order;
      });
      var endpoints = entries.slice(0);
      endpoints.sort(function (left, right) {
        if (left.end !== right.end) {
          return left.end - right.end;
        }
        return left.order - right.order;
      });

      function heapPush(heap, order) {
        var position = heap.length;
        heap.push(order);
        while (position > 0) {
          var parent = Math.floor((position - 1) / 2);
          if (heap[parent] <= order) {
            break;
          }
          heap[position] = heap[parent];
          position = parent;
        }
        heap[position] = order;
      }

      function heapRemoveRoot(heap) {
        var last = heap.pop();
        if (!heap.length) {
          return;
        }
        var position = 0;
        while (position * 2 + 1 < heap.length) {
          var child = position * 2 + 1;
          var right = child + 1;
          if (right < heap.length && heap[right] < heap[child]) {
            child = right;
          }
          if (heap[child] >= last) {
            break;
          }
          heap[position] = heap[child];
          position = child;
        }
        heap[position] = last;
      }

      var active = [];
      var heap = [];
      var segments = [];
      var startIndex = 0;
      var endIndex = 0;
      while (startIndex < entries.length || endIndex < endpoints.length) {
        var nextStart = startIndex < entries.length ?
          entries[startIndex].start : null;
        var nextEnd = endIndex < endpoints.length ?
          endpoints[endIndex].end : null;
        var position = nextEnd === null ||
          (nextStart !== null && nextStart < nextEnd) ?
            nextStart : nextEnd;
        while (endIndex < endpoints.length &&
          endpoints[endIndex].end === position) {
          active[endpoints[endIndex].order] = false;
          endIndex += 1;
        }
        while (startIndex < entries.length &&
          entries[startIndex].start === position) {
          var entry = entries[startIndex];
          if (entry.end > entry.start) {
            active[entry.order] = true;
            heapPush(heap, entry.order);
          }
          startIndex += 1;
        }
        while (heap.length && !active[heap[0]]) {
          heapRemoveRoot(heap);
        }
        nextStart = startIndex < entries.length ?
          entries[startIndex].start : null;
        nextEnd = endIndex < endpoints.length ?
          endpoints[endIndex].end : null;
        var limit = nextEnd === null ||
          (nextStart !== null && nextStart < nextEnd) ?
            nextStart : nextEnd;
        if (heap.length && limit !== null && limit > position) {
          var activePiece = pieces[heap[0]];
          var previous = segments.length ?
            segments[segments.length - 1] : null;
          if (previous &&
            previous.end === position &&
            previous.piece === activePiece) {
            previous.end = limit;
          } else {
            segments.push({
              start: position,
              end: limit,
              piece: activePiece
            });
          }
        }
      }
      return {
        segments: segments,
        endpoints: endpoints
      };
    }

    function legacyWordCpAtFileOffset(pieceIndex, fileOffset) {
      var segments = pieceIndex.segments;
      var low = 0;
      var high = segments.length - 1;
      var segmentIndex = -1;
      while (low <= high) {
        var middle = Math.floor((low + high) / 2);
        if (segments[middle].start <= fileOffset) {
          segmentIndex = middle;
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }
      if (segmentIndex >= 0 &&
        fileOffset < segments[segmentIndex].end) {
        var piece = segments[segmentIndex].piece;
        var delta = fileOffset - piece.fileOffset;
        if (delta % piece.bytesPerCharacter !== 0) {
          return null;
        }
        return piece.cpStart + delta / piece.bytesPerCharacter;
      }

      var endpoints = pieceIndex.endpoints;
      low = 0;
      high = endpoints.length - 1;
      var endpointIndex = -1;
      while (low <= high) {
        middle = Math.floor((low + high) / 2);
        if (endpoints[middle].end <= fileOffset) {
          endpointIndex = middle;
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }
      if (endpointIndex < 0 ||
        endpoints[endpointIndex].end !== fileOffset) {
        return null;
      }
      return endpoints[endpointIndex].piece.cpEnd;
    }

    function legacyWordPieceAtCp(pieces, cp) {
      var low = 0;
      var high = pieces.length - 1;
      var candidate = -1;
      while (low <= high) {
        var middle = Math.floor((low + high) / 2);
        if (pieces[middle].cpStart <= cp) {
          candidate = middle;
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }
      if (candidate >= 0 && cp < pieces[candidate].cpEnd) {
        return pieces[candidate];
      }
      return null;
    }

    async function parseLegacyWordPapxBte(tableBytes, pair) {
      var bytes = requireWordTablePair(tableBytes, pair);
      ensure(bytes.length >= 4 && (bytes.length - 4) % 8 === 0,
        "WORD_PAPX", "PlcfBtePapxのサイズが不正です。");
      var bteCount = (bytes.length - 4) / 8;
      ensure(bteCount <= 1000000, "WORD_PAPX_LIMIT",
        "PlcfBtePapxの定義数が安全上限を超えています。");
      var fcBytes = (bteCount + 1) * 4;
      ensure(
        Number.isSafeInteger(fcBytes) && fcBytes <= bytes.length,
        "WORD_PAPX_LIMIT",
        "PlcfBtePapxのFC配列サイズが不正です。"
      );
      var fcs = [];
      var previousFc = -1;
      var index;
      for (index = 0; index <= bteCount; index += 1) {
        var fc = u32(bytes, index * 4);
        ensure(fc > previousFc, "WORD_PAPX",
          "PlcfBtePapxのFCが昇順ではないか、重複しています。");
        fcs.push(fc);
        previousFc = fc;
        await cooperativeYield(index + 1, 1000);
      }
      var entries = [];
      for (index = 0; index < bteCount; index += 1) {
        entries.push({
          index: index,
          startFc: fcs[index],
          endFc: fcs[index + 1],
          pageNumber: u32(bytes, fcBytes + index * 4) & 0x003FFFFF
        });
        await cooperativeYield(index + 1, 1000);
      }
      return entries;
    }

    async function parseLegacyWordPapxFkp(
      wordBytes,
      entry,
      pieces,
      pieceFileIndex,
      mainCharacterCount
    ) {
      ensure(
        entry &&
          Number.isSafeInteger(entry.startFc) &&
          Number.isSafeInteger(entry.endFc) &&
          Number.isSafeInteger(entry.pageNumber) &&
          entry.startFc >= 0 &&
          entry.endFc > entry.startFc &&
          entry.pageNumber >= 0,
        "WORD_PAPX_SAFETY",
        "PAPX BTEのFC範囲またはページ番号が不正です。"
      );
      var pageOffset = entry.pageNumber * 512;
      ensure(Number.isSafeInteger(pageOffset), "WORD_PAPX_LIMIT",
        "PAPX FKPのoffsetが安全な整数範囲外です。");
      requireRange(wordBytes, pageOffset, 512, "Word PAPX FKP");
      var page = wordBytes.subarray(pageOffset, pageOffset + 512);
      var paragraphCount = page[511];
      ensure(paragraphCount >= 1,
        "WORD_PAPX", "PAPX FKPの段落数が不正です。");
      ensure(paragraphCount <= 0x1D,
        "WORD_PAPX_LIMIT", "PAPX FKPの段落数が安全上限を超えています。");
      var fcBytes = (paragraphCount + 1) * 4;
      var bxBytes = paragraphCount * 13;
      var dataStart = fcBytes + bxBytes;
      ensure(
        Number.isSafeInteger(fcBytes) &&
          Number.isSafeInteger(bxBytes) &&
          Number.isSafeInteger(dataStart),
        "WORD_PAPX_LIMIT",
        "PAPX FKPのrgfcまたはBxPapサイズを安全に計算できません。"
      );
      ensure(dataStart <= 511, "WORD_PAPX_SAFETY",
        "PAPX FKPのrgfcまたはBxPapが範囲外です。");
      requireRange(page, 0, fcBytes, "PAPX FKP aFC");
      requireRange(page, fcBytes, bxBytes, "PAPX FKP BX");

      var fcs = [];
      var previousFc = -1;
      var index;
      for (index = 0; index <= paragraphCount; index += 1) {
        var fc = u32(page, index * 4);
        ensure(fc > previousFc, "WORD_PAPX",
          "PAPX FKPのFCが昇順ではないか、重複しています。");
        fcs.push(fc);
        previousFc = fc;
        await cooperativeYield(index + 1, 1000);
      }

      var pageBreakCps = Object.create(null);
      var unmappedCount = 0;
      var outOfRangeParagraphCount = 0;
      var disjointBteCount = 0;
      var fkpFirstFc = fcs[0];
      var fkpLastFc = fcs[fcs.length - 1];
      if (!(fkpLastFc > entry.startFc &&
        fkpFirstFc < entry.endFc)) {
        disjointBteCount = 1;
        return {
          pageBreakCps: pageBreakCps,
          unmappedCount: unmappedCount,
          outOfRangeParagraphCount: outOfRangeParagraphCount,
          disjointBteCount: disjointBteCount
        };
      }
      for (index = 0; index < paragraphCount; index += 1) {
        var paragraphEndFc = fcs[index + 1];
        if (!(paragraphEndFc > entry.startFc &&
          paragraphEndFc <= entry.endFc)) {
          outOfRangeParagraphCount += 1;
          await cooperativeYield(index + 1, 1000);
          continue;
        }
        var bOffset = page[fcBytes + index * 13];
        var pageBreakBefore = false;
        if (bOffset) {
          var papxOffset = bOffset * 2;
          ensure(
            Number.isSafeInteger(papxOffset) &&
              papxOffset >= dataStart &&
              papxOffset < 511,
            "WORD_PAPX_SAFETY",
            "PAPX FKPのPapxInFkp offsetが不正です。"
          );
          var cb = page[papxOffset];
          var grpprlAndIstdOffset;
          var grpprlAndIstdLength;
          if (cb === 0) {
            ensure(papxOffset + 1 < 511, "WORD_PAPX_SAFETY",
              "PAPX FKPの拡張長がありません。");
            var cbPrime = page[papxOffset + 1];
            ensure(cbPrime >= 1, "WORD_PAPX",
              "PAPX FKPの拡張長が不正です。");
            grpprlAndIstdOffset = papxOffset + 2;
            grpprlAndIstdLength = cbPrime * 2;
          } else {
            grpprlAndIstdOffset = papxOffset + 1;
            grpprlAndIstdLength = cb * 2 - 1;
          }
          ensure(grpprlAndIstdLength >= 2, "WORD_PAPX",
            "PAPX FKPのGrpPrlAndIstd長が不正です。");
          ensure(
            Number.isSafeInteger(
              grpprlAndIstdOffset + grpprlAndIstdLength
            ) &&
              grpprlAndIstdOffset + grpprlAndIstdLength <= 511,
            "WORD_PAPX_SAFETY",
            "PAPX FKPのGrpPrlAndIstdが範囲外です。"
          );
          pageBreakBefore = legacyWordPageBreakBeforeFromGrpprl(
            page,
            grpprlAndIstdOffset + 2,
            grpprlAndIstdOffset + grpprlAndIstdLength,
            false,
            "WORD_PAPX",
            "Word PAPX grpprl"
          );
        }
        var paragraphCp = legacyWordCpAtFileOffset(
          pieceFileIndex,
          fcs[index]
        );
        var paragraphLimitCp = legacyWordCpAtFileOffset(
          pieceFileIndex,
          fcs[index + 1]
        );
        if (paragraphCp === null) {
          if (pageBreakBefore) {
            unmappedCount += 1;
          }
        } else if (paragraphLimitCp === null ||
          paragraphLimitCp <= paragraphCp) {
          unmappedCount += 1;
        } else {
          pageBreakBefore = legacyWordPcdPageBreakBefore(
            legacyWordPieceAtCp(pieces, paragraphLimitCp - 1),
            pageBreakBefore
          );
          if (pageBreakBefore &&
            paragraphCp >= 0 &&
            paragraphCp < mainCharacterCount) {
            pageBreakCps[String(paragraphCp)] = true;
          }
        }
        await cooperativeYield(index + 1, 1000);
      }
      return {
        pageBreakCps: pageBreakCps,
        unmappedCount: unmappedCount,
        outOfRangeParagraphCount: outOfRangeParagraphCount,
        disjointBteCount: disjointBteCount
      };
    }

    async function collectLegacyWordPageBreakBeforeCps(
      wordBytes,
      tableBytes,
      fib,
      pieces,
      warnings,
      warningSeen
    ) {
      var pageBreakCps = Object.create(null);
      var pair = fib.fcLcb.plcfBtePapx;
      if (!pair || !pair.present || !pair.lcb) {
        return pageBreakCps;
      }
      var entries = await parseLegacyWordPapxBte(tableBytes, pair);
      var pieceFileIndex = buildLegacyWordFilePieceIndex(pieces);
      var partialErrorCount = 0;
      var unmappedCount = 0;
      var outOfRangeParagraphCount = 0;
      var disjointBteCount = 0;
      var index;
      for (index = 0; index < entries.length; index += 1) {
        try {
          var fkpResult = await parseLegacyWordPapxFkp(
            wordBytes,
            entries[index],
            pieces,
            pieceFileIndex,
            fib.counts.main
          );
          var cpKeys = Object.keys(fkpResult.pageBreakCps);
          var cpIndex;
          for (cpIndex = 0; cpIndex < cpKeys.length; cpIndex += 1) {
            pageBreakCps[cpKeys[cpIndex]] = true;
            await cooperativeYield(cpIndex + 1, 1000);
          }
          unmappedCount += fkpResult.unmappedCount;
          outOfRangeParagraphCount +=
            fkpResult.outOfRangeParagraphCount;
          disjointBteCount += fkpResult.disjointBteCount;
        } catch (error) {
          throwIfLegacyWordAnnotationSafetyError(error);
          partialErrorCount += 1;
        }
        await cooperativeYield(index + 1, 500);
      }
      if (disjointBteCount || outOfRangeParagraphCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-papx-bte-range",
          warningValue(
            "warning.legacyPapxRangeMismatch",
            null,
            "",
            disjointBteCount + outOfRangeParagraphCount
          )
        );
      }
      if (entries.length &&
        partialErrorCount === entries.length &&
        !Object.keys(pageBreakCps).length) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-papx-fallback",
          warningValue("warning.legacyPapxFallback")
        );
      } else if (partialErrorCount || unmappedCount) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "legacy-papx-partial",
          warningValue("warning.legacyPapxPartial")
        );
      }
      return pageBreakCps;
    }

    function readWordSectionBreakKind(wordBytes, fcSepx) {
      if (fcSepx < 0) {
        return 2;
      }
      requireRange(wordBytes, fcSepx, 2, "Word Sepx");
      var length = u16(wordBytes, fcSepx);
      requireRange(wordBytes, fcSepx + 2, length, "Word Sepx grpprl");
      var position = fcSepx + 2;
      var end = position + length;
      while (position < end) {
        requireRange(wordBytes, position, 2, "Word section sprm");
        var sprm = u16(wordBytes, position);
        position += 2;
        var operandLength = wordSprmOperandLength(
          wordBytes,
          position,
          end,
          sprm
        );
        if (sprm === 0x3009) {
          var kind = wordBytes[position];
          ensure(kind >= 0 && kind <= 4, "WORD_SECTION_PLC",
            "sprmSBkcの値が不正です。");
          return kind;
        }
        position += operandLength;
      }
      return 2;
    }

    function parseLegacyWordSections(wordBytes, tableBytes, fib) {
      var pair = fib.fcLcb.plcfSed;
      if (!pair || !pair.present || !pair.lcb) {
        return {
          available: false,
          sections: []
        };
      }
      var bytes = requireWordTablePair(tableBytes, pair);
      ensure(bytes.length >= 4 && (bytes.length - 4) % 16 === 0,
        "WORD_SECTION_PLC", "PlcfSedのサイズが不正です。");
      var sectionCount = (bytes.length - 4) / 16;
      ensure(sectionCount > 0 && sectionCount <= 1000000,
        "WORD_SECTION_PLC", "PlcfSedのセクション数が不正です。");
      var cpBytes = (sectionCount + 1) * 4;
      var cps = [];
      var previousCp = -1;
      var index;
      for (index = 0; index <= sectionCount; index += 1) {
        var cp = u32(bytes, index * 4);
        ensure(cp > previousCp, "WORD_SECTION_PLC",
          "PlcfSedのCPが昇順ではないか、重複しています。");
        if (index < sectionCount) {
          ensure(cp <= fib.counts.main, "WORD_SECTION_PLC",
            "PlcfSedの開始CPが本文範囲外です。");
        }
        cps.push(cp);
        previousCp = cp;
      }
      var sections = [];
      for (index = 0; index < sectionCount; index += 1) {
        if (index > 0 && index % 1000 === 0) {
          checkCancelled();
        }
        var sedOffset = cpBytes + index * 12;
        var fcSepx = i32(bytes, sedOffset + 2);
        sections.push({
          start: cps[index],
          end: cps[index + 1],
          breakKind: readWordSectionBreakKind(wordBytes, fcSepx)
        });
      }
      return {
        available: true,
        sections: sections
      };
    }

    function renderLegacyWordStory(raw) {
      var rawBlocks = String(raw || "").split(/\u000C/g);
      var normalizedBlocks = [];
      var index;
      for (index = 0; index < rawBlocks.length; index += 1) {
        normalizedBlocks.push(normalizeWordBinaryControls(rawBlocks[index]));
      }
      return deduplicateTextBlocks(normalizedBlocks).join("\n\n");
    }

    function legacyWordSectionBoundaryMaps(sectionInfo, mainCharacterCount) {
      var maps = {
        byCharacter: Object.create(null),
        byBoundary: Object.create(null)
      };
      if (!sectionInfo.available) {
        return maps;
      }
      var index;
      for (index = 1; index < sectionInfo.sections.length; index += 1) {
        var boundary = sectionInfo.sections[index].start;
        if (boundary > 0 && boundary <= mainCharacterCount) {
          maps.byBoundary[String(boundary)] =
            sectionInfo.sections[index].breakKind;
          maps.byCharacter[String(boundary - 1)] =
            sectionInfo.sections[index].breakKind;
        }
      }
      return maps;
    }

    async function renderLegacyWordPages(
      rawMain,
      footnotes,
      endnotes,
      textboxes,
      sectionInfo,
      paragraphPageBreakCps,
      warnings,
      warningSeen
    ) {
      var footnoteByCp = Object.create(null);
      var endnoteByCp = Object.create(null);
      footnotes.references.forEach(function (reference) {
        footnoteByCp[String(reference.cp)] = reference;
      });
      endnotes.references.forEach(function (reference) {
        endnoteByCp[String(reference.cp)] = reference;
      });
      var boundaries = legacyWordSectionBoundaryMaps(
        sectionInfo,
        rawMain.length
      );
      var pages = [];
      var currentRaw = "";
      var currentFootnotes = [];
      var referencedEndnotes = [];
      var seenFootnotes = Object.create(null);
      var seenEndnotes = Object.create(null);
      var pageBreakSeen = false;
      var lastPageBreakCp = -1;
      var paragraphTextboxes = [];

      function appendCurrentParagraphTextboxes() {
        currentRaw = appendLegacyWordParagraphTextboxes(
          currentRaw,
          paragraphTextboxes
        );
        paragraphTextboxes = [];
      }

      function finishPage() {
        appendCurrentParagraphTextboxes();
        var text = normalizeWordBinaryControls(currentRaw);
        var pageAdded = false;
        if (hasMeaningfulText(text) || currentFootnotes.length) {
          pages.push({
            text: text,
            footnotes: currentFootnotes
          });
          pageAdded = true;
        }
        currentRaw = "";
        currentFootnotes = [];
        return pageAdded;
      }

      function finishPageAtCp(cp) {
        if (!Number.isSafeInteger(cp) ||
          cp <= 0 ||
          lastPageBreakCp === cp) {
          return false;
        }
        var pageAdded = finishPage();
        if (pageAdded) {
          pageBreakSeen = true;
          lastPageBreakCp = cp;
        }
        return pageAdded;
      }

      function register(reference, annotationSet, isFootnote) {
        var seen = isFootnote ? seenFootnotes : seenEndnotes;
        var key = String(reference.index);
        if (seen[key]) {
          return;
        }
        seen[key] = true;
        var definition = annotationSet.definitions[reference.index];
        if (!definition) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "legacy-missing-definition:" +
              (isFootnote ? "footnote:" : "endnote:") + key,
            warningValue("warning.legacyAnnotationIndexMissing", {
              subject: translationValue(
                isFootnote ?
                  "warning.subject.footnote" : "warning.subject.endnote"
              ),
              id: key
            })
          );
          return;
        }
        definition.referenced = true;
        if (isFootnote) {
          currentFootnotes.push(definition);
        } else {
          referencedEndnotes.push(definition);
        }
      }

      var cp;
      for (cp = 0; cp < rawMain.length; cp += 1) {
        await cooperativeYield(cp + 1, 10000);
        if (paragraphPageBreakCps &&
          paragraphPageBreakCps[String(cp)]) {
          finishPageAtCp(cp);
        }
        var boundaryKind = boundaries.byBoundary[String(cp)];
        if (typeof boundaryKind === "number" &&
          rawMain.charCodeAt(cp - 1) !== 0x0C) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "legacy-section-without-marker",
            warningValue("warning.legacySectionBoundaryApproximation")
          );
          if (boundaryKind >= 2) {
            finishPageAtCp(cp);
          }
        }

        collectLegacyWordTextboxesAtCp(
          textboxes,
          cp,
          paragraphTextboxes
        );
        var footnoteReference = footnoteByCp[String(cp)];
        var endnoteReference = endnoteByCp[String(cp)];
        if (footnoteReference) {
          register(footnoteReference, footnotes, true);
        }
        if (endnoteReference) {
          register(endnoteReference, endnotes, false);
        }

        var code = rawMain.charCodeAt(cp);
        var automaticReference = (
          (footnoteReference && footnoteReference.automatic) ||
          (endnoteReference && endnoteReference.automatic)
        );
        if (automaticReference && code === 0x02) {
          continue;
        }
        if (automaticReference && code !== 0x02) {
          addWordExtractionWarning(
            warnings,
            warningSeen,
            "legacy-reference-character",
            warningValue("warning.legacyReferenceCharacter")
          );
        }

        if (code === 0x0C) {
          var sectionKind = boundaries.byCharacter[String(cp)];
          if (typeof sectionKind === "number") {
            currentRaw += "\r";
            appendCurrentParagraphTextboxes();
            if (sectionKind >= 2) {
              finishPageAtCp(cp);
            }
          } else {
            appendCurrentParagraphTextboxes();
            finishPageAtCp(cp);
          }
          continue;
        }
        currentRaw += rawMain.charAt(cp);
        if (code === 0x0D) {
          appendCurrentParagraphTextboxes();
        }
      }
      finishPage();
      if (!pages.length) {
        pages.push({ text: "", footnotes: currentFootnotes });
      }

      if ((footnotes.definitions.length || footnotes.references.length) &&
        !pageBreakSeen) {
        addWordExtractionWarning(
          warnings,
          warningSeen,
          "no-page-break:legacy",
          warningValue("warning.wordNoPageBreak", {
            document: translationValue("warning.document.legacyWord")
          })
        );
      }

      var renderedPages = [];
      var pageIndex;
      for (pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
        var page = pages[pageIndex];
        var value = page.text;
        if (page.footnotes.length) {
          value += (hasMeaningfulText(value) ? "\n\n" : "") +
            "----- 脚注 -----\n" +
            page.footnotes.map(function (definition) {
              return definition.text;
            }).join("\n\n");
        }
        if (hasMeaningfulText(value)) {
          renderedPages.push(value);
        }
      }
      return {
        text: renderedPages.join("\n\n"),
        referencedEndnotes: referencedEndnotes
      };
    }

    function addLegacyWordStorySection(sections, label, raw) {
      var value = renderLegacyWordStory(raw);
      if (hasMeaningfulText(value)) {
        sections.push("===== " + label + " =====\n" + value);
      }
    }

    async function extractLegacyWordText(cfb, warnings) {
      var outputWarnings = warnings || [];
      var warningSeen = Object.create(null);
      var wordBytes = cfb.rootStreamBytes("WordDocument");
      var fib = wordFibInfo(wordBytes);
      ensure(!fib.encrypted, "ENCRYPTED",
        "Word FIBに暗号化フラグがあります。");
      var tableEntry = cfb.firstRootStream(fib.tableName);
      ensure(tableEntry, "WORD_TABLE",
        "FIBが参照する" + fib.tableName + " streamがありません。");
      var tableBytes = cfb.getStream(tableEntry);
      var pieces = parseWordPieceTable(wordBytes, tableBytes, fib);
      var storyOrder = [
        "main",
        "footnote",
        "header",
        "macro",
        "annotation",
        "endnote",
        "textbox",
        "headerTextbox"
      ];
      var stories = Object.create(null);
      var cp = 0;
      var storyIndex;
      for (storyIndex = 0; storyIndex < storyOrder.length; storyIndex += 1) {
        var storyKey = storyOrder[storyIndex];
        var storyCount = fib.counts[storyKey];
        stories[storyKey] = {
          key: storyKey,
          baseCp: cp,
          count: storyCount,
          raw: storyCount ? wordPiecesRange(pieces, cp, storyCount) : ""
        };
        cp += storyCount;
      }

      var mainTextboxes = await readLegacyWordTextboxSet(
        "main",
        fib,
        tableBytes,
        stories.textbox,
        stories.main,
        outputWarnings,
        warningSeen
      );
      var headerTextboxes = await readLegacyWordTextboxSet(
        "header",
        fib,
        tableBytes,
        stories.headerTextbox,
        stories.header,
        outputWarnings,
        warningSeen
      );
      var footnotes = readLegacyWordAnnotationSet(
        "footnote",
        fib,
        tableBytes,
        pieces,
        stories.footnote,
        outputWarnings,
        warningSeen
      );
      var endnotes = readLegacyWordAnnotationSet(
        "endnote",
        fib,
        tableBytes,
        pieces,
        stories.endnote,
        outputWarnings,
        warningSeen
      );
      var sectionInfo;
      try {
        sectionInfo = parseLegacyWordSections(wordBytes, tableBytes, fib);
      } catch (error) {
        throwIfLegacyWordAnnotationSafetyError(error);
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "legacy-section-fallback",
          warningValue(
            "warning.legacySectionFallback",
            null,
            error
          )
        );
        sectionInfo = {
          available: false,
          sections: []
        };
      }

      var paragraphPageBreakCps;
      try {
        paragraphPageBreakCps = await collectLegacyWordPageBreakBeforeCps(
          wordBytes,
          tableBytes,
          fib,
          pieces,
          outputWarnings,
          warningSeen
        );
      } catch (error) {
        throwIfLegacyWordAnnotationSafetyError(error);
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "legacy-papx-fallback",
          warningValue(
            "warning.legacyPapxFallbackWithReason",
            null,
            error
          )
        );
        paragraphPageBreakCps = Object.create(null);
      }

      var mainResult = await renderLegacyWordPages(
        stories.main.raw,
        footnotes,
        endnotes,
        mainTextboxes,
        sectionInfo,
        paragraphPageBreakCps,
        outputWarnings,
        warningSeen
      );
      var sections = [];
      if (hasMeaningfulText(mainResult.text)) {
        sections.push(mainResult.text);
      }
      var renderedHeader = await renderLegacyWordHeaderStory(
        stories.header.raw,
        headerTextboxes
      );
      if (hasMeaningfulText(renderedHeader)) {
        sections.push(
          "===== ヘッダー／フッター =====\n" + renderedHeader
        );
      }
      addLegacyWordStorySection(sections, "コメント", stories.annotation.raw);

      var finalCp = pieces.length ? pieces[pieces.length - 1].cpEnd : 0;
      if (cp < finalCp) {
        var remainder = normalizeWordBinaryControls(
          wordPiecesRange(pieces, cp, finalCp - cp)
        );
        if (hasMeaningfulText(remainder)) {
          addWordExtractionWarning(
            outputWarnings,
            warningSeen,
            "legacy-remainder",
            warningValue("warning.legacyRemainder")
          );
          sections.push("===== その他のテキスト =====\n" + remainder);
        }
      }

      var mainTextboxFallback = await collectLegacyWordTextboxFallback(
        mainTextboxes
      );
      if (hasMeaningfulText(mainTextboxFallback.text)) {
        var mainUnplacedCount = mainTextboxFallback.count;
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "legacy-textbox-output:main",
          mainUnplacedCount ?
            warningValue(
              "warning.legacyTextboxUnplaced",
              {
                subject: translationValue("warning.subject.mainTextbox"),
                section: translationValue(
                  "warning.section.unplacedMainTextbox"
                )
              },
              "",
              mainUnplacedCount
            ) :
            warningValue("warning.legacyTextboxUnplacedUnknown", {
              subject: translationValue("warning.subject.mainTextbox"),
              section: translationValue(
                "warning.section.unplacedMainTextbox"
              )
            })
        );
        sections.push(
          "===== 配置を復元できないテキストボックス =====\n" +
          mainTextboxFallback.text
        );
      }

      var headerTextboxFallback = await collectLegacyWordTextboxFallback(
        headerTextboxes
      );
      if (hasMeaningfulText(headerTextboxFallback.text)) {
        var headerUnplacedCount = headerTextboxFallback.count;
        addWordExtractionWarning(
          outputWarnings,
          warningSeen,
          "legacy-textbox-output:header",
          headerUnplacedCount ?
            warningValue(
              "warning.legacyTextboxUnplaced",
              {
                subject: translationValue(
                  "warning.subject.headerFooterTextbox"
                ),
                section: translationValue(
                  "warning.section.unplacedHeaderFooterTextbox"
                )
              },
              "",
              headerUnplacedCount
            ) :
            warningValue("warning.legacyTextboxUnplacedUnknown", {
              subject: translationValue(
                "warning.subject.headerFooterTextbox"
              ),
              section: translationValue(
                "warning.section.unplacedHeaderFooterTextbox"
              )
            })
        );
        sections.push(
          "===== 配置を復元できないヘッダー／フッター内テキストボックス =====\n" +
          headerTextboxFallback.text
        );
      }

      if (!footnotes.usable) {
        if (hasMeaningfulText(footnotes.fallbackText)) {
          sections.push(
            "===== 脚注（参照位置を復元できませんでした） =====\n" +
            footnotes.fallbackText
          );
        }
      } else {
        var unreferencedFootnotes = footnotes.definitions.filter(
          function (definition) {
            return !definition.referenced;
          }
        );
        if (unreferencedFootnotes.length) {
          addWordExtractionWarning(
            outputWarnings,
            warningSeen,
            "legacy-unreferenced:footnotes",
            warningValue(
              "warning.unreferencedAnnotation",
              {
                subject: translationValue("warning.subject.legacyFootnote")
              },
              "",
              unreferencedFootnotes.length
            )
          );
          sections.push(
            "===== 未参照の脚注 =====\n" +
            unreferencedFootnotes.map(function (definition) {
              return definition.text;
            }).join("\n\n")
          );
        }
      }

      if (!endnotes.usable) {
        if (hasMeaningfulText(endnotes.fallbackText)) {
          sections.push("===== 文末脚注 =====\n" + endnotes.fallbackText);
        }
      } else {
        var renderedEndnotes = mainResult.referencedEndnotes.slice();
        var unreferencedEndnoteCount = 0;
        endnotes.definitions.forEach(function (definition) {
          if (!definition.referenced) {
            unreferencedEndnoteCount += 1;
            renderedEndnotes.push(definition);
          }
        });
        if (unreferencedEndnoteCount) {
          addWordExtractionWarning(
            outputWarnings,
            warningSeen,
            "legacy-unreferenced:endnotes",
            warningValue(
              "warning.unreferencedAnnotation",
              {
                subject: translationValue("warning.subject.legacyEndnote")
              },
              "",
              unreferencedEndnoteCount
            )
          );
        }
        if (renderedEndnotes.length) {
          sections.push(
            "===== 文末脚注 =====\n" +
            renderedEndnotes.map(function (definition) {
              return definition.text;
            }).join("\n\n")
          );
        }
      }
      return sections.join("\n\n");
    }

    function parseBiffRecords(bytes) {
      var records = [];
      var offset = 0;
      var limit = 2000000;
      while (offset + 4 <= bytes.length) {
        if (records.length > 0 && records.length % 1000 === 0) {
          checkCancelled();
        }
        ensure(records.length < limit, "BIFF_RECORDS",
          "BIFF record数が安全上限を超えています。");
        var id = u16(bytes, offset);
        var length = u16(bytes, offset + 2);
        requireRange(bytes, offset + 4, length, "BIFF record");
        records.push({
          id: id,
          offset: offset,
          data: bytes.subarray(offset + 4, offset + 4 + length)
        });
        offset += 4 + length;
      }
      ensure(offset === bytes.length, "BIFF_TRAILING",
        "Workbook stream末尾に不完全なBIFF recordがあります。");
      return records;
    }

    function BiffSegmentCursor(segments, initialOffset) {
      this.segments = segments;
      this.segmentIndex = 0;
      this.offset = initialOffset || 0;
    }

    BiffSegmentCursor.prototype.advance = function () {
      this.segmentIndex += 1;
      this.offset = 0;
      ensure(this.segmentIndex < this.segments.length, "BIFF_SST",
        "SSTがCONTINUEの途中で終了しています。");
    };

    BiffSegmentCursor.prototype.readByteRaw = function () {
      while (this.offset >= this.segments[this.segmentIndex].length) {
        this.advance();
      }
      return this.segments[this.segmentIndex][this.offset++];
    };

    BiffSegmentCursor.prototype.readBytesRaw = function (count) {
      ensure(count >= 0 && count <= 0x7FFFFFFF, "BIFF_SST",
        "SSTの読み取りサイズが不正です。");
      var parts = [];
      var remaining = count;
      while (remaining > 0) {
        if (this.offset >= this.segments[this.segmentIndex].length) {
          this.advance();
        }
        var segment = this.segments[this.segmentIndex];
        var take = Math.min(remaining, segment.length - this.offset);
        parts.push(segment.subarray(this.offset, this.offset + take));
        this.offset += take;
        remaining -= take;
      }
      return concatBytes(parts, count);
    };

    BiffSegmentCursor.prototype.readU16Raw = function () {
      var bytes = this.readBytesRaw(2);
      return u16(bytes, 0);
    };

    BiffSegmentCursor.prototype.readU32Raw = function () {
      var bytes = this.readBytesRaw(4);
      return u32(bytes, 0);
    };

    BiffSegmentCursor.prototype.skipRaw = function (count) {
      this.readBytesRaw(count);
    };

    BiffSegmentCursor.prototype.readCharacters = function (count, highByte) {
      var output = [];
      var remaining = count;
      var unicode = highByte;
      while (remaining > 0) {
        if (this.offset >= this.segments[this.segmentIndex].length) {
          this.advance();
          var continuationOptions = this.readByteRaw();
          unicode = (continuationOptions & 0x01) !== 0;
        }
        var segment = this.segments[this.segmentIndex];
        var width = unicode ? 2 : 1;
        var available = Math.floor((segment.length - this.offset) / width);
        if (available === 0) {
          ensure(this.offset === segment.length, "BIFF_SST",
            "SSTの文字がbyte境界で分割されています。");
          continue;
        }
        var take = Math.min(remaining, available);
        var bytes = segment.subarray(this.offset, this.offset + take * width);
        output.push(unicode ? decodeUtf16Le(bytes) : decodeCompressedUnicode(bytes));
        this.offset += take * width;
        remaining -= take;
      }
      return output.join("");
    };

    function readBiffUnicodeRichString(cursor) {
      var characterCount = cursor.readU16Raw();
      var options = cursor.readByteRaw();
      var highByte = (options & 0x01) !== 0;
      var richRuns = (options & 0x08) !== 0 ? cursor.readU16Raw() : 0;
      var extensionSize = (options & 0x04) !== 0 ? cursor.readU32Raw() : 0;
      ensure(characterCount <= 0x1000000 && richRuns <= 0x100000,
        "BIFF_STRING", "BIFF文字列の長さ情報が不正です。");
      var value = cursor.readCharacters(characterCount, highByte);
      cursor.skipRaw(richRuns * 4 + extensionSize);
      return value;
    }

    function parseBiffSst(records, sstIndex) {
      var segments = [records[sstIndex].data];
      var index = sstIndex + 1;
      while (index < records.length && records[index].id === 0x003C) {
        segments.push(records[index].data);
        index += 1;
      }
      ensure(segments[0].length >= 8, "BIFF_SST", "SST headerが不足しています。");
      var total = u32(segments[0], 0);
      var unique = u32(segments[0], 4);
      ensure(unique <= total && unique <= 1000000, "BIFF_SST",
        "SST文字列数が不正です。");
      var cursor = new BiffSegmentCursor(segments, 8);
      var strings = [];
      var stringIndex;
      for (stringIndex = 0; stringIndex < unique; stringIndex += 1) {
        strings.push(readBiffUnicodeRichString(cursor));
      }
      return strings;
    }

    function readBiffStringFromBytes(bytes, offset) {
      requireRange(bytes, offset, 3, "BIFF Unicode string");
      var characterCount = u16(bytes, offset);
      var options = bytes[offset + 2];
      var position = offset + 3;
      var richRuns = 0;
      var extensionSize = 0;
      if ((options & 0x08) !== 0) {
        richRuns = u16(bytes, position);
        position += 2;
      }
      if ((options & 0x04) !== 0) {
        extensionSize = u32(bytes, position);
        position += 4;
      }
      var byteLength = characterCount * ((options & 0x01) ? 2 : 1);
      requireRange(bytes, position, byteLength + richRuns * 4 + extensionSize,
        "BIFF Unicode string");
      var textBytes = bytes.subarray(position, position + byteLength);
      return {
        text: (options & 0x01) ?
          decodeUtf16Le(textBytes) :
          decodeCompressedUnicode(textBytes),
        end: position + byteLength + richRuns * 4 + extensionSize
      };
    }

    function readBoundSheetName(data) {
      requireRange(data, 0, 8, "BoundSheet8");
      var characterCount = data[6];
      var options = data[7];
      var byteLength = characterCount * ((options & 0x01) ? 2 : 1);
      requireRange(data, 8, byteLength, "BoundSheet8 name");
      var bytes = data.subarray(8, 8 + byteLength);
      return (options & 0x01) ? decodeUtf16Le(bytes) : decodeCompressedUnicode(bytes);
    }

    function decodeRk(value) {
      var divided = (value & 0x01) !== 0;
      var integer = (value & 0x02) !== 0;
      var result;
      if (integer) {
        result = (value >> 2);
      } else {
        var buffer = new Uint8Array(8);
        writeU32(buffer, 4, value & 0xFFFFFFFC);
        result = f64(buffer, 0);
      }
      return divided ? result / 100 : result;
    }

    function biffErrorText(code) {
      var values = {
        0x00: "#NULL!",
        0x07: "#DIV/0!",
        0x0F: "#VALUE!",
        0x17: "#REF!",
        0x1D: "#NAME?",
        0x24: "#NUM!",
        0x2A: "#N/A"
      };
      return values[code] || ("#ERROR(" + code + ")");
    }

    function biffFormulaCachedValue(data) {
      requireRange(data, 6, 8, "Formula result");
      if (data[12] === 0xFF && data[13] === 0xFF) {
        var type = data[6];
        var value = data[8];
        if (type === 0) {
          return { value: "", expectsString: true };
        }
        if (type === 1) {
          return { value: value ? "TRUE" : "FALSE", expectsString: false };
        }
        if (type === 2) {
          return { value: biffErrorText(value), expectsString: false };
        }
        return { value: "", expectsString: false };
      }
      return { value: f64(data, 6), expectsString: false };
    }

    function readBiffTxoText(records, txoIndex) {
      var data = records[txoIndex].data;
      requireRange(data, 0, 16, "TxO");
      var characterCount = u16(data, 10);
      var formattingByteCount = u16(data, 12);
      if (characterCount === 0) {
        ensure(formattingByteCount === 0, "BIFF_TXO",
          "空のTxOに文字書式データがあります。");
        return "";
      }
      ensure(formattingByteCount >= 16 && formattingByteCount % 8 === 0,
        "BIFF_TXO", "TxOの文字書式サイズが不正です。");
      var remaining = characterCount;
      var parts = [];
      var recordIndex = txoIndex + 1;
      while (remaining > 0) {
        ensure(recordIndex < records.length && records[recordIndex].id === 0x003C,
          "BIFF_TXO", "TxOの文字列CONTINUEが不足しています。");
        var segment = records[recordIndex].data;
        ensure(segment.length >= 2, "BIFF_TXO",
          "TxOの文字列CONTINUEが空です。");
        var highByte = (segment[0] & 0x01) !== 0;
        var width = highByte ? 2 : 1;
        var payloadLength = segment.length - 1;
        if (remaining * width > payloadLength) {
          ensure(payloadLength % width === 0, "BIFF_TXO",
            "TxOのUnicode文字がbyte境界で分割されています。");
        }
        var available = Math.floor(payloadLength / width);
        ensure(available > 0, "BIFF_TXO",
          "TxOの文字列CONTINUEに文字がありません。");
        var take = Math.min(remaining, available);
        var textBytes = segment.subarray(1, 1 + take * width);
        parts.push(highByte ?
          decodeUtf16Le(textBytes) :
          decodeCompressedUnicode(textBytes));
        remaining -= take;
        recordIndex += 1;
      }
      return parts.join("");
    }

    function parseBiffWorkbook(cfb, warnings) {
      var workbookEntry = cfb.firstRootStream("Workbook") ||
        cfb.firstRootStream("Book");
      ensure(workbookEntry, "BIFF_WORKBOOK", "WorkbookまたはBook streamがありません。");
      var bytes = cfb.getStream(workbookEntry);
      var records = parseBiffRecords(bytes);
      ensure(records.length && records[0].id === 0x0809, "BIFF_BOF",
        "BIFF8 Workbook BOFがありません。");
      ensure(records[0].data.length >= 4 &&
        u16(records[0].data, 0) === 0x0600 &&
        u16(records[0].data, 2) === 0x0005,
        "BIFF_VERSION", "BIFF8 Workbook形式ではありません。");
      ensure(!records.some(function (record) { return record.id === 0x002F; }),
        "ENCRYPTED", "BIFF FILEPASS recordがあるため暗号化ファイルです。");

      var globalsEnd = records.findIndex(function (record) {
        return record.id === 0x000A;
      });
      ensure(globalsEnd >= 0, "BIFF_GLOBALS", "Workbook Globals EOFがありません。");
      var globals = records.slice(0, globalsEnd + 1);
      var date1904 = false;
      var sheets = [];
      var sst = [];
      var customFormats = Object.create(null);
      var xfs = [];
      var index;
      for (index = 0; index < globals.length; index += 1) {
        var record = globals[index];
        if (record.id === 0x0022 && record.data.length >= 2) {
          date1904 = u16(record.data, 0) !== 0;
        } else if (record.id === 0x0085) {
          requireRange(record.data, 0, 8, "BoundSheet8");
          sheets.push({
            offset: u32(record.data, 0),
            hidden: record.data[4] !== 0,
            type: record.data[5],
            name: readBoundSheetName(record.data)
          });
        } else if (record.id === 0x00FC) {
          sst = parseBiffSst(globals, index);
        } else if (record.id === 0x041E && record.data.length >= 5) {
          var formatId = u16(record.data, 0);
          customFormats[formatId] = readBiffStringFromBytes(record.data, 2).text;
        } else if (record.id === 0x00E0 && record.data.length >= 4) {
          xfs.push(u16(record.data, 2));
        }
      }
      ensure(sheets.length > 0, "BIFF_SHEETS", "BoundSheet8がありません。");
      sheets = sheets.filter(function (sheet) {
        if (sheet.type !== 0x00) {
          warnings.push(warningValue(
            "warning.excelNonWorksheetExcluded",
            { name: sheet.name }
          ));
          return false;
        }
        ensure(sheet.offset < bytes.length, "BIFF_SHEET_OFFSET",
          "Worksheet BOF位置が範囲外です: " + sheet.name);
        return true;
      });
      ensure(sheets.length > 0, "BIFF_SHEETS", "処理可能なWorksheetがありません。");

      function styleIsDate(xfIndex) {
        if (xfIndex < 0 || xfIndex >= xfs.length) {
          return false;
        }
        var formatId = xfs[xfIndex];
        return !!BUILTIN_DATE_FORMATS[formatId] ||
          looksLikeDateFormat(customFormats[formatId]);
      }

      function valueForStyle(value, xfIndex) {
        if (typeof value === "number") {
          if (styleIsDate(xfIndex)) {
            var converted = excelSerialToText(value, date1904);
            if (converted !== null) {
              return converted;
            }
          }
          return normalizeExcelNumericText(value);
        }
        return String(value == null ? "" : value);
      }

      var sections = [];
      sheets.forEach(function (sheet) {
        checkCancelled();
        var startIndex = records.findIndex(function (record) {
          return record.offset === sheet.offset;
        });
        ensure(startIndex >= 0 && records[startIndex].id === 0x0809,
          "BIFF_SHEET_BOF", "Worksheet BOFが見つかりません: " + sheet.name);
        var cells = [];
        var shapeTexts = [];
        var pendingFormula = null;
        var recordIndex;
        for (recordIndex = startIndex + 1; recordIndex < records.length; recordIndex += 1) {
          var cellRecord = records[recordIndex];
          if (cellRecord.id === 0x000A) {
            break;
          }
          var data = cellRecord.data;
          var row;
          var column;
          var xf;
          if (cellRecord.id === 0x00FD && data.length >= 10) {
            row = u16(data, 0);
            column = u16(data, 2);
            xf = u16(data, 4);
            var sstIndex = u32(data, 6);
            if (sstIndex < sst.length) {
              cells.push({ row: row, column: column, xf: xf, value: sst[sstIndex], formula: "" });
            } else {
              warnings.push(warningValue(
                "warning.excelSstOutOfRange",
                { name: sheet.name }
              ));
            }
          } else if (cellRecord.id === 0x0203 && data.length >= 14) {
            cells.push({
              row: u16(data, 0),
              column: u16(data, 2),
              xf: u16(data, 4),
              value: f64(data, 6),
              formula: ""
            });
          } else if (cellRecord.id === 0x027E && data.length >= 10) {
            cells.push({
              row: u16(data, 0),
              column: u16(data, 2),
              xf: u16(data, 4),
              value: decodeRk(i32(data, 6)),
              formula: ""
            });
          } else if (cellRecord.id === 0x00BD && data.length >= 12) {
            row = u16(data, 0);
            var firstColumn = u16(data, 2);
            var lastColumn = u16(data, data.length - 2);
            var rkCount = lastColumn - firstColumn + 1;
            ensure(rkCount >= 1 && 4 + rkCount * 6 + 2 === data.length,
              "BIFF_MULRK", "MulRK record lengthが不正です。");
            var rkIndex;
            for (rkIndex = 0; rkIndex < rkCount; rkIndex += 1) {
              cells.push({
                row: row,
                column: firstColumn + rkIndex,
                xf: u16(data, 4 + rkIndex * 6),
                value: decodeRk(i32(data, 6 + rkIndex * 6)),
                formula: ""
              });
            }
          } else if (cellRecord.id === 0x0205 && data.length >= 8) {
            cells.push({
              row: u16(data, 0),
              column: u16(data, 2),
              xf: u16(data, 4),
              value: data[7] ? biffErrorText(data[6]) : (data[6] ? "TRUE" : "FALSE"),
              formula: ""
            });
          } else if (cellRecord.id === 0x0006 && data.length >= 22) {
            var cached = biffFormulaCachedValue(data);
            var formulaCell = {
              row: u16(data, 0),
              column: u16(data, 2),
              xf: u16(data, 4),
              value: cached.value,
              formula: "[数式を完全に復元できません]"
            };
            cells.push(formulaCell);
            pendingFormula = cached.expectsString ? formulaCell : null;
          } else if (cellRecord.id === 0x0207 && pendingFormula) {
            pendingFormula.value = readBiffStringFromBytes(data, 0).text;
            pendingFormula = null;
          } else if (cellRecord.id === 0x0204 && data.length >= 9) {
            row = u16(data, 0);
            column = u16(data, 2);
            xf = u16(data, 4);
            cells.push({
              row: row,
              column: column,
              xf: xf,
              value: readBiffStringFromBytes(data, 6).text,
              formula: ""
            });
          } else if (cellRecord.id === 0x00D6 && data.length >= 9) {
            cells.push({
              row: u16(data, 0),
              column: u16(data, 2),
              xf: u16(data, 4),
              value: readBiffStringFromBytes(data, 6).text,
              formula: ""
            });
          } else if (cellRecord.id === 0x01B6) {
            try {
              var shapeText = readBiffTxoText(records, recordIndex);
              if (hasMeaningfulText(shapeText)) {
                shapeTexts.push(shapeText);
              }
            } catch (error) {
              warnings.push(warningValue(
                "warning.excelShapeTextFailed",
                { name: sheet.name },
                error
              ));
            }
          }
        }
        var tableCells = [];
        cells.forEach(function (cell) {
          var value = valueForStyle(cell.value, cell.xf);
          if (value !== "" || cell.formula !== "") {
            tableCells.push({
              row: cell.row,
              column: cell.column,
              value: value !== "" ? value : cell.formula
            });
          }
        });
        var legacyPart = "Workbook#" + sheet.name;
        var legacyBlocks = excelTableRowBlocks(tableCells, legacyPart);
        shapeTexts.forEach(function (shapeText, shapeIndex) {
          var shapeBlock = createPlacementBlock(
            "shape",
            shapeText,
            legacyPart,
            legacyBlocks.length + shapeIndex
          );
          shapeBlock.sourceId = legacyPart + "#txo-" + shapeIndex;
          legacyBlocks.push(shapeBlock);
        });
        var legacyMetrics = {
          defaultRowHeightEmu: 15 * 12700,
          defaultColumnWidthEmu: 8.43 * 7 * 9525,
          rowHeights: Object.create(null)
        };
        sortExcelPlacementBlocks(legacyBlocks, legacyMetrics);
        var lines = [
          "===== Sheet: " + sheet.name + (sheet.hidden ? "（非表示）" : "") + " ====="
        ];
        if (legacyBlocks.length) {
          lines.push(renderPlacementBlocks(legacyBlocks));
        }
        sections.push(lines.join("\n"));
      });
      return sections.join("\n\n");
    }

    function extractLegacyExcelText(cfb, warnings) {
      return parseBiffWorkbook(cfb, warnings);
    }

    var PPT_RECORD = {
      DOCUMENT: 1000,
      DOCUMENT_ATOM: 1001,
      SLIDE: 1006,
      SLIDE_ATOM: 1007,
      NOTES: 1008,
      NOTES_ATOM: 1009,
      SLIDE_PERSIST_ATOM: 1011,
      SLIDE_SHOW_INFO_ATOM: 1017,
      TEXT_HEADER_ATOM: 3999,
      TEXT_CHARS_ATOM: 4000,
      STYLE_TEXT_PROP_ATOM: 4001,
      TEXT_BYTES_ATOM: 4008,
      SLIDE_LIST_WITH_TEXT: 4080,
      USER_EDIT_ATOM: 4085,
      CURRENT_USER_ATOM: 4086,
      EX_OLE_OBJ_STG: 4113,
      PERSIST_DIRECTORY_ATOM: 6002
    };

    function pptRecordAt(bytes, offset, limit) {
      var endLimit = typeof limit === "number" ? limit : bytes.length;
      ensure(isRangeValid(bytes.length, offset, 8) && offset + 8 <= endLimit,
        "PPT_RECORD", "PowerPoint record headerが範囲外です。");
      var verInstance = u16(bytes, offset);
      var length = u32(bytes, offset + 4);
      ensure(length <= endLimit - offset - 8, "PPT_RECORD",
        "PowerPoint record lengthが範囲外です。");
      return {
        offset: offset,
        version: verInstance & 0x000F,
        instance: verInstance >>> 4,
        type: u16(bytes, offset + 2),
        length: length,
        payloadStart: offset + 8,
        end: offset + 8 + length,
        children: []
      };
    }

    function parsePptSequence(bytes, start, end, depth, budget) {
      var currentDepth = depth || 0;
      var sharedBudget = budget || { count: 0 };
      ensure(currentDepth <= 128, "PPT_DEPTH",
        "PowerPoint recordの入れ子が深すぎます。");
      var records = [];
      var position = start;
      while (position < end) {
        if (sharedBudget.count > 0 && sharedBudget.count % 1000 === 0) {
          checkCancelled();
        }
        ensure(end - position >= 8, "PPT_RECORD",
          "PowerPoint record末尾が不完全です。");
        sharedBudget.count += 1;
        ensure(sharedBudget.count <= 2000000, "PPT_RECORDS",
          "PowerPoint record数が安全上限を超えています。");
        var record = pptRecordAt(bytes, position, end);
        if (record.version === 0x0F && record.length > 0) {
          record.children = parsePptSequence(
            bytes,
            record.payloadStart,
            record.end,
            currentDepth + 1,
            sharedBudget
          );
        }
        records.push(record);
        position = record.end;
      }
      ensure(position === end, "PPT_RECORD", "PowerPoint record境界が一致しません。");
      return records;
    }

    function flattenPptRecords(records, output) {
      var target = output || [];
      records.forEach(function (record) {
        target.push(record);
        if (record.children.length) {
          flattenPptRecords(record.children, target);
        }
      });
      return target;
    }

    function recordsOfType(records, type) {
      return flattenPptRecords(records, []).filter(function (record) {
        return record.type === type;
      });
    }

    function parsePptPersistState(powerPointBytes, currentUserBytes, warnings) {
      var state = {
        persistOffsets: Object.create(null),
        documentPersistId: null,
        encrypted: false,
        edits: []
      };
      if (!currentUserBytes || currentUserBytes.length < 20) {
        warnings.push(warningValue("warning.powerPointCurrentUserMissing"));
        return state;
      }
      var currentUserRecord = pptRecordAt(currentUserBytes, 0, currentUserBytes.length);
      ensure(currentUserRecord.type === PPT_RECORD.CURRENT_USER_ATOM,
        "PPT_CURRENT_USER", "Current User streamのrecord typeが不正です。");
      var headerToken = u32(currentUserBytes, 12);
      if (headerToken === 0xF3D1C4DF) {
        state.encrypted = true;
        return state;
      }
      if (headerToken !== 0xE391C05F) {
        warnings.push(warningValue("warning.powerPointHeaderTokenUnknown"));
      }
      var currentEditOffset = u32(currentUserBytes, 16);
      var seenEdits = Object.create(null);
      var editCount = 0;
      while (currentEditOffset !== 0 && currentEditOffset !== 0xFFFFFFFF) {
        ensure(editCount < 10000, "PPT_EDIT_CHAIN",
          "UserEditAtom chainが安全上限を超えています。");
        ensure(!seenEdits[currentEditOffset], "PPT_EDIT_CHAIN",
          "UserEditAtom chainが循環しています。");
        seenEdits[currentEditOffset] = true;
        var edit = pptRecordAt(powerPointBytes, currentEditOffset, powerPointBytes.length);
        ensure(edit.type === PPT_RECORD.USER_EDIT_ATOM &&
          (edit.length === 0x1C || edit.length === 0x20),
          "PPT_USER_EDIT", "UserEditAtomが不正です。");
        var payload = edit.payloadStart;
        var previousEdit = u32(powerPointBytes, payload + 8);
        var persistDirectoryOffset = u32(powerPointBytes, payload + 12);
        var documentPersistId = u32(powerPointBytes, payload + 16);
        if (state.documentPersistId === null) {
          state.documentPersistId = documentPersistId;
        }
        if (edit.length === 0x20) {
          var encryptionPersistId = u32(powerPointBytes, payload + 28);
          if (encryptionPersistId !== 0 && encryptionPersistId !== 0xFFFFFFFF) {
            state.encrypted = true;
          }
        }
        state.edits.push(edit);
        var persistRecord = pptRecordAt(
          powerPointBytes,
          persistDirectoryOffset,
          powerPointBytes.length
        );
        ensure(persistRecord.type === PPT_RECORD.PERSIST_DIRECTORY_ATOM,
          "PPT_PERSIST", "PersistDirectoryAtomが見つかりません。");
        var position = persistRecord.payloadStart;
        while (position < persistRecord.end) {
          requireRange(powerPointBytes, position, 4, "PersistDirectoryEntry");
          var packed = u32(powerPointBytes, position);
          position += 4;
          var persistId = packed & 0x000FFFFF;
          var count = packed >>> 20;
          ensure(count > 0 && count <= 0x0FFF, "PPT_PERSIST",
            "PersistDirectoryEntry countが不正です。");
          requireRange(powerPointBytes, position, count * 4, "PersistDirectoryEntry offsets");
          var entryIndex;
          for (entryIndex = 0; entryIndex < count; entryIndex += 1) {
            var id = persistId + entryIndex;
            var recordOffset = u32(powerPointBytes, position + entryIndex * 4);
            if (state.persistOffsets[id] === undefined) {
              ensure(recordOffset < powerPointBytes.length, "PPT_PERSIST",
                "Persist offsetがPowerPoint Document範囲外です。");
              state.persistOffsets[id] = recordOffset;
            }
          }
          position += count * 4;
        }
        currentEditOffset = previousEdit;
        editCount += 1;
      }
      return state;
    }

    function pptTextFromRecord(record, bytes, codePage) {
      var payload = bytes.subarray(record.payloadStart, record.end);
      if (record.type === PPT_RECORD.TEXT_CHARS_ATOM) {
        return decodeUtf16Le(payload);
      }
      if (record.type === PPT_RECORD.TEXT_BYTES_ATOM) {
        return decodeSingleByte(payload, codePage);
      }
      return "";
    }

    function normalizeLegacyPptText(text, notesMode) {
      var repaired = repairSurrogates(String(text || ""))
        .replace(/\r\n?/g, "\n")
        .replace(/\u000B/g, "\n")
        .replace(/\u0000/g, "")
        .replace(/[\u0001-\u0008\u000C\u000E-\u001F\u007F]/g, "");
      var sourceLines = repaired.split("\n");
      var lines = [];
      var previous = null;
      var previousBlank = true;
      var index;
      for (index = 0; index < sourceLines.length; index += 1) {
        var line = sourceLines[index].replace(/[ \t\u3000]+$/g, "");
        var blank = line.length === 0;
        if (blank && previousBlank) {
          continue;
        }
        if (!blank && line === previous) {
          continue;
        }
        lines.push(line);
        previous = blank ? null : line;
        previousBlank = blank;
      }
      while (lines.length && lines[0] === "") {
        lines.shift();
      }
      while (lines.length && lines[lines.length - 1] === "") {
        lines.pop();
      }
      if (notesMode && lines.length > 1 && lines[lines.length - 1] === "*") {
        var hasEarlierText = false;
        for (index = 0; index < lines.length - 1; index += 1) {
          if (hasMeaningfulText(lines[index])) {
            hasEarlierText = true;
            break;
          }
        }
        if (hasEarlierText) {
          lines.pop();
        }
      }
      return lines.join("\n");
    }

    function lineSequenceIndex(haystack, needle) {
      if (!needle.length || needle.length > haystack.length) {
        return -1;
      }
      var start;
      for (start = 0; start <= haystack.length - needle.length; start += 1) {
        var matched = true;
        var index;
        for (index = 0; index < needle.length; index += 1) {
          if (haystack[start + index] !== needle[index]) {
            matched = false;
            break;
          }
        }
        if (matched) {
          return start;
        }
      }
      return -1;
    }

    function mergeLegacyPptTextCandidates(primaryText, secondaryText) {
      var primary = normalizeLegacyPptText(primaryText, false);
      var secondary = normalizeLegacyPptText(secondaryText, false);
      if (!primary) {
        return secondary;
      }
      if (!secondary || primary === secondary) {
        return primary;
      }
      var primaryLines = primary.split("\n");
      var secondaryLines = secondary.split("\n");
      if (lineSequenceIndex(primaryLines, secondaryLines) >= 0) {
        return primary;
      }
      if (lineSequenceIndex(secondaryLines, primaryLines) >= 0) {
        return secondary;
      }
      var overlap = 0;
      var limit = Math.min(primaryLines.length, secondaryLines.length);
      var count;
      for (count = 1; count <= limit; count += 1) {
        if (primaryLines.slice(primaryLines.length - count).join("\n") ===
          secondaryLines.slice(0, count).join("\n")) {
          overlap = count;
        }
      }
      if (overlap > 0) {
        return normalizeLegacyPptText(
          primaryLines.concat(secondaryLines.slice(overlap)).join("\n"),
          false
        );
      }
      var commonPrefix = 0;
      while (commonPrefix < limit &&
        primaryLines[commonPrefix] === secondaryLines[commonPrefix]) {
        commonPrefix += 1;
      }
      if (commonPrefix > 0) {
        return normalizeLegacyPptText(
          primaryLines.concat(secondaryLines.slice(commonPrefix)).join("\n"),
          false
        );
      }
      var commonSuffix = 0;
      while (commonSuffix < limit &&
        primaryLines[primaryLines.length - commonSuffix - 1] ===
          secondaryLines[secondaryLines.length - commonSuffix - 1]) {
        commonSuffix += 1;
      }
      if (commonSuffix > 0) {
        return normalizeLegacyPptText(
          primaryLines.concat(
            secondaryLines.slice(0, secondaryLines.length - commonSuffix)
          ).join("\n"),
          false
        );
      }
      return normalizeLegacyPptText(primary + "\n" + secondary, false);
    }

    function normalizePptText(text) {
      return normalizeLegacyPptText(text, false);
    }

    function collectPptText(records, bytes, codePage, textState) {
      var values = [];
      flattenPptRecords(records, []).forEach(function (record) {
        if (record.type === PPT_RECORD.TEXT_CHARS_ATOM ||
          record.type === PPT_RECORD.TEXT_BYTES_ATOM) {
          if (record.type === PPT_RECORD.TEXT_BYTES_ATOM && textState) {
            textState.usedTextBytes = true;
          }
          var value = normalizePptText(pptTextFromRecord(record, bytes, codePage));
          if (hasMeaningfulText(value)) {
            values.push(value);
          }
        }
      });
      return values.join("\n");
    }

    function collectSlideListEntries(container, bytes, codePage, textState) {
      var entries = [];
      var current = null;
      container.children.forEach(function (record) {
        if (record.type === PPT_RECORD.SLIDE_PERSIST_ATOM && record.length >= 4) {
          current = {
            persistId: u32(bytes, record.payloadStart),
            slideId: record.length >= 16 ? u32(bytes, record.payloadStart + 12) : 0,
            flags: record.length >= 8 ? u32(bytes, record.payloadStart + 4) : 0,
            textParts: []
          };
          entries.push(current);
        } else if (current) {
          var value = collectPptText([record], bytes, codePage, textState);
          if (hasMeaningfulText(value)) {
            current.textParts.push(value);
          }
        }
      });
      entries.forEach(function (entry) {
        entry.text = entry.textParts.join("\n");
      });
      return entries;
    }

    function pptSlideHidden(records, bytes) {
      var infoRecords = recordsOfType(records, PPT_RECORD.SLIDE_SHOW_INFO_ATOM);
      return infoRecords.some(function (record) {
        if (record.length >= 12) {
          return (u16(bytes, record.payloadStart + 10) & 0x0004) !== 0;
        }
        return false;
      });
    }

    function safePptRecordAtOffset(bytes, offset) {
      try {
        var record = pptRecordAt(bytes, offset, bytes.length);
        if (record.version === 0x0F && record.length > 0) {
          record.children = parsePptSequence(bytes, record.payloadStart, record.end, 1);
        }
        return record;
      } catch (error) {
        return null;
      }
    }

    function extractLegacyPowerPointText(cfb, warnings) {
      var powerPointBytes = cfb.rootStreamBytes("PowerPoint Document");
      var currentUserEntry = cfb.firstRootStream("Current User");
      var currentUserBytes = currentUserEntry ? cfb.getStream(currentUserEntry) : null;
      var state = parsePptPersistState(powerPointBytes, currentUserBytes, warnings);
      ensure(!state.encrypted, "ENCRYPTED",
        "PowerPointのUserEditAtomに暗号化persist IDがあります。");
      var codePage = 932;
      var textState = { usedTextBytes: false };
      var slideListEntries = [];
      var notesListEntries = [];
      var documentRecord = null;
      if (state.documentPersistId !== null &&
        state.persistOffsets[state.documentPersistId] !== undefined) {
        documentRecord = safePptRecordAtOffset(
          powerPointBytes,
          state.persistOffsets[state.documentPersistId]
        );
      }
      if (documentRecord && documentRecord.type === PPT_RECORD.DOCUMENT) {
        var lists = recordsOfType([documentRecord], PPT_RECORD.SLIDE_LIST_WITH_TEXT);
        lists.forEach(function (list) {
          var entries = collectSlideListEntries(
            list,
            powerPointBytes,
            codePage,
            textState
          );
          if (list.instance === 0) {
            slideListEntries = slideListEntries.concat(entries);
          } else if (list.instance === 2) {
            notesListEntries = notesListEntries.concat(entries);
          }
        });
      }
      var notesBySlideId = Object.create(null);
      notesListEntries.forEach(function (noteEntry) {
        var noteOffset = state.persistOffsets[noteEntry.persistId];
        var noteRecord = noteOffset !== undefined ?
          safePptRecordAtOffset(powerPointBytes, noteOffset) :
          null;
        if (!noteRecord || noteRecord.type !== PPT_RECORD.NOTES) {
          return;
        }
        var notesAtoms = recordsOfType([noteRecord], PPT_RECORD.NOTES_ATOM);
        if (!notesAtoms.length || notesAtoms[0].length < 4) {
          return;
        }
        var slideIdRef = u32(powerPointBytes, notesAtoms[0].payloadStart);
        if (slideIdRef !== 0) {
          notesBySlideId[String(slideIdRef)] = {
            entry: noteEntry,
            record: noteRecord
          };
        }
      });

      var sections = [];
      if (slideListEntries.length) {
        slideListEntries.forEach(function (slideEntry, index) {
          checkCancelled();
          var recordOffset = state.persistOffsets[slideEntry.persistId];
          var slideRecord = recordOffset !== undefined ?
            safePptRecordAtOffset(powerPointBytes, recordOffset) :
            null;
          var slideRecordText = "";
          if (slideRecord && slideRecord.type === PPT_RECORD.SLIDE) {
            slideRecordText = collectPptText(
              [slideRecord],
              powerPointBytes,
              codePage,
              textState
            );
          }
          var slideText = mergeLegacyPptTextCandidates(
            slideRecordText,
            slideEntry.text
          );
          var hidden = slideRecord ? pptSlideHidden([slideRecord], powerPointBytes) : false;
          var section = "===== スライド " + (index + 1) +
            (hidden ? "（非表示）" : "") + " =====\n" +
            slideText;
          var noteData = notesBySlideId[String(slideEntry.slideId)];
          if (noteData) {
            var noteRecordText = "";
            if (noteData.record) {
              noteRecordText = collectPptText(
                [noteData.record],
                powerPointBytes,
                codePage,
                textState
              );
            }
            var noteText = normalizeLegacyPptText(
              mergeLegacyPptTextCandidates(noteRecordText, noteData.entry.text),
              true
            );
            if (hasMeaningfulText(noteText)) {
              section += "\n\n----- 発表者ノート -----\n" + noteText;
            }
          }
          sections.push(section);
        });
      } else {
        warnings.push(warningValue("warning.powerPointSlideOrderFallback"));
        if (!documentRecord) {
          var topLevel = parsePptSequence(powerPointBytes, 0, powerPointBytes.length, 0);
          var documentCandidates = recordsOfType(topLevel, PPT_RECORD.DOCUMENT);
          documentRecord = documentCandidates.length ?
            documentCandidates[documentCandidates.length - 1] :
            null;
        }
        ensure(documentRecord, "PPT_DOCUMENT", "DocumentContainerが見つかりません。");
        var slides = recordsOfType([documentRecord], PPT_RECORD.SLIDE);
        slides.forEach(function (slide, index) {
          checkCancelled();
          sections.push(
            "===== スライド " + (index + 1) +
            (pptSlideHidden([slide], powerPointBytes) ? "（非表示）" : "") +
            " =====\n" +
            normalizeLegacyPptText(
              collectPptText([slide], powerPointBytes, codePage, textState),
              false
            )
          );
        });
        if (!slides.length) {
          var fallbackText = collectPptText(
            [documentRecord],
            powerPointBytes,
            codePage,
            textState
          );
          if (hasMeaningfulText(fallbackText)) {
            sections.push("===== スライドテキスト（保存順） =====\n" + fallbackText);
          }
        }
      }
      if (textState.usedTextBytes) {
        warnings.push(warningValue("warning.powerPointLegacyCodePage"));
      }
      return sections.join("\n\n");
    }

    var OFFICE_ART_BLIP_TYPES = {
      0xF01A: { extension: "emf", metafile: true },
      0xF01B: { extension: "wmf", metafile: true },
      0xF01C: { extension: "pct", metafile: true },
      0xF01D: { extension: "jpg", metafile: false },
      0xF01E: { extension: "png", metafile: false },
      0xF01F: { extension: "dib", metafile: false },
      0xF029: { extension: "tif", metafile: false },
      0xF02A: { extension: "jpg", metafile: false }
    };

    function officeArtRecordAt(bytes, offset, limit) {
      var endLimit = typeof limit === "number" ? limit : bytes.length;
      if (!isRangeValid(bytes.length, offset, 8) || offset + 8 > endLimit) {
        return null;
      }
      var verInstance = u16(bytes, offset);
      var type = u16(bytes, offset + 2);
      var length = u32(bytes, offset + 4);
      if (type < 0xF000 || type > 0xF2FF ||
        length > endLimit - offset - 8) {
        return null;
      }
      return {
        offset: offset,
        version: verInstance & 0x000F,
        instance: verInstance >>> 4,
        type: type,
        payloadStart: offset + 8,
        end: offset + 8 + length,
        length: length
      };
    }

    function collectOfficeArtCandidates(bytes) {
      var candidates = [];
      var seen = Object.create(null);
      var visitedContainers = Object.create(null);

      function add(record) {
        var key = record.offset + ":" + record.end;
        if (!seen[key]) {
          seen[key] = true;
          candidates.push(record);
        }
      }

      function walkRecord(record, depth) {
        if (!record || depth > 128) {
          return;
        }
        if (OFFICE_ART_BLIP_TYPES[record.type]) {
          add(record);
          return;
        }
        var key = record.offset + ":" + record.end;
        if (visitedContainers[key]) {
          return;
        }
        visitedContainers[key] = true;
        if (record.type === 0xF007 && record.length >= 36) {
          var embedded = officeArtRecordAt(bytes, record.payloadStart + 36, record.end);
          if (embedded) {
            walkRecord(embedded, depth + 1);
          }
        }
        if (record.version === 0x0F ||
          record.type === 0xF000 ||
          record.type === 0xF001 ||
          record.type === 0xF002 ||
          record.type === 0xF003) {
          var position = record.payloadStart;
          while (position + 8 <= record.end) {
            var child = officeArtRecordAt(bytes, position, record.end);
            if (!child) {
              break;
            }
            walkRecord(child, depth + 1);
            position = child.end;
          }
        }
      }

      var offset;
      for (offset = 0; offset + 8 <= bytes.length; offset += 1) {
        if (offset > 0 && offset % 65536 === 0) {
          checkCancelled();
        }
        var record = officeArtRecordAt(bytes, offset, bytes.length);
        if (!record) {
          continue;
        }
        if (OFFICE_ART_BLIP_TYPES[record.type] ||
          (record.version === 0x0F &&
            (record.type === 0xF000 ||
             record.type === 0xF001 ||
             record.type === 0xF002 ||
             record.type === 0xF003))) {
          walkRecord(record, 0);
          if (record.end > offset + 8) {
            offset = record.end - 1;
          }
        }
      }
      return candidates;
    }

    function detectRasterSignature(payload) {
      var limit = Math.min(payload.length, 96);
      var offset;
      for (offset = 0; offset < limit; offset += 1) {
        if (offset + 8 <= payload.length &&
          payload[offset] === 0x89 &&
          payload[offset + 1] === 0x50 &&
          payload[offset + 2] === 0x4E &&
          payload[offset + 3] === 0x47 &&
          payload[offset + 4] === 0x0D &&
          payload[offset + 5] === 0x0A &&
          payload[offset + 6] === 0x1A &&
          payload[offset + 7] === 0x0A) {
          return { offset: offset, extension: "png" };
        }
        if (offset + 3 <= payload.length &&
          payload[offset] === 0xFF &&
          payload[offset + 1] === 0xD8 &&
          payload[offset + 2] === 0xFF) {
          return { offset: offset, extension: "jpg" };
        }
        if (offset + 6 <= payload.length &&
          payload[offset] === 0x47 &&
          payload[offset + 1] === 0x49 &&
          payload[offset + 2] === 0x46 &&
          payload[offset + 3] === 0x38 &&
          (payload[offset + 4] === 0x37 || payload[offset + 4] === 0x39) &&
          payload[offset + 5] === 0x61) {
          return { offset: offset, extension: "gif" };
        }
        if (offset + 4 <= payload.length &&
          ((payload[offset] === 0x49 && payload[offset + 1] === 0x49 &&
            payload[offset + 2] === 0x2A && payload[offset + 3] === 0x00) ||
           (payload[offset] === 0x4D && payload[offset + 1] === 0x4D &&
            payload[offset + 2] === 0x00 && payload[offset + 3] === 0x2A))) {
          return { offset: offset, extension: "tif" };
        }
        if (offset + 2 <= payload.length &&
          payload[offset] === 0x42 && payload[offset + 1] === 0x4D) {
          return { offset: offset, extension: "bmp" };
        }
      }
      return null;
    }

    async function extractOfficeArtBlip(bytes, record, warnings) {
      var info = OFFICE_ART_BLIP_TYPES[record.type];
      var payload = bytes.subarray(record.payloadStart, record.end);
      if (info.metafile) {
        var uidLengths = [16, 32];
        var choice = null;
        var index;
        for (index = 0; index < uidLengths.length; index += 1) {
          var uidLength = uidLengths[index];
          if (payload.length < uidLength + 34) {
            continue;
          }
          var uncompressedSize = u32(payload, uidLength);
          var compressedSize = u32(payload, uidLength + 28);
          var compression = payload[uidLength + 32];
          var filter = payload[uidLength + 33];
          var dataOffset = uidLength + 34;
          if (compressedSize <= payload.length - dataOffset &&
            (compression === 0x00 || compression === 0xFE) &&
            filter === 0xFE &&
            uncompressedSize <= 0x7FFFFFFF) {
            choice = {
              uncompressedSize: uncompressedSize,
              compressedSize: compressedSize,
              compression: compression,
              dataOffset: dataOffset
            };
            break;
          }
        }
        if (!choice) {
          warnings.push(warningValue("warning.officeArtMetafileHeader"));
          return null;
        }
        var data = payload.subarray(
          choice.dataOffset,
          choice.dataOffset + choice.compressedSize
        );
        if (choice.compression === 0x00) {
          try {
            data = await decompressBytes(data, "deflate", choice.uncompressedSize);
          } catch (error) {
            throwIfCancelled(error);
            warnings.push(warningValue(
              "warning.officeArtMetafileDecompression"
            ));
            return null;
          }
        } else if (choice.uncompressedSize &&
          data.length !== choice.uncompressedSize) {
          warnings.push(warningValue("warning.officeArtMetafileSize"));
          return null;
        }
        return { extension: info.extension, data: data };
      }

      var signature = detectRasterSignature(payload);
      if (signature) {
        return {
          extension: signature.extension,
          data: payload.subarray(signature.offset)
        };
      }
      if (info.extension === "dib") {
        var offsets = [17, 33];
        var dibIndex;
        for (dibIndex = 0; dibIndex < offsets.length; dibIndex += 1) {
          var dibOffset = offsets[dibIndex];
          if (payload.length >= dibOffset + 4) {
            var headerSize = u32(payload, dibOffset);
            if (headerSize === 12 || headerSize === 40 ||
              headerSize === 52 || headerSize === 56 ||
              headerSize === 108 || headerSize === 124) {
              return {
                extension: "dib",
                data: payload.subarray(dibOffset)
              };
            }
          }
        }
      }
      warnings.push(warningValue("warning.officeArtBlipLocation"));
      return null;
    }

    function uniqueBinaryPush(output, candidate, index) {
      var checksum = crc32(candidate.data);
      var key = candidate.data.length + ":" +
        zeroPad(checksum.toString(16).toUpperCase(), 8);
      var matches = index[key] || [];
      var matchIndex;
      candidate.crc = checksum;
      for (matchIndex = 0; matchIndex < matches.length; matchIndex += 1) {
        // CRC narrows candidates; byte equality still protects against collisions.
        if (bytesEqual(matches[matchIndex].data, candidate.data)) {
          return false;
        }
      }
      output.push(candidate);
      matches.push(candidate);
      index[key] = matches;
      return true;
    }

    function officeArtSourcesForLegacy(cfb, family, warnings) {
      var sources = [];
      // Parent OfficeArt sources must not be selected from embedded child storages.
      if (family === "word") {
        ["Data", "0Table", "1Table"].forEach(function (name) {
          var entry = cfb.firstRootStream(name);
          if (entry) {
            sources.push(cfb.getStream(entry));
          }
        });
      } else if (family === "excel") {
        var workbookEntry = cfb.firstRootStream("Workbook") ||
          cfb.firstRootStream("Book");
        if (workbookEntry) {
          var workbookBytes = cfb.getStream(workbookEntry);
          var records = parseBiffRecords(workbookBytes);
          var index;
          for (index = 0; index < records.length; index += 1) {
            if (records[index].id === 0x00EB || records[index].id === 0x00EC) {
              var parts = [records[index].data];
              var next = index + 1;
              while (next < records.length && records[next].id === 0x003C) {
                parts.push(records[next].data);
                next += 1;
              }
              sources.push(concatBytes(parts));
              index = next - 1;
            }
          }
        }
      } else {
        var picturesEntry = cfb.firstRootStream("Pictures");
        if (picturesEntry) {
          sources.push(cfb.getStream(picturesEntry));
        } else {
          warnings.push(warningValue("warning.powerPointPicturesMissing"));
        }
      }
      return sources;
    }

    async function extractLegacyMedia(cfb, family, warnings) {
      var sources = officeArtSourcesForLegacy(cfb, family, warnings);
      var extracted = [];
      var deduplicationIndex = Object.create(null);
      var sourceIndex;
      for (sourceIndex = 0; sourceIndex < sources.length; sourceIndex += 1) {
        checkCancelled();
        var source = sources[sourceIndex];
        var candidates = collectOfficeArtCandidates(source);
        var candidateIndex;
        for (candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
          checkCancelled();
          var mediaItem = await extractOfficeArtBlip(
            source,
            candidates[candidateIndex],
            warnings
          );
          if (mediaItem && mediaItem.data.length) {
            uniqueBinaryPush(extracted, mediaItem, deduplicationIndex);
            ensure(extracted.length <= MAX_MEDIA_FILES, "MEDIA_LIMIT",
              "メディア数が安全上限を超えたため、メディアの抽出を中止しました。");
          }
          await cooperativeYield(candidateIndex + 1, 20);
        }
        await cooperativeYield(sourceIndex + 1, 1);
      }
      return extracted.map(function (mediaItem, index) {
        return {
          name: "image_" + zeroPad(index + 1, 4) + "." + mediaItem.extension,
          data: mediaItem.data
        };
      });
    }

    function readNullTerminated(bytes, start, maximum) {
      var endLimit = Math.min(bytes.length, start + maximum);
      var end = start;
      while (end < endLimit && bytes[end] !== 0) {
        end += 1;
      }
      ensure(end < endLimit, "OLE_NATIVE", "Ole10NativeのNUL終端文字列が不正です。");
      return {
        text: decodeSingleByte(bytes.subarray(start, end), 932),
        next: end + 1
      };
    }

    function parseOle10Native(bytes) {
      requireRange(bytes, 0, 8, "Ole10Native");
      var declared = u32(bytes, 0);
      ensure(declared <= bytes.length - 4 || declared === bytes.length,
        "OLE_NATIVE", "Ole10Nativeの全体サイズが不正です。");
      var starts = [6, 4];
      var attempt;
      for (attempt = 0; attempt < starts.length; attempt += 1) {
        try {
          var position = starts[attempt];
          var label = readNullTerminated(bytes, position, 4096);
          position = label.next;
          var fileName = readNullTerminated(bytes, position, 4096);
          position = fileName.next;
          requireRange(bytes, position, 4, "Ole10Native flags");
          position += 4;
          var temporary = readNullTerminated(bytes, position, 32768);
          position = temporary.next;
          requireRange(bytes, position, 4, "Ole10Native payload size");
          var dataSize = u32(bytes, position);
          position += 4;
          requireRange(bytes, position, dataSize, "Ole10Native payload");
          return {
            name: fileName.text || label.text || leafName(temporary.text) || "embedding.bin",
            data: bytes.subarray(position, position + dataSize)
          };
        } catch (error) {
          if (attempt === starts.length - 1) {
            throw error;
          }
        }
      }
      fail("OLE_NATIVE", "Ole10Nativeを解析できません。");
    }

    function scanPptRecordsByType(bytes, type) {
      var records = [];
      var seen = Object.create(null);
      var offset;
      for (offset = 0; offset + 8 <= bytes.length; offset += 1) {
        if (offset > 0 && offset % 65536 === 0) {
          checkCancelled();
        }
        try {
          var record = pptRecordAt(bytes, offset, bytes.length);
          if (record.type === type && !seen[offset]) {
            seen[offset] = true;
            records.push(record);
            offset = record.end - 1;
          }
        } catch (error) {
          // Record header candidates are expected to fail while locating top-level atoms.
        }
      }
      return records;
    }

    function pushLimitedEmbedding(output, candidate) {
      ensure(output.length < MAX_EMBEDDING_FILES, "EMBEDDING_LIMIT",
        "埋め込み数が安全上限を超えたため、埋め込みの抽出を中止しました。");
      output.push(candidate);
    }

    async function extractLegacyEmbeddings(cfb, family, warnings) {
      var output = [];
      var allocator = new NameAllocator();
      var handledEntries = Object.create(null);
      var streams = cfb.streams();
      var index;
      for (index = 0; index < streams.length; index += 1) {
        checkCancelled();
        await cooperativeYield(index + 1, 50);
        var entry = streams[index];
        var loweredName = entry.name.toLowerCase();
        if (loweredName.indexOf("ole10native") >= 0) {
          try {
            var native = parseOle10Native(cfb.getStream(entry));
            pushLimitedEmbedding(output, {
              name: allocator.allocate(native.name, true),
              data: native.data
            });
            handledEntries[entry.id] = true;
          } catch (error) {
            throwIfCategoryLimitOrCancelled(error);
            warnings.push(warningValue("warning.oleNativeFallback"));
            pushLimitedEmbedding(output, {
              name: allocator.allocate("embedding.bin", true),
              data: cfb.getStream(entry)
            });
            handledEntries[entry.id] = true;
          }
        }
      }

      for (index = 0; index < streams.length; index += 1) {
        checkCancelled();
        await cooperativeYield(index + 1, 50);
        var stream = streams[index];
        if (handledEntries[stream.id]) {
          continue;
        }
        var path = stream.path.toLowerCase();
        var name = stream.name.toLowerCase();
        var relevant = path.indexOf("objectpool/") === 0 ||
          path.indexOf("/objectpool/") >= 0 ||
          name === "package" ||
          name === "contents" ||
          path.indexOf("embedding") >= 0;
        if (relevant && stream.size > 0) {
          pushLimitedEmbedding(output, {
            name: allocator.allocate(
              sanitizeStem(stream.name.replace(/^[\u0000-\u001F]+/, ""), "embedding") + ".bin",
              true
            ),
            data: cfb.getStream(stream)
          });
        }
      }

      if (family === "powerpoint") {
        var powerPointBytes = cfb.rootStreamBytes("PowerPoint Document");
        var oleRecords = scanPptRecordsByType(powerPointBytes, PPT_RECORD.EX_OLE_OBJ_STG);
        var oleIndex;
        for (oleIndex = 0; oleIndex < oleRecords.length; oleIndex += 1) {
          checkCancelled();
          await cooperativeYield(oleIndex + 1, 20);
          var record = oleRecords[oleIndex];
          if (record.instance === 0) {
            var uncompressedOle = powerPointBytes.subarray(record.payloadStart, record.end);
            if (uncompressedOle.length) {
              pushLimitedEmbedding(output, {
                name: allocator.allocate("embedding_" +
                  zeroPad(oleIndex + 1, 4) + ".ole", true),
                data: uncompressedOle
              });
            }
            continue;
          }
          if (record.instance !== 1 || record.length <= 4) {
            warnings.push(warningValue(
              "warning.oleRecordInstanceUnsupported"
            ));
            continue;
          }
          var expectedSize = u32(powerPointBytes, record.payloadStart);
          var compressed = powerPointBytes.subarray(record.payloadStart + 4, record.end);
          if (expectedSize === 0 || expectedSize > 0x7FFFFFFF) {
            warnings.push(warningValue("warning.oleExpandedSizeInvalid"));
            continue;
          }
          try {
            var oleBytes = await decompressBytes(compressed, "deflate", expectedSize);
            pushLimitedEmbedding(output, {
              name: allocator.allocate("embedding_" +
                zeroPad(oleIndex + 1, 4) + ".ole", true),
              data: oleBytes
            });
          } catch (error) {
            throwIfCategoryLimitOrCancelled(error);
            warnings.push(warningValue("warning.oleStorageRestoreFailed"));
          }
        }
      }

      var unique = [];
      var deduplicationIndex = Object.create(null);
      var uniqueIndex;
      for (uniqueIndex = 0; uniqueIndex < output.length; uniqueIndex += 1) {
        uniqueBinaryPush(unique, output[uniqueIndex], deduplicationIndex);
        await cooperativeYield(uniqueIndex + 1, 20);
      }
      return unique;
    }

    function legacyEncryptionStatus(cfb, family) {
      if (cfb.isEncryptedPackage()) {
        return true;
      }
      if (family === "word") {
        return wordFibInfo(cfb.rootStreamBytes("WordDocument")).encrypted;
      }
      if (family === "excel") {
        var workbookEntry = cfb.firstRootStream("Workbook") ||
          cfb.firstRootStream("Book");
        if (!workbookEntry) {
          return false;
        }
        return parseBiffRecords(cfb.getStream(workbookEntry)).some(function (record) {
          return record.id === 0x002F;
        });
      }
      if (cfb.hasRootStream("EncryptedSummary")) {
        return true;
      }
      var currentUserEntry = cfb.firstRootStream("Current User");
      if (currentUserEntry) {
        try {
          var warnings = [];
          var state = parsePptPersistState(
            cfb.rootStreamBytes("PowerPoint Document"),
            cfb.getStream(currentUserEntry),
            warnings
          );
          return state.encrypted;
        } catch (error) {
          return false;
        }
      }
      return false;
    }

    async function extractLegacyText(cfb, family, warnings) {
      if (family === "word") {
        return await extractLegacyWordText(cfb, warnings);
      }
      if (family === "excel") {
        return extractLegacyExcelText(cfb, warnings);
      }
      return extractLegacyPowerPointText(cfb, warnings);
    }

    var ui = {
      languageJa: document.getElementById("language-ja"),
      languageEn: document.getElementById("language-en"),
      addFiles: document.getElementById("add-files"),
      addFolder: document.getElementById("add-folder"),
      removeSelected: document.getElementById("remove-selected"),
      clearFiles: document.getElementById("clear-files"),
      exportZip: document.getElementById("export-zip"),
      cancelProcessing: document.getElementById("cancel-processing"),
      fileInput: document.getElementById("file-input"),
      folderInput: document.getElementById("folder-input"),
      dropZone: document.getElementById("drop-zone"),
      pageDropOverlay: document.getElementById("page-drop-overlay"),
      includeSubfolders: document.getElementById("include-subfolders"),
      includeText: document.getElementById("include-text"),
      includeEmbeddings: document.getElementById("include-embeddings"),
      selectAll: document.getElementById("select-all"),
      fileList: document.getElementById("file-list"),
      progressTrack: document.querySelector(".progress-track"),
      progressBar: document.getElementById("progress-bar"),
      progressNumbers: document.getElementById("progress-numbers"),
      progressCurrent: document.getElementById("progress-current"),
      log: document.getElementById("log"),
      clearLog: document.getElementById("clear-log"),
      compatibilityWarning: document.getElementById("compatibility-warning")
    };

    var appState = {
      items: [],
      nextId: 1,
      busy: false,
      inspecting: 0,
      inspectionQueue: Promise.resolve(),
      lastProgress: 0,
      progress: {
        completed: 0,
        total: 0,
        percentage: 0,
        key: "progress.waiting",
        params: {}
      },
      logStarted: false,
      logEntries: [{
        level: "info",
        key: "log.addPrompt",
        params: {},
        pluralCount: null,
        renderedLineCount: 1
      }],
      logNodes: ui.log.firstChild ? [ui.log.firstChild] : [],
      logDomDirty: false,
      cancelRequested: false,
      cachedInputBytes: 0,
      rowByItemId: Object.create(null),
      pageDragDepth: 0,
      pageDropOverlayVisible: false
    };

    function dataTransferHasFiles(dataTransfer) {
      if (!dataTransfer || !dataTransfer.types) {
        return false;
      }
      var types = dataTransfer.types;
      if (typeof types === "string") {
        return types === "Files";
      }
      try {
        if (typeof types.contains === "function" && types.contains("Files")) {
          return true;
        }
      } catch (containsError) {
        // Fall back to indexed access for nonstandard DOMStringList objects.
      }
      var length;
      try {
        length = Number(types.length);
      } catch (lengthError) {
        return false;
      }
      if (!Number.isFinite(length) || length < 0) {
        return false;
      }
      var index;
      for (index = 0; index < length; index += 1) {
        var type = null;
        try {
          if (typeof types.item === "function") {
            type = types.item(index);
          }
        } catch (itemMethodError) {
          type = null;
        }
        if (type === null || typeof type === "undefined") {
          try {
            type = types[index];
          } catch (itemAccessError) {
            type = null;
          }
        }
        if (String(type) === "Files") {
          return true;
        }
      }
      return false;
    }

    function showPageDropOverlay() {
      if (appState.pageDropOverlayVisible || appState.busy) {
        return;
      }
      appState.pageDropOverlayVisible = true;
      ui.pageDropOverlay.hidden = false;
      ui.pageDropOverlay.setAttribute("aria-hidden", "false");
      ui.dropZone.classList.add("dragging");
    }

    function hidePageDropOverlay() {
      appState.pageDropOverlayVisible = false;
      ui.pageDropOverlay.hidden = true;
      ui.pageDropOverlay.setAttribute("aria-hidden", "true");
      ui.dropZone.classList.remove("dragging");
    }

    function resetPageDropState() {
      appState.pageDragDepth = 0;
      hidePageDropOverlay();
    }

    function renderLogLines() {
      try {
        ui.log.textContent = "";
        appState.logNodes = [];
        var entryIndex;
        var renderedEntries = [];
        var totalLines = 0;
        for (entryIndex = 0;
          entryIndex < appState.logEntries.length;
          entryIndex += 1) {
          var entry = appState.logEntries[entryIndex];
          var lines = formatLogEntryLines(entry, currentLanguage);
          entry.renderedLineCount = lines.length;
          renderedEntries.push(lines);
          totalLines += lines.length;
        }
        while (totalLines > MAX_LOG_LINES && appState.logEntries.length > 1) {
          var removedEntry = appState.logEntries.shift();
          renderedEntries.shift();
          totalLines -= removedEntry.renderedLineCount;
        }
        for (entryIndex = 0;
          entryIndex < renderedEntries.length;
          entryIndex += 1) {
          lines = renderedEntries[entryIndex];
          var lineIndex;
          for (lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
            var node = document.createTextNode(
              (appState.logNodes.length ? "\n" : "") + lines[lineIndex]
            );
            ui.log.appendChild(node);
            appState.logNodes.push(node);
          }
        }
        appState.logDomDirty = false;
      } catch (error) {
        appState.logDomDirty = true;
      }
    }

    function formatLogEntryLines(entry, language) {
      var params = {};
      var name;
      for (name in entry.params) {
        if (Object.prototype.hasOwnProperty.call(entry.params, name)) {
          params[name] = entry.params[name];
        }
      }
      var messageKey = entry.key;
      if (typeof entry.pluralCount === "number") {
        messageKey = pluralKey(entry.key, entry.pluralCount, language);
        params.count = formatNumber(entry.pluralCount, language);
      }
      var messageLines = t(messageKey, params, language)
        .replace(/\r\n?/g, "\n")
        .split("\n");
      var levelKey = entry.level === "ok" ? "success" : entry.level;
      if (levelKey !== "warning" && levelKey !== "error" &&
        levelKey !== "success") {
        levelKey = "info";
      }
      var label = t("log.level." + levelKey, null, language);
      var lines = [];
      var lineIndex;
      for (lineIndex = 0; lineIndex < messageLines.length; lineIndex += 1) {
        lines.push(
          (lineIndex === 0 ?
            "[" + label + "] " :
            "  ") + messageLines[lineIndex]
        );
      }
      return lines;
    }

    function logMessage(level, key, params, pluralCount) {
      var entry = {
        level: level,
        key: key,
        params: params || {},
        pluralCount: typeof pluralCount === "number" ? pluralCount : null,
        renderedLineCount: 0
      };
      if (!appState.logStarted) {
        appState.logEntries = [];
        appState.logNodes = [];
        appState.logStarted = true;
        try {
          ui.log.textContent = "";
        } catch (error) {
          appState.logDomDirty = true;
        }
      }
      var lines = formatLogEntryLines(entry, currentLanguage);
      entry.renderedLineCount = lines.length;
      appState.logEntries.push(entry);
      var totalLines = appState.logNodes.length + lines.length;
      var removedCount = 0;
      while (totalLines > MAX_LOG_LINES && appState.logEntries.length > 1) {
        var removedEntry = appState.logEntries.shift();
        totalLines -= removedEntry.renderedLineCount;
        removedCount += removedEntry.renderedLineCount;
      }
      try {
        // Keep normal updates incremental; full rendering is only a failure fallback.
        while (!appState.logDomDirty && removedCount > 0 &&
          appState.logNodes.length) {
          var removedNode = appState.logNodes.shift();
          if (removedNode.parentNode === ui.log) {
            ui.log.removeChild(removedNode);
          }
          removedCount -= 1;
        }
        if (!appState.logDomDirty && appState.logNodes.length) {
          appState.logNodes[0].nodeValue =
            appState.logNodes[0].nodeValue.replace(/^\n/, "");
        }
        if (!appState.logDomDirty) {
          var lineIndex;
          for (lineIndex = 0; lineIndex < lines.length; lineIndex += 1) {
            var node = document.createTextNode(
              (appState.logNodes.length ? "\n" : "") + lines[lineIndex]
            );
            ui.log.appendChild(node);
            appState.logNodes.push(node);
          }
        }
      } catch (error) {
        appState.logDomDirty = true;
      }
      if (appState.logDomDirty ||
        appState.logNodes.length !== totalLines) {
        renderLogLines();
      }
      ui.log.scrollTop = ui.log.scrollHeight;
    }

    function localizeRuntimeWarning(message, language) {
      var selectedLanguage = language === "ja" || language === "en" ?
        language : currentLanguage;
      if (message && typeof message.key === "string") {
        var params = {};
        var name;
        for (name in message.params) {
          if (Object.prototype.hasOwnProperty.call(message.params, name)) {
            params[name] = message.params[name];
          }
        }
        if (message.technicalMessage &&
          !Object.prototype.hasOwnProperty.call(params, "message")) {
          params.message = localizeWarningTechnicalMessage(
            message,
            selectedLanguage
          );
        }
        var messageKey = message.key;
        if (typeof message.pluralCount === "number") {
          messageKey = pluralKey(
            message.key,
            message.pluralCount,
            selectedLanguage
          );
          params.count = formatNumber(message.pluralCount, selectedLanguage);
        }
        return t(messageKey, params, selectedLanguage);
      }
      var raw = String(message || "");
      if (selectedLanguage === "ja") {
        return raw;
      }
      if (raw.indexOf("抽出テキストが安全上限") >= 0) {
        return t("warning.textOmitted", null, selectedLanguage);
      }
      if (raw.indexOf("有効な文書テキスト") >= 0) {
        return t("warning.noText", null, selectedLanguage);
      }
      if (raw.indexOf("埋め込み") >= 0 || raw.indexOf("Ole10Native") >= 0 ||
        raw.indexOf("ExOleObjStg") >= 0 || raw.indexOf("OLE") >= 0) {
        return t("warning.embedded", null, selectedLanguage);
      }
      if (raw.indexOf("メディア") >= 0 || raw.indexOf("OfficeArt") >= 0 ||
        raw.indexOf("BLIP") >= 0 || raw.indexOf("Pictures stream") >= 0) {
        return t("warning.media", null, selectedLanguage);
      }
      if (raw.indexOf("Excel") >= 0 || raw.indexOf("Worksheet") >= 0 ||
        raw.indexOf("LabelSst") >= 0 || raw.indexOf("sheet") >= 0) {
        return t("warning.excel", null, selectedLanguage);
      }
      if (raw.indexOf("PowerPoint") >= 0 || raw.indexOf("Current User") >= 0 ||
        raw.indexOf("CurrentUserAtom") >= 0 || raw.indexOf("Persist") >= 0) {
        return t("warning.powerpoint", null, selectedLanguage);
      }
      if (raw.indexOf("Word") >= 0 || raw.indexOf("脚注") >= 0 ||
        raw.indexOf("文末脚注") >= 0 || raw.indexOf("テキストボックス") >= 0 ||
        raw.indexOf("PAPX") >= 0 || raw.indexOf("FTXBXS") >= 0) {
        return t("warning.word", null, selectedLanguage);
      }
      if (raw.indexOf("テキスト") >= 0) {
        return t("warning.text", null, selectedLanguage);
      }
      return t("warning.generic", null, selectedLanguage);
    }

    function localizeWarningTechnicalMessage(message, language) {
      var raw = String(message.technicalMessage || "");
      if (!raw || language === "ja") {
        return raw;
      }
      if (message.technicalMessageKey) {
        var localized = t(message.technicalMessageKey, null, language);
        return message.technicalCode ?
          message.technicalCode + ": " + localized : localized;
      }
      if (/[\u3040-\u30FF\u3400-\u9FFF]/.test(raw)) {
        return (message.technicalName || "Error") + ": " +
          t("error.unexpected", null, language);
      }
      return raw;
    }

    function friendlyError(error, language) {
      var selectedLanguage = language === "ja" || language === "en" ?
        language : currentLanguage;
      if (error instanceof AppError) {
        if (selectedLanguage === "ja") {
          return error.rawMessage;
        }
        return t(error.messageKey, error.messageParams, selectedLanguage);
      }
      if (error && error.name === "RangeError") {
        return t("error.range", null, selectedLanguage);
      }
      if (error && typeof error.message === "string" && error.message) {
        return t("error.unexpectedWithMessage", {
          message: error.message
        }, selectedLanguage);
      }
      return t("error.unexpected", null, selectedLanguage);
    }

    function statusInfo(
      labelKey,
      cssClass,
      processable,
      detailKey,
      labelParams,
      detailParams
    ) {
      return {
        labelKey: labelKey,
        labelParams: labelParams || {},
        cssClass: cssClass,
        processable: !!processable,
        detailKey: detailKey || "",
        detailParams: detailParams || {}
      };
    }

    function applyStaticTranslations() {
      document.documentElement.setAttribute("lang", currentLanguage);
      var textNodes = document.querySelectorAll("[data-i18n]");
      Array.prototype.forEach.call(textNodes, function (node) {
        node.textContent = t(node.getAttribute("data-i18n"));
      });
      var ariaNodes = document.querySelectorAll("[data-i18n-aria-label]");
      Array.prototype.forEach.call(ariaNodes, function (node) {
        node.setAttribute(
          "aria-label",
          t(node.getAttribute("data-i18n-aria-label"))
        );
      });
      var titleNodes = document.querySelectorAll("[data-i18n-title]");
      Array.prototype.forEach.call(titleNodes, function (node) {
        node.setAttribute("title", t(node.getAttribute("data-i18n-title")));
      });
      var contentNodes = document.querySelectorAll("[data-i18n-content]");
      Array.prototype.forEach.call(contentNodes, function (node) {
        node.setAttribute("content", t(node.getAttribute("data-i18n-content")));
      });
      ui.languageJa.setAttribute(
        "aria-pressed",
        currentLanguage === "ja" ? "true" : "false"
      );
      ui.languageEn.setAttribute(
        "aria-pressed",
        currentLanguage === "en" ? "true" : "false"
      );
      ui.languageJa.classList.toggle("active", currentLanguage === "ja");
      ui.languageEn.classList.toggle("active", currentLanguage === "en");
    }

    function renderProgress() {
      var progress = appState.progress;
      var completed = formatNumber(progress.completed, currentLanguage);
      var total = formatNumber(progress.total, currentLanguage);
      var percentage = formatNumber(Math.round(progress.percentage), currentLanguage);
      ui.progressNumbers.textContent = completed + " / " + total + "　" +
        percentage + "%";
      ui.progressCurrent.textContent = t(progress.key, progress.params);
    }

    function setLanguage(language) {
      var nextLanguage = language === "ja" || language === "en" ? language : "en";
      if (nextLanguage !== language && typeof console !== "undefined" && console.warn) {
        console.warn("Unsupported display language: " + language);
      }
      currentLanguage = nextLanguage;
      applyStaticTranslations();
      renderFileList();
      renderProgress();
      renderLogLines();
    }

    function isZipSignature(bytes) {
      if (bytes.length < 4) {
        return false;
      }
      var signature = u32(bytes, 0);
      return signature === 0x04034B50 ||
        signature === 0x06054B50 ||
        signature === 0x08074B50;
    }

    async function inspectBuffer(bytes, extension) {
      var expected = SUPPORTED_EXTENSIONS[extension];
      if (!expected) {
        return statusInfo(
          "status.unsupported",
          "warning",
          false,
          "status.detail.unsupported"
        );
      }
      ensure(bytes.length > 0, "EMPTY_FILE", "0バイトのファイルです。");
      if (isZipSignature(bytes)) {
        var zip = new ZipArchive(bytes);
        var zipFamily = await detectOoxmlFamily(zip);
        if (expected.container !== "zip" || expected.family !== zipFamily) {
          return statusInfo(
            "status.formatMismatch",
            "error",
            false,
            "status.detail.formatMismatchOoxml"
          );
        }
        return statusInfo(
          "status.ready",
          "ok",
          true,
          "status.detail.ooxml",
          null,
          { family: zipFamily }
        );
      }
      if (startsWithBytes(bytes, CFB_SIGNATURE)) {
        var cfb = new CompoundFile(bytes);
        if (cfb.isEncryptedPackage()) {
          return statusInfo(
            "status.encrypted",
            "warning",
            false,
            "status.detail.encryptedOoxml"
          );
        }
        var cfbFamily = detectCompoundFamily(cfb);
        if (legacyEncryptionStatus(cfb, cfbFamily)) {
          return statusInfo(
            "status.encrypted",
            "warning",
            false,
            "status.detail.encryptedLegacy"
          );
        }
        if (expected.container !== "cfb" || expected.family !== cfbFamily) {
          return statusInfo(
            "status.formatMismatch",
            "error",
            false,
            "status.detail.formatMismatchCfb"
          );
        }
        return statusInfo(
          "status.ready",
          "ok",
          true,
          "status.detail.cfb",
          null,
          { family: cfbFamily }
        );
      }
      return statusInfo(
        "status.formatMismatch",
        "error",
        false,
        "status.detail.missingSignature"
      );
    }

    function itemStillExists(item) {
      return appState.items.some(function (candidate) {
        return candidate.id === item.id;
      });
    }

    function releaseItemCache(item) {
      if (item && item.cachedBytes) {
        appState.cachedInputBytes = Math.max(
          0,
          appState.cachedInputBytes - item.cachedBytes.length
        );
        item.cachedBytes = null;
      }
    }

    function cacheItemBytes(item, bytes) {
      releaseItemCache(item);
      if (!item || !bytes || bytes.length > MAX_CACHED_INPUT_BYTES) {
        return;
      }
      var nextTotal = appState.cachedInputBytes + bytes.length;
      // Inspection succeeds even when the shared cache budget is already exhausted.
      if (!Number.isSafeInteger(nextTotal) ||
        nextTotal > MAX_TOTAL_CACHED_INPUT_BYTES) {
        return;
      }
      item.cachedBytes = bytes;
      appState.cachedInputBytes = nextTotal;
    }

    function displayedItemStatus(item) {
      return item.processingStatus || item.inspectionStatus;
    }

    function queueInspection(item) {
      appState.inspecting += 1;
      updateControls();
      appState.inspectionQueue = appState.inspectionQueue
        .catch(function () {
          // A previous item failure must not stop later inspections.
        })
        .then(async function () {
        try {
          if (!itemStillExists(item)) {
            return;
          }
          var bytes = new Uint8Array(await item.file.arrayBuffer());
          var result = await inspectBuffer(bytes, item.extension);
          if (itemStillExists(item)) {
            item.inspectionStatus = result;
            if (result.processable === true) {
              cacheItemBytes(item, bytes);
            } else {
              releaseItemCache(item);
            }
          }
        } catch (error) {
          if (itemStillExists(item)) {
            releaseItemCache(item);
            var code = error instanceof AppError ? error.code : "";
            if (code === "ENCRYPTED" || code === "ZIP_ENCRYPTED") {
              item.inspectionStatus = statusInfo(
                "status.encrypted",
                "warning",
                false,
                "message.value",
                null,
                { message: error }
              );
            } else {
              item.inspectionStatus = statusInfo(
                "status.error",
                "error",
                false,
                "message.value",
                null,
                { message: error }
              );
            }
          }
        } finally {
          appState.inspecting = Math.max(0, appState.inspecting - 1);
          updateFileRow(item);
          updateControls();
        }
      });
    }

    function relativePathAllowed(path) {
      if (ui.includeSubfolders.checked) {
        return true;
      }
      var parts = String(path || "").replace(/\\/g, "/").split("/").filter(Boolean);
      return parts.length <= 2;
    }

    function addFileEntries(entries) {
      var currentBatchEntries = [];
      var currentBatchFiles = [];
      var totalInputBytes = 0;
      appState.items.forEach(function (item) {
        totalInputBytes += item.file.size;
      });
      var added = 0;
      var skippedSubfolders = 0;
      entries.forEach(function (entry) {
        var file = entry.file;
        if (!file) {
          return;
        }
        if (currentBatchEntries.indexOf(entry) >= 0 ||
          currentBatchFiles.indexOf(file) >= 0) {
          return;
        }
        currentBatchEntries.push(entry);
        currentBatchFiles.push(file);
        var path = String(entry.path || file.webkitRelativePath || file.name)
          .replace(/\\/g, "/");
        if (!relativePathAllowed(path)) {
          skippedSubfolders += 1;
          return;
        }
        if (file.size === 0) {
          logMessage("warning", "log.zeroByteSkipped", { name: file.name });
          return;
        }
        if (file.size > MAX_INPUT_FILE_BYTES) {
          logMessage("warning", "log.fileTooLarge", { name: file.name });
          return;
        }
        if (!Number.isSafeInteger(totalInputBytes + file.size) ||
          totalInputBytes + file.size > MAX_TOTAL_INPUT_BYTES) {
          logMessage("warning", "log.totalTooLarge", { name: file.name });
          return;
        }
        totalInputBytes += file.size;
        var extension = extensionOf(file.name);
        var supported = SUPPORTED_EXTENSIONS[extension];
        var browserUnsupported = supported &&
          supported.container === "zip" &&
          typeof DecompressionStream !== "function";
        var item = {
          id: appState.nextId++,
          file: file,
          path: path,
          extension: extension,
          selected: false,
          cachedBytes: null,
          processingStatus: null,
          inspectionStatus: browserUnsupported ?
            statusInfo(
              "status.browserUnsupported",
              "warning",
              false,
              "status.detail.browserUnsupported"
            ) :
            (supported ?
              statusInfo(
                "status.checking",
                "pending",
                false,
                "status.detail.checking"
              ) :
              statusInfo(
                "status.unsupported",
                "warning",
                false,
                "status.detail.unsupported"
              ))
        };
        appState.items.push(item);
        added += 1;
        if (supported && !browserUnsupported) {
          queueInspection(item);
        }
      });
      renderFileList();
      updateControls();
      if (added) {
        logMessage("info", "log.filesAdded", null, added);
      }
      if (skippedSubfolders) {
        logMessage("info", "log.subfoldersSkipped", null, skippedSubfolders);
      }
      if (!added && !skippedSubfolders) {
        logMessage("info", "log.noFilesAdded");
      }
    }

    function updateFileRowContents(row, item) {
      row._selectCheckbox.checked = item.selected;
      row._selectCheckbox.disabled = appState.busy;
      row._selectCheckbox.setAttribute(
        "aria-label",
        t("table.selectItem", { name: item.file.name })
      );
      row._selectCheckbox.title = t("table.selectItemTitle");
      row._nameCell.textContent = item.file.name;
      row._pathCell.textContent = item.path;
      row._extensionCell.textContent = item.extension ? "." + item.extension : "—";
      row._sizeCell.textContent = formatBytes(item.file.size);
      var itemStatus = displayedItemStatus(item);
      row._status.className = "status " + itemStatus.cssClass;
      row._status.textContent = t(itemStatus.labelKey, itemStatus.labelParams);
      if (itemStatus.detailKey) {
        row._status.title = t(itemStatus.detailKey, itemStatus.detailParams);
      } else {
        row._status.removeAttribute("title");
      }
    }

    function createFileRow(item) {
      var row = document.createElement("tr");
      var selectCell = document.createElement("td");
      var checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.title = t("table.selectItemTitle");
      checkbox.addEventListener("change", function () {
        item.selected = checkbox.checked;
        updateFileRow(item);
        updateSelectionState();
      });
      selectCell.appendChild(checkbox);

      var nameCell = document.createElement("td");
      nameCell.className = "file-name";
      var pathCell = document.createElement("td");
      pathCell.className = "path";
      var extensionCell = document.createElement("td");
      var sizeCell = document.createElement("td");
      var statusCell = document.createElement("td");
      var status = document.createElement("span");
      statusCell.appendChild(status);

      row.appendChild(selectCell);
      row.appendChild(nameCell);
      row.appendChild(pathCell);
      row.appendChild(extensionCell);
      row.appendChild(sizeCell);
      row.appendChild(statusCell);
      row._selectCheckbox = checkbox;
      row._nameCell = nameCell;
      row._pathCell = pathCell;
      row._extensionCell = extensionCell;
      row._sizeCell = sizeCell;
      row._status = status;
      updateFileRowContents(row, item);
      return row;
    }

    function renderFileList() {
      ui.fileList.textContent = "";
      appState.rowByItemId = Object.create(null);
      if (!appState.items.length) {
        var emptyRow = document.createElement("tr");
        emptyRow.className = "empty-row";
        var emptyCell = document.createElement("td");
        emptyCell.colSpan = 6;
        emptyCell.textContent = t("table.empty");
        emptyRow.appendChild(emptyCell);
        ui.fileList.appendChild(emptyRow);
        ui.selectAll.checked = false;
        ui.selectAll.indeterminate = false;
        return;
      }
      appState.items.forEach(function (item) {
        var row = createFileRow(item);
        appState.rowByItemId[item.id] = row;
        ui.fileList.appendChild(row);
      });
      updateSelectionState();
    }

    function updateFileRow(item) {
      if (!item) {
        return false;
      }
      var row = appState.rowByItemId[item.id];
      if (!row || row.parentNode !== ui.fileList) {
        // Recover from an unexpected DOM mismatch without stopping inspection or export.
        if (itemStillExists(item)) {
          renderFileList();
        }
        return false;
      }
      updateFileRowContents(row, item);
      return true;
    }

    function updateAllFileRows() {
      var index;
      for (index = 0; index < appState.items.length; index += 1) {
        var item = appState.items[index];
        var row = appState.rowByItemId[item.id];
        if (!row || row.parentNode !== ui.fileList) {
          renderFileList();
          return;
        }
        updateFileRowContents(row, item);
      }
      updateSelectionState();
    }

    function updateSelectionState() {
      var selectedCount = appState.items.filter(function (item) {
        return item.selected;
      }).length;
      ui.selectAll.checked = appState.items.length > 0 &&
        selectedCount === appState.items.length;
      ui.selectAll.indeterminate = selectedCount > 0 &&
        selectedCount < appState.items.length;
      ui.removeSelected.disabled = appState.busy || selectedCount === 0;
    }

    function updateControls() {
      var hasItems = appState.items.length > 0;
      var hasProcessable = appState.items.some(function (item) {
        return item.inspectionStatus.processable;
      });
      ui.addFiles.disabled = appState.busy;
      ui.addFolder.disabled = appState.busy;
      ui.fileInput.disabled = appState.busy;
      ui.folderInput.disabled = appState.busy;
      ui.clearFiles.disabled = appState.busy || !hasItems;
      ui.exportZip.disabled = appState.busy || appState.inspecting > 0 || !hasProcessable;
      ui.includeSubfolders.disabled = appState.busy;
      ui.includeText.disabled = appState.busy;
      ui.includeEmbeddings.disabled = appState.busy;
      ui.selectAll.disabled = appState.busy || !hasItems;
      ui.clearLog.disabled = appState.busy;
      ui.cancelProcessing.disabled = !appState.busy || appState.cancelRequested;
      updateSelectionState();
    }

    function setProgress(completed, total, percentage, key, params) {
      var safeTotal = Math.max(0, total);
      var safeCompleted = Math.max(0, Math.min(completed, safeTotal));
      var next = Math.max(appState.lastProgress, Math.min(100, Math.max(0, percentage)));
      appState.lastProgress = next;
      appState.progress = {
        completed: safeCompleted,
        total: safeTotal,
        percentage: next,
        key: key || "progress.waiting",
        params: params || {}
      };
      ui.progressBar.style.width = next.toFixed(2) + "%";
      ui.progressTrack.setAttribute("aria-valuenow", String(Math.round(next)));
      renderProgress();
    }

    function resetProgress() {
      appState.lastProgress = 0;
      setProgress(0, 0, 0, "progress.waiting");
    }

    function summarizeResult(fileName, mediaCount, characterCount, embeddingCount) {
      return translationValue("summary.result", {
        name: fileName,
        media: pluralValue("summary.media", mediaCount),
        characters: pluralValue("summary.characters", characterCount),
        embeddings: pluralValue("summary.embeddings", embeddingCount)
      });
    }

    async function processOfficeBuffer(item, bytes, options, currentOutputBytes, stage) {
      var expected = SUPPORTED_EXTENSIONS[item.extension];
      ensure(expected, "UNSUPPORTED", "未対応形式です。");
      checkCancelled();
      ensure(bytes.length > 0 && bytes.length <= MAX_INPUT_FILE_BYTES, "INPUT_SIZE",
        "入力ファイルのサイズが安全上限外です。");
      var warnings = [];
      var media = [];
      var embeddings = [];
      var textResult = null;

      stage("progress.checkingInternalFormat", 0.12);
      var archive = null;
      var cfb = null;
      var family;
      if (isZipSignature(bytes)) {
        archive = new ZipArchive(bytes);
        var outputBytes = typeof currentOutputBytes === "number" ?
          currentOutputBytes :
          0;
        var estimatedWorkingBytes = bytes.length +
          archive.totalUncompressedSize +
          outputBytes;
        ensure(
          Number.isSafeInteger(outputBytes) &&
          outputBytes >= 0 &&
          Number.isSafeInteger(estimatedWorkingBytes) &&
          estimatedWorkingBytes <= MAX_ESTIMATED_WORKING_BYTES,
          "ESTIMATED_MEMORY",
          "展開後の推定メモリ使用量が安全上限を超えるため処理できません。\n" +
            "ファイル数またはファイルサイズを減らして再実行してください。"
        );
        family = await detectOoxmlFamily(archive);
        ensure(expected.container === "zip" && expected.family === family,
          "FORMAT_MISMATCH", "拡張子と内部形式が一致しません。");
      } else if (startsWithBytes(bytes, CFB_SIGNATURE)) {
        cfb = new CompoundFile(bytes);
        if (cfb.isEncryptedPackage()) {
          fail("ENCRYPTED", "暗号化またはパスワード保護されたファイルです。");
        }
        family = detectCompoundFamily(cfb);
        ensure(expected.container === "cfb" && expected.family === family,
          "FORMAT_MISMATCH", "拡張子と内部形式が一致しません。");
        ensure(!legacyEncryptionStatus(cfb, family), "ENCRYPTED",
          "暗号化またはパスワード保護されたファイルです。");
      } else {
        fail("FORMAT_SIGNATURE", "Officeファイルの内部シグネチャが一致しません。");
      }

      checkCancelled();
      stage("progress.extractingMedia", 0.3);
      try {
        if (archive) {
          var mediaResult = await extractZipPrefix(
            archive,
            ooxmlPrefix(family, "media"),
            "media"
          );
          media = mediaResult.files;
          Array.prototype.push.apply(warnings, mediaResult.warnings);
        } else {
          media = await extractLegacyMedia(cfb, family, warnings);
        }
      } catch (error) {
        throwIfCategoryLimitOrCancelled(error);
        warnings.push(warningValue(
          "warning.categoryExtractionFailed",
          { subject: translationValue("warning.subject.media") },
          error
        ));
        media = [];
      }

      if (options.text) {
        checkCancelled();
        stage("progress.extractingText", 0.52);
        try {
          var extractedText = archive ?
            await extractOoxmlText(archive, family, warnings) :
            await extractLegacyText(cfb, family, warnings);
          textResult = textToOutputBytes(extractedText);
          if (textResult && textResult.omitted) {
            warnings.push(textResult.warning);
            textResult = null;
          } else if (!textResult) {
            warnings.push(warningValue("warning.noText"));
          }
        } catch (error) {
          throwIfCategoryLimitOrCancelled(error);
          warnings.push(warningValue(
            "warning.categoryExtractionFailed",
            { subject: translationValue("warning.subject.text") },
            error
          ));
          textResult = null;
        }
      }

      if (options.embeddings) {
        checkCancelled();
        stage("progress.extractingEmbeddings", 0.72);
        try {
          if (archive) {
            var embeddingResult = await extractZipPrefix(
              archive,
              ooxmlPrefix(family, "embeddings"),
              "embeddings"
            );
            embeddings = embeddingResult.files;
            Array.prototype.push.apply(warnings, embeddingResult.warnings);
          } else {
            embeddings = await extractLegacyEmbeddings(cfb, family, warnings);
          }
        } catch (error) {
          throwIfCategoryLimitOrCancelled(error);
          warnings.push(warningValue(
            "warning.categoryExtractionFailed",
            { subject: translationValue("warning.subject.embedding") },
          error
          ));
          embeddings = [];
        }
      }

      checkCancelled();
      stage("progress.addingToZip", 0.9);
      return {
        media: media,
        text: textResult,
        embeddings: embeddings,
        warnings: warnings
      };
    }

    async function exportAll() {
      if (appState.busy || appState.inspecting > 0) {
        return;
      }
      var targets = appState.items.filter(function (item) {
        return item.inspectionStatus.processable;
      });
      if (!targets.length) {
        logMessage("warning", "log.noProcessableFiles");
        return;
      }

      appState.busy = true;
      appState.cancelRequested = false;
      appState.lastProgress = 0;
      var options = {
        text: ui.includeText.checked,
        embeddings: ui.includeEmbeddings.checked
      };
      var zipBuilder = new ZipBuilder();
      var rootAllocator = new OutputRootAllocator();
      var parentPathMap = null;
      var extractedCount = 0;
      var emptyCount = 0;
      var failedCount = 0;
      targets.forEach(function (item) {
        item.processingStatus = null;
      });
      updateControls();
      updateAllFileRows();
      try {
        parentPathMap = new ParentPathMap(targets);
        var targetIndex;
        for (targetIndex = 0; targetIndex < targets.length; targetIndex += 1) {
          checkCancelled();
          var item = targets[targetIndex];
          var bytes = null;
          var rootName = null;
          var baseProgress = targetIndex / targets.length * 96;
          var span = 96 / targets.length;
          item.processingStatus = statusInfo(
            "status.processing",
            "pending",
            false,
            "status.detail.processing"
          );
          updateFileRow(item);
          setProgress(
            targetIndex,
            targets.length,
            baseProgress,
            "progress.readingFile",
            { name: item.file.name }
          );
          try {
            bytes = item.cachedBytes ||
              new Uint8Array(await item.file.arrayBuffer());
            checkCancelled();
            var result = await processOfficeBuffer(
              item,
              bytes,
              options,
              zipBuilder.totalBytes,
              function (label, fraction) {
                setProgress(
                  targetIndex,
                  targets.length,
                  baseProgress + span * fraction,
                  "progress.fileStage",
                  {
                    name: item.file.name,
                    stage: translationValue(label)
                  }
                );
              }
            );
            var extractedFiles = result.media.length +
              (result.text ? 1 : 0) +
              result.embeddings.length;
            var warningIndex;
            for (warningIndex = 0;
              warningIndex < result.warnings.length;
              warningIndex += 1) {
              logMessage(
                "warning",
                "log.fileWarning",
                {
                  name: item.file.name,
                  warning: runtimeWarningValue(result.warnings[warningIndex])
                }
              );
              await cooperativeYield(warningIndex + 1, 20);
            }
            if (extractedFiles === 0) {
              emptyCount += 1;
              item.processingStatus = statusInfo(
                "status.noResults",
                "warning",
                false,
                "status.detail.noResults"
              );
              logMessage("warning", "log.noExtractableContent", {
                name: item.file.name
              });
            } else {
              var sourceStem = buildSourceOutputStem(item.file.name);
              rootName = buildOutputRootPath(
                item,
                rootAllocator,
                sourceStem,
                parentPathMap
              );
              var mediaNameLimit = outputRelativeFileNameLimit(
                rootName,
                "media"
              );
              var mediaAllocator = new NameAllocator(mediaNameLimit);
              var mediaIndex;
              for (mediaIndex = 0; mediaIndex < result.media.length; mediaIndex += 1) {
                checkCancelled();
                var mediaFile = result.media[mediaIndex];
                var mediaName = allocatePrefixedMediaFileName(
                  mediaAllocator,
                  sourceStem,
                  mediaFile.name,
                  mediaNameLimit
                );
                zipBuilder.add(
                  buildOutputRelativePath(rootName, "media", mediaName),
                  mediaFile.data
                );
                await cooperativeYield(mediaIndex + 1, 25);
              }
              if (result.text) {
                var textNameLimit = outputRelativeFileNameLimit(rootName, "");
                var textName = buildTextOutputFileName(
                  sourceStem,
                  textNameLimit
                );
                zipBuilder.add(
                  buildOutputRelativePath(rootName, "", textName),
                  result.text.bytes
                );
              }
              var embeddingNameLimit = outputRelativeFileNameLimit(
                rootName,
                "embeddings"
              );
              var embeddingAllocator = new NameAllocator(embeddingNameLimit);
              var embeddingIndex;
              for (embeddingIndex = 0;
                embeddingIndex < result.embeddings.length;
                embeddingIndex += 1) {
                checkCancelled();
                var embeddingFile = result.embeddings[embeddingIndex];
                var embeddingName = embeddingAllocator.allocate(
                  embeddingFile.name,
                  true,
                  embeddingNameLimit
                );
                zipBuilder.add(
                  buildOutputRelativePath(
                    rootName,
                    "embeddings",
                    embeddingName
                  ),
                  embeddingFile.data
                );
                await cooperativeYield(embeddingIndex + 1, 25);
              }
              extractedCount += 1;
              item.processingStatus = statusInfo(
                result.warnings.length ?
                  "status.completedWithWarnings" : "status.completed",
                result.warnings.length ? "warning" : "ok",
                false,
                result.warnings.length ?
                  "status.detail.completedWithWarnings" :
                  "status.detail.completed"
              );
              logMessage(
                result.warnings.length ? "warning" : "ok",
                "message.value",
                {
                  message: summarizeResult(
                    item.file.name,
                    result.media.length,
                    result.text ? result.text.characterCount : 0,
                    result.embeddings.length
                  )
                }
              );
            }
          } catch (error) {
            if (error instanceof AppError && error.code === "CANCELLED") {
              item.processingStatus = statusInfo(
                "status.cancelled",
                "warning",
                false,
                "status.detail.cancelled"
              );
              throw error;
            }
            if (isProcessingSafetyError(error) ||
              (error instanceof AppError &&
                (error.code === "OUTPUT_SIZE" ||
                 error.code === "OUTPUT_ZIP64" ||
                 error.code === "OUTPUT_PATH" ||
                 error.code === "OUTPUT_DUPLICATE" ||
                 error.code === "NAME_COLLISION" ||
                 error.code === "MEMORY"))) {
              item.processingStatus = statusInfo(
                "status.processingFailed",
                "error",
                false,
                "message.value",
                null,
                { message: error }
              );
              throw error;
            }
            failedCount += 1;
            item.processingStatus = statusInfo(
              "status.processingFailed",
              "error",
              false,
              "message.value",
              null,
              { message: error }
            );
            logMessage("error", "log.fileError", {
              message: error,
              name: item.file.name
            });
          } finally {
            releaseItemCache(item);
            bytes = null;
            updateFileRow(item);
          }
          setProgress(
            targetIndex + 1,
            targets.length,
            (targetIndex + 1) / targets.length * 96,
            "progress.fileFinished",
            { name: item.file.name }
          );
          await delayTurn();
        }

        checkCancelled();
        if (!zipBuilder.entries.length) {
          logMessage("warning", "log.noZipResults");
          setProgress(targets.length, targets.length, 100, "progress.completed");
          return;
        }
        setProgress(targets.length, targets.length, 97, "progress.generatingZip");
        await delayTurn();
        checkCancelled();
        var blob = zipBuilder.toBlob();
        await delayTurn();
        checkCancelled();
        var outputName = "document_content_" +
          timestampForFileName(new Date()) + ".zip";
        downloadBlob(blob, outputName);
        setProgress(targets.length, targets.length, 100, "progress.completed");
        logMessage("ok", "log.zipGenerated", {
          name: outputName,
          success: numberValue(extractedCount),
          empty: numberValue(emptyCount),
          failed: numberValue(failedCount)
        });
      } catch (error) {
        if (error instanceof AppError && error.code === "CANCELLED") {
          logMessage("info", "log.cancelled");
          setProgress(
            0,
            targets.length,
            appState.lastProgress,
            "progress.cancelled"
          );
        } else {
          logMessage("error", "message.value", { message: error });
          setProgress(
            0,
            targets.length,
            appState.lastProgress,
            "progress.error"
          );
        }
      } finally {
        targets.forEach(releaseItemCache);
        appState.cancelRequested = false;
        appState.busy = false;
        updateControls();
        updateAllFileRows();
      }
    }

    function fileEntryPromise(entry) {
      return new Promise(function (resolve, reject) {
        entry.file(resolve, reject);
      });
    }

    function readDirectoryBatch(reader) {
      return new Promise(function (resolve, reject) {
        reader.readEntries(resolve, reject);
      });
    }

    async function readAllDirectoryEntries(directoryEntry) {
      var reader = directoryEntry.createReader();
      var all = [];
      while (true) {
        var batch = await readDirectoryBatch(reader);
        if (!batch.length) {
          break;
        }
        all = all.concat(batch);
        ensure(all.length <= MAX_ZIP_ENTRIES, "FOLDER_ENTRIES",
          "フォルダ内の項目数が安全上限を超えています。");
      }
      return all;
    }

    async function walkDroppedEntry(entry, prefix, includeSubfolders, depth) {
      ensure(depth <= 128, "FOLDER_DEPTH", "フォルダ階層が深すぎます。");
      if (entry.isFile) {
        var file = await fileEntryPromise(entry);
        return [{ file: file, path: prefix + file.name }];
      }
      if (!entry.isDirectory) {
        return [];
      }
      var children = await readAllDirectoryEntries(entry);
      var output = [];
      var childIndex;
      for (childIndex = 0; childIndex < children.length; childIndex += 1) {
        var child = children[childIndex];
        if (child.isDirectory && !includeSubfolders) {
          continue;
        }
        var childPrefix = prefix + entry.name + "/";
        var nested = await walkDroppedEntry(
          child,
          childPrefix,
          includeSubfolders,
          depth + 1
        );
        output = output.concat(nested);
      }
      return output;
    }

    async function walkDroppedHandle(handle, prefix, includeSubfolders, depth, state) {
      ensure(depth <= 128, "FOLDER_DEPTH", "フォルダ階層が深すぎます。");
      if (handle.kind === "file") {
        var file = await handle.getFile();
        return [{ file: file, path: prefix + file.name }];
      }
      if (handle.kind !== "directory") {
        return [];
      }
      ensure(typeof handle.values === "function", "FOLDER_HANDLE",
        "フォルダハンドルの列挙APIを利用できません。");
      var iterator = handle.values();
      ensure(iterator && typeof iterator.next === "function", "FOLDER_HANDLE",
        "フォルダハンドルを列挙できません。");
      var output = [];
      var childPrefix = prefix + handle.name + "/";
      while (true) {
        var next = await iterator.next();
        if (next.done) {
          break;
        }
        var child = next.value;
        state.count += 1;
        ensure(state.count <= MAX_ZIP_ENTRIES, "FOLDER_ENTRIES",
          "フォルダ内の項目数が安全上限を超えています。");
        if (!child || (child.kind !== "file" && child.kind !== "directory")) {
          continue;
        }
        if (child.kind === "directory" && !includeSubfolders) {
          continue;
        }
        var nested = await walkDroppedHandle(
          child,
          childPrefix,
          includeSubfolders,
          depth + 1,
          state
        );
        output = output.concat(nested);
      }
      return output;
    }

    function snapshotDroppedItem(item) {
      var file = null;
      var entry = null;
      var entryError = null;
      var handlePromise = null;
      if (!item || (item.kind && item.kind !== "file")) {
        return {
          file: null,
          entry: null,
          entryError: null,
          handlePromise: null
        };
      }
      if (typeof item.getAsFileSystemHandle === "function") {
        try {
          handlePromise = Promise.resolve(item.getAsFileSystemHandle()).then(
            function (handle) {
              return { handle: handle, error: null };
            },
            function (error) {
              return { handle: null, error: error };
            }
          );
        } catch (handleError) {
          handlePromise = Promise.resolve({
            handle: null,
            error: handleError
          });
        }
      }
      if (typeof item.getAsFile === "function") {
        try {
          file = item.getAsFile();
        } catch (ignoreFileError) {
          file = null;
        }
      }
      var entryGetter = null;
      if (typeof item.getAsEntry === "function") {
        entryGetter = item.getAsEntry;
      } else if (typeof item.webkitGetAsEntry === "function") {
        entryGetter = item.webkitGetAsEntry;
      }
      if (entryGetter) {
        try {
          entry = entryGetter.call(item);
        } catch (error) {
          entryError = error;
        }
      }
      return {
        file: file,
        entry: entry,
        entryError: entryError,
        handlePromise: handlePromise
      };
    }

    async function handleDrop(dataTransfer) {
      var entries = [];
      var items = Array.prototype.slice.call(dataTransfer.items || []);
      var fallbackFiles = Array.prototype.slice.call(dataTransfer.files || []);
      var snapshots = items.map(function (item) {
        return snapshotDroppedItem(item);
      });
      var failedDirectories = 0;
      var failedItems = 0;
      var handleWalkState = { count: 0 };
      var index;
      for (index = 0; index < snapshots.length; index += 1) {
        var snapshot = snapshots[index];
        var handleResult = snapshot.handlePromise ?
          await snapshot.handlePromise :
          { handle: null, error: null };
        var handle = handleResult.handle;
        var modernDirectory = handle && handle.kind === "directory";
        var legacyDirectory = snapshot.entry && snapshot.entry.isDirectory;
        if (modernDirectory || legacyDirectory) {
          var directoryRead = false;
          if (modernDirectory) {
            try {
              var handleFiles = await walkDroppedHandle(
                handle,
                "",
                ui.includeSubfolders.checked,
                0,
                handleWalkState
              );
              entries = entries.concat(handleFiles);
              directoryRead = true;
            } catch (handleDirectoryError) {
              directoryRead = false;
            }
          }
          if (!directoryRead && legacyDirectory) {
            try {
              var entryFiles = await walkDroppedEntry(
                snapshot.entry,
                "",
                ui.includeSubfolders.checked,
                0
              );
              entries = entries.concat(entryFiles);
              directoryRead = true;
            } catch (entryDirectoryError) {
              directoryRead = false;
            }
          }
          if (!directoryRead) {
            failedDirectories += 1;
          }
        } else if (snapshot.file) {
          entries.push({ file: snapshot.file, path: snapshot.file.name });
        } else if (handle && handle.kind === "file") {
          try {
            var handleFile = await handle.getFile();
            entries.push({ file: handleFile, path: handleFile.name });
          } catch (handleFileError) {
            failedItems += 1;
          }
        } else if (snapshot.entry && snapshot.entry.isFile) {
          try {
            var entryFile = await fileEntryPromise(snapshot.entry);
            entries.push({ file: entryFile, path: entryFile.name });
          } catch (fileEntryError) {
            failedItems += 1;
          }
        } else if (snapshot.entryError || handleResult.error) {
          failedItems += 1;
        }
      }
      entries = entries.concat(fallbackFiles.map(function (file) {
        if (!file) {
          return null;
        }
        return { file: file, path: file.name };
      }).filter(function (entry) {
        return entry !== null;
      }));
      addFileEntries(entries);
      if (failedDirectories) {
        logMessage(
          "warning",
          "log.folderReadFailed",
          null,
          failedDirectories
        );
      }
      if (failedItems && !fallbackFiles.length) {
        logMessage("warning", "log.dropItemsFailed", null, failedItems);
      }
    }

    function openFilePicker() {
      if (ui.fileInput.disabled) {
        return;
      }
      ui.fileInput.click();
    }

    ui.languageJa.addEventListener("click", function () {
      setLanguage("ja");
    });

    ui.languageEn.addEventListener("click", function () {
      setLanguage("en");
    });

    ui.addFiles.addEventListener("click", openFilePicker);

    ui.dropZone.addEventListener("click", openFilePicker);

    ui.addFolder.addEventListener("click", function () {
      ui.folderInput.click();
    });

    ui.fileInput.addEventListener("change", function () {
      addFileEntries(Array.prototype.slice.call(ui.fileInput.files).map(function (file) {
        return { file: file, path: file.name };
      }));
      ui.fileInput.value = "";
    });

    ui.folderInput.addEventListener("change", function () {
      addFileEntries(Array.prototype.slice.call(ui.folderInput.files).map(function (file) {
        return { file: file, path: file.webkitRelativePath || file.name };
      }));
      ui.folderInput.value = "";
    });

    ui.removeSelected.addEventListener("click", function () {
      var removedItems = appState.items.filter(function (item) {
        return item.selected;
      });
      removedItems.forEach(releaseItemCache);
      appState.items = appState.items.filter(function (item) {
        return !item.selected;
      });
      renderFileList();
      updateControls();
      logMessage("info", "log.removed", null, removedItems.length);
    });

    ui.clearFiles.addEventListener("click", function () {
      var count = appState.items.length;
      appState.items.forEach(releaseItemCache);
      appState.items = [];
      renderFileList();
      updateControls();
      resetProgress();
      logMessage("info", "log.listCleared", null, count);
    });

    ui.selectAll.addEventListener("change", function () {
      appState.items.forEach(function (item) {
        item.selected = ui.selectAll.checked;
        updateFileRow(item);
      });
      updateControls();
    });

    ui.exportZip.addEventListener("click", exportAll);

    ui.cancelProcessing.addEventListener("click", function () {
      if (!appState.busy || appState.cancelRequested) {
        return;
      }
      appState.cancelRequested = true;
      ui.cancelProcessing.disabled = true;
      logMessage("info", "log.cancelRequested");
    });

    ui.clearLog.addEventListener("click", function () {
      appState.logEntries = [{
        level: "info",
        key: "log.cleared",
        params: {},
        pluralCount: null,
        renderedLineCount: 1
      }];
      appState.logNodes = [];
      appState.logDomDirty = false;
      appState.logStarted = true;
      renderLogLines();
      ui.log.scrollTop = ui.log.scrollHeight;
    });

    document.addEventListener("dragenter", function (event) {
      if (!dataTransferHasFiles(event.dataTransfer)) {
        return;
      }
      event.preventDefault();
      appState.pageDragDepth += 1;
      if (appState.busy) {
        hidePageDropOverlay();
      } else {
        showPageDropOverlay();
      }
    }, true);

    document.addEventListener("dragover", function (event) {
      if (!dataTransferHasFiles(event.dataTransfer)) {
        return;
      }
      event.preventDefault();
      try {
        event.dataTransfer.dropEffect = appState.busy ? "none" : "copy";
      } catch (dropEffectError) {
        // Some browser drag sources expose dropEffect as read-only.
      }
      if (appState.busy) {
        hidePageDropOverlay();
      } else {
        showPageDropOverlay();
      }
    }, true);

    document.addEventListener("dragleave", function () {
      if (appState.pageDragDepth <= 0 &&
        !appState.pageDropOverlayVisible) {
        return;
      }
      appState.pageDragDepth = Math.max(0, appState.pageDragDepth - 1);
      if (appState.pageDragDepth === 0) {
        resetPageDropState();
      }
    }, true);

    document.addEventListener("drop", function (event) {
      var dataTransfer = event.dataTransfer;
      if (!dataTransferHasFiles(dataTransfer) &&
        appState.pageDragDepth <= 0) {
        return;
      }
      event.preventDefault();
      event.stopPropagation();
      resetPageDropState();
      if (appState.busy) {
        return;
      }
      handleDrop(dataTransfer).catch(function (error) {
        resetPageDropState();
        logMessage("error", "log.dropReadFailed", { message: error });
      });
    }, true);

    document.addEventListener("dragend", resetPageDropState, true);
    window.addEventListener("blur", resetPageDropState);

    ui.dropZone.addEventListener("keydown", function (event) {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openFilePicker();
      }
    });

    applyStaticTranslations();
    renderFileList();
    updateControls();
    resetProgress();
    renderLogLines();
    if (typeof DecompressionStream !== "function") {
      ui.compatibilityWarning.hidden = false;
      logMessage("warning", "log.browserUnsupported");
    }
    document.documentElement.setAttribute("data-document-content-extractor-ready", "true");

  }());
