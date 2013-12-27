/*===================================================================================================
  File Name: 見開きメーカー.jsx
  Title: 見開きメーカー
  Version: 1.0.0
  Author: show555
  Description: 複数ページのPDFから指定したページの見開き画像を生成する
  Includes: Underscore.js,
            Underscore.string.js
===================================================================================================*/

#target photoshop

// Photoshopの設定単位を保存
var originalRulerUnits = app.preferences.rulerUnits;
// Photoshopの設定単位をピクセルに変更
app.preferences.rulerUnits = Units.PIXELS;

// 初期設定
var settings = {
	filePath: '',     // 対象フォルダのパスの初期値
	colorMode: 'RGB', // カラーモードの初期値
	_quality: {
		jpgWeb: {
			init: 90,     // 保存画質（Web用JPG）の初期値
			min:  0,
			max:  100
		},
		jpgDtp: {
			init: 10,     // 保存画質（DTP用JPG）の初期値
			min:  0,
			max:  12
		},
	},
	pageWidth: 224,   // 1ページの幅の初期値
	offsetPage: 1,    // PDFページ番号のオフセット
	save: {
		init: 'JPG',    // 保存形式の初期値
		type: {
			jpgWeb: { label: 'JPG（WEB用）', extension: 'jpg' },
			jpgDtp: { label: 'JPG', extension: 'jpg' },
			eps:    { label: 'EPS', extension: 'eps' },
			png:    { label: 'PNG', extension: 'png' }
		},
		dir: 'spread_images', // 保存先のディレクトリ名
		fileName: 'spread_'   // 書き出すファイル名
	},
	fileTypes: [],
	saveType: '',
	quality: '',
	folderObj: {}
};

// ----------------------▼ Underscore.js & Underscore.string.js ▼----------------------
#include "underscore.inc"
#include "underscore.string.inc"
_.mixin(_.string.exports());
// ----------------------▲ Underscore.js & Underscore.string.js ▲----------------------

// 保存関数
var saveFunctions = {
	jpgWeb: function( theDoc, newFile, settings ) {
		var jpegOpt       = new ExportOptionsSaveForWeb();
		jpegOpt.format    = SaveDocumentType.JPEG;
		jpegOpt.optimized = true;
		jpegOpt.quality   = settings.quality;
		theDoc.exportDocument( newFile, ExportType.SAVEFORWEB, jpegOpt );
	},
	jpgDtp: function( theDoc, newFile, settings ) {
		var jpegOpt     = new JPEGSaveOptions();
		jpegOpt.quality = settings.quality;
		theDoc.saveAs( newFile, jpegOpt, true, Extension.LOWERCASE );
	},
	eps: function( theDoc, newFile, settings ) {
		var epsOpt               = new EPSSaveOptions();
		epsOpt.embedColorProfile = true;
		epsOpt.encoding          = SaveEncoding.JPEGMAXIMUM;
		epsOpt.halftoneScreen    = false;
		epsOpt.interpolation     = false;
		epsOpt.preview           = Preview.MACOSEIGHTBIT;
		epsOpt.psColorManagement = false;
		epsOpt.transferFunction  = false;
		epsOpt.transparentWhites = false;
		epsOpt.vectorData        = false;
		theDoc.saveAs( newFile, epsOpt, true, Extension.LOWERCASE );
	},
	png: function( theDoc, newFile, settings ) {
		var pngOpt        = new PNGSaveOptions();
		pngOpt.interlaced = false;
		theDoc.saveAs( newFile, pngOpt, true, Extension.LOWERCASE );
	}
};

// 実行フラグ
var do_flag = true;

// ---------------------------------- ダイアログ作成 ----------------------------------
// ダイアログオブジェクト
var uDlg = new Window( 'dialog', '見開きメーカー', { x:0, y:0, width:400, height:410 } );

// ダイアログを画面に対して中央揃えに
uDlg.center();

// パネル 対象PDFファイル
uDlg.filePnl           = uDlg.add( "panel",    { x:10,  y:10, width:380, height:60 }, "対象PDFファイル" );
uDlg.filePnl.path      = uDlg.add( "edittext", { x:25,  y:30, width:270, height:25 }, settings.filePath );
uDlg.filePnl.selectBtn = uDlg.add( "button",   { x:300, y:30, width:75,  height:25 }, "選択" );
// 対象フォルダ選択ボタンが押された時の処理
uDlg.filePnl.selectBtn.onClick = function() {
	var oldPath = uDlg.filePnl.path.text;
	uDlg.filePnl.path.text = File.openDialog( 'PDFファイルを選択してください' ) || oldPath;
}

// パネル ページ幅
uDlg.pageWidthPnl                     = uDlg.add( "panel",      { x:10,  y:80,  width:380, height:60 }, "ページ幅" );
uDlg.pageWidthPnl.pageWidthText       = uDlg.add( "statictext", { x:25,  y:105, width:80,  height:20 }, "1ページの幅:" );
uDlg.pageWidthPnl.pageWidth           = uDlg.add( "edittext",   { x:115, y:105, width:50,  height:22 }, settings.pageWidth );
uDlg.pageWidthPnl.pageWidthUnit       = uDlg.add( "statictext", { x:170, y:105, width:20,  height:20 }, "px" );
uDlg.pageWidthPnl.spreadPageWidthText = uDlg.add( "statictext", { x:200, y:105, width:80,  height:20 }, "見開きの幅:" );
uDlg.pageWidthPnl.spreadPageWidth     = uDlg.add( "statictext", { x:280, y:105, width:50,  height:20 }, settings.pageWidth*2 );
uDlg.pageWidthPnl.spreadPageWidthUnit = uDlg.add( "statictext", { x:335, y:105, width:20,  height:20 }, "px" );
uDlg.pageWidthPnl.pageWidthText.justify = uDlg.pageWidthPnl.spreadPageWidth.justify = uDlg.pageWidthPnl.spreadPageWidthText.justify = 'right';

// 画質を入力した時の処理
uDlg.pageWidthPnl.pageWidth.onChange = function() {
	var pageWidth = parseInt( uDlg.pageWidthPnl.pageWidth.text )*2;
	uDlg.pageWidthPnl.spreadPageWidth.text = _.isNaN( pageWidth ) ? '' : pageWidth;
}

// パネル 見開きページ指定
uDlg.spreadsPnl             = uDlg.add( "panel",      { x:10, y:150,  width:380, height:80 }, "ページ指定" );
uDlg.spreadsPnl.spreadsText = uDlg.add( "statictext", { x:25, y:170, width:280,  height:20 }, "指定ページをカンマ区切りで入力してください" );
uDlg.spreadsPnl.spreads     = uDlg.add( "edittext",   { x:25, y:195, width:350,  height:20 }, "001-002,003-004" );

// パネル 書き出し設定
var saveTypeList = _.pluck( settings.save.type, 'label' );
uDlg.resizePnl               = uDlg.add( "panel",        { x:10,  y:240, width:380, height:120 }, "書き出し設定" );
uDlg.resizePnl.colorModeText = uDlg.add( "statictext",   { x:25,  y:265, width:60,  height:20  }, "モード:" );
uDlg.resizePnl.saveTypeText  = uDlg.add( "statictext",   { x:25,  y:295, width:60,  height:20  }, "保存形式:" );
uDlg.resizePnl.saveType      = uDlg.add( "dropdownlist", { x:95,  y:293, width:110, height:22  }, saveTypeList );
uDlg.resizePnl.qualityText   = uDlg.add( "statictext",   { x:25,  y:325, width:60,  height:20  }, "画質:" );
uDlg.resizePnl.quality       = uDlg.add( "edittext",     { x:95,  y:323, width:50,  height:22  }, settings._quality.jpgDtp.init );
uDlg.resizePnl.qualityRange  = uDlg.add( "statictext",   { x:150, y:325, width:60,  height:20  }, "(0〜" + settings._quality.jpgDtp.max + ")" );
uDlg.resizePnl.qualitySlider = uDlg.add( "slider",       { x:205, y:320, width:170, height:20  }, settings._quality.jpgDtp.init, settings._quality.jpgDtp.min, settings._quality.jpgDtp.max );
uDlg.resizePnl.colorModeText.justify = uDlg.resizePnl.saveTypeText.justify = uDlg.resizePnl.qualityText.justify = 'right';
// カラーモード選択ラジオボタンの追加
uDlg.resizePnl.colorMode      = uDlg.add( "group", { x:95, y:265, width:245, height:20 } );
uDlg.resizePnl.colorMode.RGB  = uDlg.resizePnl.colorMode.add( "radiobutton",  { x:0,  y:0, width:50, height:20 }, "RGB" );
uDlg.resizePnl.colorMode.CMYK = uDlg.resizePnl.colorMode.add( "radiobutton",  { x:55, y:0, width:70, height:20 }, "CMYK" );
// カラーモードの初期値を設定
uDlg.resizePnl.colorMode[settings.colorMode].value = true;

// 保存形式の初期値を設定
uDlg.resizePnl.saveType.selection = _.indexOf( saveTypeList, settings.save.init );
setSaveTypeQuality( settings.save.init );
// 保存形式が変更された時の処理
uDlg.resizePnl.saveType.onChange = function() {
	setSaveTypeQuality( uDlg.resizePnl.saveType.selection.text );
}
// 画質のスライダーを動かしている時の処理
uDlg.resizePnl.qualitySlider.onChanging = function() {
	uDlg.resizePnl.quality.text = parseInt( uDlg.resizePnl.qualitySlider.value );
}
// 画質を入力した時の処理
uDlg.resizePnl.quality.onChange = function() {
	uDlg.resizePnl.qualitySlider.value = parseInt( uDlg.resizePnl.quality.text );
}

// キャンセルボタン
uDlg.cancelBtn = uDlg.add( "button", { x:95, y:370, width:100, height:25 }, "キャンセル", { name: "cancel" } );
// キャンセルボタンが押されたらキャンセル処理（ESCキー含む）
uDlg.cancelBtn.onClick = function() {
	// 実行フラグにfalseを代入
	do_flag = false;
	// ダイアログを閉じる
	uDlg.close();
}

// OKボタン
uDlg.okBtn = uDlg.add( "button", { x:205, y:370, width:100, height:25 }, "作成", { name: "ok" } );
// OKボタンが押されたら各設定項目に不備がないかチェック
uDlg.okBtn.onClick = function() {
	// 各種項目の値を格納
	settings.filePath  = uDlg.filePnl.path.text;
	settings.fileObj   = new File( settings.filePath );
	settings.pageWidth = parseInt( uDlg.pageWidthPnl.pageWidth.text );
	settings.spreads   = parseSpreads( uDlg.spreadsPnl.spreads.text );
	settings.colorMode = uDlg.resizePnl.colorMode.RGB.value ? 'RGB' : 'CMYK';
	settings.saveType  = getSaveTypeKey( uDlg.resizePnl.saveType.selection.text );

	// 対象フォルダが選択されているかチェック
	if ( !settings.filePath ) {
		alert( '対象PDFファイルが選択されていません' );
		return false;
	}
	// 対象フォルダが存在するかチェック
	if ( !settings.fileObj.exists ) {
		alert( '対象PDFファイルが存在しません' );
		return false;
	}
	// ページ幅が入力されているかチェック
	if ( _.isNaN( settings.pageWidth ) ) {
		alert( '1ページの幅を整数で入力して下さい' );
		return false;
	}
	// ページ幅が0より大きいかチェック
	if ( settings.pageWidth < 1 ) {
		alert( '0より大きい1ページの幅を入力してください' );
		return false;
	}
	// カラーモードがCMYKの時保存形式がPNGになっていないかチェック
	if ( settings.colorMode == 'CMYK' ) {
		if ( settings.saveType == 'jpgWeb' ) {
			alert( 'JPG（WEB用）形式で保存するためにはカラーモードはRGBでなければいけません' );
			return false;
		}
		if ( settings.saveType == 'png' ) {
			alert( 'PNG形式で保存するためにはカラーモードはRGBでなければいけません' );
			return false;
		}
	}
	// 保存形式がJPG／JPG（WEB用）の場合
	if ( _.contains( [ 'jpgWeb', 'jpgDtp' ], settings.saveType ) ) {
		// 画質が入力されているかチェック
		settings.quality = parseInt( uDlg.resizePnl.quality.text );
		if ( _.isNaN( settings.quality ) ) {
			alert( '画質を整数で入力して下さい' );
			return false;
		}
	}
	// 不備がなかった場合処理続行
	uDlg.close();
}

// ダイアログ表示
uDlg.show();

// ---------------------------------- メイン書き出し処理 ----------------------------------
if ( do_flag ) {
	// alert( 'フォルダ:' + uDlg.filePnl.path.text + "\n" + 'ページ幅：' + settings.pageWidth + "\n" + 'モード：' + settings.colorMode + "\n" + '保存形式：' + settings.saveType + "\n" + '画質：' + settings.quality );

	var pdfOpenOptions = new PDFOpenOptions();
	pdfOpenOptions.antiAlias        = true;
	pdfOpenOptions.mode             = OpenDocumentMode[settings.colorMode];
	pdfOpenOptions.bitsPerChannel   = BitsPerChannelType.EIGHT;
	pdfOpenOptions.resolution       = settings.pageWidth * 25.4 / getPageWidth();
	pdfOpenOptions.suppressWarnings = true;
	pdfOpenOptions.cropPage         = CropToType.BLEEDBOX;
	pdfOpenOptions.usePageNumber    = true;
	var i = 1;

	// 対象ページを結合しリサイズ→保存のループ処理
	_.each( settings.spreads, function( spread ) {
		var theBaseDoc,
		    j = 1;

		_.each( spread, function( page ) {
			pdfOpenOptions.page = page + settings.offsetPage;
			// PDFファイルをオープン
			if ( j === 1 ) {
				theBaseDoc = app.open( settings.fileObj, pdfOpenOptions );
			} else {
				theBaseDoc.resizeCanvas( settings.pageWidth * j, theBaseDoc.height, AnchorPosition.TOPRIGHT );
				var theDoc = app.open( settings.fileObj, pdfOpenOptions );
				theDoc.flatten();
				theDoc.selection.selectAll();
				theDoc.selection.copy();
				theDoc.close( SaveOptions.DONOTSAVECHANGES );
				theBaseDoc.paste();
				theBaseDoc.activeLayer.translate( theBaseDoc.activeLayer.bounds[0] * -1, 0 );
			}
			j++;
		} );

		// 画像を統合
		theBaseDoc.flatten();
		// 保存先フォルダを作成
		var saveDir = new Folder( settings.fileObj.path + '/' + settings.save.dir );
		if( !saveDir.exists ){
			saveDir.create();
		}
		// 保存用の新規オブジェクト作成
		var newFile = new File( settings.fileObj.path + '/' + settings.save.dir + '/' + settings.save.fileName + _.pad( i, 3, '0' ) + '.' + settings.save.type[settings.saveType].extension );
		// 保存形式ごとの関数を呼び出し
		saveFunctions[settings.saveType]( theBaseDoc, newFile, settings );
		// ファイルクローズ
		theBaseDoc.close( SaveOptions.DONOTSAVECHANGES );
		i++;
	} );
}

// Photoshopの設定単位を復元
app.preferences.rulerUnits = originalRulerUnits;

// ------------------------------------------ 関数 -----------------------------------------
function setSaveTypeQuality( saveType ) {
	var key = getSaveTypeKey( saveType );
	if ( _.contains( [ 'jpgWeb', 'jpgDtp' ], key ) ) {
		uDlg.resizePnl.quality.text           = settings._quality[key].init;
		uDlg.resizePnl.qualityRange.text      = "(" + settings._quality[key].min + "〜" + settings._quality[key].max + ")";
		uDlg.resizePnl.qualitySlider.minvalue = settings._quality[key].min;
		uDlg.resizePnl.qualitySlider.maxvalue = settings._quality[key].max;
		uDlg.resizePnl.qualitySlider.value    = settings._quality[key].init;
		uDlg.resizePnl.qualityText.enabled    = true;
		uDlg.resizePnl.qualityRange.enabled   = true;
		uDlg.resizePnl.quality.enabled        = true;
		uDlg.resizePnl.qualitySlider.enabled  = true;
	} else {
		uDlg.resizePnl.qualityText.enabled    = false;
		uDlg.resizePnl.qualityRange.enabled   = false;
		uDlg.resizePnl.quality.enabled        = false;
		uDlg.resizePnl.qualitySlider.enabled  = false;
	}
}

function getSaveTypeKey( saveType ) {
	var saveTypeKey;
	_.some( settings.save.type, function( value, key ) {
		if ( value.label == saveType ) {
			saveTypeKey = key;
			return true;
		}
		return false;
	} );
	return saveTypeKey;
}

function parseSpreads( str ) {
	var str = _.trim( str );
	return spreads = _.map( str.split( ',' ), function( spread ) {
		var pages = [];
		_.each( spread.split( '-' ), function( page ) {
			var page = _.trim( page );
			page = parseInt( page.replace( /^0+/, '' ) );
			if ( !_.isNaN( page ) ) {
				pages.push( page );
			}
		} );
		return pages;
	} );
}

function getPageWidth() {
	var pdfOpenOptions = new PDFOpenOptions();
	pdfOpenOptions.antiAlias        = true;
	pdfOpenOptions.bitsPerChannel   = BitsPerChannelType.EIGHT;
	pdfOpenOptions.resolution       = 25.4;
	pdfOpenOptions.usePageNumber    = true;
	pdfOpenOptions.suppressWarnings = true;
	pdfOpenOptions.cropPage         = CropToType.BLEEDBOX;
	pdfOpenOptions.page             = 1;
	var page = app.open( settings.fileObj, pdfOpenOptions );
	var pageWidth = page.width.value;
	page.close( SaveOptions.DONOTSAVECHANGES );
	return pageWidth;
}