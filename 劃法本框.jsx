main();
function main(){
	ZhiKuan=420;              //紙寬mm
	ZhiGao=90;                //紙高mm
	BiaoTiYeBaiBian=80;  //標題頁白邊寬mm
	BaiBian=40;              //白邊寬mm
	HangGao=8;              //行高mm
	BT_NeiKuangHangGao=2*HangGao;//標題頁內框高=兩倍行高：2*HangGao，也可以自定
	Hua_NeiKuangHangGao=4*HangGao;//花頁內框高=四倍行高：4*HangGao，，也可以自定
	Chang_NeiKuangHangGao=6*HangGao;//常頁內框高=六倍行高：6*HangGao，也可以自定
	BT_TuKuan=3*HangGao//標題頁圖框寬=三倍行高：3*HangGao，也可以自定
	TuKuan=3.5*HangGao;//圖框寬=四倍行高：4*HangGao，，也可以自定
	WuTuKuan=2*HangGao;//無圖框寬=兩倍行高：2*HangGao，也可以自定
	XianCu=1;           //線粗磅數=1磅，也可以自定
	xianJianJu=0.6;    //雙線的間距=0.6mm，也可以自定
	YeMaJu=5;           //頁碼距離外框的距離
//	BookName ="ཡིག་མཚན";//書名
	//劃法本框===============================================
	var myDocument = app.documents.add();//打開新文檔
	myDocument.viewPreferences.horizontalMeasurementUnits = MeasurementUnits.MILLIMETERS;//標尺單位mm
	myDocument.viewPreferences.verticalMeasurementUnits = MeasurementUnits.MILLIMETERS;
	myDocument.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;//每一頁有自己的原點
	with(myDocument.documentPreferences){//設定紙張的大小
		pageHeight = ZhiGao;
		pageWidth =ZhiKuan;
		pageOrientation =PageOrientation.landscape;
		pagesPerDocument = 1;
		documentBleedUniformSize = true;
		documentBleedBottomOffset = "0";
		documentSlugUniformSize =true;
		slugTopOffset="0";
		font ="aril"
	}
	for(i=0; i<8; i++){//加9新主頁
		var myMasterSpread = myDocument.masterSpreads.add(); 
	}
	//畫框線
	HuaKuang(0, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(0, 1, 1, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(1, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(1, 1, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(2, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(2, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(3, 0, 3, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(3, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(4, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(4, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(5, 0, 5, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(5, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(6, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(6, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(7, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(7, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(8, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaKuang(8, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
//畫正文框	
	HuaXian(0, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(0, 1, 1, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(1, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(1, 1, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(2, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(2, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(3, 0, 3, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(3, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(4, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(4, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(5, 0, 5, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(5, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(6, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(6, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(7, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(7, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(8, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	HuaXian(8, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu);
	app.documents.item(0).masterSpreads.item(1).pages.item(0).textFrames.item(0).nextTextFrame = app.documents.item(0).masterSpreads.item(1).pages.item(1).textFrames.item(0);//鏈接兩個文本框，只連接常用框，以免亂跑
/* //畫書名框
	HuaShuMingKuang(0, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(1, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(2, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(3, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(4, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(5, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(6, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(7, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
	HuaShuMingKuang(8, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName);
//畫頁碼框
	HuaYeMaKuang(0, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(0, 1, 1, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(1, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(1, 1, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(2, 0, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(2, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(3, 0, 3, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(3, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(4, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(4, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(5, 0, 5, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(5, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(6, 0, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(6, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(7, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(7, 1, 2, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(8, 0, 6, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
	HuaYeMaKuang(8, 1, 4, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, YeMaJu);
 */}
/* //功能：畫頁碼
function HuaYeMaKuang(ZhuYeHao, BianHao, KuangLie, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName){
	var myDocument = app.documents.item(0)
	var myPage = myDocument.masterSpreads.item(ZhuYeHao).pages.item(BianHao);
	myTextFrame= myPage.textFrames.add();
	myTextFrame.geometricBounds = [ZhiGao/2-HangGao/2, ZhiKuan-BaiBian+YeMaJu-45, ZhiGao/2-HangGao/2+HangGao, ZhiKuan-BaiBian+YeMaJu-45+ZhiGao];
	myTextFrame.absoluteRotationAngle = 90;
	myTextFrame.contents = "PageNumber";
	myTextFrame.parentStory.characters.item(0).justification = Justification.centerAlign;
	with(myTextFrame.textFramePreferences){
		firstBaselineOffset = FirstBaseline.LEADING_OFFSET
	}
}
//功能：畫書名框
function HuaShuMingKuang(ZhuYeHao, BianHao, KuangLie, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu, BookName){
		if (KuangLie==1){//標題
			TK=BT_TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan無圖框寬
			NKG=BT_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高Hua_NeiKuangHangGao花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BiaoTiYeBaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian白邊寬
		}
		if (KuangLie==2){//0圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==3){//1圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==4){//2圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==5){//3圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==6 ){
		TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
		NKG=Chang_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
		BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
	var myDocument = app.documents.item(0)
	var myPage = myDocument.masterSpreads.item(ZhuYeHao).pages.item(BianHao);
	myTextFrame= myPage.textFrames.add();
	myTextFrame.geometricBounds = [ZhiGao/2-HangGao/2, BB+HangGao/2-ZhiGao/2, (ZhiGao/2-HangGao/2)+HangGao, (BB+HangGao/2-ZhiGao/2)+ZhiGao];
	myTextFrame.absoluteRotationAngle = 270;
	myTextFrame.contents = BookName;
	myTextFrame.parentStory.characters.item(0).justification = Justification.centerAlign;
	with(myTextFrame.textFramePreferences){
		firstBaselineOffset = FirstBaseline.LEADING_OFFSET
	}
} */
//功能：畫正文 框
function HuaXian(ZhuYeHao, BianHao, KuangLie, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu){
		if (KuangLie==1){//標題
			TK=BT_TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan無圖框寬
			NKG=BT_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高Hua_NeiKuangHangGao花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BiaoTiYeBaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian白邊寬
		}
		if (KuangLie==2){//0圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==3){//1圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==4){//2圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==5){//3圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==6 ){
		TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
		NKG=Chang_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
		BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
	var myDocument = app.documents.item(0)
	var myPage = myDocument.masterSpreads.item(ZhuYeHao).pages.item(BianHao);//page.item(1)是右邊，.item(0)是左邊
	myTextFrame= myPage.textFrames.add();
	if (KuangLie<6){
		myTextFrame.geometricBounds = [((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao, BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu+xianJianJu+HangGao, ((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao+NKG, BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu+xianJianJu+HangGao+(ZhiKuan-2*BB-8*xianJianJu-2*TK-4*HangGao)];
	}else{
		myTextFrame.geometricBounds = [(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu, BB+2*xianJianJu+HangGao, ((ZhiGao-NKG-2*xianJianJu)/2+xianJianJu)+NKG, (BB+2*xianJianJu+HangGao)+(ZhiKuan-2*BB-4*xianJianJu-2*HangGao)];
	}
	myTextFrame.label = "ZhongXinWenZiKuang";
	with (myTextFrame.textFramePreferences){
		firstBaselineOffset = FirstBaseline.LEADING_OFFSET
	}
}
//功能-畫框：（框的種類，紙寬，紙高，標題頁白邊寬，白邊寬，行高，標題頁內框高，花頁內框高，常頁內框高，標題頁圖框寬，圖框寬，無圖框寬，線粗磅數，雙線的間距）
function HuaKuang(ZhuYeHao, BianHao, KuangLie, ZhiKuan, ZhiGao, HangGao, BiaoTiYeBaiBian, BaiBian, HangGao, BT_NeiKuangHangGao, Hua_NeiKuangHangGao, Chang_NeiKuangHangGao, BT_TuKuan, TuKuan, WuTuKuan, XianCu, xianJianJu){
	var myDocument = app.documents.item(0)
	var myPage = myDocument.masterSpreads.item(ZhuYeHao).pages.item(BianHao);//page.item(1)是右邊，.item(0)是左邊
	if (KuangLie<6){
		if (KuangLie==1){//標題
			TK=BT_TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan無圖框寬
			NKG=BT_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高Hua_NeiKuangHangGao花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BiaoTiYeBaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian白邊寬
		}
		if (KuangLie==2){//0圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==3){//1圖
			TK=WuTuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==4){//2圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		if (KuangLie==5){//3圖
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
			NKG=Hua_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
			BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		}
		bt_Y=(ZhiGao-2*HangGao-NKG-4*xianJianJu)/2;
		bt_X=BB;
		bt_YK=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+4*xianJianJu+2*HangGao+NKG;
		bt_XK=BB+(ZhiKuan-2*BB);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//2============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+xianJianJu;
		bt_X=BB+xianJianJu;
		bt_YK=bt_Y+2*HangGao+2*xianJianJu+NKG;
		bt_XK=bt_X+(ZhiKuan-2*BB-2*xianJianJu);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//3============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+xianJianJu+HangGao;
		bt_X=BB+xianJianJu+HangGao;
		bt_YK=bt_Y+2*xianJianJu+NKG;
		bt_XK=bt_X+(ZhiKuan-2*BB-2*xianJianJu-2*HangGao);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//4============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
		bt_X=BB+xianJianJu+HangGao+xianJianJu;
		bt_YK=bt_Y+NKG;
		bt_XK=bt_X+TK;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//5============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
		bt_X=BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu;
		bt_YK=bt_Y+NKG;
		bt_XK=bt_X+HangGao;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//6============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
		bt_X=BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu+xianJianJu+HangGao;
		bt_YK=bt_Y+NKG;
		bt_XK=bt_X+(ZhiKuan-2*BB-8*xianJianJu-2*TK-4*HangGao);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//7============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
		bt_X=ZhiKuan-BB-3*xianJianJu-2*HangGao-TK;
		bt_YK=bt_Y+NKG;
		bt_XK=bt_X+HangGao;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//8============================================================================================
		bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
		bt_X=ZhiKuan-BB-2*xianJianJu-HangGao-TK;
		bt_YK=bt_Y+NKG;
		bt_XK=bt_X+TK;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		with(myPage.marginPreferences){
			bottom =((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao
			left =BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu+xianJianJu+HangGao
			right =BB+xianJianJu+HangGao+xianJianJu+TK+xianJianJu+xianJianJu+HangGao
			top =((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao
		}
		//9============================================================================================
		if (KuangLie==3 | KuangLie==5 ){
			TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬			
			bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
			bt_X=(ZhiKuan-TK-2*HangGao)/2;
			bt_YK=bt_Y+NKG;
			bt_XK=bt_X+TK+2*HangGao;
			myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
			myRectangle.strokeWeight = XianCu;
//			myRectangle.textWrapPreference.textWrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP
			//10============================================================================================
			bt_Y=((ZhiGao-2*HangGao-NKG-4*xianJianJu)/2)+2*xianJianJu+HangGao;
			bt_X=(ZhiKuan-TK)/2;
			bt_YK=bt_Y+NKG;
			bt_XK=bt_X+TK;
			myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
			myRectangle.strokeWeight = XianCu;
//			myRectangle.tagsextWrapPreference.textWrapMode = TextWrapModes.BOUNDING_BOX_TEXT_WRAP
		}
	}
	if (KuangLie==6 ){
		TK=TuKuan;//圖框寬//BT_TuKuan標題頁圖框寬 TuKuan 圖框寬 WuTuKuan 無圖框寬
		NKG=Chang_NeiKuangHangGao;//內框高: BT_NeiKuangHangGao標題頁內框高 Hua_NeiKuangHangGao 花頁內框高 Chang_NeiKuangHangGao 常頁內框高
		BB=BaiBian;//BiaoTiYeBaiBian標題頁白邊寬 BaiBian 白邊寬
		//外框
		bt_Y=(ZhiGao-NKG-2*xianJianJu)/2;
		bt_X=BB;
		bt_YK=((ZhiGao-NKG-2*xianJianJu)/2)+2*xianJianJu+NKG;
		bt_XK=BB+(ZhiKuan-2*BB);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//左小框
		bt_Y=(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu;
		bt_X=BB+xianJianJu;
		bt_YK=((ZhiGao-NKG-2*xianJianJu)/2+xianJianJu)+NKG;
		bt_XK=BB+xianJianJu+HangGao;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//中大框
		bt_Y=(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu;
		bt_X=BB+2*xianJianJu+HangGao;
		bt_YK=((ZhiGao-NKG-2*xianJianJu)/2+xianJianJu)+NKG;
		bt_XK=(BB+2*xianJianJu+HangGao)+(ZhiKuan-2*BB-4*xianJianJu-2*HangGao);
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
		//右小框
		bt_Y==(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu;
		bt_X=ZhiKuan-BB-xianJianJu-HangGao;
		bt_YK=((ZhiGao-NKG-2*xianJianJu)/2+xianJianJu)+NKG;
		bt_XK=(ZhiKuan-BB-xianJianJu-HangGao)+HangGao;
		myRectangle = myPage.rectangles.add({geometricBounds:[bt_Y, bt_X, bt_YK, bt_XK]});
		myRectangle.strokeWeight = XianCu;
	with(myPage.marginPreferences){
			bottom =(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu
			left =BB+2*xianJianJu+HangGao
			right =BB+2*xianJianJu+HangGao
			top =(ZhiGao-NKG-2*xianJianJu)/2+xianJianJu
	}
	}
}