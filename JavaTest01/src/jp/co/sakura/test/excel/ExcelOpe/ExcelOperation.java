package jp.co.sakura.test.excel.ExcelOpe;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelDataDto;
import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelDataManager;
import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelProp;
import jp.co.sakura.test.excel.ExcelOpe.util.ExcelOutput;
import jp.co.sakura.test.excel.ExcelOpe.util.commonUtil;


/**
 * 建築工程表取り込み準備用プログラム
 *
 * @author rsaito
 *　2011/05/12
 */
public class ExcelOperation {

	static Logger log = Logger.getLogger(ExcelOperation.class);

	/**
	 * 建築工程表操作  - main -
	 * @param args
	 */
	public static void main(String[] args) {
		log.info("建築工程表読み取り 処理開始");

		// 引数チェック
		if(args.length == 0) {
			log.debug("設定ファイルのパスを引数に渡してください。");
		}

		// 設定ファイルの読み込み
		ExcelProp.init(args[0]);

		// AD認証OPEN
		// iRtn = commonUtil.ActiveDirectoryAuth();

		FileInputStream is;
		String inputFile = ExcelProp.getInputFile();
		String outputFile = ExcelProp.getOutputFile();

		try {

			ExcelDataDto dataDto = null;
			ExcelDataManager dataManager = new ExcelDataManager();

			is = new FileInputStream(inputFile);

			boolean clearFlag = true;

			// Excel読み込み
			log.info("Excelファイル読み取り中");
//			Workbook book = WorkbookFactory.create(is);

			XSSFWorkbook  book = new XSSFWorkbook(is);
			Sheet 	sheet = book.getSheetAt(0);    				/* シートは一番左のもの限定 */

			// 最大行数
			int maxSheetRow = sheet.getLastRowNum() + 1;
			log.info("最大行数:" + maxSheetRow);

			List<String> headInfo = new ArrayList<String>();

			// 行でループ
			for(int row = 0; row < maxSheetRow; row++) {
				dataDto = new ExcelDataDto();						/* Excel 1行分保存 */
				Row hssfRow = sheet.getRow(row);
				boolean headerCheck = false;
				boolean flag = false;

				if(hssfRow == null) {
					headerCheck = false;
					headInfo = new ArrayList<String>();
				} else {

					// 行の最大列数取得
					// *************************
					// 現状28項目
					// int maxCellNum = hssfRow.getLastCellNum();
					int maxCellNum = 29;

					/*
					 * 列のループ
					 */
					for(int cell = 0; cell < maxCellNum; cell++) {

						Cell hssfCell = hssfRow.getCell(cell);

						String answer = "";
						if(hssfCell != null) {

							try {

								switch (hssfCell.getCellType()) {

								case HSSFCell.CELL_TYPE_BLANK:
									break;

								case HSSFCell.CELL_TYPE_ERROR:
									break;

								case HSSFCell.CELL_TYPE_FORMULA:
								case HSSFCell.CELL_TYPE_NUMERIC:

									if(HSSFDateUtil.isCellDateFormatted(hssfCell)){
										SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy/MM/dd");
										answer = sdf1.format(hssfCell.getDateCellValue());
									} else {
										if(cell == 1) {
											DecimalFormat df2 = new DecimalFormat("00");
											answer = df2.format(((int)hssfCell.getNumericCellValue()));
										} else {
											DecimalFormat df2 = new DecimalFormat("000000");
											answer = df2.format(((int)hssfCell.getNumericCellValue()));
										}
									}

									break;

								case HSSFCell.CELL_TYPE_STRING:

									String cellString = hssfCell.getStringCellValue();

									// 置換
									cellString=cellString.replaceAll("①", "");
									cellString=cellString.replaceAll("②", "");
									cellString=cellString.replaceAll("③", "");
									cellString=cellString.replaceAll("④", "");
									cellString=cellString.replaceAll("⑤", "");
									cellString=cellString.replaceAll("⑥", "");
									cellString=cellString.replaceAll("⑦", "");
									cellString=cellString.replaceAll("⑧", "");
									cellString=cellString.replaceAll("⑨", "");
									cellString=cellString.replaceAll("⑩", "");

									cellString=cellString.replaceAll("2-10", "10");
									cellString=cellString.replaceAll("2-11", "11");
									cellString=cellString.replaceAll("2-12", "12");
									cellString=cellString.replaceAll("2-13", "13");
									cellString=cellString.replaceAll("2-1", "01");
									cellString=cellString.replaceAll("2-2", "02");
									cellString=cellString.replaceAll("2-3", "03");
									cellString=cellString.replaceAll("2-4", "04");
									cellString=cellString.replaceAll("2-5", "05");
									cellString=cellString.replaceAll("2-6", "06");
									cellString=cellString.replaceAll("2-7", "07");
									cellString=cellString.replaceAll("2-8", "08");
									cellString=cellString.replaceAll("2-9", "09");

									// データ格納
									answer = commonUtil.exchangeSmall(cellString);
									break;
								}
							} catch (IllegalStateException ie) {
								answer = "0000/00/00";
								log.info("日付にエラー発生:" + row);
							}

							if((cell == 0 && answer.equals("物件名")) || (cell == 1 && answer.equals("号棟"))) {
								headerCheck = true;

								if(clearFlag) {
									headInfo = new ArrayList<String>();
									clearFlag = false;
								}
							}

							if (headerCheck) {
								// header行はリストに格納する。
								// headerCheck = true;
								if(headInfo.size() == 0 && answer.equals("号棟")) {
									headInfo.add(0, "物件名");
								}
								headInfo.add(answer);

							} else if(!headerCheck && headInfo.size() > 0) {
								flag = true;
								clearFlag = true;

								String key = headInfo.get(cell);
								setData(dataDto, key, answer);
							} else if(!headerCheck && cell == 0 && answer != "") {
								String key = "物件名";
								setData(dataDto, key, answer);
							}
						} else {
							headerCheck = false;
						}

						// 列ループ終了

					}


					// リストに格納
					// header 空白行は除く
					if(flag) {
						dataManager.addExcelData(dataDto);
					}
				}
			} // 行ループ終了

			// データ書き出し
			ExcelOutput excelOutput = new ExcelOutput();
			excelOutput.dataOutput(dataManager, outputFile);

		} catch (MalformedURLException e1) {


		} catch (FileNotFoundException e) {
			log.error("ファイルがオープンできません。:" + e);
			e.printStackTrace();
//		} catch (InvalidFormatException e) {
//			log.error("ファイルのフォーマットが不正です:" + e);
//			e.printStackTrace();
		} catch (IOException e) {
			log.error("ファイル操作でエラーが発生しました。:" + e);
			e.printStackTrace();
		} catch (OutOfMemoryError e) {
			log.error("メモリエラー");
			e.printStackTrace();
		}

	} // main終了

	/**
	 * Excelデータを格納
	 * @param dataDto
	 * @param key ⇒ Excel表記が変更されたらエラーとする。
	 * @param data
	 */
	public static void setData(ExcelDataDto dataDto, String key, String data) {

		if(key.equals("物件名")) {

			// 物件臨時対応
			if(data.equals("馬絹3-2")){
				data = "馬絹3（二期）";
			} else if (data.equals("犬蔵2-3")) {
				data = "犬蔵2(三期)";
			} else if (data.equals("黒川2")) {
				data = "黒川2（ﾗｸﾞﾗｽはるひ野）";
			} else if (data.equals("大曽根2")) {
				data = "大曽根2(ﾗｸﾞﾗｽ大倉山2)";
			} else if (data.equals("大曽根3")) {
				data = "大曽根3(ﾗｸﾞﾗｽ大倉山3)";
			} else if (data.equals("大曽根4")) {
				data = "大曽根4(ﾗｸﾞﾗｽ大倉山4)";
			} else if (data.equals("大曽根5")) {
				data = "大曽根5(ﾗｸﾞﾗｽ大倉山5)";
			} else if (data.equals("大曽根6")) {
				data = "大曽根6(ﾗｸﾞﾗｽ大倉山6)";
			} else if (data.equals("一色")) {
				data = "葉山一色";
			} else if (data.equals("梶ヶ谷2")) {
				data = "梶ケ谷2";
			} else if (data.equals("彩光社宅")) {
				data = "有限会社彩光社宅";
			} else if (data.equals("高石3-2")) {
				data = "高石3-2期（ﾗｸﾞﾗｽ新百合ヶ丘）";
			}

			dataDto.setBukkenName(data);
		} else if (key.equals("号棟")) {
			if(dataDto.getBukkenName().equals("有限会社彩光社宅") ) {
				dataDto.setGoutou("(注文)");
			} else if (dataDto.getBukkenName().equals("潮見台") && data.equals("04")) {
				dataDto.setGoutou("(注文)");
				dataDto.setBukkenName("潮見台 4号棟　林様邸");
			} else {
				dataDto.setGoutou(data);
			}
		} else if (key.equals("ＣＡＤ") || key.equals("CAD") || key.equals("設計担当")){
			dataDto.setCadName(data);
		} else if (key.equals("建築担当")) {
			dataDto.setKentikuTantou(data);
		} else if (key.equals("ｺｰﾃﾞｨﾈｰﾀｰ") || key.equals("コーディネーター")) {
			dataDto.setCordinator(data);
		} else if (key.equals("プラン開始")) {
			dataDto.setPlanStart(data);
		} else if (key.equals("プラン決定")) {
			dataDto.setPlanDecide(data);
		} else if (key.equals("外観") || key.equals("図面依頼")) {
			dataDto.setGaikan(data);
		} else if (key.equals("外構計画")) {
			dataDto.setGaisouKeikaku(data);
		} else if (key.equals("開発検済") || key.equals("*")) {
			dataDto.setKaihatuKensai(data);
		} else if (key.equals("宅造許可")) {
			dataDto.setTakuzoKyoka(data);
		} else if (key.equals("完了公告")) {
			dataDto.setKanryoKoukoku(data);
		} else if (key.equals("位置指定公告")) {
			dataDto.setItiSiteiKoukoku(data);
		} else if (key.equals("確認申請")) {
			dataDto.setKakuninSinsei(data);
		} else if (key.equals("建築確認")) {
			dataDto.setKentikuKakunin(data);
		} else if (key.equals("地盤調査")) {
			dataDto.setJibanTyosa(data);
		} else if (key.equals("地盤改良")) {
			dataDto.setJibanKairyo(data);
		} else if (key.equals("基礎着工")) {
			dataDto.setKisoTyakkou(data);
		} else if (key.equals("配筋検査")) {
			dataDto.setHaikinKensa(data);
		} else if (key.equals("上棟")) {
			dataDto.setJoutou(data);
		} else if (key.equals("社内構造")) {
			dataDto.setSyanaikouzou(data);
		} else if (key.equals("高耐久")) {
			dataDto.setKoutaikyu(data);
		} else if (key.equals("中間検査")) {
			dataDto.setTyukanKensa(data);
		} else if (key.equals("木完.") || key.equals("木完") ) {
			dataDto.setMokkan(data);
		} else if (key.equals("社内検査")) {
			dataDto.setSyanaiKensa(data);
		} else if (key.equals("外構着手")) {
			dataDto.setGaikoTyakusyu(data);
		} else if (key.equals("完了検査")) {
			dataDto.setKanryoKensa(data);
		} else if (key.equals("外構完了")) {
			dataDto.setGaikoKanryo(data);
		} else if (key.equals("完成立会可能日")) {
			dataDto.setHikiwatasiKensa(data);
		} else if (key.equals("引渡予定日")) {
			dataDto.setHikiwatasiKanou(data);
			if(data.equals("完成")) {
				dataDto.setKanryoflag(data);
			}
		} else if (key.equals("引渡約定日")) {
			dataDto.setHikiwatasiYakujo(data);
		} else if (key.equals("風致許可")) {
			dataDto.setFutikyoka(data);
		} else if (key.equals("物件ｺｰﾄﾞ")) {
			dataDto.setBukkenCode(data);
		} else {
			//どれにもマッチしないものは、エラーチェックフラグを1にする。
			log.info("物件名:" + dataDto.getBukkenName() + " でエラーが発生しました。");
			dataDto.setErrorCheck("1");
		}
	}

}

