package jp.co.sakura.test.excel.ExcelOpe.util;

import java.io.IOException;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import org.apache.log4j.Logger;

import jcifs.Config;
import jcifs.smb.SmbFile;
import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelDataDto;
import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelDataManager;
import jp.co.sakura.test.excel.ExcelOpe.dto.ExcelProp;

/**
 * CSV出力
 * @author rsaito
 *
 */
public class ExcelOutput {
	private static Logger log = Logger.getLogger(ExcelOutput.class);

	/**
	 * ファイル出力メソッド
	 * @param dataManager
	 * @param outputFile
	 */
	public void dataOutput(ExcelDataManager dataManager, String outputFile) {
		log.info("ファイル出力処理　開始");
		log.info("ファイル名:" + outputFile);

//		FileWriter file;
		try {

			Properties prop = new Properties();
	        prop.setProperty("jcifs.smb.client.username", ExcelProp.getAdUser());
	        prop.setProperty("jcifs.smb.client.password", ExcelProp.getAdPass());
	        //prop.setProperty("jcifs.smb.client.domain", "sakura.local");
	        Config.setProperties(prop);

	        SmbFile smfile = new SmbFile(outputFile);
	        OutputStream os = smfile.getOutputStream();
	        PrintWriter pw = new PrintWriter(os);

//			file = new FileWriter(outputFile);

//			BufferedWriter outBuffer = new BufferedWriter(file);

			List<ExcelDataDto> list = dataManager.getExcelDataList();

			Iterator<ExcelDataDto> ite = list.iterator();

			while(ite.hasNext()) {
				String line = "";
				ExcelDataDto dataDto = ite.next();

				// データの書き出し
				if(dataDto.getBukkenName() != "") {
					line = dataDto.getBukkenName();
					line = line + "," + dataDto.getGoutou();
					line = line + "," + dataDto.getCadName();
					line = line + "," + dataDto.getKentikuTantou();
					line = line + "," + dataDto.getCordinator();
					line = line + "," + dataDto.getPlanStart();
					line = line + "," + dataDto.getPlanDecide();
					line = line + "," + dataDto.getGaikan();
					line = line + "," + dataDto.getGaisouKeikaku();
					line = line + "," + dataDto.getKaihatuKensai();
					line = line + "," + dataDto.getTakuzoKyoka();
					line = line + "," + dataDto.getKanryoKoukoku();
					line = line + "," + dataDto.getItiSiteiKoukoku();
					line = line + "," + dataDto.getFutikyoka();
					line = line + "," + dataDto.getKakuninSinsei();
					line = line + "," + dataDto.getKentikuKakunin();
					line = line + "," + dataDto.getJibanTyosa();
					line = line + "," + dataDto.getJibanKairyo();
					line = line + "," + dataDto.getKisoTyakkou();
					line = line + "," + dataDto.getHaikinKensa();
					line = line + "," + dataDto.getJoutou();
					line = line + "," + dataDto.getSyanaikouzou();
					line = line + "," + dataDto.getKoutaikyu();
					line = line + "," + dataDto.getTyukanKensa();
					line = line + "," + dataDto.getMokkan();
					line = line + "," + dataDto.getSyanaiKensa();
					line = line + "," + dataDto.getGaikoTyakusyu();
					line = line + "," + dataDto.getKanryoKensa();
					line = line + "," + dataDto.getGaikoKanryo();
					line = line + "," + dataDto.getHikiwatasiKensa();
					line = line + "," + dataDto.getHikiwatasiKanou();
					line = line + "," + dataDto.getHikiwatasiYakujo();
					line = line + "," + dataDto.getBukkenCode();

					if(dataDto.getErrorCheck().equals("")) {
						line = line + "," + "0";
					} else {
						line = line + "," + dataDto.getErrorCheck();
					}

					line = line + "," + dataDto.getKanryoflag();
					pw.println(line);
//					outBuffer.write(line + "\n");
				}
			}
			pw.flush();
			pw.close();
//			outBuffer.flush();
//			outBuffer.close();

			log.info("ファイル出力処理　終了");
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
