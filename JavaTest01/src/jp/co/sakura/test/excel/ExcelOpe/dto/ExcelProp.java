package jp.co.sakura.test.excel.ExcelOpe.dto;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Properties;

import org.apache.log4j.Logger;

/**
 * 設定ファイル 操作クラス
 * @author rsaito 2011/05/13
 *
 */
public class ExcelProp {

	private static Logger log = Logger.getLogger(ExcelProp.class);

	// setting
	private static final int OUTPUT_FILE = 0;
	private static final int INPUT_FILE = 1;
	private static final int AD_USER = 2;
	private static final int AD_PASS = 3;
	private static final int LDAP_PATH = 4;
	private static final int DOMAIN_NAME = 5;


	private static final String[] arraykey = {"OUTPUT_FILE", "INPUT_FILE",
		"AD_USER","AD_PASS","LDAP_PATH","DOMAIN_NAME"};


	public final static HashMap<String, String> useData = new HashMap<String, String>();

	private static String getData(String key) {
		return (String)useData.get(key);
	}

	// OUTPUT_FILE名
	public static String getOutputFile() {
		return getData(arraykey[OUTPUT_FILE]);
	}

	// INPUT_FILE名
	public static String getInputFile() {
		return getData(arraykey[INPUT_FILE]);
	}

	// AD_USER名
	public static String getAdUser() {
		return getData(arraykey[AD_USER]);
	}

	// AD_PASS名
	public static String getAdPass() {
		return getData(arraykey[AD_PASS]);
	}

	// DOMAIN_NAME名
	public static String getDomainName() {
		return getData(arraykey[DOMAIN_NAME]);
	}

	// LDAP_PATH名
	public static String getLdapPath() {
		return getData(arraykey[LDAP_PATH]);
	}

	/**
	 * プロパティファイル初期化
	 * @param propertieFile
	 */
	public static void init(String propertieFile) {
		log.info("**********************************************");
		log.info("ExcelOpe Start ");
		log.info("**********************************************");
		log.info("設定ファイル初期化　開始");

		Properties pro = new Properties();

		try {
			if (propertieFile == "") {
				pro.load(new FileInputStream(".\\conf\\ExcelOperation.properties"));
			}
			else {
				pro.load(new FileInputStream(propertieFile));
			}

			String key = "";
			for(int i = 0; i < arraykey.length; i++) {
				key = arraykey[i];
				useData.put(key,pro.getProperty(key));
			}

			log.info("読み込みファイル:" + getInputFile());
			log.info("出力ファイル:" + getOutputFile());
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			log.error(e);
		} catch (IOException e) {
			e.printStackTrace();
			log.error(e);
		}
		log.info("設定ファイル初期化 完了");
	}
}
