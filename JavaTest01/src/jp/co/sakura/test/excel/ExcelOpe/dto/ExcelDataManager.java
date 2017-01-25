package jp.co.sakura.test.excel.ExcelOpe.dto;

import java.util.ArrayList;
import java.util.List;

/**
 * Excelデータクラスの管理用
 * @author rsaito
 *
 */
public class ExcelDataManager {

	// ExcelDataDto格納用
	private List<ExcelDataDto> excelDataList = new ArrayList<ExcelDataDto>();

	/**
	 * DTO取得用
	 * @return
	 */
	public List<ExcelDataDto> getExcelDataList() {
		return excelDataList;
	}

	/**
	 *
	 * @param excelDataList
	 */
	public void setExcelDataList(List<ExcelDataDto> excelDataList) {
		this.excelDataList = excelDataList;
	}

	/**
	 * DTO追加処理
	 * @param excelDataDtos
	 */
	public void addExcelData(ExcelDataDto excelDataDtos) {
		this.excelDataList.add(excelDataDtos);
	}

}
