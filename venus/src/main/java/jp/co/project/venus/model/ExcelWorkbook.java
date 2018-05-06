/**
 *
 */
package jp.co.project.venus.model;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jp.co.project.venus.enumeration.ExcelFileType;
import jp.co.project.venus.exception.ExcelWrokBookCopyException;
import lombok.Getter;

/**
 * @author M.Tamura
 *
 */
public class ExcelWorkbook {

	/** ワークブック. */
	@Getter
	private Workbook workbook;
	/** ワークシート. */
	@Getter
	private Sheet sheet;

	@Getter
	private String fileName;

	/**
	 * Excel ブックを新規作成
	 *
	 * @param excelFileType
	 *            ファイルタイプ
	 */
	public void createWorkbook(ExcelFileType excelFileType) {

		switch (excelFileType) {
		case XLS:
			this.workbook = new HSSFWorkbook();
			break;
		case XLSX:
			this.workbook = new XSSFWorkbook();
			break;
		default:
			break;
		}
	}

	/**
	 * Excel ブックを読み込み
	 *
	 * @param argFileName
	 *            String
	 * @throws IOException
	 *             java.io.IOException
	 * @throws InvalidFormatException
	 *             org.apache.poi.openxml4j.exceptions.InvalidFormatException
	 */
	public void load(String argFileName) throws IOException, InvalidFormatException {
		this.fileName = argFileName;
		InputStream in = null;
		try {
			in = new FileInputStream(argFileName);
			load(in);
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * Excel ブックを読み込み
	 *
	 * @param stream
	 *            InputStream
	 * @throws IOException
	 *             java.io.IOException
	 * @throws InvalidFormatException
	 *             org.apache.poi.openxml4j.exceptions.InvalidFormatException
	 */
	public void load(InputStream stream) throws IOException, InvalidFormatException {
		try {
			workbook = WorkbookFactory.create(stream);
			selectSheet(0);
		} catch (EncryptedDocumentException e) {
			e.printStackTrace();
			throw e;
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw e;
		} catch (IOException e) {
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * Excel ブックを保存
	 *
	 * @param stream
	 *            OutputStream
	 * @throws IOException
	 *             java.io.IOException
	 */
	public void save(OutputStream stream) throws IOException {

		try {
			workbook.write(stream);
		} catch (IOException e) {
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * ワークシートを選択
	 *
	 * @param sheetNo
	 *            int
	 */
	public void selectSheet(int sheetNo) {
		if (0 <= sheetNo && sheetNo < workbook.getNumberOfSheets()) {
			this.sheet = workbook.getSheetAt(sheetNo);
		} else {
			throw new IllegalArgumentException("argeument sheetNo out of range");
		}
	}

	/**
	 * ワークシートを選択
	 *
	 * @param sheetName
	 *            String
	 */
	public void selectSheet(String sheetName) {
		if (StringUtils.isNotBlank(sheetName)) {
			this.sheet = workbook.getSheet(sheetName);
		} else {
			throw new IllegalArgumentException("argeument sheetName not null or blank");
		}
	}

	/**
	 * シートを削除
	 *
	 * @param sheetNo
	 *            int
	 */
	public void removeSheet(int sheetNo) {
		this.workbook.removeSheetAt(sheetNo);
	}

	/**
	 * シート名を設定
	 *
	 * @param sheetNo
	 *            int
	 * @param sheetName
	 *            String
	 */
	public void setSheetName(int sheetNo, String sheetName) {
		this.workbook.setSheetName(sheetNo, sheetName);
	}

	/**
	 * オートフィルタを設定
	 *
	 * @param address
	 *            String
	 */
	public void setAutoFilter(String address) {
		this.sheet.setAutoFilter(CellRangeAddress.valueOf(address));
	}

	/**
	 * シートを保護
	 *
	 * @param password
	 *            String
	 */
	public void protectSheet(String password) {
		this.sheet.protectSheet(password);
	}

	/**
	 * シートの表示/隠しを設定
	 *
	 * @param sheetNo
	 *            int
	 * @param hidden
	 *            SheetVisibility
	 */
	public void setSheetHidden(int sheetNo, SheetVisibility hidden) {
		this.workbook.setSheetVisibility(sheetNo, hidden);
	}

	/**
	 * シートの強制再計算設定
	 *
	 * @param formulaRecalculation
	 *            boolean
	 */
	public void setForceFormulaRecalculation(boolean formulaRecalculation) {
		this.sheet.setForceFormulaRecalculation(formulaRecalculation);
	}

	/**
	 * ヘルパーを取得
	 *
	 * @return CreationHelper
	 */
	public CreationHelper getCreationHelper() {
		return workbook.getCreationHelper();
	}

	/**
	 * 最終行を取得
	 *
	 * @return 最終行
	 */
	public int getLastRowNum() {
		return this.sheet.getLastRowNum();
	}

	/**
	 * 最終列を取得
	 *
	 * @param rowNo
	 *            int
	 * @return int
	 */
	public int getLastCellNum(int rowNo) {
		int lastCellNum = 0;
		Row row = getRow(rowNo);

		if (row != null) {
			lastCellNum = row.getLastCellNum();
		}

		return lastCellNum;
	}

	/**
	 * 行を削除
	 *
	 * @param rowNo
	 *            int
	 */
	public void removeRow(int rowNo) {
		Row row = getRow(rowNo);
		if (row != null) {
			sheet.removeRow(row);
		}
	}

	/**
	 * 行をグループ化
	 *
	 * @param fromRowNo
	 *            int
	 * @param toRowNo
	 *            int
	 */
	public void groupRow(int fromRowNo, int toRowNo) {
		this.sheet.groupRow(fromRowNo, toRowNo);
	}

	/**
	 * 行グループの折りたたみ/展開を設定
	 *
	 * @param argRow
	 *            int
	 * @param collapse
	 *            boolean
	 */
	public void setRowGroupCollapsed(int argRow, boolean collapse) {
		this.sheet.setRowGroupCollapsed(argRow, collapse);
	}

	/**
	 * レンジアドレスを取得(ex."A1:B2")
	 *
	 * @param startColIdx
	 *            int
	 * @param startRowIdx
	 *            int
	 * @param endColIdx
	 *            int
	 * @param endRowIdx
	 *            int
	 * @return String
	 */
	public String getRangeAddress(int startColIdx, int startRowIdx, int endColIdx, int endRowIdx) {
		CellRangeAddress address = new CellRangeAddress(startRowIdx, endRowIdx, startColIdx, endColIdx);
		return address.formatAsString();
	}

	/**
	 * 行の取得
	 *
	 * @param rowIdx
	 *            int
	 * @return Row
	 */
	public Row getRow(int rowIdx) {
		Row row = null;

		row = this.sheet.getRow(rowIdx);
		if (row == null) {
			// 対象の行が存在しない場合は、新しく作成する
			row = this.sheet.createRow(rowIdx);
		}

		return row;
	}

	/**
	 * セルの取得
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return Cell
	 */
	public Cell getCell(int colIdx, int rowIdx) {
		Row row = null;
		Cell cell = null;

		row = getRow(rowIdx);
		cell = row.getCell(colIdx);
		if (cell == null) {
			cell = row.createCell(colIdx);
		}

		return cell;
	}

	/**
	 * セル値(真理値)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            boolean
	 */
	public void setCellValue(int colIdx, int rowIdx, boolean value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value);
	}

	/**
	 * セル値(カレンダ)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            Calendar
	 */
	public void setCellValue(int colIdx, int rowIdx, Calendar value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value);
	}

	/**
	 * セル値(日付)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            Date
	 */
	public void setCellValue(int colIdx, int rowIdx, Date value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value);
	}

	/**
	 * セル値(数値)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            double
	 */
	public void setCellValue(int colIdx, int rowIdx, double value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value);
	}

	/**
	 * セル値(数値)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            BigDecimal
	 */
	public void setCellValue(int colIdx, int rowIdx, BigDecimal value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value.doubleValue());
	}

	/**
	 * セル値(文字列)を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param value
	 *            String
	 */
	public void setCellValue(int colIdx, int rowIdx, String value) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellValue(value);
	}

	/**
	 * セル計算式を設定
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param formula
	 *            String
	 */
	public void setCellFormula(int colIdx, int rowIdx, String formula) {
		Cell cell = getCell(colIdx, rowIdx);
		cell.setCellFormula(formula);
	}

	/**
	 * セルから値を取得
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return String
	 */
	public String getStringCellValue(int colIdx, int rowIdx) {
		Cell cell = getCell(colIdx, rowIdx);
		return cell.getStringCellValue();
	}

	/**
	 * セルから値を取得
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return Date
	 */
	public Date getDateCellValue(int colIdx, int rowIdx) {
		Cell cell = getCell(colIdx, rowIdx);
		return cell.getDateCellValue();
	}

	/**
	 * セルから値を取得
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return boolean
	 */
	public boolean getBooleanCellValue(int colIdx, int rowIdx) {
		Cell cell = getCell(colIdx, rowIdx);
		return cell.getBooleanCellValue();
	}

	/**
	 * セルから値を取得
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return double
	 */
	public double getNumericCellValue(int colIdx, int rowIdx) {
		Cell cell = getCell(colIdx, rowIdx);
		return cell.getNumericCellValue();
	}

	/**
	 * 対象シートをアクティブ
	 *
	 * @param sheetIndex
	 *            int
	 */
	public void setActiveSheet(int sheetIndex) {
		this.workbook.setActiveSheet(sheetIndex);
	}

	/**
	 * 対象シートの位置を変更する
	 *
	 * @param sheetName
	 *            String
	 * @param position
	 *            int
	 */
	public void setSheetOrder(String sheetName, int position) {
		this.workbook.setSheetOrder(sheetName, position);
	}

	/**
	 * シート名からシートインデックスを取得する
	 *
	 * @param sheetName
	 *            String
	 * @return int
	 */
	public int getSheetIndex(String sheetName) {
		return this.workbook.getSheetIndex(sheetName);
	}

	/**
	 * 行のコピー
	 *
	 * @param srcRowIdx
	 *            int
	 * @param distRowIdx
	 *            int
	 */
	public void copyRow(int srcRowIdx, int distRowIdx) {
		int lastColIdx = 0;
		Row srcRow = null;
		Row distRow = null;
		if (srcRowIdx != distRowIdx) {
			srcRow = getRow(srcRowIdx);
			lastColIdx = srcRow.getLastCellNum();

			distRow = getRow(distRowIdx);
			if (distRow == null) {
				// 対象行が存在しない場合は、行を作成する
				this.sheet.createRow(distRowIdx);
			}

			for (int colIdx = 0; colIdx < lastColIdx; colIdx++) {
				// セルのコピー
				copyCell(colIdx, srcRowIdx, colIdx, distRowIdx);
			}
			// 行の高さをコピー元と揃える
			distRow.setHeight(srcRow.getHeight());
		} else {
			throw new ExcelWrokBookCopyException("The copy destination row is the same as the copy source row");
		}
	}

	/**
	 * セルのコピー
	 *
	 * @param srcColIdx
	 *            int
	 * @param srcRowIdx
	 *            int
	 * @param distColIdx
	 *            int
	 * @param distRowIdx
	 *            int
	 */
	public void copyCell(int srcColIdx, int srcRowIdx, int distColIdx, int distRowIdx) {
		Cell srcCell = null;
		Cell distCell = null;

		srcCell = getCell(srcColIdx, srcRowIdx);
		distCell = getCell(distColIdx, distRowIdx);

		switch (srcCell.getCellTypeEnum()) {
		case NUMERIC:
			distCell.setCellType(srcCell.getCellTypeEnum());
			distCell.setCellValue(srcCell.getNumericCellValue());
			break;
		case STRING:
			distCell.setCellType(srcCell.getCellTypeEnum());
			distCell.setCellValue(srcCell.getRichStringCellValue());
			break;
		case FORMULA:
			distCell.setCellType(srcCell.getCellTypeEnum());
			distCell.setCellValue(srcCell.getCellFormula());
			break;
		case BOOLEAN:
			distCell.setCellType(srcCell.getCellTypeEnum());
			distCell.setCellValue(srcCell.getBooleanCellValue());
			break;
		case ERROR:
			distCell.setCellType(srcCell.getCellTypeEnum());
			distCell.setCellValue(srcCell.getErrorCellValue());
			break;
		default:
			break;
		}
		// スタイルのコピー
		distCell.setCellStyle(srcCell.getCellStyle());
	}

	/**
	 * ExcelのカラムIndexからExcelの列名(A.B...)に変換
	 *
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @return String
	 */
	public static String toCellAddress(int colIdx, int rowIdx) {
		// アルファベットの先頭
		int firstAlphabet = 'A';
		// アルファベットの数
		int alphabetSize = 26;
		// ワーク
		int workColumnIndex = 0;
		int divsion = 0;
		int remainder = 0;
		StringBuffer sb = new StringBuffer();

		if (0 <= colIdx) {
			if (alphabetSize <= colIdx) {
				// ワークの初期値として引数のIndexを設定
				workColumnIndex = colIdx;
				while (true) {
					// 商
					divsion = workColumnIndex / alphabetSize;
					// 余
					remainder = workColumnIndex % alphabetSize;

					// アルファベットに変換
					sb.append(String.valueOf((char) (firstAlphabet + divsion)));

					if (divsion <= 0) {
						break;
					}
					// 余りをワークへ
					workColumnIndex = remainder;
				}
			} else {
				// アルファベットに変換
				sb.append(String.valueOf((char) (firstAlphabet + colIdx)));
			}
		}
		return sb.reverse().toString() + String.valueOf((rowIdx + 1));
	}
}
