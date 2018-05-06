package jp.co.project.venus.transfer;

import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import jp.co.project.venus.annotation.CellType;
import jp.co.project.venus.annotation.ExcelFormula;
import jp.co.project.venus.annotation.ExcelHeader;
import jp.co.project.venus.annotation.ExcelRelatedValidation;
import jp.co.project.venus.annotation.ExcelValidation;
import jp.co.project.venus.annotation.SaveFlag;
import jp.co.project.venus.enumeration.ExcelCellType;
import jp.co.project.venus.enumeration.StandardType;
import jp.co.project.venus.enumeration.ValidationResultBean;
import jp.co.project.venus.formula.CellFormula;
import jp.co.project.venus.model.BaseDto;
import jp.co.project.venus.model.ExcelWorkbook;
import jp.co.project.venus.model.TransferReflectionBean;
import jp.co.project.venus.validation.ExcelRelationValidator;
import jp.co.project.venus.validation.ExcelValidator;
import lombok.Getter;
import lombok.Setter;

/**
 *
 * @author M.Tamura
 *
 */
public abstract class ExcelService {

	/** ロケール. */
	protected Locale locale;

	/** Excel ワークブック. */
	protected ExcelWorkbook uplodWorkbook;

	@Getter
	@Setter
	private List<ValidationResultBean> errorResultList;

	private static final int CLASS_NAME_ROW = 0;
	private static final int METHOD_NAME_ROW = 1;
	public static final int COL_MAX = 16383;
	public static final int ROW_MAX = 1048575;

	/**
	 * コンストラクタ
	 */
	public ExcelService() {
		this.locale = Locale.JAPAN;
	}

	/**
	 * コンストラクタ
	 *
	 * @param locale
	 *            Locale
	 */
	public ExcelService(Locale locale) {
		this.locale = locale;
	}

	/**
	 * エラー有無
	 *
	 * @return boolean
	 */
	public boolean hasError() {
		return CollectionUtils.isNotEmpty(errorResultList);
	}

	/**
	 * Excel ブックを読み込み
	 *
	 * @param stream
	 *            InputStream
	 * @return Workbook
	 * @throws IOException
	 *             java.io.IOException
	 * @throws InvalidFormatException
	 *             org.apache.poi.openxml4j.exceptions.InvalidFormatException
	 */
	protected Workbook load(InputStream stream) throws IOException, InvalidFormatException {
		Workbook workbook = null;

		try {
			workbook = WorkbookFactory.create(stream);
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

		return workbook;
	}

	/**
	 * Excelの読み込み
	 *
	 * @param stream
	 *            InputStream
	 * @throws Exception
	 *             java.long.Exception
	 */
	private void loadUploadExcel(InputStream stream) throws Exception {

		if (uplodWorkbook == null) {
			try {
				// Excelの読み込み
				this.uplodWorkbook = new ExcelWorkbook();
				this.uplodWorkbook.load(stream);
			} catch (InvalidFormatException e) {
				throw e;
			} catch (IOException e) {
				throw e;
			}
		} else {
			throw new Exception("The target file has already been opened.");
		}

	}

	/**
	 * Excelの読込み
	 *
	 * @param stream
	 *            InputStream
	 * @return Map<String, List<U>>
	 * @throws Exception
	 *             java.lang.Exception
	 */
	@SuppressWarnings("unchecked")
	public <U extends BaseDto> Map<String, List<U>> readExcel(final InputStream stream) throws Exception {
		Set<String> shettNameSet = null;
		Map<String, Integer> sheetInfo = null;
		Map<String, List<U>> entityMap = new HashMap<>();
		String className = null;
		Class<U> entityClass = null;
		int rowIdx = 0;
		try {
			// 対象Excelを読み込む
			loadUploadExcel(stream);

			// シート情報のマップを取得
			sheetInfo = getSheetInfo();
			shettNameSet = sheetInfo.keySet();

			for (String sheetName : shettNameSet) {
				// 明細行の開始行
				rowIdx = sheetInfo.get(sheetName).intValue();
				className = this.uplodWorkbook.getStringCellValue(0, 0);

				// アップロード対象シートかを判断する（クラス名がセットされているか)
				if (!StringUtils.isEmpty(className)) {
					// Excelに書かれたクラス名より実態化できるか確認
					entityClass = (Class<U>) Class.forName(className);
					if (entityClass.newInstance() instanceof BaseDto) {
						// Excelからデータを読み取り、リストを追加
						entityMap.put(sheetName, readExcelDetail(sheetName, rowIdx));
					}
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		}

		return entityMap;
	}

	/**
	 * 明細行の読み込み
	 *
	 * @param sheetName
	 *            String
	 * @param startRowIdx
	 *            int
	 * @return List<U>
	 * @throws Exception
	 *             java.lang.Exceptopn
	 */
	@SuppressWarnings("unchecked")
	protected <T, U extends BaseDto> List<U> readExcelDetail(final String sheetName, final int startRowIdx)
			throws Exception {
		int lastColNum;
		Class<U> dtoClass = null;

		List<U> transferList = null;
		List<TransferReflectionBean> reflectionFiledList = null;

		// ブックから対象シートを取得する
		uplodWorkbook.selectSheet(sheetName);

		// 先頭行を取得する(0固定)
		lastColNum = uplodWorkbook.getLastCellNum(METHOD_NAME_ROW);

		try {
			for (int colIdx = 0; colIdx <= lastColNum; colIdx++) {
				String className;

				// クラス名を検索していく.
				className = uplodWorkbook.getStringCellValue(colIdx, CLASS_NAME_ROW);
				if (StringUtils.isEmpty(className)) {
					continue;
				}
				// Excelに書かれたクラス名を取得
				dtoClass = (Class<U>) Class.forName(className);
				if (dtoClass.newInstance() instanceof BaseDto) {
					reflectionFiledList = createReflection(uplodWorkbook, colIdx);
					transferList = transferMakeDetailList(startRowIdx, dtoClass, reflectionFiledList);
					break;
				}
			}
		} catch (InstantiationException e) {
			throw e;
		} catch (IllegalAccessException e) {
			throw e;
		} catch (ClassNotFoundException e) {
			throw e;
		} catch (Exception e) {
			throw e;
		} finally {
			if (transferList == null) {
				transferList = new ArrayList<>();
			}
		}

		return transferList;
	}

	/**
	 * リフレクションフィールドのリストを作成
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param targetColIdx
	 *            int
	 * @return List<TransferReflectionBean>
	 */
	protected List<TransferReflectionBean> createReflection(ExcelWorkbook workbook, int targetColIdx) {
		Cell cell = null;
		int colIdx = 0;
		int lastCol = 0;
		int lastColNum = workbook.getLastCellNum(METHOD_NAME_ROW);
		List<TransferReflectionBean> reflectionFiledNameList = new ArrayList<>();

		for (colIdx = targetColIdx + 1; colIdx < lastColNum; colIdx++) {
			if (StringUtils.isNotEmpty(workbook.getStringCellValue(colIdx, CLASS_NAME_ROW))) {
				break;
			}
		}
		lastCol = colIdx;

		if (lastCol == lastColNum) {
			int fieldLastCol = workbook.getLastCellNum(METHOD_NAME_ROW);
			for (int fieldColIdx = lastCol + 1; fieldColIdx < COL_MAX; fieldColIdx++) {
				cell = workbook.getCell(fieldColIdx, METHOD_NAME_ROW);
				if (cell == null) {
					if (fieldColIdx < fieldLastCol) {
						continue;
					}
					lastCol = fieldColIdx;
					break;
				}
			}
		}

		// リフレクションする名称のオフセットを作成
		for (colIdx = targetColIdx; colIdx < lastCol; colIdx++) {
			String filedName;
			cell = workbook.getCell(colIdx, METHOD_NAME_ROW);
			filedName = cell.getStringCellValue();
			if (filedName.length() >= 2 && filedName.charAt(0) > '#') {
				TransferReflectionBean reflectionBean = new TransferReflectionBean();
				reflectionBean.setFiledName(filedName);
				reflectionBean.setOffset(colIdx);
				reflectionFiledNameList.add(reflectionBean);
			}
		}

		return reflectionFiledNameList;
	}

	/**
	 * Excelの明細部を読込む
	 *
	 * @param startRowIdx
	 *            int
	 * @param dtoClass
	 *            Class<U>
	 * @param reflectionFiledNameList
	 *            List<TransferReflectionBean>
	 * @return List<U>
	 * @throws Exception
	 *             java.lang.Exception
	 */
	protected <U extends BaseDto> List<U> transferMakeDetailList(int startRowIdx, Class<U> dtoClass,
			List<TransferReflectionBean> reflectionFiledNameList) throws Exception {

		int endRowIdx = uplodWorkbook.getLastRowNum();
		String cellVaule = null;
		U baseDto = null;
		Cell cell = null;
		ValidationResultBean resultBean = null;
		FormulaEvaluator feval = uplodWorkbook.getCreationHelper().createFormulaEvaluator();
		List<U> detailList = new ArrayList<>();
		this.errorResultList = new ArrayList<>();

		for (int rowIdx = startRowIdx; rowIdx <= endRowIdx; rowIdx++) {
			for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
				cell = uplodWorkbook.getCell(reflectionBean.getOffset(), rowIdx);
				cellVaule = getCellVaule(cell, feval);
				reflectionBean.setFiledValue(cellVaule);
			}

			try {
				baseDto = dtoClass.newInstance();
				for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
					resultBean = null;
					Field field = baseDto.getClass().getDeclaredField(reflectionBean.getFiledName());
					field.setAccessible(true);
					resultBean = fieldValidator(rowIdx, field, reflectionBean);
					if (resultBean != null) {
						this.errorResultList.add(resultBean);
					} else {
						setReflectValue(baseDto, field, reflectionBean.getFiledValue());
					}
				}
				resultBean = relationValidator(rowIdx, baseDto);
				if (resultBean != null) {
					this.errorResultList.add(resultBean);
				}
				detailList.add(baseDto);
			} catch (InstantiationException | IllegalAccessException | NoSuchFieldException | SecurityException e) {
				e.printStackTrace();
				throw e;
			}
		}
		return detailList;
	}

	/**
	 * 対象のセルから値を取得
	 *
	 * @param cell
	 *            Cell
	 * @param feval
	 *            FormulaEvaluator
	 * @return String
	 */
	private String getCellVaule(Cell cell, FormulaEvaluator feval) {
		String cellVaule = null;
		switch (cell.getCellTypeEnum()) {
		case NUMERIC:
			BigDecimal decimal = new BigDecimal(cell.getNumericCellValue());
			cellVaule = decimal.toString();
			break;
		case STRING:
			cellVaule = cell.getStringCellValue();
			break;
		case FORMULA:
			Cell formulaCell = feval.evaluateInCell(cell);
			getCellVaule(formulaCell, feval);
			break;
		default:
			break;
		}
		return cellVaule;
	}

	/**
	 * 対象フィールド変数に設定
	 *
	 * @param baseDto
	 *            BaseDto
	 * @param field
	 *            Field
	 * @param reflectionValue
	 *            String
	 * @throws IllegalArgumentException
	 *             java.lang.IllegalArgumentException
	 * @throws IllegalAccessException
	 *             java.lang.IllegalAccessException
	 */
	private void setReflectValue(BaseDto baseDto, Field field, String reflectionValue)
			throws IllegalArgumentException, IllegalAccessException {
		Object filedValue = null;
		Class<?> clazz = field.getType();
		String className = clazz.getSimpleName();

		if (String.class.getSimpleName().equals(className)) {
			field.set(baseDto, reflectionValue);
		} else if (Character.class.getSimpleName().equals(className)) {
			field.set(baseDto, Character.valueOf(reflectionValue.charAt(0)));
		} else if (StandardType.CHAR.getType().equals(className)) {
			field.setChar(baseDto, reflectionValue.charAt(0));
		} else if (StandardType.BYTE.getType().equals(className)) {
			field.setByte(baseDto, Byte.parseByte(reflectionValue));
		} else if (StandardType.SHORT.getType().equals(className)) {
			field.setShort(baseDto, Short.parseShort(reflectionValue));
		} else if (StandardType.INT.getType().equals(className)) {
			field.setInt(baseDto, Integer.parseInt(reflectionValue));
		} else if (StandardType.LONG.getType().equals(className)) {
			field.setLong(baseDto, Long.parseLong(reflectionValue));
		} else if (StandardType.FLOAT.getType().equals(className)) {
			field.setFloat(baseDto, Float.parseFloat(reflectionValue));
		} else if (StandardType.DOUBLE.getType().equals(className)) {
			field.setDouble(baseDto, Double.parseDouble(reflectionValue));
		} else if (Byte.class.getSimpleName().equals(className)) {
			field.set(baseDto, Byte.valueOf(reflectionValue));
		} else if (Short.class.getSimpleName().equals(className)) {
			field.set(baseDto, Short.valueOf(reflectionValue));
		} else if (Integer.class.getSimpleName().equals(className)) {
			field.set(baseDto, Integer.valueOf(reflectionValue));
		} else if (Long.class.getSimpleName().equals(className)) {
			field.set(baseDto, Long.valueOf(reflectionValue));
		} else if (Float.class.getSimpleName().equals(className)) {
			field.set(baseDto, Float.valueOf(reflectionValue));
		} else if (Double.class.getSimpleName().equals(className)) {
			field.set(baseDto, Double.valueOf(reflectionValue));
		} else if (BigDecimal.class.getSimpleName().equals(className)) {
			field.set(baseDto, new BigDecimal(reflectionValue));
		} else if (Boolean.class.getSimpleName().equals(className) || StandardType.BOOLEAN.equals(className)) {
			if (Boolean.TRUE.equals(reflectionValue) || BigDecimal.ONE.toString().equals(reflectionValue)) {
				filedValue = Boolean.TRUE.toString();
			} else if (Boolean.FALSE.equals(reflectionValue) || BigDecimal.ZERO.toString().equals(reflectionValue)) {
				filedValue = Boolean.FALSE.toString();
			} else {
				filedValue = null;
			}
			if (Boolean.class.getSimpleName().equals(className)) {
				field.set(baseDto, filedValue);
			} else {
				field.setBoolean(baseDto, ((Boolean) filedValue).booleanValue());
			}
		} else {
			field.set(baseDto, null);
		}
	}

	/**
	 * 入力チェック
	 *
	 * @param rowIdx
	 *            int
	 * @param field
	 *            Field
	 * @param reflectionBean
	 *            TransferReflectionBean
	 * @return ValidationResultBean
	 * @throws ReflectiveOperationException
	 *             java.lang.ReflectiveOperationException
	 */
	@SuppressWarnings("unchecked")
	private <T, A extends Annotation> ValidationResultBean fieldValidator(int rowIdx, Field field,
			TransferReflectionBean reflectionBean) throws ReflectiveOperationException {
		ValidationResultBean resultBean = null;
		try {
			for (Annotation annotation : field.getAnnotations()) {
				for (Annotation metaAnnotation : annotation.annotationType().getAnnotations()) {
					if (ExcelValidation.class.equals(metaAnnotation.annotationType())) {
						ExcelValidation validation = (ExcelValidation) annotation.annotationType()
								.getAnnotation(metaAnnotation.annotationType());
						Class<? extends ExcelValidator<?, ?>>[] clazzs = validation.validatedBy();
						for (Class<? extends ExcelValidator<?, ?>> clazz : clazzs) {
							ExcelValidator<A, T> validator;
							validator = (ExcelValidator<A, T>) clazz.newInstance();
							validator.initialize((A) annotation);
							resultBean = validator.validator(reflectionBean.getOffset(), rowIdx,
									(T) reflectionBean.getFiledValue());
						}
						break;
					}
				}
			}
		} catch (InstantiationException | IllegalAccessException e) {
			throw e;
		}

		return resultBean;
	}

	/**
	 * 関連チェックの入力チェック
	 *
	 * @param rowIdx
	 *            int
	 * @param baseDto
	 *            BaseDto
	 * @return ValidationResultBean
	 * @throws ReflectiveOperationException
	 *             java.lang.ReflectiveOperationException
	 */
	@SuppressWarnings("unchecked")
	private <A extends Annotation> ValidationResultBean relationValidator(int rowIdx, BaseDto baseDto)
			throws ReflectiveOperationException {
		ValidationResultBean resultBean = null;
		try {
			for (Annotation annotation : baseDto.getClass().getAnnotations()) {
				for (Annotation metaAnnotation : annotation.annotationType().getAnnotations()) {
					if (ExcelRelatedValidation.class.equals(metaAnnotation.annotationType())) {
						ExcelRelatedValidation validation = (ExcelRelatedValidation) annotation.annotationType()
								.getAnnotation(metaAnnotation.annotationType());
						Class<? extends ExcelRelationValidator<?, ?>>[] clazzs = validation.validatedBy();
						for (Class<? extends ExcelRelationValidator<?, ?>> clazz : clazzs) {
							ExcelRelationValidator<A, BaseDto> validator;
							validator = (ExcelRelationValidator<A, BaseDto>) clazz.newInstance();
							validator.initialize((A) annotation);
							resultBean = validator.validator(rowIdx, baseDto);
						}
						break;
					}
				}
			}
		} catch (InstantiationException | IllegalAccessException e) {
			e.printStackTrace();
			throw e;
		}
		return resultBean;
	}

	/**
	 * 読み込み対象のシートと開始行のマップ
	 *
	 * @return Map<String, Integer>
	 */
	protected abstract Map<String, Integer> getSheetInfo();

	/**
	 * ダウンロード.
	 *
	 * @param stream
	 *            InputStream
	 * @param model
	 *            Map<String, Object>
	 * @return ExcelWorkbook
	 * @throws Exception
	 *             java.lang.Exception
	 */
	public ExcelWorkbook makeWorkbook(InputStream stream, final Map<String, Object> model) throws Exception {
		final String[] sheetNames = { getInputSheetName(), getBeforeSheetName() };
		ExcelWorkbook workbook = new ExcelWorkbook();
		List<TransferReflectionBean> reflectionFiledNameList = null;

		// テンプレートExcelを読み込み
		workbook.load(stream);
		workbook.selectSheet(sheetNames[0]);
		reflectionFiledNameList = createReflection(workbook, 1);

		for (String sheetName : sheetNames) {
			workbook.selectSheet(sheetName);
			// ヘッダー部作成
			makeInputSheetHeader(model, workbook);

			// 明細部作成
			makeInputSheetDetail(workbook, reflectionFiledNameList, model);
		}

		return workbook;
	}

	/**
	 * 入力シートのヘッダを作成
	 *
	 * @param model
	 *            Map<String, Object>
	 * @throws Exception
	 *             java.lang.Exception
	 */
	protected void makeInputSheetHeader(Map<String, Object> model, ExcelWorkbook workbook) throws Exception {
		Object headerObj = null;
		headerObj = model.get(ExcelMapKey.INPUT_HEADER);
		if (headerObj != null) {
			for (Field field : headerObj.getClass().getFields()) {
				for (Annotation annotation : field.getAnnotations()) {
					if (annotation instanceof ExcelHeader) {
						ExcelHeader header = field.getAnnotation(ExcelHeader.class);
						setCellValue(workbook, header.col(), header.row(), headerObj, field);
					}
				}
			}
		}
	}

	/**
	 * 対象セルに値を設定
	 *
	 * @param <T>
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param field
	 *            Field
	 * @throws Exception
	 *             java.lang.Exception
	 */
	protected <T> void setCellValue(ExcelWorkbook workbook, int colIdx, int rowIdx, T instance, Field field)
			throws Exception {
		try {
			field.setAccessible(true);
			Object obj = field.get(instance);
			if (obj instanceof String || obj instanceof Character) {
				workbook.setCellValue(colIdx, rowIdx, obj.toString());
			} else if (obj instanceof Number) {
				workbook.setCellValue(colIdx, rowIdx, new BigDecimal(obj.toString()));
			} else if (obj instanceof Date) {
				workbook.setCellValue(colIdx, rowIdx, (Date) obj);
			} else if (obj instanceof Calendar) {
				workbook.setCellValue(colIdx, rowIdx, (Calendar) obj);
			} else if (StandardType.CHAR.getType().equals(obj.getClass().getSimpleName())) {
				workbook.setCellValue(colIdx, rowIdx, obj.toString());
			} else if (StandardType.BYTE.getType().equals(obj.getClass().getSimpleName())
					|| StandardType.SHORT.getType().equals(obj.getClass().getSimpleName())
					|| StandardType.INT.getType().equals(obj.getClass().getSimpleName())
					|| StandardType.LONG.getType().equals(obj.getClass().getSimpleName())
					|| StandardType.FLOAT.getType().equals(obj.getClass().getSimpleName())
					|| StandardType.DOUBLE.getType().equals(obj.getClass().getSimpleName())) {
				workbook.setCellValue(colIdx, rowIdx, new BigDecimal(obj.toString()));

			} else {
				workbook.setCellValue(colIdx, rowIdx, new Boolean(obj.toString()));
			}
		} catch (IllegalArgumentException | IllegalAccessException e) {
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * 数式を設定
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param colIdx
	 *            int
	 * @param rowIdx
	 *            int
	 * @param field
	 *            Field
	 * @param reflectionFiledNameList
	 *            List<TransferReflectionBean>
	 */
	protected void setFormula(ExcelWorkbook workbook, int colIdx, int rowIdx, Field field,
			List<TransferReflectionBean> reflectionFiledNameList) {
		int formulaIdx = 0;
		CellType cellType = field.getAnnotation(CellType.class);
		String formula = cellType.formula();

		for (String formlaField : cellType.formulaFieldList()) {
			for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
				if (formlaField.equals(reflectionBean.getFiledName())) {
					formula = formula.replaceAll("{" + formulaIdx + "}",
							workbook.getCell(reflectionBean.getOffset(), rowIdx).getAddress().formatAsString());
					formulaIdx++;
					break;
				}
			}
		}
		// 数式を設定
		workbook.setCellFormula(colIdx, rowIdx, formula);
	}

	/**
	 * 入力シートの明細を作成します.
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param reflectionFiledNameList
	 *            List<TransferReflectionBean>
	 * @param model
	 *            Map<String, Object>
	 * @throws Exception
	 *             java.lang.Exception
	 */
	@SuppressWarnings("unchecked")
	protected <T, U extends BaseDto> void makeInputSheetDetail(ExcelWorkbook workbook,
			List<TransferReflectionBean> reflectionFiledNameList, Map<String, Object> model) throws Exception {
		int writeRow = getDetailRow();
		int copyRow = getLockRow();
		List<U> detailList = (List<U>) model.get(ExcelMapKey.INPUT_DETAIL);

		// コピー行に数式を設定
		setCopyRowFomula(workbook, copyRow, reflectionFiledNameList, detailList);

		for (U baseDto : detailList) {
			Class<? extends BaseDto> clazz = baseDto.getClass();

			// コピー行から書き込み対象の行へコピー
			workbook.copyRow(copyRow, writeRow);
			for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
				Field field = clazz.getDeclaredField(reflectionBean.getFiledName());
				field.setAccessible(true);
				for (Annotation annotation : field.getAnnotations()) {
					if (annotation instanceof CellType) {
						CellType cellType = field.getAnnotation(CellType.class);
						switch (cellType.cellType()) {
						case STRING:
						case NUMERIC:
							setCellValue(workbook, reflectionBean.getOffset(), writeRow, baseDto, field);
							break;
						case FORMULA:
							setFormula(workbook, reflectionBean.getOffset(), writeRow, field, reflectionFiledNameList);
							break;
						default:
							break;
						}
						break;
					} else if (annotation.annotationType().getAnnotation(ExcelFormula.class) != null) {
						setDistinctiveFormula(workbook, annotation, reflectionBean, reflectionFiledNameList, writeRow);
						break;
					} else {
						continue;
					}
				}
			}
			writeRow++;
		}
		workbook.setForceFormulaRecalculation(true);
	}

	/**
	 * セルに対して数式のコピーを設定
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param copyRowIdx
	 *            int
	 * @param reflectionFiledNameList
	 *            List<TransferReflectionBean>
	 * @param detailList
	 *            List<U>
	 * @throws NoSuchFieldException
	 *             java.lang.NoSuchFieldException
	 * @throws InstantiationException
	 *             java.lang.InstantiationException
	 * @throws IllegalAccessException
	 *             java.lang.IllegalAccessException
	 */
	@SuppressWarnings("unchecked")
	protected <U extends BaseDto> void setCopyRowFomula(ExcelWorkbook workbook, int copyRowIdx,
			List<TransferReflectionBean> reflectionFiledNameList, List<U> detailList)
			throws NoSuchFieldException, InstantiationException, IllegalAccessException {
		U dto = detailList.get(0);
		Class<U> clazz = null;

		try {
			clazz = (Class<U>) dto.getClass();
			for (TransferReflectionBean reflectionBean : reflectionFiledNameList) {
				Field field = clazz.getDeclaredField(reflectionBean.getFiledName());
				CellType cellType = field.getAnnotation(CellType.class);
				SaveFlag saveFlag = field.getAnnotation(SaveFlag.class);
				if (cellType != null && ExcelCellType.FORMULA.equals(cellType.cellType())) {
					setFormula(workbook, reflectionBean.getOffset(), copyRowIdx, field, reflectionFiledNameList);
				} else if (saveFlag != null) {
					setDistinctiveFormula(workbook, saveFlag, reflectionBean, reflectionFiledNameList, copyRowIdx);
				} else {
					continue;
				}
			}
		} catch (NoSuchFieldException e) {
			e.printStackTrace();
			throw e;
		} catch (SecurityException e) {
			e.printStackTrace();
			throw e;
		} catch (InstantiationException e) {
			e.printStackTrace();
			throw e;
		} catch (IllegalAccessException e) {
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * 特有な数機をセルに設定
	 *
	 * @param workbook
	 *            ExcelWorkbook
	 * @param annotation
	 *            Annotation
	 * @param reflectionBean
	 *            TransferReflectionBean
	 * @param reflectionFiledNameList
	 *            List<TransferReflectionBean>
	 * @param rowIdx
	 *            int
	 * @throws InstantiationException
	 *             java.lang.InstantiationException
	 * @throws IllegalAccessException
	 *             java.lang.IllegalAccessException
	 */
	protected void setDistinctiveFormula(ExcelWorkbook workbook, Annotation annotation,
			TransferReflectionBean reflectionBean, List<TransferReflectionBean> reflectionFiledNameList, int rowIdx)
			throws InstantiationException, IllegalAccessException {
		ExcelFormula excelFormula = annotation.annotationType().getAnnotation(ExcelFormula.class);
		Class<?> formulaClass = excelFormula.formulaBy();
		@SuppressWarnings("unchecked")
		CellFormula<Annotation> cellFormula = (CellFormula<Annotation>) formulaClass.newInstance();
		cellFormula.initialize(annotation);
		String formula = cellFormula.formula(rowIdx, reflectionFiledNameList, getBeforeSheetName());
		workbook.setCellFormula(reflectionBean.getOffset(), rowIdx, formula);
	}

	/**
	 * 数式を作成
	 *
	 * @param rowIdx
	 *            int
	 */
	protected abstract void makeFormula(int rowIdx);

	/**
	 * 更新フラグの数式を作成
	 *
	 * @param rowIdx
	 *            int
	 */
	protected abstract void makeFormulaSaveFlg(int rowIdx);

	/**
	 * 通常行
	 *
	 * @return int
	 */
	protected abstract int getNormalRow();

	/**
	 * ロック行
	 *
	 * @return int
	 */
	protected abstract int getLockRow();

	/**
	 * 明細開始行
	 *
	 * @return int
	 */
	protected abstract int getDetailRow();

	/**
	 * 入力シート名を取得
	 *
	 * @return String
	 */
	protected abstract String getInputSheetName();

	/**
	 * 更新前シート名を取得
	 *
	 * @return String
	 */
	protected abstract String getBeforeSheetName();

	/**
	 * コード表を取得
	 *
	 * @return String
	 */
	protected abstract String getCodeTableSheetName();
}
