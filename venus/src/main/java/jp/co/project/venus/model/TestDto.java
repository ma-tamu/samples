/**
 *
 */
package jp.co.project.venus.model;

import java.math.BigDecimal;

import jp.co.project.venus.annotation.CellNotEmpty;
import jp.co.project.venus.annotation.CellNumber;
import jp.co.project.venus.annotation.CellType;
import jp.co.project.venus.annotation.SaveFlag;
import jp.co.project.venus.enumeration.ExcelCellType;
import lombok.Getter;
import lombok.Setter;

/**
 * @author tamu
 *
 */
public class TestDto implements BaseDto {

	@Getter
	@Setter
	@CellType(cellType = ExcelCellType.STRING)
	@CellNotEmpty(label = "Cell1")
	private String test1;

	@Getter
	@Setter
	@CellType(cellType = ExcelCellType.NUMERIC)
	@CellNotEmpty(label = "Cell2")
	@CellNumber(label = "Cell2")
	private BigDecimal decimal;

	@Getter
	@Setter
	@SaveFlag(formula = "IF(EXACT({0}&\"|\"&{1},{2}!{0}&\"|\"&{2}!{1}), 0, 1)", formulaFieldList = {"test1", "decimal"})
	private int saveFlag;
}
