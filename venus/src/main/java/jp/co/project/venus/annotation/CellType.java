package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.enumeration.ExcelCellType;

/**
 * DL
 * @author M.Tamura
 *
 */
@Documented
@Retention(RUNTIME)
@Target(FIELD)
public @interface CellType {

	/** Excelセルタイプ */
	ExcelCellType cellType();

	int scale() default 0;

	/** 数式 */
	String formula() default "";

	/** 数式で使用するセルリスト(DTOのフィールド名) */
	String[] formulaFieldList() default {};

}
