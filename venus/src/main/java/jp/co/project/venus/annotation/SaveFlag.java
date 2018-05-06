/**
 *
 */
package jp.co.project.venus.annotation;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import jp.co.project.venus.formula.DefaultSaveFlagFormula;

/**
 * @author M.Tamura
 *
 */
@Documented
@Retention(RUNTIME)
@Target(FIELD)
@ExcelFormula(formulaBy=DefaultSaveFlagFormula.class)
public @interface SaveFlag {
	/** 数式 */
	String formula();

	/** 数式で使用するセルリスト(DTOのフィールド名) */
	String[] formulaFieldList();
}
