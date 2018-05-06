/**
 *
 */
package jp.co.project.venus.enumeration;

import lombok.Getter;
import lombok.Setter;

/**
 * @author M.Tamura
 * @param <T>
 *
 */
public class ValidationResultBean {

	/**
	 * 列インデックス
	 */
	@Getter
	@Setter
	private int colIdx;

	/**
	 * 行インデックス
	 */
	@Getter
	@Setter
	private int rowIdx;

	@Getter
	@Setter
	private String cellAddress;

	/**
	 * メッセージに表示するラベル
	 */
	@Getter
	@Setter
	private String label;

	/**
	 * メッセージ
	 */
	@Getter
	@Setter
	private String message;
}
