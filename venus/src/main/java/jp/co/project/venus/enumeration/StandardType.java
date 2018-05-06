/**
 *
 */
package jp.co.project.venus.enumeration;

/**
 * @author M.Tamura
 *
 */
public enum StandardType {

	BOOLEAN("boolean"),
	CHAR("char"),
	BYTE("byte"),
	SHORT("shot"),
	INT("int"),
	LONG("long"),
	FLOAT("float"),
	DOUBLE("double");

	private String type;

	StandardType(String artType) {
		this.type = artType;
	}

	public String getType() {
		return this.type;
	}
}
