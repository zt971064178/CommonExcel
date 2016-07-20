package cn.itcast.common.excel.utils;

public class StringUtils extends org.apache.commons.lang3.StringUtils {
	public static String[] oldValue = new String[]{"&lt;", "&gt;","&amp;","&apos;","&quot;"} ;
	public static String[] newValue = new String[]{"<", ">","&","'","\""} ;
	
	/**
	 * 替换特殊字符
	 * @param str
	 * @return
	 * @author zhangtian
	 */
	public static String replaceEscapeChar(String str) {
		if(StringUtils.isNotBlank(str)) {
			return StringUtils.replaceEach(str, oldValue, newValue) ;
		} else {
			return str ;
		}
	}
	
	/**
	 * 替换特殊字符
	 * @param str
	 * @return
	 * @author zhangtian
	 */
	public static String replaceEscapeCharRollback(String str) {
		if(StringUtils.isNotBlank(str)) {
			return StringUtils.replaceEach(str, newValue, oldValue) ;
		} else {
			return str ;
		}
	}
}
