package com.elim.htmlexcel.core;

/**
 * 工具类
 * @author Elim
 * 2017年9月6日
 */
public class Utils {

	/**
	 * 获取像素形式的字符串中包含的数字
	 * @param pxNum 可以包含px的数字形式
	 * @return 解析出来的数字
	 */
	public static Float getPxNumber(String pxNum) {
		if (pxNum.trim().endsWith("px")) {
			pxNum = pxNum.substring(0, pxNum.length() - 2);
		}
		return Float.valueOf(pxNum.trim());
	}
	
}
