package com.elim.htmlexcel.core;

import lombok.Data;

/**
 * 字体信息
 * @author Elim
 * 2017年9月6日
 */
@Data
public class Font {

	/**
	 * 字体类型
	 */
	private String family;
	/**
	 * 字体颜色
	 */
	private String color;
	/**
	 * 尺寸，单位是像素
	 */
	private Integer size;
	/**
	 * 是否需要加粗
	 */
	private boolean bold;
	
}
