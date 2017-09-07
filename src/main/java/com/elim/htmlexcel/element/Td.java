/**
 * 
 */
package com.elim.htmlexcel.element;

import com.elim.htmlexcel.core.DataType;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * 对应于HTML中的Td
 * @author Elim
 * 2017年9月6日
 */
@Data
@EqualsAndHashCode(callSuper=true)
public class Td extends StyleElement {

	/**
	 * 是否需要自动调整高度
	 */
	private boolean autoHeight;
	/**
	 * 数据类型，默认是文本形式
	 */
	private DataType dataType = DataType.STRING;
	/**
	 * 直接包含的文本
	 */
	private String value;
	/**
	 * 直接包含的图片信息
	 */
	private Img image;
	
}
