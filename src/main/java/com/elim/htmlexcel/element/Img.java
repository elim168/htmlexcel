package com.elim.htmlexcel.element;

import lombok.Data;

/**
 * 对应于HTML的img元素
 * @author Elim
 * 2017年9月6日
 */
@Data
public class Img {

	private String src;
	/**
	 * 图片高度
	 */
	private String height;
	/**
	 * 图片类型
	 */
	private String type;
	
}
