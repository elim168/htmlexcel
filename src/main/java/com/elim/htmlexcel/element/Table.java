/**
 * 
 */
package com.elim.htmlexcel.element;

import java.util.List;

import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlElement;

import lombok.Data;

/**
 * 对应于HTML的table
 * @author Elim
 * 2017年9月6日
 */
@Data
public class Table {

	/**
	 * 标题，会被用作sheet
	 */
	@XmlAttribute
	private String title;
	
	/**
	 * 边框，大于0表示需要边框
	 */
	@XmlAttribute
	private Integer border;
	/**
	 * 是否需要添加筛选
	 */
	@XmlAttribute
	private boolean autoFilter;
	
	@XmlElement(name="tr")
	private List<Tr> rows;
	
}
