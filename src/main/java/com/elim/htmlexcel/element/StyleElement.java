package com.elim.htmlexcel.element;

import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlTransient;
import javax.xml.bind.annotation.adapters.XmlJavaTypeAdapter;

import com.elim.htmlexcel.core.Style;
import com.elim.htmlexcel.core.StyleAdapter;

import lombok.Data;

/**
 * 能够包含样式信息的元素
 * @author Elim
 * 2017年9月6日
 */
@Data
public abstract class StyleElement {

	/**
	 * 样式信息
	 */
	@XmlAttribute
	@XmlJavaTypeAdapter(StyleAdapter.class)
	protected Style style;
	
	/**
	 * 父级元素
	 */
	@XmlTransient
	private StyleElement parent;
	
	/**
	 * 设置父样式元素
	 * @param parent
	 */
	public void setParent(StyleElement parent) {
		if (this.style != null) {
			this.style.setParent(parent.getStyle());
		}
	}
	
	/**
	 * 获取{@link Style}，存在父级元素的Style时将设置父级元素的Style
	 * @return
	 */
	public Style getStyle() {
		if (this.style == null && this.parent != null) {
			return this.parent.getStyle();
		}
		return this.style;
	}
	
}
