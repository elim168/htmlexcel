package com.elim.htmlexcel.core;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;

import lombok.Getter;
import lombok.Setter;

/**
 * 样式信息
 * @author Elim
 * 2017年9月6日
 */
public class Style {

	private static final String BACKGROUND_COLOR = "background-color";
	private static final String TEXT_ALIGN = "text-align";
	private static final String VERTICAL_ALIGN = "vertical-align";
	private static final String WIDTH = "width";
	private static final String HEIGHT = "height";
	private static final String COLOR = "color";
	private static final String FONT_FAMILY = "font-family";
	private static final String FONT_SIZE = "font-size";
	private static final String FONT_WEIGHT = "font-weight";
	
	/**
	 * 单个样式信息组成的Map
	 */
	private Map<String, String> styleItemMap = new HashMap<>();
	/**
	 * 字体信息
	 */
	private Font font;
	/**
	 * 标记字体是否已经解析过了
	 */
	private boolean fontResolved;
	/**
	 * 是否拥有边框
	 */
	@Setter
	@Getter
	private boolean border;
	/**
	 * 父级元素的样式信息
	 */
	@Setter
	private Style parent;
	
	public Style(Map<String, String> itemMap) {
		if (itemMap != null) {
			this.styleItemMap.putAll(itemMap);
		}
	}
	
	/**
	 * 获取背景色
	 * @return
	 */
	public String getBackgroundColor() {
		return this.get(BACKGROUND_COLOR);
	}
	
	/**
	 * 获取横向的文本对齐方式
	 * @return
	 */
	public String getAlign() {
		return this.get(TEXT_ALIGN);
	}
	
	/**
	 * 获取纵向的文本对齐方式
	 * @return
	 */
	public String getVerticalAlign() {
		return this.get(VERTICAL_ALIGN);
	}
	
	/**
	 * 获取宽度
	 * @return
	 */
	public String getWidth() {
		return this.get(WIDTH);
	}
	
	/**
	 * 获取高度
	 * @return
	 */
	public String getHeight() {
		return this.get(HEIGHT);
	}
	
	public Font getFont() {
		if (!this.fontResolved) {
			this.resolveFont();
		}
		return this.font;
	}
	
	/**
	 * 解析字体信息
	 */
	private void resolveFont() {
		Font font = new Font();
		boolean definedFont = false;//定义了颜色相关信息
		if (StringUtils.isNotBlank(this.get(COLOR))) {//定义了颜色
			font.setColor(this.get(COLOR));
			definedFont = true;
		}
		if (StringUtils.isNotBlank(this.get(FONT_FAMILY))) {
			font.setFamily(this.get(FONT_FAMILY));
			definedFont = true;
		}
		if (StringUtils.isNotBlank(this.get(FONT_SIZE))) {
			font.setSize(Utils.getPxNumber(this.get(FONT_SIZE)).intValue());
			definedFont = true;
		}
		if (StringUtils.isNotBlank(this.get(FONT_WEIGHT)) && this.get(FONT_WEIGHT).contains("bold")) {
			font.setBold(true);
			definedFont = true;
		}
		if (definedFont) {
			this.font = font;
		}
		this.fontResolved = true;
	}
	
	/**
	 * 获取Key对应的样式信息
	 * @param key
	 * @return
	 */
	private String get(String key) {
		String value = this.styleItemMap.get(key);
		if (StringUtils.isBlank(value) && this.parent != null) {
			value = this.parent.get(key);
		}
		return value;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + (border ? 1231 : 1237);
		result = prime * result + ((parent == null) ? 0 : parent.hashCode());
		result = prime * result + ((styleItemMap == null) ? 0 : styleItemMap.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		Style other = (Style) obj;
		if (border != other.border)
			return false;
		if (parent == null) {
			if (other.parent != null)
				return false;
		} else if (!parent.equals(other.parent))
			return false;
		if (styleItemMap == null) {
			if (other.styleItemMap != null)
				return false;
		} else if (!styleItemMap.equals(other.styleItemMap))
			return false;
		return true;
	}

}
