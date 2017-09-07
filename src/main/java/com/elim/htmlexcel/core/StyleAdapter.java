package com.elim.htmlexcel.core;

import java.util.HashMap;
import java.util.Map;

import javax.xml.bind.annotation.adapters.XmlAdapter;

import org.apache.commons.lang3.StringUtils;

/**
 * 解析样式信息。标准的样式信息如：<code>color:red; width: 20px;</code>
 * @author Elim
 * 2017年9月6日
 */
public class StyleAdapter extends XmlAdapter<String, Style> {

	private static final String ITEM_DELIM = ";";
	private static final String KEY_VALUE_DELIM = ":";
	
	@Override
	public Style unmarshal(String value) throws Exception {
		if (StringUtils.isBlank(value)) {
			return null;
		}
		
		String[] items = StringUtils.split(value, ITEM_DELIM);
		Map<String, String> itemMap = new HashMap<>();
		for (String item : items) {
			String[] keyValues = StringUtils.split(item, KEY_VALUE_DELIM);
			if (keyValues.length != 2) {
				continue;
			}
			String key = keyValues[0];
			if (StringUtils.isBlank(key)) {
				continue;
			}
			itemMap.put(key.trim(), keyValues[1].trim());
		}
		Style style = new Style(itemMap);
		return style;
	}

	@Override
	public String marshal(Style value) throws Exception {
		return null;
	}

}
