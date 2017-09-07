package com.elim.htmlexcel.element;

import java.io.StringReader;
import java.util.List;

import javax.xml.bind.JAXB;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;

import org.apache.commons.lang3.StringUtils;

import lombok.Data;

/**
 * 用来作为table的载体
 * @author Elim
 * 2017年9月6日
 */
@XmlRootElement(name="body")
@Data
public class Body {

	@XmlElement(name="table")
	private List<Table> tables;
	
	/**
	 * 解析HTML文本为对应的Body对象
	 * @param html 需要解析的HTML文本，其中只能包含table元素
	 * @return
	 */
	public static Body parse(String html) {
		if (StringUtils.isBlank(html)) {
			throw new IllegalArgumentException("参数HTML不能为空");
		}
		html = "<body>" + html + "</body>";
		return JAXB.unmarshal(new StringReader(html), Body.class);
	}
	
}
