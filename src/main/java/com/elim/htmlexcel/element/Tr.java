/**
 * 
 */
package com.elim.htmlexcel.element;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Pattern;

import javax.xml.bind.Unmarshaller;
import javax.xml.bind.annotation.XmlAnyElement;
import javax.xml.bind.annotation.XmlTransient;

import org.apache.commons.lang3.StringUtils;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.Text;

import com.elim.htmlexcel.core.DataType;
import com.elim.htmlexcel.core.Style;
import com.elim.htmlexcel.core.StyleAdapter;

import lombok.Getter;
import lombok.extern.slf4j.Slf4j;

/**
 * 对应于HTML table的tr
 * @author Elim
 * 2017年9月6日
 */
@Slf4j
public class Tr extends StyleElement {

	private static final String IMAGE_SRC_PATTERN = "data:image/(jpeg|png);base64,.+";
	
	@XmlTransient
	@Getter
	private List<Td> cells;
	@XmlAnyElement
	private List<Element> childs;
	
	public void afterUnmarshal(Unmarshaller unmarshaller, Object parent) throws Exception {
		if (this.childs == null || this.childs.isEmpty()) {
			return;
		}
		cells = new ArrayList<>(childs.size());
		for (Element child : childs) {
			if (!"td".equalsIgnoreCase(child.getNodeName())) {
				continue;
			}
			cells.add(this.parseCell(child));
		}
	}
	
	/**
	 * 解析单元格
	 * @param tdElement
	 * @return
	 * @throws Exception 
	 */
	private Td parseCell(Element tdElement) throws Exception {
		Td cell = new Td();
		NodeList nodes = tdElement.getChildNodes();
		if (nodes.getLength() > 0) {
			Node node = nodes.item(0);
			if (node.getNodeType() == Node.TEXT_NODE) {
				String value = ((Text)node).getData();
				cell.setValue(value);
			} else if (node.getNodeType() == Node.ELEMENT_NODE) {
				Element element = (Element) node;
				if ("img".equalsIgnoreCase(element.getTagName())) {
					Img image = this.parseImg(element);
					cell.setImage(image);
				}
			}
		}
		String autoHeight = tdElement.getAttribute("autoHeight");
		String dataType = tdElement.getAttribute("dataType");
		String rowspan = tdElement.getAttribute("rowspan");
		String colspan = tdElement.getAttribute("colspan");
		String style = tdElement.getAttribute("style");
		if (StringUtils.isNotBlank(dataType)) {
			try {
				cell.setDataType(DataType.valueOf(dataType.trim().toUpperCase()));
			} catch (Exception e) {}//忽略枚举解析错误的场景
		}
		cell.setAutoHeight(Boolean.parseBoolean(autoHeight));
		if (StringUtils.isNotBlank(style)) {
			Style styleObj = new StyleAdapter().unmarshal(style);
			cell.setStyle(styleObj);
		}
		if (StringUtils.isNotBlank(rowspan)) {
			cell.setRowspan(Integer.parseInt(rowspan.trim()));
		}
		if (StringUtils.isNotBlank(colspan)) {
			cell.setColspan(Integer.parseInt(colspan.trim()));
		}
		cell.setParent(this);
		return cell;
	}
	
	/**
	 * 解析图片信息
	 * @param imgElement
	 * @return
	 */
	private Img parseImg(Element imgElement) {
		String src = imgElement.getAttribute("src");
		if (StringUtils.isBlank(src)) {
			log.warn("img的src为空，将忽略该img");
			return null;
		}
		if (!Pattern.matches(IMAGE_SRC_PATTERN, src)) {
			log.warn("给定的图片地址[{}]不符合规范[{}]，将忽略该图片", src, IMAGE_SRC_PATTERN);
			return null;
		}
		Img image = new Img();
		String height = imgElement.getAttribute("height");
		image.setSrc(src);
		image.setHeight(height);
		if (StringUtils.isNotBlank(src)) {
			//src要求是base64格式，格式为data:image/jpeg;base64,base64Str
			int index1 = src.indexOf("image/");
			int index2 = src.indexOf(";", index1);
			String type = src.substring(index1 + 6, index2);
			image.setType(type);
		}
		return image;
	}
	
}
