package com.elim.htmlexcel.core;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.internal.ZipHelper;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.elim.htmlexcel.element.Body;
import com.elim.htmlexcel.element.Img;
import com.elim.htmlexcel.element.Table;
import com.elim.htmlexcel.element.Td;
import com.elim.htmlexcel.element.Tr;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;

/**
 * Excel生成器，基于XML的方式生成Excel， 参考<a href=
 * "https://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/xssf/usermodel/examples/BigGridDemo.java">
 * POI官方示例</a> 所写。
 * 
 * @author Elim 2017年9月6日
 */
@Slf4j
public class ExcelGenerator {

	private static final String XML_ENCODING = System.getProperty("file.encoding");
	private static final String TEMPLATE = "template";
	private static final String SUFFIX = ".xlsx";
	private static final boolean WRAP_TEXT = true;

	private static final ColorResolver colorResolver = new ColorResolver();
	
	private Map<Style, CellStyle> cellStyles = new HashMap<>();
	private List<ImageCell> imageCells = new ArrayList<>();
	/**
	 * 通用的单元格样式
	 */
	private Map<Table, CellStyle> commonCellStyles = new HashMap<>();

	/**
	 * 生成Excel文件
	 * 
	 * @param html
	 *            用来生成Excel文件的HTML文本
	 * @param out
	 *            生成的Excel文件的输出目标
	 * @throws IOException
	 */
	public void generate(String html, OutputStream out) throws IOException {
		Body body = Body.parse(html);
		File template = this.prepare(body);
		if (!this.imageCells.isEmpty()) {
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			this.build(body, template, baos);
			try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
				this.imageCells.forEach(imageCell -> {
					this.addImage(wb, imageCell);
				});
				wb.write(out);
			}
		} else {
			this.build(body, template, out);
		}
		template.delete();
	}

	/**
	 * 准备好模板，同时返回对应的Excel模板文件
	 * 
	 * @param body
	 * @return
	 * @throws IOException
	 */
	private File prepare(Body body) throws IOException {
		File template = File.createTempFile(TEMPLATE, SUFFIX);
		try (XSSFWorkbook wb = new XSSFWorkbook(); OutputStream out = new FileOutputStream(template);) {
			this.createStyles(body, wb);
			List<Table> tables = body.getTables();
			for (int i=0; i<tables.size(); i++) {
				Table table = tables.get(i);
				if (StringUtils.isNotBlank(table.getTitle())) {
					wb.createSheet(table.getTitle());
				} else {
					wb.createSheet();
				}
			}
			wb.write(out);
		}
		return template;
	}
	
	/**
	 * 添加图片
	 * @param wb
	 * @param imageCell
	 */
	private void addImage(Workbook wb, ImageCell imageCell) {
		Img image = imageCell.getImage();
		String src = image.getSrc();
		String base64 = src.substring(src.lastIndexOf(",") + 1);
		byte[] bytes = Base64.getDecoder().decode(base64);
		int format = Workbook.PICTURE_TYPE_JPEG;
		if ("png".equalsIgnoreCase(image.getType())) {
			format = Workbook.PICTURE_TYPE_PNG;
		}
		int pictureIndex = wb.addPicture(bytes, format);
		Sheet sheet = wb.getSheetAt(imageCell.getSheet());
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		ClientAnchor clientAnchor = wb.getCreationHelper().createClientAnchor();
		clientAnchor.setCol1(imageCell.getCol1());
		clientAnchor.setCol2(imageCell.getCol2() + 1);
		clientAnchor.setRow1(imageCell.getRow1());
		clientAnchor.setRow2(imageCell.getRow2() + 1);
		drawing.createPicture(clientAnchor, pictureIndex);
	}

	/**
	 * 构建Excel文件
	 * @param body
	 * @param template
	 * @param out
	 * @throws IOException
	 */
	private void build(Body body, File template, OutputStream out) throws IOException {
		List<Table> tables = body.getTables();
		try (ZipFile zip = ZipHelper.openZipFile(template); ZipOutputStream zos = new ZipOutputStream(out);) {
			Enumeration<? extends ZipEntry> entries = zip.entries();
			while (entries.hasMoreElements()) {
				ZipEntry zipEntry = entries.nextElement();
				if (!zipEntry.getName().startsWith("xl/worksheets/")) {
					zos.putNextEntry(new ZipEntry(zipEntry.getName()));
					InputStream is = zip.getInputStream(zipEntry);
					this.copyStream(is, zos);
					is.close();
				}
			}
			int sheet = 1;
			for (Table table : tables) {
				String entry = "xl/worksheets/sheet" + sheet++ + ".xml";
				ZipEntry zipEntry = new ZipEntry(entry);
				zos.putNextEntry(zipEntry);
				String content = this.buildSheetData(table);
				zos.write(content.getBytes(XML_ENCODING));
			}
		}
	}
	
	/**
	 * 构建sheet数据
	 * @param table
	 * @return
	 * @throws IOException
	 */
	private String buildSheetData(Table table) throws IOException {
		SpreadsheetWriter sw = new SpreadsheetWriter(table);
		return sw.buildData();
	}
	
	/**
	 * 预先准备好样式信息
	 * @param body
	 * @param wb
	 */
	private void createStyles(Body body, XSSFWorkbook wb) {
		List<Table> tables = body.getTables();
		for (int i=0; i<tables.size(); i++ ) {
			Table table = tables.get(i);
			List<Tr> rows = table.getRows();
			boolean border = table.getBorder() != null && table.getBorder() > 0;
			for (int rowNum=0; rowNum<rows.size(); rowNum++) {
				Tr row = rows.get(rowNum);
				List<Td> cells = row.getCells();
				for (int colNum=0; colNum<cells.size(); colNum++) {
					Td cell = cells.get(colNum);
					if (cell.getStyle() != null) {
						Style style = cell.getStyle();
						style.setBorder(border);
						this.createStyles(style, wb);
					}
					this.addImageCell(i, rowNum, colNum, cell);
				}
			}
			CellStyle commonCellStyle = wb.createCellStyle();
			commonCellStyle.setWrapText(WRAP_TEXT);
			if (border) {
				this.setBorder(commonCellStyle);
			}
			this.commonCellStyles.put(table, commonCellStyle);
		}
	}

	/**
	 * 保存一份图片信息
	 * @param sheetIndex
	 * @param rowNum
	 * @param colNum
	 * @param cell
	 */
	private void addImageCell(int sheetIndex, int rowNum, int colNum, Td cell) {
		if (cell.getImage() != null) {
			int row1 = rowNum;
			int row2 = row1;
			int col1 = colNum;
			int col2 = col1;
			if (cell.getRowspan() > 1) {
				row2 += cell.getRowspan() - 1;
			}
			if (cell.getColspan() > 1) {
				col2 += cell.getColspan() - 1;
			}
			ImageCell imageCell = new ImageCell(sheetIndex, row1, row2, col1, col2, cell.getImage());
			this.imageCells.add(imageCell);
		}
	}

	/**
	 * 创建样式信息
	 * @param style
	 * @param wb
	 */
	private void createStyles(Style style, XSSFWorkbook wb) {
		if (this.cellStyles.containsKey(style)) {
			return;
		}
		XSSFCellStyle cellStyle = wb.createCellStyle();
		if (style.getFont() != null) {
			Font fontDef = style.getFont();
			XSSFFont font = wb.createFont();
			if (StringUtils.isNotBlank(fontDef.getColor())) {
				font.setColor(colorResolver.resolve(fontDef.getColor()));
			}
			if (StringUtils.isNotBlank(fontDef.getColor())) {
				font.setFontName(fontDef.getColor());
			}
			if (fontDef.getSize() != null && fontDef.getSize() > 0) {
				font.setFontHeightInPoints(fontDef.getSize().shortValue());
			}
			font.setBold(fontDef.isBold());
			cellStyle.setFont(font);
		}
		if (StringUtils.isNotBlank(style.getBackgroundColor())) {
			cellStyle.setFillForegroundColor(colorResolver.resolve(style.getBackgroundColor()));
			cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		}
		if (StringUtils.isNotBlank(style.getAlign())) {
			try {
				cellStyle.setAlignment(HorizontalAlignment.valueOf(style.getAlign().toUpperCase()));
			} catch (Exception e) {
				log.warn("不支持的横向对齐方式[{}]", style.getAlign());
			}
		}
		if (StringUtils.isNotBlank(style.getVerticalAlign())) {
			if ("middle".equalsIgnoreCase(style.getVerticalAlign())) {
				cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			} else {
				try {
					cellStyle.setVerticalAlignment(VerticalAlignment.valueOf(style.getVerticalAlign()));
				} catch (Exception e) {
					log.warn("不支持的纵向对齐方式[{}]", style.getVerticalAlign());
				}
			}
		}
		cellStyle.setWrapText(WRAP_TEXT);
		if (style.isBorder()) {
			this.setBorder(cellStyle);
		}
		this.cellStyles.put(style, cellStyle);
	}
	
	/**
	 * 设置边框信息
	 * @param cellStyle
	 */
	private void setBorder(CellStyle cellStyle) {
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
	}
	
	/**
	 * 把输入流中的内容写入到输出流中
	 * @param in
	 * @param out
	 * @throws IOException
	 */
	private void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] chunk = new byte[1024];
		int count;
		while ((count = in.read(chunk)) >= 0) {
			out.write(chunk, 0, count);
		}
	}
	
	/**
	 * 用来写单个sheet数据的
	 * @author Elim
	 * 2017年9月6日
	 */
	@RequiredArgsConstructor
	private class SpreadsheetWriter {
		/**
		 * 默认字体时一行显示的字符数
		 */
		private static final int DEFAULT_LINE_CHARS = 8;
		/**
		 * 默认行高
		 */
		private static final float DEFAULT_ROW_HEIGHT = 13.5F;
		/**
		 * 最大行高
		 */
		private static final float MAX_ROW_HEIGHT = 409F;
		/**
		 * 默认的字体大小
		 */
		private static final int DEFAULT_FONT_SIZE = 11;
		/**
		 * 字体的最大值
		 */
		private static final int MAX_FONT_SIZE = 72;
		/**
		 * 最后一个单元格
		 */
		private String lastCell;
		/**
		 * 记录列宽
		 */
		private Map<Integer, Integer> columnWidths = new HashMap<>();
		/**
		 * 记录行高
		 */
		private Map<Integer, Float> rowHeights = new HashMap<>();
		/**
		 * 合并区域
		 */
		private List<MergeArea> mergeAreas = new ArrayList<>();
		/**
		 * 记录列的合并信息，Key是列索引，Value是剩余的需要合并的行
		 */
		private Map<Integer, Integer> colRowSpans = new HashMap<>();
		
		private final Table table;
		private final StringWriter out = new StringWriter();
		private int commonStyleIndex;
		
		/**
		 * 构建表单数据
		 * @return
		 */
		public String buildData() {
			this.commonStyleIndex = ExcelGenerator.this.commonCellStyles.get(this.table).getIndex();
			this.calculateColumnWidths();
			this.calculateRowHeights();
			this.beginSheet();
			int rowNum = 0;
			for (Tr row : this.table.getRows()) {
				this.insertRow(rowNum++, row);
			}
			this.endSheet();
			return this.out.toString();
		}
		
		private void beginSheet() {
            out.write("<?xml version=\"1.0\" encoding=\""+XML_ENCODING+"\"?>" +
                    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" );
            this.insertColumnWidth();
            out.write("<sheetData>\n");
        }

		private void endSheet() {
			out.write("</sheetData>");
			if (this.table.isAutoFilter()) {
				out.write("<autoFilter ref=\"A1:");
				out.write(this.lastCell);
				out.write("\"/>");
			}
			if (!this.mergeAreas.isEmpty()) {
				out.write("<mergeCells count=\"" + this.mergeAreas.size() + "\">");
				for (MergeArea area : this.mergeAreas) {
					out.write("<mergeCell ref=\"" + area.getAreaInfo() + "\"/>");
				}
				out.write("</mergeCells>");
			}
			out.write("</worksheet>");
        }
		
		/**
		 * 插入列宽信息
		 */
		private void insertColumnWidth() {
			if (this.columnWidths.isEmpty()) {
				return;
			}
			out.write("<cols>");
			for (Map.Entry<Integer, Integer> entry : this.columnWidths.entrySet()) {
				Integer column = entry.getKey() + 1;
				Integer width = entry.getValue();
				out.write("<col min=\"" + column + "\" max=\"" + column + "\" width=\"" + width + "\" customWidth=\"1\"/>");
			}
			out.write("</cols>");
		}
		
		/**
		 * 插入行
		 * @param rowNum
		 * @param row
		 */
		private void insertRow(int rowNum, Tr row) {
			out.write("<row r=\"" + (rowNum + 1) + "\"");
			float rowHeight = 0F;
			if (row.getStyle() != null) {
				if (StringUtils.isNotBlank(row.getStyle().getHeight())) {
					rowHeight = Utils.getPxNumber(row.getStyle().getHeight());
				}
			}
			if (this.rowHeights.containsKey(rowNum)) {
				rowHeight = Math.max(rowHeight, this.rowHeights.get(rowNum));
			}
			if (rowHeight > 0) {
				if (rowHeight > MAX_ROW_HEIGHT) {
					rowHeight = MAX_ROW_HEIGHT;
				}
				out.write(" ht=\"" + rowHeight + "\" customHeight=\"1\"");
			}
			out.write(">\n");
			int columnIndex = 0;
			for (Td cell : row.getCells()) {
				columnIndex = this.handleRowspanRemain(rowNum, columnIndex);
				this.prepareMergeArea(rowNum, columnIndex, cell);
				int styleIndex;
				if (cell.getStyle() != null) {
					styleIndex = ExcelGenerator.this.cellStyles.get(cell.getStyle()).getIndex();
				} else {
					styleIndex = this.commonStyleIndex;
				}
				switch (cell.getDataType()) {
				case NUMBER:
					this.lastCell = this.createCell(rowNum, columnIndex, Double.parseDouble(cell.getValue()), styleIndex);
					break;
				default:
					this.lastCell = this.createCell(rowNum, columnIndex, cell.getValue(), styleIndex);
				}
				if (cell.getColspan() > 1) {//需要合并列
					for (int i=1; i<cell.getColspan(); i++) {
						this.lastCell = this.createCell(rowNum, columnIndex + i, "", this.commonStyleIndex);
					}
					columnIndex += cell.getColspan() - 1;
				}
				columnIndex++;
			}
			this.handleRowspanRemain(rowNum, columnIndex);
			out.write("</row>\n");
		}

		/**
		 * 处理合并区间
		 * @param rowNum
		 * @param columnIndex
		 * @param cell
		 */
		private void prepareMergeArea(int rowNum, int columnIndex, Td cell) {
			if (cell.getRowspan() > 1) {
				this.colRowSpans.put(columnIndex, cell.getRowspan() - 1);
				int lastCol = columnIndex;
				if (cell.getColspan() > 1) {
					lastCol += cell.getColspan() - 1;
					for (int i=1; i<cell.getColspan(); i++) {
						this.colRowSpans.put(columnIndex + 1, cell.getRowspan() - 1);
					}
				}
				this.mergeAreas.add(new MergeArea(rowNum, rowNum + cell.getRowspan() -1, columnIndex, lastCol));
			} else if (cell.getColspan() > 1) {
				this.mergeAreas.add(new MergeArea(rowNum, rowNum, columnIndex, columnIndex + cell.getColspan() - 1));
			}
		}

		/**
		 * 处理剩余的需要合并的行
		 * @param rowNum 行号
		 * @param columnIndex 当前列索引
		 * @return 处理了合并的列后下一个单元格的列索引
		 */
		private int handleRowspanRemain(int rowNum, int columnIndex) {
			while (this.colRowSpans.containsKey(columnIndex)) {
				Integer remainRowspan = this.colRowSpans.get(columnIndex);
				remainRowspan--;
				if (remainRowspan == 0) {
					this.colRowSpans.remove(columnIndex);
				} else {
					this.colRowSpans.put(columnIndex, remainRowspan);
				}
				this.createCell(rowNum, columnIndex++, "", this.commonStyleIndex);
			}
			return columnIndex;
		}
		
		/**
		 * 创建文本内容的单元格
		 * @param rowNum
		 * @param colNum
		 * @param value
		 * @param styleIndex
		 * @return 
		 */
		private String createCell(int rowNum, int colNum, String value, int styleIndex) {
			if (value == null) {
				value = "";
			}
			String ref = new CellReference(rowNum, colNum).formatAsString();
			out.write("<c r=\""+ref+"\" t=\"s\"");
            out.write(" s=\""+styleIndex+"\"");
            out.write(">");
            out.write("<is><t>"+value+"</t></is>");
            out.write("</c>");
            return ref;
		}
		
		/**
		 * 创建数字内容的单元格
		 * @param rowNum
		 * @param colNum
		 * @param value
		 * @param styleIndex
		 * @return 
		 */
		private String createCell(int rowNum, int colNum, double value, int styleIndex) {
			String ref = new CellReference(rowNum, colNum).formatAsString();
			out.write("<c r=\""+ref+"\" t=\"n\"");
            out.write(" s=\""+styleIndex+"\"");
            out.write(">");
            out.write("<v>"+value+"</v>");
            out.write("</c>");
            return ref;
		}
		
		/**
		 * 计算列宽
		 */
		private void calculateColumnWidths() {
			this.table.getRows().forEach(row -> {
				for (int colNum=0; colNum<row.getCells().size(); colNum++) {
					Td cell = row.getCells().get(colNum);
					if (cell.getStyle() == null || StringUtils.isBlank(cell.getStyle().getWidth())) {
						continue;
					}
					Float width = Utils.getPxNumber(cell.getStyle().getWidth());
					this.columnWidths.compute(colNum, (k, v) -> {
						int words = Float.valueOf((width-5)/7).intValue();
						if (v == null) {
							return words;
						} else {
							return Math.max(words, v);
						}
					});
				}
			});
		}
		
		/**
		 * 根据文本内容计算行高
		 */
		private void calculateRowHeights() {
			List<Tr> rows = this.table.getRows();
			for (int rowNum=0; rowNum<rows.size(); rowNum++) {
				Tr row = rows.get(rowNum);
				List<Td> cells = row.getCells();
				float rowHeight = 0F;//计算文本的高度
				float maxImageHeight = 0F;//计算图片指定的高度
				for (int colNum=0; colNum<cells.size(); colNum++) {
					Td cell = cells.get(colNum);
					if (cell.getImage() != null) {
						if (StringUtils.isNotBlank(cell.getImage().getHeight())) {
							Float imageHeight = Utils.getPxNumber(cell.getImage().getHeight());
							maxImageHeight = Math.max(imageHeight, maxImageHeight);
						}
						continue;
					}
					if (!cell.isAutoHeight() || StringUtils.isBlank(cell.getValue())) {
						continue;
					}
					String cellValue = cell.getValue();
					int fontSize = this.getFontSize(cell);
					int chars = (cellValue.length() + cellValue.getBytes().length)/2;
					int lineChars = this.columnWidths.getOrDefault(colNum, DEFAULT_LINE_CHARS);
					if (fontSize > DEFAULT_FONT_SIZE) {
						if (fontSize <= 40) {
							lineChars = ((Double)(lineChars * (new Double(DEFAULT_FONT_SIZE)/fontSize))).intValue();
						} else {
							lineChars = lineChars/DEFAULT_LINE_CHARS;//字体大小超过了40后，默认宽度只能容下一个字符了。
						}
					}
					if (chars > lineChars) {
						int currentLine = chars%lineChars == 0 ? chars/lineChars : (chars/lineChars + 1);
						float currentHeight = fontSize * DEFAULT_ROW_HEIGHT / DEFAULT_FONT_SIZE * currentLine;
						rowHeight = Math.max(rowHeight, currentHeight);
					}
				}
				if (rowHeight > DEFAULT_ROW_HEIGHT || maxImageHeight > 0) {
					if (maxImageHeight > 0) {
						rowHeight = Math.max(rowHeight, maxImageHeight);
					}
					this.rowHeights.put(rowNum, rowHeight);
				}
			}
		}

		/**
		 * 获取单元格的字体尺寸
		 * @param cell
		 * @return
		 */
		private int getFontSize(Td cell) {
			int fontSize = DEFAULT_FONT_SIZE;
			if (cell.getStyle() != null && cell.getStyle().getFont() != null) {
				Integer size = cell.getStyle().getFont().getSize();
				if (size != null && size > 0) {
					if (size > MAX_FONT_SIZE) {
						size = MAX_FONT_SIZE;
					}
					fontSize = size;
				}
			}
			return fontSize;
		}
		
	}
	
	/**
	 * 包含图片的单元格信息
	 * 
	 * @author Elim 2017年9月6日
	 */
	@Data
	private static class ImageCell {
		private final int sheet;
		private final int row1;
		private final int row2;
		private final int col1;
		private final int col2;
		private final Img image;
	}
	
	/**
	 * 合并范围
	 * @author Elim
	 * 2017年9月7日
	 */
	@RequiredArgsConstructor
	private static class MergeArea {
		private final int row1;
		private final int row2;
		private final int col1;
		private final int col2;
		
		/**
		 * 获取合并的范围的字符串表示
		 * @return
		 */
		public String getAreaInfo() {
			String ref1 = new CellReference(row1, col1).formatAsString();
			String ref2 = new CellReference(row2, col2).formatAsString();
			return ref1 + ":" + ref2;
		}
	}

}
