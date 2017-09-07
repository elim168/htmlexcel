package com.elim.htmlexcel.core;

import java.awt.Color;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFColor;

import lombok.extern.slf4j.Slf4j;

/**
 * 颜色解析器
 * @author Elim
 * 2017年9月6日
 */
@Slf4j
public class ColorResolver {

    private final static String WHITE = "white";
    private final static String GRAY = "gray";
    private final static String BLACK = "black";
    private final static String RED = "red";
    private final static String PINK = "pink";
    private final static String ORANGE = "orange";
    private final static String YELLOW = "yellow";
    private final static String GREEN = "green";
    private final static String MAGENTA = "magenta";
    private final static String CYAN = "cyan";
    private final static String BLUE = "blue";
 
    private static final Map<String, XSSFColor> COLORS = new HashMap<>(11);
    /**
     * 默认颜色
     */
    private static final XSSFColor DEFAULT = new XSSFColor(Color.BLACK);
    
    static {
    	COLORS.put(WHITE, new XSSFColor(Color.WHITE));
    	COLORS.put(GRAY, new XSSFColor(Color.GRAY));
    	COLORS.put(BLACK, new XSSFColor(Color.BLACK));
    	COLORS.put(RED, new XSSFColor(Color.RED));
    	COLORS.put(PINK, new XSSFColor(Color.PINK));
    	COLORS.put(ORANGE, new XSSFColor(Color.ORANGE));
    	COLORS.put(YELLOW, new XSSFColor(Color.YELLOW));
    	COLORS.put(GREEN, new XSSFColor(Color.GREEN));
    	COLORS.put(MAGENTA, new XSSFColor(Color.MAGENTA));
    	COLORS.put(CYAN, new XSSFColor(Color.CYAN));
    	COLORS.put(BLUE, new XSSFColor(Color.BLUE));
    }
	
    /**
     * 解析颜色
     * @param color
     * @return
     */
    public XSSFColor resolve(String color) {
    	if (color.startsWith("#")) {
    		color = color.substring(1).trim();
    		String redStr = null;
    		String greenStr = null;
    		String blueStr = null;
    		if (color.length() == 3) {
    			redStr = color.substring(0, 1);
    			greenStr = color.substring(1, 2);
    			blueStr = color.substring(2);
    			redStr += redStr;
    			greenStr += greenStr;
    			blueStr += blueStr;
    		} else if (color.length() == 6) {
    			redStr = color.substring(0, 2);
    			greenStr = color.substring(2, 4);
    			blueStr = color.substring(4);
    		} else {
    			log.warn("指定了不规范的颜色[{}]，系统将使用默认颜色", color);
    		}
    		try {
    			int red = Integer.valueOf(redStr, 16);
    			int green = Integer.valueOf(greenStr, 16);
    			int blue = Integer.valueOf(blueStr, 16);
    			return new XSSFColor(new Color(red, green, blue));
    		} catch (Exception e) {
    			log.warn("指定了不规范的颜色[{}]，系统将使用默认颜色", color);
    			return DEFAULT;
    		}
    	} else {
    		return COLORS.getOrDefault(color, DEFAULT);
    	}
    }
    
}
