package com.jiudao.poi;

import java.util.ArrayList;
import java.util.List;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class POITest {
	public static void main(String[] args) {
		POITest poiTest = new POITest();
		poiTest.generateXlsFile();
	}
	
	// 使用POI生成xls文件
	public void generateXlsFile() {
		// 单元格字段名
		List<String> cellFields = new ArrayList<String>();
		cellFields.add("姓名");
		cellFields.add("年龄");
		cellFields.add("性别");
		cellFields.add("姓名");
		cellFields.add("年龄");
		cellFields.add("性别");
		
		// 第一步创建workbook
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFCellStyle style = workbook.createCellStyle();
		// 第二步创建sheet
		HSSFSheet hssfSheet = workbook.createSheet("sheet");
		// 第三步创建行row: 添加表头0行
		HSSFRow row = hssfSheet.createRow(0);
		
		// 第四步创建单元格
		HSSFCell cell = row.createCell(0);
		
		for (int i = 0; i < cellFields.size(); i++) {
			cell = row.createCell(i);
			cell.setCellValue(cellFields.get(i));
			cell.setCellStyle(style);
		}
		
		
		// 第五步插入数据
		for (int i = 0; i < 100; i++) {
			// 创建行
			row = hssfSheet.createRow(i+1);
			// 创建单元格并添加数据
			for (int j = 0; j < cellFields.size(); j++) {
				row.createCell(j).setCellValue(cellFields.get(j) + j);
			}
		}
		
		// 将生成的 xls 文件保存到指定路径下
		String filePath = "C:\\Users\\12459\\Desktop\\test.xls";
		try {
			FileOutputStream fileOutputStream = new FileOutputStream(filePath);
			workbook.write(fileOutputStream);
			fileOutputStream.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 
		System.out.println("xls 文件生成成功!");
	}
}