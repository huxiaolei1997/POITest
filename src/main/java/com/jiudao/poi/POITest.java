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
	
	// ʹ��POI����xls�ļ�
	public void generateXlsFile() {
		// ��Ԫ���ֶ���
		List<String> cellFields = new ArrayList<String>();
		cellFields.add("����");
		cellFields.add("����");
		cellFields.add("�Ա�");
		cellFields.add("����");
		cellFields.add("����");
		cellFields.add("�Ա�");
		
		// ��һ������workbook
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFCellStyle style = workbook.createCellStyle();
		// �ڶ�������sheet
		HSSFSheet hssfSheet = workbook.createSheet("sheet");
		// ������������row: ��ӱ�ͷ0��
		HSSFRow row = hssfSheet.createRow(0);
		
		// ���Ĳ�������Ԫ��
		HSSFCell cell = row.createCell(0);
		
		for (int i = 0; i < cellFields.size(); i++) {
			cell = row.createCell(i);
			cell.setCellValue(cellFields.get(i));
			cell.setCellStyle(style);
		}
		
		
		// ���岽��������
		for (int i = 0; i < 100; i++) {
			// ������
			row = hssfSheet.createRow(i+1);
			// ������Ԫ���������
			for (int j = 0; j < cellFields.size(); j++) {
				row.createCell(j).setCellValue(cellFields.get(j) + j);
			}
		}
		
		// �����ɵ� xls �ļ����浽ָ��·����
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
		System.out.println("xls �ļ����ɳɹ�!");
	}
}