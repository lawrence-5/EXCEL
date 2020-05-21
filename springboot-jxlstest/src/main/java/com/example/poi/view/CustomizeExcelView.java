package com.example.poi.view;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.servlet.view.document.AbstractXlsxView;

public class CustomizeExcelView extends AbstractXlsxView {

	@Override
	protected void buildExcelDocument(Map<String, Object> model, Workbook workbook, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		Sheet sheet = workbook.createSheet("report");
		
		@SuppressWarnings("unchecked")
		List<List<Object>> dataList=(List<List<Object>>) model.get("data");
		@SuppressWarnings("unchecked")
		List<String> headersList=(List<String>) model.get("header");
		// header
		Row row = sheet.createRow(0);
		Cell cell=null;
		int i=0;
		for (String title:headersList) {
			cell = row.createCell(i);
			cell.setCellValue(title);
			i++;
		}
		int y=1;
		Row row1=null;
		for(List<Object> list:dataList) {
			row1 = sheet.createRow(y);
			int z=0;
			for (Object obj:list) {
				row1.createCell(z).setCellValue(obj.toString());;
				z++;
			}
			y++;
		}
	}

}
