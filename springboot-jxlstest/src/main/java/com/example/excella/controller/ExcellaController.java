package com.example.excella.controller;

import com.example.excella.view.ExcellaExcelView;
import com.example.excella.view.OutputStreamExporter;
import com.example.jxls.bean.Person;

import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.ExcelOutputStreamExporter;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.bbreak.excella.reports.tag.BlockColRepeatParamParser;
import org.bbreak.excella.reports.tag.BlockRowRepeatParamParser;
import org.bbreak.excella.reports.tag.ColRepeatParamParser;
import org.bbreak.excella.reports.tag.RowRepeatParamParser;
import org.bbreak.excella.reports.tag.SingleParamParser;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import java.math.BigDecimal;
import java.util.*;

import javax.servlet.http.HttpServletResponse;

@Controller
@EnableAutoConfiguration
public class ExcellaController {


    @GetMapping("/export/excella1")
    public ModelAndView excella1(HttpServletResponse response) {
//        
//        Map<String, Object> map = new HashMap<String, Object>();
//        
//     // テストデータの準備
//    	List<Person> personList = new ArrayList<Person>();
//    	personList.add(new Person("A","M",new BigDecimal(10)));
//    	personList.add(new Person("B","F",new BigDecimal(20)));
//    	personList.add(new Person("C","M",new BigDecimal(30)));
//    	personList.add(new Person("D","M",new BigDecimal(40)));
//    	personList.add(new Person("E","M",new BigDecimal(50)));
//    	
//    	List<List<Object>> dataList = new ArrayList<List<Object>>();
//    	for (Person p:personList) {
//    		List<Object> tmpList = new ArrayList<Object>();
//    		tmpList.add(p.getName());
//    		tmpList.add(p.getGender());
//    		tmpList.add(p.getAge());
//    		dataList.add(tmpList);
//    	}
    	
    	
    	
    	
    	
    	
//    	Map<String, Object> map = new HashMap<String, Object>();
        
       	List<String> headerlist=new ArrayList<String>();
       	
//       	List<List<Object>> dataList = new ArrayList<List<Object>>();
//       	List<Object> tmpDataList = new ArrayList<Object>();
       	
       	List<Object[]> dataList = new ArrayList<Object[]>();
       	List<Object> tmpDataList = new ArrayList<Object>();
       	for(int y=0; y<30;y++) {
       		headerlist.add("コラム"+y);
       		tmpDataList.add("value"+y);
       	}
//       	map.put("headers", headerlist);
       	
       	for(int i=0; i < 2; i++) {
       		dataList.add(tmpDataList.toArray());
       	}
//        map.put("data", dataList.toArray());
           
           
           
    	
    	ReportBook outputBook = new ReportBook("src/main/resources/templates/template02.xlsx", "out", ExcelExporter.FORMAT_TYPE);
        ReportSheet outputSheet = new ReportSheet("report");
//        int data = 1;
//        outputSheet.addParam(SingleParamParser.DEFAULT_TAG, "DATA1", data);

//        outputSheet.addParams(SingleParamParser.DEFAULT_TAG, map);
        
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"headers",headerlist.toArray());
        outputSheet.addParam(RowRepeatParamParser.DEFAULT_TAG,"datas",dataList.toArray());
        
        outputBook.addReportSheet(outputSheet);
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.addReportBookExporter(new OutputStreamExporter(response,"ExcellaTest"));
        try {
          reportProcessor.process(outputBook);
        } catch (Exception e) {
          e.printStackTrace();
        }
        
        
//        // レスポンス設定
//        response.setContentType("application/vnd.ms-excel");
//        // Excelファイルのダウンロード
//        response.setHeader("Content-disposition","attachment; filename=ExcellaTest.xlsx");
//        ReportProcessor reportProcessor = new ReportProcessor();
//        ReportBook outputBook = new ReportBook( "src/main/resources/templates/template02.xlsx", "out", ExcelOutputStreamExporter.FORMAT_TYPE);
//        try {
//        reportProcessor.addReportBookExporter( new ExcelOutputStreamExporter( response.getOutputStream()));
//        reportProcessor.process( outputBook);
//        }catch (Exception e) {
//        	e.printStackTrace();
//        }
        
        
        return null;
        
//        return new ModelAndView(new ExcellaExcelView("templates/template02.xlsx", "result1"), map);
        
        
        
    }

}
