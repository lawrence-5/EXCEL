package com.example.excella.controller;

import com.example.bean.Person;
import com.example.bean.Person30;
import com.example.excella.view.ExcellaExcelView;
import com.example.excella.view.OutputStreamExporter;
import com.example.excella.view.RowColLoopParamParser;

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

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

import javax.servlet.http.HttpServletResponse;

@Controller
@EnableAutoConfiguration
public class ExcellaController {


    @GetMapping("/export/excella1")
    public ModelAndView excella1(HttpServletResponse response) {
    	
    	
        
       	List<String> headerlist=new ArrayList<String>();
       	
       	List<Object> dataList = new ArrayList<Object>();
       	List<Object> tmpDataList1 = new ArrayList<Object>();
       	List<Object> tmpDataList2 = new ArrayList<Object>();
       	for(int y=0; y<30;y++) {
       		headerlist.add("コラム"+y);
//       		headerlist.add("${test}");
       		tmpDataList1.add("value1_"+y);
       		tmpDataList2.add("value2_"+y);
       	}
       	
       	
       	dataList.add("$REMOVE{row}");
       	dataList.add("$C[]{data1}");
   		dataList.add("$C[]{data2}");
       	dataList.add("$C[]{data1}");
   		dataList.add("$C[]{data2}");
           
    	
    	ReportBook outputBook = new ReportBook("src/main/resources/templates/template02.xlsx", "out", ExcelExporter.FORMAT_TYPE);
        ReportSheet outputSheet = new ReportSheet("report");
        
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"headers",headerlist.toArray());
//        outputSheet.addParam(RowColLoopParamParser.DEFAULT_TAG,"datas",dataList.toArray());
//        outputSheet.addParam(SingleParamParser.DEFAULT_TAG, "test", "TEST");
        
        outputSheet.addParam(RowRepeatParamParser.DEFAULT_TAG,"datas",dataList.toArray());
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"data1",tmpDataList1.toArray());
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"data2",tmpDataList2.toArray());
        
        outputBook.addReportSheet(outputSheet);
        ReportProcessor reportProcessor = new ReportProcessor();
        
        // $RCLにRowColLoopParamParserを関連付ける
//        reportProcessor.addReportsTagParser( new RowColLoopParamParser("$RCL"));
        
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
    
    @GetMapping("/export/excella2")
    public ModelAndView excella2(HttpServletResponse response) {
    	
    	ReportBook outputBook = new ReportBook("src/main/resources/templates/template02.xlsx", "", ExcelExporter.FORMAT_TYPE);
    	ReportSheet outputSheet = new ReportSheet("report");
    	
     // テストデータの準備
    	List<Person> personList = new ArrayList<Person>();
    	for (int y =0; y<1000;y++) {
	    	personList.add(new Person("NAME"+y,"GENDER"+y,new BigDecimal(y)));
    	}
    	List<Object> dataList = new ArrayList<Object>();
    	dataList.add("$REMOVE{row}");
       	int i =0;
    	for (Person p:personList) {
    		List<Object> tmpList = new ArrayList<Object>();
    		tmpList.add(p.getName());
    		tmpList.add(p.getGender());
    		tmpList.add(p.getAge());
    		outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"data"+i,tmpList.toArray());
    		dataList.add("$C[]{data"+i+"}");
    		i++;
    	}
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"headers",new String[]{"Name","Gender","Age"});
        outputSheet.addParam(RowRepeatParamParser.DEFAULT_TAG,"datas",dataList.toArray());
        outputBook.addReportSheet(outputSheet);
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.addReportBookExporter(new OutputStreamExporter(response,"ExcellaTest"));
        try {
          reportProcessor.process(outputBook);
        } catch (Exception e) {
          e.printStackTrace();
        }
        return null;
    }
    
    @GetMapping("/export/excella3")
    public ModelAndView excella3(HttpServletResponse response) throws IllegalArgumentException, IllegalAccessException {
    	
    	ReportBook outputBook = new ReportBook("src/main/resources/templates/template02.xlsx", "", ExcelExporter.FORMAT_TYPE);
    	ReportSheet outputSheet = new ReportSheet("report");
    	
    	// テストデータの準備
    	List<Person30> personList = new ArrayList<Person30>();
    	for (int y =0; y<10000;y++) {
    		Person30 p = new Person30();
    		for(Field f:Person30.class.getDeclaredFields()) {
    			f.setAccessible(true);
    			f.set(p, "Vlue"+f.getName()+y);
    		}
	    	personList.add(p);
    	}
    	List<Object> dataList = new ArrayList<Object>();
    	dataList.add("$REMOVE{row}");
       	int i =0;
    	for (Person30 p:personList) {
    		List<Object> tmpList = new ArrayList<Object>();
    		for(Field f:Person30.class.getDeclaredFields()) {
    			f.setAccessible(true);
    			tmpList.add(f.get(p));
    		}
    		
    		outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"data"+i,tmpList.toArray());
    		dataList.add("$C[]{data"+i+"}");
    		i++;
    	}
    	
    	List<String> headersList = new ArrayList<String>();
    	for(Field f:Person30.class.getDeclaredFields()) {
			f.setAccessible(true);
			headersList.add(f.getName());
		}
    	
        outputSheet.addParam(ColRepeatParamParser.DEFAULT_TAG,"headers",headersList.toArray());
        outputSheet.addParam(RowRepeatParamParser.DEFAULT_TAG,"datas",dataList.toArray());
        outputBook.addReportSheet(outputSheet);
        ReportProcessor reportProcessor = new ReportProcessor();
        reportProcessor.addReportBookExporter(new OutputStreamExporter(response,"ExcellaTest"));
        try {
          reportProcessor.process(outputBook);
        } catch (Exception e) {
          e.printStackTrace();
        }
        return null;
    }

}
