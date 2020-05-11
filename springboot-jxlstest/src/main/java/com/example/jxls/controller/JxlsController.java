package com.example.jxls.controller;

import com.example.jxls.bean.Person;
import com.example.jxls.view.AutoColumnWidthCommand;
import com.example.jxls.view.JxlsExcelView;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.TransformerFactory;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.*;

@Controller
@EnableAutoConfiguration
public class JxlsController {


    @GetMapping("/export/jxls1")
    public ModelAndView jxls1() {
        
        Map<String, Object> map = new HashMap<String, Object>();
        
     // テストデータの準備
    	List<Person> personList = new ArrayList<Person>();
    	personList.add(new Person("A","M",new BigDecimal(10)));
    	personList.add(new Person("B","F",new BigDecimal(20)));
    	personList.add(new Person("C","M",new BigDecimal(30)));
    	
    	List<List<Object>> dataList = new ArrayList<List<Object>>();
    	for (Person p:personList) {
    		List<Object> tmpList = new ArrayList<Object>();
    		tmpList.add(p.getName());
    		tmpList.add(p.getGender());
    		tmpList.add(p.getAge());
    		dataList.add(tmpList);
    	}
    	map.put("headers", Arrays.asList("Name" , "Gender" , "Age"));
        map.put("dataList", dataList);
        
        return new ModelAndView(new JxlsExcelView("templates/template01.xlsx", "result1"), map);
        
        
        
    }
    
//    @GetMapping("/export/jxls2")
//    public ModelAndView jxls2() {
//    	
//    	// テストデータの準備
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
//
//    	// テンプレート
//        //InputStream is = TestJxls.class.getClass().getResourceAsStream("/templates/template01.xlsx");
//        FileInputStream is = new FileInputStream(new File("src/main/resources/templates/template01.xlsx"));
//        Context context = new Context();
//        context.putVar("headers", Arrays.asList("Name" , "Gender" , "Age"));
//        context.putVar("dataList", dataList);
//        OutputStream os = new FileOutputStream(new File("C:\\temp\\result.xlsx"));
//        // Excelファイルの生成
//        
//     // static method call:
//        
//        
//        
//        Transformer transformer = TransformerFactory.createTransformer(is, os);
//        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
//        XlsCommentAreaBuilder.addCommandMapping("autoColumnWidth", AutoColumnWidthCommand.class);
//        List<Area> xlsAreaList = areaBuilder.build();
//        Area xlsArea = xlsAreaList.get(0);
////        xlsArea.applyAt(new CellRef("Template!A1"), context);        
//        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
//        transformer.write();
//        
//        
////        XlsCommentAreaBuilder.addCommandMapping("autoColumnWidth", AutoColumnWidthCommand.class);
////        JxlsHelper.getInstance().processTemplate(is, os, context);
//        os.flush();
//        os.close();
//        is.close();
//    }

}
