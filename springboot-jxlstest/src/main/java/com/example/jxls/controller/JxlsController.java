package com.example.jxls.controller;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import com.example.jxls.bean.Person;
import com.example.jxls.view.JxlsExcelView;

@Controller
@EnableAutoConfiguration
public class JxlsController {


    @GetMapping("/export/jxls1")
    public ModelAndView jxls1() {

        Map<String, Object> map = new HashMap<String, Object>();
        
     // テストデータの準備
//    	List<Person> personList = new ArrayList<Person>();
    	
//    	map.put("headers", Arrays.asList("Name" , "Gender" , "Age"));
//    	for(int i=0; i < 1000; i++) {
//    		personList.add(new Person("A1A2A3A1A2A3A1A2A3A1A2A3A1A2A3A1A2A3","MAN",new BigDecimal(1000.123)));
//    	}
    	
//    	map.put("headers", Arrays.asList("コラム" , "Gender" , "Age"));
    	List<String> headerlist=new ArrayList<String>();
    	
    	List<List<Object>> dataList = new ArrayList<List<Object>>();
    	List<Object> tmpDataList = new ArrayList<Object>();
    	for(int y=0; y<30;y++) {
    		headerlist.add("コラム"+y);
    		tmpDataList.add("value"+y);
    	}
    	map.put("headers", headerlist);
    	
    	for(int i=0; i < 10000; i++) {
    		dataList.add(tmpDataList);
    	}
    	
//    	List<List<Object>> dataList = new ArrayList<List<Object>>();
//    	for (Person p:personList) {
//    		List<Object> tmpList = new ArrayList<Object>();
//    		tmpList.add(p.getName());
//    		tmpList.add(p.getGender());
//    		tmpList.add(p.getAge());
//    		dataList.add(tmpList);
//    	}
        map.put("dataList", dataList);

        return new ModelAndView(new JxlsExcelView("templates/template01.xlsx", "result1"), map);

    }
}
