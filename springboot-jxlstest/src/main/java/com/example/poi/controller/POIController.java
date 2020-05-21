package com.example.poi.controller;

import com.example.bean.Person;
import com.example.bean.Person30;
import com.example.poi.view.CustomizeExcelView;

import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

import javax.servlet.http.HttpServletResponse;

@Controller
@EnableAutoConfiguration
public class POIController {


	@GetMapping("/export/poi")
    public ModelAndView poi(HttpServletResponse response,Model model) throws IllegalArgumentException, IllegalAccessException {
    	
    	
    	// テストデータの準備
    	List<Person30> personList = new ArrayList<Person30>();
    	for (int y =0; y<1000;y++) {
    		Person30 p = new Person30();
    		for(Field f:Person30.class.getDeclaredFields()) {
    			f.setAccessible(true);
    			f.set(p, "Vlue"+f.getName()+y);
    		}
	    	personList.add(p);
    	}
    	// データ部分
    	List<List<Object>> dataList = new ArrayList<List<Object>>();
    	for (Person30 p:personList) {
    		List<Object> tmpList = new ArrayList<Object>();
    		for(Field f:Person30.class.getDeclaredFields()) {
    			f.setAccessible(true);
    			tmpList.add(f.get(p));
    		}
    		
    		dataList.add(tmpList);
    	}
    	//　ヘーダ部分
    	List<String> headersList = new ArrayList<String>();
    	for(Field f:Person30.class.getDeclaredFields()) {
			f.setAccessible(true);
			headersList.add(f.getName());
		}
    	
    	// 出力
    	
    	model.addAttribute("header", headersList);
    	model.addAttribute("data", dataList);
    	
        return new ModelAndView(new CustomizeExcelView());
    }
    

}
