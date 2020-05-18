package com.example.excella.bean;

import java.util.Calendar;
import java.util.Date;

public class Test {

	public static void main(String[] args) {
//		Double a=(double) 7.9909;
//        int b = a.intValue();
//        double c = a-b;
//        System.out.println(a);
//        System.out.println(b);
//        System.out.println(c);
        
        
        Calendar cal = Calendar.getInstance();
        cal.set(Calendar.YEAR, 1900);
//        cal.setTime(new Date("16:00:00"));	
        System.out.println(cal.getTime());

	}

}
