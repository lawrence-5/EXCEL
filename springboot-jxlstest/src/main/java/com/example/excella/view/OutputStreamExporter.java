package com.example.excella.view;

import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.exporter.ExcelExporter;
import org.bbreak.excella.reports.exporter.ReportBookExporter;
import org.bbreak.excella.reports.model.ConvertConfiguration;

public class OutputStreamExporter extends ReportBookExporter {

    private HttpServletResponse response;
    private String fileName;

    public OutputStreamExporter(HttpServletResponse response,String fileName) {
        this.response = response;
        this.fileName=fileName;
    }

    @Override
    public String getExtention() {
        return null;
    }

    @Override
    public String getFormatType() {
        return ExcelExporter.FORMAT_TYPE;
    }

    @Override
    public void output(Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {

        System.out.println(book.getFirstVisibleTab());
        System.out.println(book.getSheetName(0));

        try {
            response.setContentType("application/vnd.ms-Excel");
//            response.setHeader("Content-Disposition", "attachment; filename=masatoExample.xlsx");
            response.setHeader("Content-Disposition", "attachment; filename="+fileName+".xlsx");
            book.write(response.getOutputStream());
            response.getOutputStream().close();
        }
        catch(Exception e) {
            System.out.println(e);
        }
    }
}//end class