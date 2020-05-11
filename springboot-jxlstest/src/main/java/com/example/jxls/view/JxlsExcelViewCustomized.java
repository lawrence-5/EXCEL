package com.example.jxls.view;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.jxls.area.Area;
import org.jxls.builder.AreaBuilder;
import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.CellRef;
import org.jxls.common.Context;
import org.jxls.transform.Transformer;
import org.jxls.util.JxlsHelper;
import org.jxls.util.TransformerFactory;
import org.springframework.web.servlet.view.AbstractView;

public class JxlsExcelViewCustomized  extends AbstractView{
	private static final String CONTENT_TYPE = "application/vnd.ms-excel";
	 
    private String templatePath;
    private String exportFileName;
    
	public JxlsExcelViewCustomized(String templatePath, String exportFileName) {
        this.templatePath = templatePath;
        if (exportFileName != null) {
            try {
                exportFileName = URLEncoder.encode(exportFileName, "UTF-8");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
        }
        this.exportFileName = exportFileName;
        setContentType(CONTENT_TYPE);
		
	}

	@Override
	protected void renderMergedOutputModel(Map<String, Object> model, HttpServletRequest request,
			HttpServletResponse response) throws Exception {
		Context context = new Context(model);
        response.setContentType(getContentType());
        response.setHeader("content-disposition",
                "attachment;filename=" + exportFileName + ".xls");
        ServletOutputStream os = response.getOutputStream();
        InputStream is = getClass().getClassLoader().getResourceAsStream(templatePath);
//        JxlsHelper.getInstance().processTemplate(is, os, context);
//        is.close();
        
        Transformer transformer = TransformerFactory.createTransformer(is, os);
        AreaBuilder areaBuilder = new XlsCommentAreaBuilder(transformer);
        XlsCommentAreaBuilder.addCommandMapping("autoColumnWidth", AutoColumnWidthCommand.class);
        List<Area> xlsAreaList = areaBuilder.build();
        Area xlsArea = xlsAreaList.get(0);
//        xlsArea.applyAt(new CellRef("Template!A1"), context);        
        xlsArea.applyAt(new CellRef("Sheet1!A1"), context);
        transformer.write();
        
        
//        XlsCommentAreaBuilder.addCommandMapping("autoColumnWidth", AutoColumnWidthCommand.class);
//        JxlsHelper.getInstance().processTemplate(is, os, context);
        os.flush();
        os.close();
        is.close();
		
	}

}
