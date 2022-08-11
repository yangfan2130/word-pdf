package com.ruoyi.project.utils;


import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.ruoyi.common.utils.DateUtils;
import com.ruoyi.framework.config.RuoYiConfig;
import com.ruoyi.project.utils.wordZpdf.WordIsPdf;
import freemarker.template.Configuration;
import freemarker.template.DefaultObjectWrapper;
import freemarker.template.Template;
import freemarker.template.TemplateException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

public class FreemarkerUtil {
	private static final Object LOCK = new Object();
	 /**
	  * word文件
	  */
	 public static final int WORD_FILE = 1;
	 /**
	  * excel文件
	  */
	 public static final int EXCEL_FILE = 2;
	 
	 private static Configuration cfg;

	private static FreemarkerUtil ftl ;

	private FreemarkerUtil(String templateFolder) throws IOException {
		cfg = new Configuration();
		cfg.setDirectoryForTemplateLoading(new File(templateFolder));
		cfg.setObjectWrapper(new DefaultObjectWrapper());
		cfg.setClassicCompatible(true);
	}

	private static void check(HttpServletRequest request) {
		if (ftl == null) {
			synchronized (LOCK) {
				try {
					File directory = new File("");// 参数为空
        			String courseFile = directory.getCanonicalPath();
        			String uploadDir = courseFile + "\\" + "src\\main\\resources\\ftl";
					ftl = new FreemarkerUtil(uploadDir);
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	public static void createFile(String templateName, String docFileName, Map<String,Object> rootMap, HttpServletRequest request, HttpServletResponse response, int fileType) throws IOException {
		//设置导出
		response.addHeader("Cache-Control","no-cache");
		response.setCharacterEncoding("UTF-8");
		if( WORD_FILE == fileType){
			response.setContentType("application/vnd.ms-word;charset=UTF-8");
		}else if(EXCEL_FILE == fileType){
			response.setContentType("application/octet-stream;charset=UTF-8");
		}else{
			response.setContentType("application/octet-stream");
		}
		String ua = request.getHeader("user-agent");
		ua = ua == null ? null : ua.toLowerCase();
		if(ua != null && (ua.indexOf("firefox") > 0 || ua.indexOf("safari")>0)){
			try {
				docFileName = new String(docFileName.getBytes(),"ISO8859-1");
				response.addHeader("Content-Disposition","attachment;filename=" + docFileName);
			} catch (Exception e) {
			}
		}else{
			try {
				docFileName = URLEncoder.encode(docFileName, "utf-8");
				response.addHeader("Content-Disposition","attachment;filename=" + docFileName);
			} catch (Exception e) {
			}
		}
		check(request);
		//解析模版
		Template temp = cfg.getTemplate(templateName, "UTF-8");


		String filePath = "D:/pdf"+"/";
		String ml = filePath+"word.doc";
		File file = new File(ml);
		if (!file.exists()) {
			if (!file.getParentFile().exists()) {
				file.getParentFile().mkdirs();
			}
		}
		PrintWriter writes = new PrintWriter(file);

		PrintWriter write = response.getWriter();
		try {
//			temp.process(rootMap, write);
			temp.process(rootMap, writes);
		} catch (TemplateException e) {
			e.printStackTrace();
		}finally {
			if(write != null){
				write.flush();
				write.close();
			}
		}
	}

	//word转pdf并预览，（word文件删除，pdf定时删除）
	public static void createFiles(String templateName, String orderCode, Map<String,Object> rootMap, HttpServletRequest request,HttpServletResponse response,int fileType) throws IOException {
		check(request);
		//解析模版
		Template temp = cfg.getTemplate(templateName, "UTF-8");
		//word文件名称
		String filePath = RuoYiConfig.getPdfLj()+"/"+orderCode+"/"+DateUtils.datePath()+"/";
		SimpleDateFormat df=new SimpleDateFormat("yyyyMMddHHmmss");
		String sj = df.format(new Date());
		String wj = filePath+sj+".doc";
		File file = new File(wj);
		if (!file.exists()) {
			if (!file.getParentFile().exists()) {
				file.getParentFile().mkdirs();
			}
		}
		PrintWriter writes = new PrintWriter(file);
		try {
			//生成word
			long startTime = System.currentTimeMillis();
			temp.process(rootMap, writes);
			long endTime = System.currentTimeMillis(); // 获取结束时间
			System.out.println("程序运行时间： " + (endTime - startTime) + "ms");

			//转pdf
			long startTimes = System.currentTimeMillis();
			String pdf = filePath+sj+".pdf";
			WordIsPdf.docToPdf(wj,pdf);
			long endTimes = System.currentTimeMillis(); // 获取结束时间
			System.out.println("程序运行时间： " + (endTimes - startTimes) + "ms");

			//pdf换流显示，并删除word
			FileInputStream in = new FileInputStream(pdf);
			//设置导出
			response.setHeader("Content-Disposition", "inline;filename="
					+ "pz.pdf");
			OutputStream out = response.getOutputStream();
			byte[] b = new byte[1024];
			while ((in.read(b))!=-1) {
				out.write(b);
			}
			out.flush();
			in.close();
			out.close();

		} catch (Exception e) {
			e.printStackTrace();
		}finally {
			if(writes != null){
				writes.flush();
				writes.close();
			}
		}
		//删除word
		file.delete();
	}
}
