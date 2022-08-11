package com.ruoyi.project.utils.wordZpdf;

import java.io.File;
import java.io.FileOutputStream;
import com.aspose.words.Document;
import com.aspose.words.FontSettings;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;
import org.apache.commons.lang3.StringUtils;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.parser.AutoDetectParser;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.sax.BodyContentHandler;

import java.io.*;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordIsPdf {

    public static void main(String[] args) {
        try {
            long startTime = System.currentTimeMillis();
            docToPdf("D:\\pdf\\wpxx.doc",
                    "D:\\pdf\\wpxx.pdf");
            long endTime = System.currentTimeMillis(); // 获取结束时间
            System.out.println("程序运行时间： " + (endTime - startTime) + "ms");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void docToPdf(String inPath, String outPath) throws Exception {
        if (!getLicense()) { // match License if not marks on document
            throw new Exception("license not correct!");
        }
        System.out.println(inPath + " -> " + outPath);
        try {
            File file = new File(outPath);
            FileOutputStream os = new FileOutputStream(file);
            Document doc = new Document(inPath); // doc/docx
            FontSettings.getDefaultInstance().setFontsFolders(new String[]{"/usr/share/fonts/truetype/chinese", "C:\\Windows\\Fonts"}, true);
            doc.save(os, SaveFormat.PDF);
            os.close();  //关闭文件
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    /*
 证书有效性验证
 */
    public static boolean getLicense() {
        boolean result = false;
        try {
            InputStream is = Document.class.getResourceAsStream("/com.aspose.words.lic_2999.xml");
            License asposeLic = new License();
            asposeLic.setLicense(is);
            result = true;
            is.close();
            return result;
        } catch (Exception e) {
            e.printStackTrace();
            return result;
        }
    }
}
