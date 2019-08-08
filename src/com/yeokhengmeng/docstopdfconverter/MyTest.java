package com.yeokhengmeng.docstopdfconverter;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;

public class MyTest {
    public static void main(String[] args) {
        // method 2

        /*try {
            String inputFile = "F:\\部署文档.docx";
            String outputFile = "F:\\部署文档.docx.pdf";
            if (args != null && args.length == 2) {
                inputFile = args[0];
                outputFile = args[1];
            }
            System.out.println("inputFile:" + inputFile + ",outputFile:" + outputFile);
            FileInputStream in = new FileInputStream(inputFile);
            XWPFDocument document = new XWPFDocument(in);
            File outFile = new File(outputFile);
            OutputStream out = new FileOutputStream(outFile);
            PdfOptions options = null;
            PdfConverter.getInstance().convert(document, out, options);
        } catch (Exception e) {
            e.printStackTrace();
        }*/

        // method 1 doc to pdf
        /* try {
            FileInputStream fileInputStream = new FileInputStream("F:\\部署文档.doc");
            FileOutputStream fileOutputStream = new FileOutputStream("F:\\部署文档_doc.pdf");
             DocToPDFConverter docToPDFConverter = new DocToPDFConverter(fileInputStream,fileOutputStream,true,true);
             docToPDFConverter.convert();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }*/



    }

    public void test(){

    }
}
