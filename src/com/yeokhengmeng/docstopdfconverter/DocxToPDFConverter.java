package com.yeokhengmeng.docstopdfconverter;

import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.ITextFontRegistry;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.io.*;

public class DocxToPDFConverter extends Converter {

    public static void main(String[] args) {
          try {
            FileInputStream fileInputStream = new FileInputStream("F:\\部署文档.docx");
            FileOutputStream fileOutputStream = new FileOutputStream("F:\\部署文档_docx.pdf");
            DocxToPDFConverter docxToPDFConverter = new DocxToPDFConverter(fileInputStream,fileOutputStream,true,true);
            docxToPDFConverter.convert();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (Exception e){
            e.printStackTrace();
        }
    }

    public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }

    @Override
    public void convert() throws Exception {
        loading();
        XWPFDocument document = new XWPFDocument(inStream);
        PdfOptions options = PdfOptions.create();
//		options.fontEncoding("宋体");
		/*options.fontProvider(new IFontProvider() {

			@Override
			public Font getFont(String familyName, String encoding, float size, int style, Color color) {
				try {
					BaseFont bfChinese = BaseFont.createFont("font/kaiu.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
					Font fontChinese = new Font(bfChinese, 12, style, color);
					if (familyName != null)
						fontChinese.setFamily(familyName);
					return fontChinese;
				} catch (Exception e) {
					e.printStackTrace();
					return null;
				}
			}
		});*/

        options.fontProvider(new ITextFontRegistry() {
            public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                try {
                    BaseFont bfChinese = BaseFont.createFont("C:\\Windows\\fonts\\simfang.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
//					BaseFont bfChinese = BaseFont.createFont("C:\\Windows\\fonts\\simsun.ttc", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font fontChinese = new Font(bfChinese, size, style, color);
                    if (familyName != null)
                        fontChinese.setFamily(familyName);
                    return fontChinese;
                } catch (Throwable e) {
                    e.printStackTrace();
                    // An error occurs, use the default font provider.
                    return ITextFontRegistry.getRegistry().getFont(familyName, encoding, size, style, color);
                }
            }
        });
        processing();
        PdfConverter.getInstance().convert(document, outStream, options);
        finished();
    }


}
