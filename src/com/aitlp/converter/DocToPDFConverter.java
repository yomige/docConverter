package com.aitlp.converter;

import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;


public class DocToPDFConverter extends Converter {


    public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages, boolean closeStreamsWhenComplete) {
        super(inStream, outStream, showMessages, closeStreamsWhenComplete);
    }


    @Override
    public void convert() throws Exception {

        loading();

        InputStream iStream = inStream;


        WordprocessingMLPackage wordMLPackage = getMLPackage(iStream);
        Mapper fontMapper = new IdentityPlusMapper();
//		fontMapper.put("Times-Bold", PhysicalFonts.get("Microsoft Yahei"));
        //中文字体转换
        String fontFamily = "SimSun";

//		URL simsunUrl = this.getClass().getResource("C:\\Windows\\fonts\\simfang.ttf"); //加载字体文件（解决linux环境下无中文字体问题）
//        PhysicalFonts.addPhysicalFont(fontFamily, new URL("file://C:\\Windows\\fonts\\STSONG.TTF"));
        PhysicalFont simsunFont = PhysicalFonts.get(fontFamily);
        fontMapper.put(fontFamily, simsunFont);
        wordMLPackage.setFontMapper(fontMapper, true);

        processing();
        Docx4J.toPDF(wordMLPackage, outStream);

        finished();

    }

    protected WordprocessingMLPackage getMLPackage(InputStream iStream) throws Exception {
        PrintStream originalStdout = System.out;

        //Disable stdout temporarily as Doc convert produces alot of output
        System.setOut(new PrintStream(new OutputStream() {
            public void write(int b) {
                //DO NOTHING
            }
        }));

//		WordprocessingMLPackage mlPackage = Doc.convert(iStream);
        WordprocessingMLPackage mlPackage = null;
//        WordprocessingMLPackage mlPackage = Doc.convert(iStream);

        System.setOut(originalStdout);
        return mlPackage;
    }

}
