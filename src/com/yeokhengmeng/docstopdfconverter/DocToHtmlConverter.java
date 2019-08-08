package com.yeokhengmeng.docstopdfconverter;

import fr.opensagres.poi.xwpf.converter.core.IXWPFConverter;
import fr.opensagres.poi.xwpf.converter.core.ImageManager;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.*;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;

public class DocToHtmlConverter {

    public static void main(String[] args) {
        officeToHtml("F://部署文档.docx");
    }

    /**
     * @param sourceFile 文件完整路径
     * @return
     */
    public static boolean officeToHtml(String sourceFile) {
        String suffix = FilenameUtils.getExtension(sourceFile);
        if (StringUtils.equals(suffix, "doc")) {
            return wordToHtml03(sourceFile);
        } else if (StringUtils.equals(suffix, "docx")) {
            return wordToHtml07(sourceFile);
        } else {
            return false;
        }
    }


    /**
     * 2007版本word转换成html
     *
     * @param sourceFile
     * @return
     */
    private static boolean wordToHtml07(String sourceFile) {
        String filepath = FilenameUtils.getFullPath(sourceFile);
        String fileName = FilenameUtils.getBaseName(sourceFile);
        String htmlFileName = fileName + ".html";
        String htmlFile = filepath + htmlFileName;
        XWPFDocument document;
        OutputStreamWriter outputStreamWriter;
        try {
            document = new XWPFDocument(new FileInputStream(sourceFile));
            //html属性器
            XHTMLOptions options = XHTMLOptions.create();
            //图片处理，第二个参数为html文件同级目录下，否则图片找不到。
            ImageManager imageManager = new ImageManager(new File(filepath), fileName);
            options.setImageManager(imageManager);
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(htmlFile), "utf-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
            // 处理转换后的html charset
            InputStream input = new FileInputStream(new File(htmlFile));
            BufferedReader br = new BufferedReader(new InputStreamReader(input, "utf-8"));
            StringBuilder sb = new StringBuilder();
            String line, htmlStr;
            while ((line = br.readLine()) != null) {
                sb.append(line);
            }
            htmlStr = sb.toString();
            htmlStr = StringUtils.replace(htmlStr, "<head>", "<head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\" />");
            OutputStreamWriter write = new OutputStreamWriter(new FileOutputStream(htmlFile), "utf-8");
            BufferedWriter writer = new BufferedWriter(write);
            writer.write(htmlStr);
            writer.close();
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("word07转html失败");
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("word07转html失败");
            return false;
        } catch (Exception ex) {
            ex.printStackTrace();
            System.out.println("word07转html失败");
            return false;
        }
    }


    /**
     * 2003版本word转换成html
     *
     * @param sourceFile
     * @return
     */
    private static boolean wordToHtml03(String sourceFile) {
        String filepath = FilenameUtils.getFullPath(sourceFile);
        final String fileName = FilenameUtils.getBaseName(sourceFile);
        final String imgPath = filepath + "\\" + fileName + "\\";
        String htmlFileName = fileName + ".html";
        String htmlFile = filepath + htmlFileName;
        InputStream input;
        try {
            input = new FileInputStream(new File(sourceFile));
            HWPFDocument wordDocument = new HWPFDocument(input);
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
            //设置图片存放的位置
            wordToHtmlConverter.setPicturesManager(new PicturesManager() {
                public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches, float heightInches) {
                    File imgFile = new File(imgPath);
                    if (!imgFile.exists()) {
                        //图片目录不存在则创建
                        imgFile.mkdirs();
                    }
                    File imgfile = new File(imgPath + suggestedName);
                    try {
                        OutputStream os = new FileOutputStream(imgfile);
                        os.write(content);
                        os.close();
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    return fileName + "\\" + suggestedName;
                }
            });

            //解析word文档
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            OutputStream outStream = new FileOutputStream(new File(htmlFile));

            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(outStream);

            TransformerFactory factory = TransformerFactory.newInstance();
            Transformer serializer = factory.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            outStream.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        } catch (IOException e1) {
            e1.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        } catch (TransformerConfigurationException e) {
            e.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        } catch (TransformerException e) {
            e.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        } catch (ParserConfigurationException e1) {
            e1.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        } catch (Exception ex) {
            ex.printStackTrace();
            System.out.println("word03转html失败");
            return false;
        }
        return true;
    }
}
