package com.git.person.export.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


/**
 * Word文档（格式为docx）转换为pdf
 */
public class DocxToPDFConverter extends Converter {

  public DocxToPDFConverter(XWPFDocument document, OutputStream outStream) throws IOException {
    super(document, outStream);
  }

  @Override
  public void convert() throws Exception {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    document.write(buffer);
    byte[] bytes = buffer.toByteArray();
    InputStream inputStream = new ByteArrayInputStream(bytes);
    
    WordprocessingMLPackage wordMLPackage = getMLPackage(inputStream);
    Docx4J.toPDF(wordMLPackage, outStream);
    
//    PdfOptions options = PdfOptions.create().fontEncoding("Unicode");
//    PdfConverter.getInstance().convert(document, outStream, options);
    
    finished();
  }

  private static WordprocessingMLPackage getMLPackage(InputStream is) throws Exception {
    WordprocessingMLPackage mlPackage = WordprocessingMLPackage.load(is);
    
    Mapper fontMapper = new IdentityPlusMapper();
    fontMapper.put("华文行楷", PhysicalFonts.get("STXingkai"));
    fontMapper.put("华文仿宋", PhysicalFonts.get("STFangsong"));
    fontMapper.put("隶书", PhysicalFonts.get("LiSu"));
    fontMapper.put("宋体", PhysicalFonts.get("SimSun"));
    fontMapper.put("微软雅黑", PhysicalFonts.get("Microsoft Yahei"));
    fontMapper.put("黑体", PhysicalFonts.get("SimHei"));
    fontMapper.put("楷体", PhysicalFonts.get("KaiTi"));
    fontMapper.put("新宋体", PhysicalFonts.get("NSimSun"));
    fontMapper.put("宋体扩展", PhysicalFonts.get("simsun-extB"));
    fontMapper.put("仿宋", PhysicalFonts.get("FangSong"));
    fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
    fontMapper.put("幼圆", PhysicalFonts.get("YouYuan"));
    fontMapper.put("华文宋体", PhysicalFonts.get("STSong"));
    fontMapper.put("华文中宋", PhysicalFonts.get("STZhongsong"));
    mlPackage.setFontMapper(fontMapper);
    return mlPackage;
  }

}
