package com.git.person.export.service;

import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.Enumeration;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.log4j.Logger;
import org.springframework.stereotype.Component;

import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import freemarker.template.Version;


/**
 * 
 * @description 用于动态生成word文档
 * @description 只能用来生成docx格式的word模板
 * 
 * @author Administration
 * 
 */
@Component
public class FreemarkerDocxService {

  static Logger logger = Logger.getLogger(FreemarkerDocxService.class);

  private Configuration configuration = null;

  /**
   * 构造函数
   * 
   * @param zipFile docx格式的Word模板
   */

  public FreemarkerDocxService() {
    configuration = new Configuration(new Version(2, 3, 25));
    configuration.setDefaultEncoding("utf-8");
    configuration.setClassicCompatible(true);
  }

  /**
   * 
   * @param xmlInputStream
   * @param docxInputStream
   * @param outputStream
   * @param contentMap
   * @throws Exception
   */
  public void createDocx(InputStream xmlInputStream, OutputStream outputStream,ZipFile zipFile, 
      Map<String, Object> contentMap) throws Exception {

    // 第一步：把map中的数据动态由freemarker传给xml模板
    InputStreamReader isr = new InputStreamReader(xmlInputStream);
    Template t = new Template("name", isr, configuration); // 获取模板文件
    Writer out = null;

    ByteArrayOutputStream tempoutputStream = new ByteArrayOutputStream();
    tempoutputStream.writeTo(outputStream);
    try {

      // 填充完数据到临时xml
      out = new BufferedWriter(new OutputStreamWriter(tempoutputStream));
      t.process(contentMap, out);

      // 第二步：把临时xml写入到docx中
      outDocx(zipFile, tempoutputStream, outputStream);

    } catch (TemplateException e) {
      logger.error("ftl文档转换为doc字节流出错", e);
    } finally {
      if (null != out) {
        out.flush();
        out.close();
      }
    }
  }

  /**
   * 
   * @param xmlTempOut
   * @param outputStream
   * @throws Exception
   */
  private void outDocx(ZipFile zipFile, ByteArrayOutputStream tempoutputStream,
      OutputStream outputStream) throws Exception {
    ZipOutputStream zipout = null;
    try {
      Enumeration<? extends ZipEntry> zipEntrys = zipFile.entries();
      zipout = new ZipOutputStream(outputStream);
      int len = -1;
      byte[] buffer = new byte[1024];
      while (zipEntrys.hasMoreElements()) {
        ZipEntry next = zipEntrys.nextElement();
        InputStream is = zipFile.getInputStream(next);
        // 把输入流的文件传到输出流中 如果是word/document.xml由我们输入
        zipout.putNextEntry(new ZipEntry(next.toString()));
        if ("word/document.xml".equals(next.toString())) {
          InputStream in = new ByteArrayInputStream(tempoutputStream.toByteArray());
          while ((len = in.read(buffer)) != -1) {
            zipout.write(buffer, 0, len);
          }
          in.close();
        } else {
          while ((len = is.read(buffer)) != -1) {
            zipout.write(buffer, 0, len);
          }
          is.close();
        }
      }
    } catch (FileNotFoundException e) {
      e.printStackTrace();
    } finally {
      zipout.close();
      zipFile.close();
    }
  }
}
