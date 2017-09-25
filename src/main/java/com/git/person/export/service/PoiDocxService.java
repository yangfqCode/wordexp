package com.git.person.export.service;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;

import org.apache.log4j.Logger;

import com.git.person.export.entity.ExportResultInfo;
import com.git.person.export.util.CustomXWPFDocument;
import com.git.person.export.util.ProcessWordTemplateUtils;


public class PoiDocxService {
  
  static Logger logger = Logger.getLogger(PoiDocxService.class);

  /**
   * 根据传入的Word模板和数据替换模板中的占位符
   * 
   * @param File file 导出的word模板
   * @param String id 导出单据的id
   * @param Map<String, Object> contentMap 用户自己需要导出到word的contentMap
   * 
   * @return word的Document
   * @throws IOException
   */
  public static <T extends Serializable> ExportResultInfo createDoc(InputStream inputStream, OutputStream outputStream,T t) throws IOException {
    CustomXWPFDocument doc = null;
    ExportResultInfo resultInfo = new ExportResultInfo();
    try {
      doc = new CustomXWPFDocument(inputStream);
      // 填充内容到模板doc中
      BeanWrapper data = new BeanWrapperImpl(t);
      ProcessWordTemplateUtils.processDoc(doc, data);
      doc.write(outputStream);
    } catch (Exception e) {
      resultInfo.setErrors(e.getMessage());
      resultInfo.setSuccess(false);
      logger.error(e.getMessage(),e);
    } finally {
      if (null != inputStream) {
        inputStream.close();
      }
      if (null != doc) {
        doc.close();
      }
    }
    return resultInfo;
  }
}
