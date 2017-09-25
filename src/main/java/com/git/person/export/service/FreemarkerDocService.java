package com.git.person.export.service;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;

import org.apache.log4j.Logger;
import org.springframework.beans.BeanWrapper;
import org.springframework.beans.BeanWrapperImpl;
import org.springframework.stereotype.Component;

import com.git.person.export.entity.ExportResultInfo;

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
public class FreemarkerDocService {

  static Logger logger = Logger.getLogger(FreemarkerDocService.class);

  private Configuration configuration = null;

  /**
   * 构造函数
   */
  public FreemarkerDocService() {
    configuration = new Configuration(new Version(2, 3, 25));
    configuration.setDefaultEncoding("utf-8");
    configuration.setClassicCompatible(true);
  }

  /**
   * @param inputStream Word模板输入流
   * @param outputStream 输出流
   * @param contentMap 需要填充模板的数据
   * @description 该方法只能用来生成doc格式的word模板
   * @throws IOException
   */
  public <T extends Serializable> ExportResultInfo  createDoc(InputStream inputStream, OutputStream outputStream, T t) throws IOException {
    ExportResultInfo resultInfo = new ExportResultInfo();
    InputStreamReader isr = new InputStreamReader(inputStream);

    Template template = new Template("name", isr, configuration); // 获取模板文件
    Writer out = null;

    try {
      out = new BufferedWriter(new OutputStreamWriter(outputStream));
      BeanWrapper data = new BeanWrapperImpl(t);
      template.process(data.getWrappedInstance(), out);
    } catch (TemplateException e) {
      resultInfo.setErrors(e.getMessage());
      resultInfo.setSuccess(false);
      logger.error("ftl文档转换为doc字节流出错", e);
    } finally {
      if (null != out) {
        out.flush();
        out.close();
      }
    }
    return resultInfo;
  }

}
