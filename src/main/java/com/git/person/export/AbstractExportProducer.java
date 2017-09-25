package com.git.person.export;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.logging.Logger;

import com.git.person.export.entity.ExportResultInfo;
import com.git.person.export.entity.Result;
import com.git.person.export.service.FreemarkerDocService;
import com.git.person.export.service.PoiDocxService;

/**
 * 导出的默认实现类
 * 
 * @param <T>
 */
public abstract class AbstractExportProducer<T extends Serializable> implements ExportProducer<Result> {

  protected final Logger logger = Logger.getLogger(this.getClass());

  /**
   * 获取导出模板文件输入流
   * 
   * @param mm
   * @return
   */
  protected abstract InputStream getExportTplInputStream(MultivaluedMap<String, String> mm)
      throws IOException;

  /**
   * 获取填充导出模板的数据
   * 
   * @param mm
   * @return
   */
  protected abstract T getExportDataSource(MultivaluedMap<String, String> mm);



  /*****************************************************************************************************
   * 以下方法为框架提供的ExportProducer接口中exportDoc几种实现方式，选其一即可，如果满足不了项目需求，自行实现
   *****************************************************************************************************/


  /**
   * 使用poi实现Word模板导出
   * 
   * @param inputStream 导出模板的输入流
   * @param outputStream 输出流
   * @param t 导出数据的bean
   * @throws IOException
   */
  protected ExportResultInfo pcreateDoc(OutputStream outputStream, MultivaluedMap<String, String> mm) throws IOException {
    return  PoiDocxService.createDoc(getExportTplInputStream(mm), outputStream, getExportDataSource(mm));
  }


  /**
   * 使用Freemarker实现Word模板导出
   * 
   * @param inputStream
   * @param outputStream
   * @param contentMap
   * @throws IOException
   */
  protected ExportResultInfo fcreateDoc(OutputStream outputStream, MultivaluedMap<String, String> mm)
      throws IOException {
    FreemarkerDocService freemarkerDocUtils = new FreemarkerDocService();
    return freemarkerDocUtils.createDoc(getExportTplInputStream(mm), outputStream, getExportDataSource(mm));
  }


}
