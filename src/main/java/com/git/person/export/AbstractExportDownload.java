/**
 * 
 */
package com.git.person.export;

import java.io.IOException;
import java.io.OutputStream;

import javax.annotation.PostConstruct;
import javax.ws.rs.WebApplicationException;
import javax.ws.rs.core.MultivaluedMap;
import javax.ws.rs.core.StreamingOutput;

import org.springframework.context.support.ApplicationObjectSupport;


/**
 * 文件导出下载抽象类；
 * 需要实现该类的抽象方法
 * 
 * @param <T> 导出数据的bean
 */
public abstract class AbstractExportDownload extends ApplicationObjectSupport implements DownloadProducer {

  protected ExportProducer<?> exportProducer;

  /**
   * 导出的文件名称
   * 
   * @param mm
   * @return
   */
  protected abstract String getExportFileName(MultivaluedMap<String, String> mm);
  

  @PostConstruct
  public void init() {
    exportProducer = (ExportProducer<?>) getApplicationContext().getBean(getType() + "-Export");
  }

  @Override
  public DownLoadStatus produce(MultivaluedMap<String, String> mm) throws IOException {
    try {
      return new DownLoadStatus(getExportFileName(mm), new StreamingOutput() {
        public void write(OutputStream outputStream) throws IOException, WebApplicationException {
          try {
            exportProducer.export(outputStream,  mm);
          } catch (Exception e) {
            logger.error(e.getMessage(),e);
          }
        }
      });
    } catch (IllegalArgumentException e) {
      logger.error("导出文件出错，组装文件数据出错！", e);

    } catch (SecurityException e) {
      logger.error("生成文件出错，组装文件数据出错！", e);
    }
    throw new BusinessAccessException(
        "error.download.util.AbstractExportDownload");
  }

  @Override
  public boolean downloadable(MultivaluedMap<String, String> mm) throws IOException {
    return true;
  }

}
