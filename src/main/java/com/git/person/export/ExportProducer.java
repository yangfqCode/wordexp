package com.git.person.export;

import java.io.IOException;
import java.io.OutputStream;

import com.git.person.export.entity.Result;

/**
 * 导出生产者接口类
 * 
 * @Author: administration
 * @Version:v1.0
 */
public interface ExportProducer <T extends Result>{

  /**
   * 处理输出流
   * 
   * @param OutputStream
   * @param mm
   * @return
   * 
   * @return word的Document
   * @throws IOException
   */
  public  T export(OutputStream outputStream, MultivaluedMap<String, String> mm)
      throws IOException;

}
