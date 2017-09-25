package com.git.person.export.util;

import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;


/**
 * 文档转换抽象类
 * 
 * @author Administration
 */
public abstract class Converter {

  protected XWPFDocument document;
  protected OutputStream outStream;


  public Converter(XWPFDocument document, OutputStream outStream) throws IOException {
    this.document = document;
    this.outStream = outStream;
  }

  public abstract void convert() throws Exception;

  protected void finished() {
    try {
      document.close();
      outStream.close();
    } catch (IOException e) {
      throw new RuntimeException(e);
    }
  }

}
