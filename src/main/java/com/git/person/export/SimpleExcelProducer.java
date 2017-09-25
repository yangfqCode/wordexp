package com.git.person.export;

import java.io.IOException;
import java.io.OutputStream;

import com.git.person.export.entity.ExportResultInfo;

/**
 * 无模板，直接输出数据到xlsx, col为默认顺序, 单行head
 * @author Think
 */
public abstract class SimpleExcelProducer implements ExportProducer<ExportResultInfo> {

  //获取Sheet名称
  protected abstract String getSheetName(MultivaluedMap<String, String> mm);

  //获取标题组
  protected abstract String[] getHeadName(MultivaluedMap<String, String> mm);
  
  /**
   * 写Head, 如为复合表头， 需重写此方法
   * @param sheet
   * @param mm
   * @throws IOException
   */
  protected void writeHead(XSSFSheet sheet, MultivaluedMap<String, String> mm)
      throws IOException {
    
    // 简单的单行Head
    int i = 0;
    
    String[] headNameList = getHeadName(mm);
    if(headNameList!=null && headNameList.length>0){
      XSSFRow headRow = sheet.createRow(i);
  
      for (int h = 0; h < headNameList.length; h++) {
        String headName = headNameList[h];
        XSSFCell cell = headRow.createCell(h);
        cell.setCellValue(headName);
      }
      i++;
    }
  }

  /**
   * 写数据， 允许改动格式及合并
   * @param sheet
   * @param mm
   */
  protected abstract void writeRecord(XSSFSheet sheet, MultivaluedMap<String, String> mm);

  
  /**
   * 写入到输出流
   */
  @Override
  public ExportResultInfo export(OutputStream out, MultivaluedMap<String, String> mm)
      throws IOException {

    ExportResultInfo result = new ExportResultInfo();


    try {
      XSSFWorkbook workbook = new XSSFWorkbook();
      
      XSSFSheet sheet = workbook.createSheet(getSheetName(mm));
      writeHead(sheet, mm);
      writeRecord(sheet, mm);

      workbook.write(out);
      workbook.close();
      result.setSuccess(true);
    } catch (Exception ex) {
      result.setSuccess(false);
      result.setErrors("");
    }

    return result;
  }
}
