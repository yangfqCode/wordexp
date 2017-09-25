package com.git.person.export.service;

import java.beans.beancontext.BeanContext;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Clob;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.git.person.export.util.CustomXWPFDocument;

/**
 * 根据表名和列名获取数据
 * 
 * 根据传入的Word模板占位符为${table.cloumn_T}这种结构的，自动根据id获取对应的值
 */

public class DocumentDataService {

  /**
   * 
   * 
   * @param File file 导出的word模板
   * @param String id 导出单据的id
   * @param Map<String, Object> contentMap
   * @param Long id 对应数据的id
   * @param String pattern 时间字段对应的格式，默认为yyyy-MM-dd
   * @return Map<String, Object>
   * @throws IOException
   */
  public static Map<String, Object> templateData(InputStream in, Map<String, Object> contentMap, Long id, String pattern) 
      throws IOException {
    
    CustomXWPFDocument doc = null;
    doc = new CustomXWPFDocument(in);
    
    // 获得占位符
    HashSet<String> templateset = getTemplatePlaceholder(doc);
    Iterator<String> iterator = templateset.iterator();

    while (iterator.hasNext()) {
      // 占位符templatename包含tableName、columnName
      String templatename = iterator.next();
      String[] namearray = templatename.substring(templatename.indexOf("{") + 1, templatename.lastIndexOf("}")).split("\\.");
      if (namearray.length == 2 && namearray[1].endsWith("_T")) {
        if (!contentMap.containsKey(templatename)) {
          String tableName = namearray[0];
          String columnName = namearray[1];
          columnName = columnName.substring(0, columnName.lastIndexOf("_T"));
          String columnvalue = getColValueByIdDynamic(columnName, tableName, id, pattern);
          contentMap.put(templatename, columnvalue.toString());
        }
      }
    }
    return contentMap;
  }

  /**
   * 获取模板的占位符，并且将占位符组装成Set集合
   * 
   * @param doc
   */
  public static HashSet<String> getTemplatePlaceholder(CustomXWPFDocument doc) {
    HashSet<String> set = new HashSet<String>();
    // 正则表达式，匹配${***}
    Pattern pattern = Pattern.compile("\\$\\{[^\u4e00-\u9fa5]+?\\}");
    StringBuffer stringBuffer = new StringBuffer();
    List<XWPFParagraph> paragraphList = doc.getParagraphs();
    if (paragraphList != null && paragraphList.size() > 0) {
      for (XWPFParagraph paragraph : paragraphList) {
        List<XWPFRun> runs = paragraph.getRuns();
        try {
          for (XWPFRun run : runs) {
            String text = run.getText(0);
            if (text != null && !text.equals("")) {
              stringBuffer.append(text);
            }
          }
        } catch (Exception e) {
          e.printStackTrace();
        }
      }
    }
    for (int i = 0; i < doc.getTables().size(); i++) {
      stringBuffer.append(doc.getTables().get(i).getText());
    }

    // 正则匹配
    Matcher matcher = pattern.matcher(stringBuffer.toString());
    while (matcher.find()) {
      String s = matcher.group();
      set.add(s);
    }
    return set;
  }



  public static String getColValueByIdDynamic(String columnName, String tableName, Long id,
      String pattern) {
    String sql = "select " + columnName + " from " + tableName + " where ID = " + id;
    JdbcTemplate jdbcTemplate = BeanContext.getBean(JdbcTemplate.class);
    Object object = jdbcTemplate.queryForObject(sql, Object.class);
    return getExportValue(object, pattern);
  }

  /**
   * 获取打印的值 将Object类型的值转换成对应的String类型的值
   * 
   * @param param
   * @return String
   */
  private static String getExportValue(Object param, String pattern) {
    String value = "";
    if (param == null || "".equals(param) || "null".equals(param)) {
      return "";
    }
    if (param instanceof Integer) {
      value = ((Integer) param).intValue() + "";
    } else if (param instanceof Date) {// 时间
      Date d = (Date) param;
      String formate = pattern == null ? "yyyy-MM-dd" : pattern;
      SimpleDateFormat sdf = new SimpleDateFormat(formate);
      value = sdf.format(d);
    } else if (param instanceof Clob) {// 大字段
      try {
        value = ((Clob) param).getSubString(1, (int) ((Clob) param).length());
      } catch (SQLException e) {
        e.printStackTrace();
      }
    } else if (param instanceof Number) {
      value = ((Number) param).toString();
    } else {
      value = param.toString();
    }
    return value;
  }
  
  /**
   * 将输入流中的数据写入字节数组
   * 
   * @param in
   * @return
   */
  public static byte[] inputStream2ByteArray(InputStream in, boolean isClose) {
    byte[] byteArray = null;
    try {
      int total = in.available();
      byteArray = new byte[total];
      in.read(byteArray);
    } catch (IOException e) {
      e.printStackTrace();
    } finally {
      if (isClose) {
        try {
          in.close();
        } catch (Exception e2) {
          System.out.println("关闭流失败");
        }
      }
    }
    return byteArray;
  }

}
