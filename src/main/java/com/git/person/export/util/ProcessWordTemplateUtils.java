package com.git.person.export.util;

import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.springframework.beans.BeanWrapper;

/**
 * 处理Word模板
 */
public class ProcessWordTemplateUtils {

  static Logger logger = Logger.getLogger(ProcessWordTemplateUtils.class);

  /**
   * 填充contentMap内容到模板doc
   * 
   * @param doc 模板doc
   * @param contentMap 填充内容
   */
  public static void processDoc(CustomXWPFDocument doc, BeanWrapper data) {
    // 处理段落
    List<XWPFParagraph> paragraphList = doc.getParagraphs();
    try {
      processParagraphs(paragraphList, data, doc);
      // 处理表格
      Iterator<XWPFTable> it = doc.getTablesIterator();
      while (it.hasNext()) {
        XWPFTable table = it.next();
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
          List<XWPFTableCell> cells = row.getTableCells();
          for (XWPFTableCell cell : cells) {
            List<XWPFParagraph> paragraphListTable = cell.getParagraphs();
            processParagraphs(paragraphListTable, data, doc);
          }
        }
      }
    } catch (Exception e) {
      logger.error(e);
    }
  }


  /**
   * 处理模板段落数据
   * 
   * @param paragraphList
   * @param data
   * @param doc
   */
  public static void processParagraphs(List<XWPFParagraph> paragraphList, BeanWrapper data,
      CustomXWPFDocument doc) {
    if (paragraphList != null && paragraphList.size() > 0) {
      for (XWPFParagraph paragraph : paragraphList) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
          XWPFRun run = runs.get(i);
          String text = run.getText(0);
          if (null != text && !text.equals("")) {
            Pattern pattern = Pattern.compile("\\$\\{[^\u4e00-\u9fa5]+?\\}");
            Matcher matcher = pattern.matcher(text);
            while (matcher.find()) {
              boolean isSetText = false;
              text = text.substring(2, text.length() - 1);
              Object value = "";
              try {
                value = data.getPropertyValue(text);
              } catch (Exception e) {
                logger.error(e);
              }
              isSetText = true;
              if (value instanceof String) {// 文本替换
                text = text.replace(text, value.toString());
              } else if (value instanceof Map) {// 图片替换
                text = text.replace(text, "");
                doc.creactPic(doc, paragraph, (Map<?, ?>) value);
              }

              if (isSetText) {
                run.setText(text, 0);
              }
            }
          }
        }
      }
    }

  }

}
