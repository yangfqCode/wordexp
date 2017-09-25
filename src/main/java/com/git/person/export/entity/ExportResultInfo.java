package com.git.person.export.entity;

/**
 * 导出的结果返回集合
 */
public class ExportResultInfo implements Result {

  /**
   * 
   */
  private static final long serialVersionUID = 1L;
  boolean success; //是否成功
  String errors;//错误消息

  @Override
  public boolean getSuccess() {
    return this.success;
  }

  @Override
  public String getErrors() {
    return this.errors;
  }

  public void setSuccess(boolean success) {
    this.success = success;
  }

  public void setErrors(String errors) {
    this.errors = errors;
  }
  
}
