package com.git.person.export.entity;

import java.io.Serializable;

/**
 * 结果集返回接口
 * 
 * @author administration
 *
 */
public interface Result extends Serializable{

  boolean getSuccess();//是否成功

  String getErrors();//错误消息


}
