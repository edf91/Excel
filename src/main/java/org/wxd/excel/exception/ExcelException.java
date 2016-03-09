package org.wxd.excel.exception;

/**
 * @Description: excel操作异常
 * @Author : wangxd
 * @Date: 2016-3-2
 * @Version 1.0
 */
public class ExcelException extends RuntimeException {
    private static final long serialVersionUID = -1743039006480308622L;

    public ExcelException() {
        super();
    }

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelException(Throwable cause) {
        super(cause);
    }
}
