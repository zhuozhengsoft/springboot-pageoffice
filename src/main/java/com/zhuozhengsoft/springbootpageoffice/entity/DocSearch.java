package com.zhuozhengsoft.springbootpageoffice.ageoffice.entity;
import java.io.Serializable;

public class DocSearch implements Serializable {
    private static final long serialVersionUID = -3368973367666536973L;
    private String fileName;
    private String content;
    private Integer id;

    public String getFileName() {
        return fileName;
    }

    public void setFileName(String fileName) {
        this.fileName = fileName;
    }

    public String getContent() {
        return content;
    }

    public void setContent(String content) {
        content = content;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    @Override
    public String toString() {
        return "DocSearch{" +
                "fileName='" + fileName + '\'' +
                ", Content='" + content + '\'' +
                ", id=" + id +
                '}';
    }
}
