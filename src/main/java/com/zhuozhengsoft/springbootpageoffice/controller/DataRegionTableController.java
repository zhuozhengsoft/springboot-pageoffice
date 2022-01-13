package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.awt.*;
import java.io.FileNotFoundException;
import java.util.Map;

import static java.lang.System.out;

@RestController
@RequestMapping(value = "/DataRegionTable/")
public class DataRegionTableController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dTable = doc.openDataRegion("PO_table");
        //设置数据区域可编辑性
        dTable.setEditing(true);

        //打开数据区域中的表格，OpenTable(index)方法中的index为word文档中表格的下标，从1开始
        Table table1 = doc.openDataRegion("PO_Table").openTable(1);
        //设置表格边框样式
        table1.getBorder().setLineColor(Color.green);
        table1.getBorder().setLineWidth(WdLineWidth.wdLineWidth050pt);
        // 设置表头单元格文本居中
        table1.openCellRC(1, 2).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(1, 3).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(2, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);
        table1.openCellRC(3, 1).getParagraphFormat().setAlignment(WdParagraphAlignment.wdAlignParagraphCenter);

        // 给表头单元格赋值
        table1.openCellRC(1, 2).setValue("产品1");
        table1.openCellRC(1, 3).setValue("产品2");
        table1.openCellRC(2, 1).setValue("A部门");
        table1.openCellRC(3, 1).setValue("B部门");

        poCtrl.setWriter(doc);
        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/DataRegionTable/test.doc", OpenModeType.docSubmitForm, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionTable/Word");
        return mv;
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dataReg = doc.openDataRegion("PO_table");
        com.zhuozhengsoft.pageoffice.wordreader.Table table = dataReg.openTable(1);

        //输出提交的table中的数据
        out.print("表格中的各个单元的格数据为：<br/><br/>");
        StringBuilder dataStr = new StringBuilder();
        for (int i = 1; i <= table.getRowsCount(); i++) {
            dataStr.append("<div style='width:220px;'>");
            for (int j = 1; j <= table.getColumnsCount(); j++) {
                dataStr.append("<div style='float:left;width:70px;border:1px solid red;'>" + table.openCellRC(i, j).getValue() + "</div>");
            }
            dataStr.append("</div>");
        }
        out.print(dataStr.toString());

        //向客户端显示提交的数据
        doc.showPage(300, 300);
        doc.close();
    }

}
