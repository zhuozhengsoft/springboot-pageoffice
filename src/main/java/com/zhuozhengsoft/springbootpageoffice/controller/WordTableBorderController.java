package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.*;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordTableBorder/")
public class WordTableBorderController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        //打开数据区域
        DataRegion dataRegion = doc.openDataRegion("PO_regTable");
        //打开table，openTable(index)方法中的index代表Word文档中table位置的索引，从1开始
        Table table = dataRegion.openTable(1);

        //给table中的单元格赋值， openCellRC(int,int)中的参数分别代表第几行、第几列，从1开始
        table.openCellRC(3, 1).setValue("A公司");
        table.openCellRC(3, 2).setValue("开发部");
        table.openCellRC(3, 3).setValue("李清");

        //插入一行，insertRowAfter方法中的参数代表在哪个单元格下面插入一个空行
        table.insertRowAfter(table.openCellRC(3, 3));

        table.openCellRC(4, 1).setValue("B公司");
        table.openCellRC(4, 2).setValue("销售部");
        table.openCellRC(4, 3).setValue("张三");

        //设置表格行的高度
        table.setRowsHeight(30.5f);

        //设置表格的边框
        Border border = table.getBorder();
        // 设置边框的类型
        border.setBorderType(WdBorderType.wdFullGrid);//包含内边框
        //设置边框的颜色
        border.setLineColor(Color.red);
        //设置边框的线条样式
        border.setLineStyle(WdLineStyle.wdLineStyleDot);
        //设置边框的粗细
        border.setLineWidth(WdLineWidth.wdLineWidth150pt);

        //设置表格内字体样式
        com.zhuozhengsoft.pageoffice.wordwriter.Font font = dataRegion.getFont();
        //设置字体的是否加粗
        font.setBold(true);
        //设置字体的颜色
        font.setColor(Color.blue);
        //设置字体是否为斜体
        font.setItalic(true);
        //设置字体名称
        font.setName("宋体");
        //设置字体大小
        font.setSize(15.5f);

        poCtrl.setWriter(doc);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/WordTableBorder/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordTableBorder/Word");
        return mv;
    }

}
