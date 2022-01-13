package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.Table;
import com.zhuozhengsoft.pageoffice.wordwriter.WdParagraphAlignment;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.awt.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/MergeWordCell/")
public class MergeWordCellController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        DataRegion dataReg = doc.openDataRegion("PO_table");
        Table table = dataReg.openTable(1);
        //合并table中的单元格
        table.openCellRC(1, 1).mergeTo(1, 4);
        //给合并后的单元格赋值
        table.openCellRC(1, 1).setValue("销售情况表");
        //设置单元格文本样式
        table.openCellRC(1, 1).getFont().setColor(Color.red);
        table.openCellRC(1, 1).getFont().setSize(24);
        table.openCellRC(1, 1).getFont().setName("楷体");
        table.openCellRC(1, 1).getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphCenter);

        poCtrl.setWriter(doc);

        poCtrl.setCustomToolbar(false);
        //打开Word文档
        poCtrl.webOpen("/doc/MergeWordCell/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("MergeWordCell/Word");
        return mv;
    }


}
