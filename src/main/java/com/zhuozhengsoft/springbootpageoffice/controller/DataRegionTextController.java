package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
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
@RequestMapping(value = "/DataRegionText/")
public class DataRegionTextController {
    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        DataRegion d1 = doc.openDataRegion("d1");
        d1.getFont().setColor(Color.BLUE);//设置数据区域文本字体颜色
        d1.getFont().setName("华文彩云");//设置数据区域文本字体样式
        d1.getFont().setSize(16);//设置数据区域文本字体大小
        d1.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphCenter);//设置数据区域文本对齐方式

        DataRegion d2 = doc.openDataRegion("d2");
        d2.getFont().setColor(Color.orange);//设置数据区域文本字体颜色
        d2.getFont().setName("黑体");//设置数据区域文本字体样式
        d2.getFont().setSize(14);//设置数据区域文本字体大小
        d2.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphLeft);//设置数据区域文本对齐方式

        DataRegion d3 = doc.openDataRegion("d3");
        d3.getFont().setColor(Color.magenta);//设置数据区域文本字体颜色
        d3.getFont().setName("华文行楷");//设置数据区域文本字体样式
        d3.getFont().setSize(12);//设置数据区域文本字体大小
        d3.getParagraphFormat().setAlignment(
                WdParagraphAlignment.wdAlignParagraphRight);//设置数据区域文本对齐方式

        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/doc/DataRegionText/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionText/Word");
        return mv;
    }

    @RequestMapping(value = "Word2", method = RequestMethod.GET)
    public ModelAndView showWord2(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //打开Word文档
        poCtrl.webOpen("/doc/DataRegionText/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataRegionText/Word2");
        return mv;
    }

}
