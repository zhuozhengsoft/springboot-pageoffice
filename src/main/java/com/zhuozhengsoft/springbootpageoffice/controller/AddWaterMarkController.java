package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;

@RestController
@RequestMapping(value = "/AddWaterMark")
public class AddWaterMarkController {
    @RequestMapping(value = "/Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        WordDocument doc = new WordDocument();
        //添加水印 ，设置水印的内容
        doc.getWaterMark().setText("PageOffice开发平台");
        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/doc/AddWaterMark/test.doc", OpenModeType.docNormalEdit, "张三");
        request.setAttribute("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("AddWaterMark/Word");
        return mv;
    }


}
