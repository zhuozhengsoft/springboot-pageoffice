package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordCompare/")
public class WordCompareController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("显示A文档", "ShowFile1View()", 0);
        poCtrl.addCustomToolButton("显示B文档", "ShowFile2View()", 0);
        poCtrl.addCustomToolButton("显示比较结果", "ShowCompareView()", 0);
        //poCtrl.wordCompare("doc/aaa1.doc", "doc/aaa2.doc", OpenModeType.docReadOnly, "张三");
        poCtrl.wordCompare("/doc/WordCompare/aaa1.doc", "/doc/WordCompare/aaa2.doc", OpenModeType.docAdmin, "张三");

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordCompare/Word");
        return mv;
    }


}
