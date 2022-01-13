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
@RequestMapping(value = "/CommandCtrl/")
public class CommandCtrlController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面

        poCtrl.setCustomToolbar(false);
        poCtrl.setOfficeToolbars(false);
        poCtrl.setAllowCopy(false);//禁止拷贝

        poCtrl.setJsFunction_AfterDocumentOpened("AfterDocumentOpened");

        //打开Word文档
        poCtrl.webOpen("/doc/CommandCtrl/test.doc", OpenModeType.docReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("CommandCtrl/Word");
        return mv;
    }


}
