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
@RequestMapping(value = "/WordLocateBKMK")
public class WordLocateBKMKController {
    @RequestMapping(value = "/Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
       //隐藏菜单栏
        poCtrl.setMenubar(false);
        //添加自定义按钮
        poCtrl.addCustomToolButton("定位光标到指定书签", "locateBookMark", 5);
        //打开Word文档
        poCtrl.webOpen("/doc/WordLocateBKMK/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordLocateBKMK/Word");
        return mv;
    }

}
