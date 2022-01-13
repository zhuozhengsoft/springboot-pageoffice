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
@RequestMapping(value = "/OpenWord/")
public class OpenWordController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //隐藏Office工具条
        poCtrl.setOfficeToolbars(false);
        //隐藏自定义工具栏
        poCtrl.setCustomToolbar(false);
        //设置页面的显示标题
        poCtrl.setCaption("演示：最简单的以只读模式打开Word文档");
        //打开Word文档
        poCtrl.webOpen("/doc/OpenWord/test.doc", OpenModeType.docReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("OpenWord/Word");
        return mv;
    }


}
