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
@RequestMapping(value = "/WordRibbonCtrl/")
public class WordRibbonCtrlController {
    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.getRibbonBar().setTabVisible("TabHome", true);//开始
        poCtrl.getRibbonBar().setTabVisible("TabPageLayoutWord", false);//页面布局
        poCtrl.getRibbonBar().setTabVisible("TabReferences", false);//引用
        poCtrl.getRibbonBar().setTabVisible("TabMailings", false);//邮件
        poCtrl.getRibbonBar().setTabVisible("TabReviewWord", false);//审阅
        poCtrl.getRibbonBar().setTabVisible("TabInsert", false);//插入
        poCtrl.getRibbonBar().setTabVisible("TabView", false);//视图
        poCtrl.getRibbonBar().setSharedVisible("FileSave", false);//office自带的保存按钮
        poCtrl.getRibbonBar().setGroupVisible("GroupClipboard", false);//分组剪贴板

        //打开Word文档
        poCtrl.webOpen("/doc/WordRibbonCtrl/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordRibbonCtrl/Word");
        return mv;
    }


}
