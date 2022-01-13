package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataTag;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;

@RestController
@RequestMapping(value = "/WordDataTag2/")
public class WordDataTag2Controller {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //定义WordDocument对象
        WordDocument doc = new WordDocument();
        //定义DataTag对象
        DataTag deptTag = doc.openDataTag("{部门名}");
        deptTag.setValue("技术");

        DataTag userTag = doc.openDataTag("{姓名}");
        userTag.setValue("李四");

        DataTag dateTag = doc.openDataTag("【时间】");
        dateTag.setValue(new SimpleDateFormat("yyyy-MM-dd").format(new Date()).toString());

        poCtrl.setWriter(doc);
        //打开Word文档
        poCtrl.webOpen("/doc/WordDataTag2/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("WordDataTag2/Word");
        return mv;
    }


}
