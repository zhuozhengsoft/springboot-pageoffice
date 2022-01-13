package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegionInsertType;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.util.Map;

@RestController
@RequestMapping(value = "/InsertPageBreak2/")
public class InsertPageBreak2Controller {
    private String dir = ResourceUtils.getURL("classpath:").getPath() + "static/doc/";

    public InsertPageBreak2Controller() throws FileNotFoundException {
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument doc = new WordDocument();
        DataRegion mydr1 = doc.createDataRegion("PO_first", DataRegionInsertType.After, "[end]");
        mydr1.selectEnd();
        doc.insertPageBreak();//插入分页符
        DataRegion mydr2 = doc.createDataRegion("PO_second", DataRegionInsertType.After, "[end]");
        mydr2.setValue("[word]doc/test2.doc[/word]");

        poCtrl.addCustomToolButton("保存", "Save()", 1);
        poCtrl.setWriter(doc);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/InsertPageBreak2/test1.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("InsertPageBreak2/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "InsertPageBreak2/" + fs.getFileName());
        fs.close();
    }

}
