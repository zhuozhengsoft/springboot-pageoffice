package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.pageoffice.wordwriter.DataRegion;
import com.zhuozhengsoft.pageoffice.wordwriter.WordDocument;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;

@RestController
@RequestMapping(value = "/SplitWord/")
public class SplitWordController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        WordDocument wordDoc = new WordDocument();
        //打开数据区域，openDataRegion方法的参数代表Word文档中的书签名称
        DataRegion dataRegion1 = wordDoc.openDataRegion("PO_test1");
        dataRegion1.setSubmitAsFile(true);
        DataRegion dataRegion2 = wordDoc.openDataRegion("PO_test2");
        dataRegion2.setSubmitAsFile(true);
        dataRegion2.setEditing(true);
        DataRegion dataRegion3 = wordDoc.openDataRegion("PO_test3");
        dataRegion3.setSubmitAsFile(true);
        poCtrl.setWriter(wordDoc);
        poCtrl.addCustomToolButton("保存", "Save()", 1);

        //设置保存页面
        poCtrl.setSaveDataPage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/SplitWord/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("SplitWord/Word");
        return mv;
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws IOException {
        String filePath = dir + "SplitWord/";
        com.zhuozhengsoft.pageoffice.wordreader.WordDocument doc = new com.zhuozhengsoft.pageoffice.wordreader.WordDocument(request, response);
        byte[] bWord;

        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr1 = doc.openDataRegion("PO_test1");

        bWord = dr1.getFileBytes();

        FileOutputStream fos1 = new FileOutputStream(filePath + "new1.doc");
        fos1.write(bWord);
        fos1.flush();
        fos1.close();
        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr2 = doc.openDataRegion("PO_test2");
        bWord = dr2.getFileBytes();
        FileOutputStream fos2 = new FileOutputStream(filePath + "new2.doc");
        fos2.write(bWord);
        fos2.flush();
        fos2.close();

        com.zhuozhengsoft.pageoffice.wordreader.DataRegion dr3 = doc.openDataRegion("PO_test3");
        bWord = dr3.getFileBytes();
        FileOutputStream fos3 = new FileOutputStream(filePath + "new3.doc");
        fos3.write(bWord);
        fos3.flush();
        fos3.close();

        doc.close();
    }

}
