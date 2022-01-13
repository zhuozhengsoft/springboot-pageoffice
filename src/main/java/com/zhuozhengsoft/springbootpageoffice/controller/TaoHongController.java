package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
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
import java.io.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/TaoHong/")
public class TaoHongController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {
        ModelAndView mv = new ModelAndView("TaoHong/index");
        return mv;
    }

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.addCustomToolButton("保存", "Save", 1);
        poCtrl.addCustomToolButton("领导圈阅", "StartHandDraw", 3);
        poCtrl.addCustomToolButton("分层显示手写批注", "ShowHandDrawDispBar", 7);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen", 4);
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/TaoHong/test.doc", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("TaoHong/Word");
        return mv;
    }

    @RequestMapping(value = "taoHong")
    public ModelAndView showtaoHong(HttpServletRequest request, Map<String, Object> map) {
        String fileName = "";
        String mbName = request.getParameter("mb");
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.setCustomToolbar(false);
        if (mbName != null && mbName.trim() != "") {
            // 选择模板后执行套红
            // 复制模板，命名为正式发文的文件名：zhengshi.doc
            fileName = "zhengshi.doc";
            String templateName = request.getParameter("mb");
            String templatePath = dir + templateName;
            String filePath = dir + fileName;
            copyFile(templatePath, filePath);

            // 填充数据和正文内容到“zhengshi.doc”
            WordDocument doc = new WordDocument();
            DataRegion copies = doc.openDataRegion("PO_Copies");
            copies.setValue("6");
            DataRegion docNum = doc.openDataRegion("PO_DocNum");
            docNum.setValue("001");
            DataRegion issueDate = doc.openDataRegion("PO_IssueDate");
            issueDate.setValue("2013-5-30");
            DataRegion issueDept = doc.openDataRegion("PO_IssueDept");
            issueDept.setValue("开发部");
            DataRegion sTextS = doc.openDataRegion("PO_STextS");
            sTextS.setValue("[word]/doc/TaoHong/test.doc[/word]");
            DataRegion sTitle = doc.openDataRegion("PO_sTitle");
            sTitle.setValue("北京某公司文件");
            DataRegion topicWords = doc.openDataRegion("PO_TopicWords");
            topicWords.setValue("Pageoffice、 套红");
            poCtrl.setWriter(doc);

        } else {
            //首次加载时，加载正文内容：test.doc
            fileName = "test.doc";

        }
        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("/doc/TaoHong/" + fileName, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("TaoHong/taoHong");
        return mv;
    }


    @RequestMapping(value = "readOnly", method = RequestMethod.GET)
    public ModelAndView showreadOnly(HttpServletRequest request, Map<String, Object> map) {
        String fileName = "zhengshi.doc"; //正式发文的文件
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        poCtrl.setCaption(fileName);
        poCtrl.addCustomToolButton("另存到本地", "ShowDialog1()", 5);
        poCtrl.addCustomToolButton("页面设置", "ShowDialog2()", 0);
        poCtrl.addCustomToolButton("打印", "ShowDialog3()", 6);
        poCtrl.addCustomToolButton("全屏/还原", "IsFullScreen()", 4);
        poCtrl.setMenubar(false);
        poCtrl.setOfficeToolbars(false);

        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法

        //打开Word文档
        poCtrl.webOpen("/doc/TaoHong/" + fileName, OpenModeType.docReadOnly, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("TaoHong/Word");
        return mv;
    }

    public void copyFile(String oldPath, String newPath) {
        try {
            int bytesum = 0;
            int byteread = 0;
            File oldfile = new File(oldPath);
            if (oldfile.exists()) { //文件存在时
                InputStream inStream = new FileInputStream(oldPath); //读入原文件
                FileOutputStream fs = new FileOutputStream(newPath);
                byte[] buffer = new byte[1444];
                int length;
                while ((byteread = inStream.read(buffer)) != -1) {
                    bytesum += byteread; //字节数 文件大小
                    //System.out.println(bytesum);
                    fs.write(buffer, 0, byteread);
                }
                inStream.close();
            }
        } catch (Exception e) {
            System.out.println("复制单个文件操作出错");
            e.printStackTrace();
        }

    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "TaoHong/" + fs.getFileName());
        fs.close();
    }

}
