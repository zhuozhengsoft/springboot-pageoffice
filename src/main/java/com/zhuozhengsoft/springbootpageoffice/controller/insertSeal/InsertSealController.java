package com.zhuozhengsoft.springbootpageoffice.controller.insertSeal;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/InsertSeal/")
public class InsertSealController {
    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) {

        ModelAndView mv = new ModelAndView("InsertSeal/index");
        return mv;
    }

    @RequestMapping(value = "refresh", method = RequestMethod.GET)
    public ModelAndView showrefresh(HttpServletRequest request, Map<String, Object> map) {

        copyFile(dir + "InsertSeal/Word/AddSeal/test1_bak.doc", dir + "InsertSeal/Word/AddSeal/test1.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test2_bak.doc", dir + "InsertSeal/Word/AddSeal/test2.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test3_bak.doc", dir + "InsertSeal/Word/AddSeal/test3.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test4_bak.doc", dir + "InsertSeal/Word/AddSeal/test4.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test5_bak.doc", dir + "InsertSeal/Word/AddSeal/test5.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test6_bak.doc", dir + "InsertSeal/Word/AddSeal/test6.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test7_bak.doc", dir + "InsertSeal/Word/AddSeal/test7.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test8_bak.doc", dir + "InsertSeal/Word/AddSeal/test8.doc");
        copyFile(dir + "InsertSeal/Word/AddSeal/test10_bak.doc", dir + "InsertSeal/Word/AddSeal/test10.doc");


        copyFile(dir + "InsertSeal/Word/AddSign/test1_bak.doc", dir + "InsertSeal/Word/AddSign/test1.doc");
        copyFile(dir + "InsertSeal/Word/AddSign/test2_bak.doc", dir + "InsertSeal/Word/AddSign/test2.doc");
        copyFile(dir + "InsertSeal/Word/AddSign/test3_bak.doc", dir + "InsertSeal/Word/AddSign/test3.doc");
        copyFile(dir + "InsertSeal/Word/AddSign/test4_bak.doc", dir + "InsertSeal/Word/AddSign/test4.doc");
        copyFile(dir + "InsertSeal/Word/AddSign/test5_bak.doc", dir + "InsertSeal/Word/AddSign/test5.doc");

        copyFile(dir + "InsertSeal/Word/BatchAddSeal/test1_bak.doc", dir + "InsertSeal/Word/BatchAddSeal/test1.doc");
        copyFile(dir + "InsertSeal/Word/BatchAddSeal/test2_bak.doc", dir + "InsertSeal/Word/BatchAddSeal/test2.doc");
        copyFile(dir + "InsertSeal/Word/BatchAddSeal/test3_bak.doc", dir + "InsertSeal/Word/BatchAddSeal/test3.doc");
        copyFile(dir + "InsertSeal/Word/BatchAddSeal/test4_bak.doc", dir + "InsertSeal/Word/BatchAddSeal/test4.doc");


        copyFile(dir + "InsertSeal/PDF/AddSeal/test1_bak.pdf", dir + "InsertSeal/PDF/AddSeal/test1.pdf");
        copyFile(dir + "InsertSeal/PDF/AddSign/test1_bak.pdf", dir + "InsertSeal/PDF/AddSign/test1.pdf");

        copyFile(dir + "InsertSeal/Excel/AddSeal/test1_bak.xls", dir + "InsertSeal/Excel/AddSeal/test1.xls");
        copyFile(dir + "InsertSeal/Excel/AddSeal/test2_bak.xls", dir + "InsertSeal/Excel/AddSeal/test2.xls");
        copyFile(dir + "InsertSeal/Excel/AddSeal/test3_bak.xls", dir + "InsertSeal/Excel/AddSeal/test3.xls");
        copyFile(dir + "InsertSeal/Excel/AddSeal/test4_bak.xls", dir + "InsertSeal/Excel/AddSeal/test4.xls");
        copyFile(dir + "InsertSeal/Excel/AddSeal/test5_bak.xls", dir + "InsertSeal/Excel/AddSeal/test5.xls");

        copyFile(dir + "InsertSeal/Excel/AddSign/test1_bak.xls", dir + "InsertSeal/Excel/AddSign/test1.xls");
        copyFile(dir + "InsertSeal/Excel/AddSign/test2_bak.xls", dir + "InsertSeal/Excel/AddSign/test2.xls");
        copyFile(dir + "InsertSeal/Excel/AddSign/test3_bak.xls", dir + "InsertSeal/Excel/AddSign/test3.xls");

        ModelAndView mv = new ModelAndView("InsertSeal/refresh");
        return mv;
    }

    //拷贝文件
    private void copyFile(String oldPath, String newPath) {
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
                    System.out.println(bytesum);
                    fs.write(buffer, 0, byteread);
                }
                fs.close();
                inStream.close();
            }
        } catch (Exception e) {
            System.out.println("复制单个文件操作出错");
            e.printStackTrace();
        }
    }


}
