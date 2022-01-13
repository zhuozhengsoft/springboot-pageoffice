package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.DocumentVersion;
import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import com.zhuozhengsoft.springbootpageoffice.entity.Doc;
import com.zhuozhengsoft.springbootpageoffice.util.GetDirPathUtil;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping(value = "/CreateWord/")
public class CreateWordController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "word-lists")
    public ModelAndView showWord20(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException, ParseException, FileNotFoundException, UnsupportedEncodingException {

        if (request.getParameter("op") != null && request.getParameter("op").length() > 0) {
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";
            Connection conn = DriverManager.getConnection(strUrl);
            Statement stmt = conn.createStatement();
            ResultSet rs = stmt.executeQuery("select Max(ID) from word");
            int newID = 1;
            if (rs.next()) {
                newID = Integer.parseInt(rs.getString(1)) + 1;
            }
            rs.close();

            String fileName = "aabb" + newID + ".doc";

            String FileSubject = "请输入文档主题";
            String getFile = (String) request.getParameter("FileSubject");
            //if (getFile != null && getFile.length() > 0)  FileSubject = new String(getFile.getBytes("iso-8859-1"), "utf-8");
            SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
            // new Date()为获取当前系统时间
            String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values(" + newID
                    + ",'" + fileName + "','" + FileSubject + "','" + df.format(new Date()) + "')";
            stmt.executeUpdate(strsql);
            stmt.close();
            conn.close();

            //拷贝文件
            //if(request.getParameter("action").equals("create")){
            String oldPath = dir + " CreateWord/template.doc";
            String newPath = dir + " CreateWord/" + fileName;
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
                    inStream.close();
                }
            } catch (Exception e) {
                System.out.println("复制单个文件操作出错");
                e.printStackTrace();
            }
            //}
            return new ModelAndView("redirect:/CreateWord/word-lists");
        }


        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from word order by id desc");
        String fileName = "";
        String subject = "";
        String submitTime = "";
        List<Doc> list = new ArrayList<Doc>();

        while (rs.next()) {
            int id = rs.getInt("ID");
            fileName = rs.getString("FileName");
            subject = rs.getString("Subject");
            submitTime = rs.getString("SubmitTime");
            if (submitTime != null && submitTime.length() > 0) {
                submitTime = new SimpleDateFormat("yyyy/MM/dd")
                        .format(new SimpleDateFormat("yyyy-MM-dd")
                                .parse(submitTime));
            }
            Doc doc = new Doc();
            doc.setId(id);
            doc.setFileName(fileName);
            doc.setSubject(subject);
            doc.setSubmitTime(submitTime);
            list.add(doc);

        }
        rs.close();
        stmt.close();
        conn.close();


        map.put("list", list);


        //--- PageOffice的调用代码 结束 -----
        ModelAndView mv = new ModelAndView("CreateWord/list");


        return mv;
    }


    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, FileNotFoundException, SQLException {
        String subject = "";
        String fileName = "";


        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String id = request.getParameter("id");
        if (!id.equals("") && !id.equals(null)) {
            ResultSet rs = stmt.executeQuery("select * from word where ID=" + id);
            subject = rs.getString("Subject");
            fileName = rs.getString("FileName");
            rs.close();
        }
        stmt.close();
        conn.close();


        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);


        //设置保存页面
        poCtrl.setSaveFilePage("save");//设置处理文件保存的请求方法


        //打开Word文档
        poCtrl.webOpen("/doc/CreateWord/" + fileName, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        map.put("subject", subject);
        ModelAndView mv = new ModelAndView("CreateWord/Word");
        return mv;
    }


    @RequestMapping(value = "creatWord", method = RequestMethod.GET)
    public ModelAndView showWord3(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, FileNotFoundException, SQLException {


        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //隐藏工具栏
        poCtrl.setCustomToolbar(false);
        poCtrl.setJsFunction_BeforeDocumentSaved("BeforeDocumentSaved()");


        //设置保存页面
        poCtrl.setSaveFilePage("SaveNewFile");//设置处理文件保存的请求方法

        //新建Word文件，webCreateNew方法中的两个参数分别指代“操作人”和“新建Word文档的版本号”
        poCtrl.webCreateNew("张佚名", DocumentVersion.Word2003);

        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("CreateWord/creatWord");
        return mv;
    }


    @RequestMapping("SaveNewFile")
    public void save2(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, FileNotFoundException, SQLException {
        FileSaver fs = new FileSaver(request, response);


        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/CreateWord.db";

        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select Max(ID) from word");
        int newID = 1;
        if (rs.next()) {
            newID = Integer.parseInt(rs.getString(1)) + 1;
        }
        rs.close();

        String FileSubject = fs.getFormField("FileSubject").trim();
        String fileName = "aabb" + newID + ".doc";
        String getFile = (String) request.getParameter("FileSubject");
        //if (getFile != null && getFile.length() > 0) FileSubject = new String(getFile.getBytes("iso-8859-1"));
        //out.print(FileSubject);
        SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");//设置日期格式
        // new Date()为获取当前系统时间
        String strsql = "Insert into word(ID,FileName,Subject,SubmitTime) values("
                + newID
                + ",'"
                + fileName
                + "','"
                + FileSubject
                + "','"
                + df.format(new Date()) + "')";
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();


        fs.saveToFile(dir + "CreateWord/" + fs.getFileName());
        //设置保存结果
        fs.setCustomSaveResult("ok");
        fs.close();
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) {
        FileSaver fs = new FileSaver(request, response);


        fs.saveToFile(dir + "CreateWord/" + fs.getFileName());
        fs.close();
    }

}
