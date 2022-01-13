package com.zhuozhengsoft.springbootpageoffice.controller;

import com.zhuozhengsoft.pageoffice.FileSaver;
import com.zhuozhengsoft.pageoffice.OpenModeType;
import com.zhuozhengsoft.pageoffice.PageOfficeCtrl;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.*;
import java.util.Map;

@RestController
@RequestMapping(value = "/DataBase/")
public class DataBaseController {

    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) {
        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //添加自定义按钮
        poCtrl.addCustomToolButton("保存", "Save", 1);
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        //设置保存页面
        poCtrl.setSaveFilePage("save?id=1");//设置处理文件保存的请求方法
        //打开Word文档
        poCtrl.webOpen("Openstream?id=1", OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
        ModelAndView mv = new ModelAndView("DataBase/Word");
        return mv;
    }

    @RequestMapping(value = "Openstream", method = RequestMethod.GET)
    public void openWord(HttpServletRequest request, HttpServletResponse response) throws SQLException, ClassNotFoundException, IOException {

        String id = "2";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            id = request.getParameter("id");
        }
        Class.forName("org.sqlite.JDBC");

        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/DataBase.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery("select * from stream where id = "
                + id);
        int newID = 1;
        if (rs.next()) {
            //******读取磁盘文件，输出文件流 开始*******************************
            byte[] imageBytes = rs.getBytes("Word");
            int fileSize = imageBytes.length;

            response.reset();
            response.setContentType("application/msword"); // application/x-excel, application/ms-powerpoint, application/pdf
            response.setHeader("Content-Disposition",
                    "attachment; filename=down.doc"); //fileN应该是编码后的(utf-8)
            response.setContentLength(fileSize);

            OutputStream outputStream = response.getOutputStream();
            outputStream.write(imageBytes);

            outputStream.flush();
            outputStream.close();
            outputStream = null;
            //******读取磁盘文件，输出文件流 结束*******************************
        }
        rs.close();
        conn.close();
    }


    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws ClassNotFoundException, FileNotFoundException, SQLException {
        FileSaver fs = new FileSaver(request, response);
        String err = "";
        if (request.getParameter("id") != null
                && request.getParameter("id").trim().length() > 0) {
            String id = request.getParameter("id").trim();
            Class.forName("org.sqlite.JDBC");
            String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/DataBase.db";
            ;
            Connection conn = DriverManager.getConnection(strUrl);
            String sql = "UPDATE  Stream SET Word=?  where ID=" + id;
            PreparedStatement pstmt = null;
            pstmt = conn.prepareStatement(sql);
            pstmt.setBytes(1, fs.getFileBytes());
            //pstmt.setBinaryStream(1, fs.getFileStream(),fs.getFileSize());
            pstmt.executeUpdate();
            pstmt.close();
            conn.close();

            fs.setCustomSaveResult("ok");
        } else {
            err = "<script>alert('未获得文件的ID，保存失败');</script>";
        }
        fs.close();
    }

}
