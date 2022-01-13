package com.zhuozhengsoft.springbootpageoffice.controller;

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
import java.io.FileNotFoundException;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@RestController
@RequestMapping(value = "/SaveAndSearch/")
public class SaveAndSearchController {

    //获取doc目录的磁盘路径
    private String dir = GetDirPathUtil.getDirPath() + "static/doc/";

    @RequestMapping(value = "index", method = RequestMethod.GET)
    public ModelAndView showindex(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, FileNotFoundException, UnsupportedEncodingException, SQLException {
        String key = request.getParameter("Input_KeyWord");
        String sql = "";

        if (key != null && key.trim().length() > 0) {
            sql = "select * from word  where Content like '%" + URLDecoder.decode(key, "UTF-8")
                    + "%' order by ID desc";
        } else {
            sql = "select * from word order by ID desc";
        }
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        ResultSet rs = stmt.executeQuery(sql);
        List<com.zhuozhengsoft.springbootpageoffice.ageoffice.entity.DocSearch> docSearchs = new ArrayList<com.zhuozhengsoft.springbootpageoffice.ageoffice.entity.DocSearch>();
        while (rs.next()) {
            com.zhuozhengsoft.springbootpageoffice.ageoffice.entity.DocSearch docSearch = new com.zhuozhengsoft.springbootpageoffice.ageoffice.entity.DocSearch();
            docSearch.setFileName(rs.getString("FileName"));
            docSearch.setContent(rs.getString("Content"));
            docSearch.setId(rs.getInt("ID"));
            docSearchs.add(docSearch);

        }
        stmt.close();
        conn.close();
        map.put("docSearchs", docSearchs);
        map.put("key", key);

        ModelAndView mv = new ModelAndView("SaveAndSearch/index");
        return mv;
    }


    @RequestMapping(value = "Word", method = RequestMethod.GET)
    public ModelAndView showWord(HttpServletRequest request, Map<String, Object> map) throws ClassNotFoundException, SQLException, FileNotFoundException {
        int id = Integer.parseInt(request.getParameter("id"));
        //根据id查询数据库中对应的文档名称
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String sql = "select * from word where id=" + id;
        ResultSet rs = stmt.executeQuery(sql);
        String FileName = "";
        while (rs.next()) {
            FileName = rs.getString("FileName");
        }
        stmt.close();
        conn.close();

        PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
        poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
        //隐藏菜单栏
        poCtrl.setMenubar(false);
        poCtrl.addCustomToolButton("保存", "Save()", 1);
        //设置保存页面
        poCtrl.setSaveFilePage("save?id=" + id);//设置处理文件保存的请求方法
        String filePath = (dir + "SaveAndSearch/" + FileName + ".doc").replace("/", "\\");
        //打开Word文档
        poCtrl.webOpen(filePath, OpenModeType.docNormalEdit, "张三");
        map.put("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));

        ModelAndView mv = new ModelAndView("SaveAndSearch/Word");
        return mv;
    }

    @RequestMapping("save")
    public void save(HttpServletRequest request, HttpServletResponse response) throws SQLException, FileNotFoundException, ClassNotFoundException {
        FileSaver fs = new FileSaver(request, response);
        fs.saveToFile(dir + "SaveAndSearch/" + fs.getFileName());
        fs.setCustomSaveResult("ok");
        String strDocumentText = fs.getDocumentText();
        //更新数据库中文档的文本内容
        int id = Integer.parseInt(request.getParameter("id"));
        Class.forName("org.sqlite.JDBC");
        String strUrl = "jdbc:sqlite:" + ResourceUtils.getURL("classpath:").getPath() + "static/demodata/SaveAndSearch.db";
        Connection conn = DriverManager.getConnection(strUrl);
        Statement stmt = conn.createStatement();
        String strsql = "update word set Content='" + strDocumentText + "' where id=" + id;
        stmt.executeUpdate(strsql);
        stmt.close();
        conn.close();
        fs.close();
    }
}
