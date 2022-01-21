package com.zhuozhengsoft.springbootpageoffice;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;

import java.io.FileNotFoundException;

@SpringBootApplication
public class SpringbootPageofficeApplication {
    @Value("${posyspath}")
    private String poSysPath;

    @Value("${popassword}")
    private String poPassword;
    public static void main(String[] args) {
        SpringApplication.run(SpringbootPageofficeApplication.class, args);
    }
    /**
     * 添加PageOffice的服务器端授权程序Servlet（必须）
     *
     * @return
     */
    @Bean
    public ServletRegistrationBean pageofficeRegistrationBean()  {
        com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
        /**如果当前项目是打成jar或者war包运行，强烈建议将license的路径更换成某个固定的绝对路径下，不要放当前项目文件夹下,为了防止每次重新发布项目导致license丢失问题。
          * 比如windows服务器下：D:/pageoffice/，linux服务器下:/root/pageoffice/
         */
        //设置PageOffice注册成功后,license.lic文件存放的目录
        poserver.setSysPath(poSysPath);//poSysPath可以在application.properties这个文件中配置，也可以直设置文件夹路径，比如：poserver.setSysPath("D:/pageoffice");
        ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
        srb.addUrlMappings("/poserver.zz");
        srb.addUrlMappings("/posetup.exe");
        srb.addUrlMappings("/pageoffice.js");
        srb.addUrlMappings("/jquery.min.js");
        srb.addUrlMappings("/pobstyle.css");
        srb.addUrlMappings("/sealsetup.exe");
        return srb;
    }

    /**
     * 添加印章管理程序Servlet（可选）
     *
     * @return
     */
    @Bean
    public ServletRegistrationBean zoomsealRegistrationBean() throws FileNotFoundException {
        com.zhuozhengsoft.pageoffice.poserver.AdminSeal adminSeal = new com.zhuozhengsoft.pageoffice.poserver.AdminSeal();
        adminSeal.setAdminPassword(poPassword);//设置印章管理员admin的登录密码（为了安全起见，强烈建议修改此密码）
        /**如果当前项目是打成jar或者war包运行，强烈建议将poseal.db文件的路径更换成某个固定的绝对路径下,不要放当前项目文件夹下,为了防止每次重新打包程序导致poseal.db被替换的问题。
         * 比如windows服务器下：D:/pageoffice/，linux服务器下:/root/pageoffice/
         */
        //设置印章数据库文件poseal.db存放的目录
        adminSeal.setSysPath(poSysPath);//poSysPath可以在application.properties这个文件中配置，也可以直设置文件夹路径，比如：poserver.setSysPath("D:/pageoffice");
        ServletRegistrationBean srb = new ServletRegistrationBean(adminSeal);
        srb.addUrlMappings("/adminseal.zz");
        srb.addUrlMappings("/sealimage.zz");
        srb.addUrlMappings("/loginseal.zz");
        return srb;
    }
}



