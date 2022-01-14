# springboot-pageoffice
## 安装说明
### 一、环境要求
Intelij IDEA,jdk1.8
### 一、示例运行步骤
1. 在D盘的根目录下新建一个lic文件夹（这个文件夹将用来存放授权文件license.lic，如果需要测试印章功能，请拷贝当前项目根目录下podata文件夹中的poseal.db文件到lic文件夹下）。
2. 导入demo项目到idea。
3. 更新项目：Maven - Update Project...。
4. 运行项目：Run As - Java Application。
5. 启动浏览器访问：http://localhost:8080/index ，即可在线打开、编辑、保存word文件 。
### 二、序列号
   PageOfficeV5.0标准版试用序列号：I2BFU-MQ89-M4ZZ-ZWY7K           
   PageOfficeV5.0专业版试用序列号：DJMTF-HYK4-BDQ3-2MBUC
### 三、集成PageOffice到您的项目中的关键步骤
   1.在您项目的pom.xml中通过下面的代码引入PageOffice依赖。

```
<dependency>
     <groupId>com.zhuozhengsoft</groupId>   
  <artifactId>pageoffice</artifactId>   
  <version>5.3.0.3</version>
</dependency>
```

2. 在您项目的启动类Application类中配置如下代码。

```
@Bean
   public ServletRegistrationBean pageofficeRegistrationBean()  {
com.zhuozhengsoft.pageoffice.poserver.Server poserver = new com.zhuozhengsoft.pageoffice.poserver.Server();
/**如果当前项目是打成jar或者war包运行，强烈建议将license的路径更换成某个固定的绝对路径下，不要放当前项目文件夹下,为了防止每次重新发布项目导致license丢失问题。
  * 比如windows服务器下：D:/lic/，linux服务器下:/root/lic/
 */
//设置PageOffice注册成功后,license.lic文件存放的目录
poserver.setSysPath(poSysPath);
ServletRegistrationBean srb = new ServletRegistrationBean(poserver);
srb.addUrlMappings("/poserver.zz");
srb.addUrlMappings("/posetup.exe");
srb.addUrlMappings("/pageoffice.js");
srb.addUrlMappings("/jquery.min.js");
srb.addUrlMappings("/pobstyle.css");
srb.addUrlMappings("/sealsetup.exe");
return srb;//
}
```

3.新建Controller并调用PageOffice，例如：

```
public class PageOfficeController {
@RequestMapping(value = "/Word", method = RequestMethod.GET)
  public ModelAndView showWord(HttpServletRequest request) {
  PageOfficeCtrl poCtrl = new PageOfficeCtrl(request);
  poCtrl.setServerPage(request.getContextPath() + "/poserver.zz");//设置服务页面
  poCtrl.webOpen("/doc/test.doc", OpenModeType.docNormalEdit, "张三");
  request.setAttribute("pageoffice", poCtrl.getHtmlCode("PageOfficeCtrl1"));
  ModelAndView mv = new ModelAndView("Word");
   return mv;
  }
}
```

4.新建View页面（PageOfficeCtroller返回的View页面，用来存放PageOffice控件)，PageOffice在View页面输出的代码如下：

```
<div style="width: auto; height: 700px;" th:utext="${pageoffice}">
```

### 四、其他说明

> 如果您的项目要用到印章功能，请在您的web项目的启动类Application中加上印章功能相关配置，具体代码请参考当前项目pageoffice-demo启动类中的代码，此处不再赘述。

### 五、技术支持

电话：010-84721198

官网：[http://www.zhuozhengsoft.com](http://www.zhuozhengsoft.com)