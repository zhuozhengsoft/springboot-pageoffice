package com.zhuozhengsoft.springbootpageoffice.controller;

import java.io.FileNotFoundException;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.web.servlet.ServletRegistrationBean;
import org.springframework.context.annotation.Bean;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.servlet.ModelAndView;

/**
 * @author Administrator
 */
@RestController
public class IndexController {
    @RequestMapping(value = "/index", method = RequestMethod.GET)
    public ModelAndView showIndex() {
        ModelAndView mv = new ModelAndView("index");
        return mv;
    }


}
