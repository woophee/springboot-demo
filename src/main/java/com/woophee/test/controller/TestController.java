package com.woophee.test.controller;

import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

@RestController
@RequestMapping(value="/xxx",method= RequestMethod.GET)
public class TestController {

    @RequestMapping(value="/hello",method= RequestMethod.GET)
    public String sayHello(){
        return "hello";
    }
}
