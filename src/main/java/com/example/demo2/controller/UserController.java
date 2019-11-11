package com.example.demo2.controller;

import com.example.demo2.service.UserService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * 作者: shiloh
 * 日期: 2019/11/8
 * 描述:
 */
@Controller
@RequestMapping("/user")
public class UserController {

    @RequestMapping("/index")
    public String index(){
        return "index";
    }

    @Autowired
    private UserService service;

    @RequestMapping("/batchSave")
    public String batchSave(HttpServletRequest request,
                          HttpServletResponse response,
                          @RequestParam(value = "excelFile", required = true) MultipartFile excelFile) {

        try {
            service.batchSaveEquipment(request,response ,excelFile);
            return "index";
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "index";
    }

}
