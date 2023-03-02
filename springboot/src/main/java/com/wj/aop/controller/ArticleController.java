package com.wj.aop.controller;

import com.wj.aop.pointcut.ViewCountAdd;
import com.wj.aop.pojo.Article;
import org.springframework.stereotype.Service;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * Created by IntelliJ IDEA
 *
 * @author wj
 * @date 2023/3/2 16:06
 */
@RestController
public class ArticleController {
    @ViewCountAdd
    @GetMapping("/article")
    public Article getArticle(){
        return new Article();
    }
}
