package com.shuwei.boot.config;

import com.shuwei.boot.bean.Pet;
import com.shuwei.boot.bean.User;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration //告诉SpringBoot这是一个配置类 == 配置文件
public class MyConfig {

    @Bean //给容器中添加组件。以方法名作为组件的id。返回类型就是组件类型。返回的值，就是组件在容器中的实例
    public User user01(){
        return new User("zhangsan",18);
    }

    @Bean
    public Pet tomcatPet(){
        return new Pet("tomcat",2);
    }
}