package com.example.demo2.dao;

import com.example.demo2.entity.User;
import org.apache.ibatis.annotations.Mapper;
import org.springframework.stereotype.Repository;

import java.util.List;

/**
 * 作者: shiloh
 * 日期: 2019/11/8
 * 描述:
 */
@Repository
@Mapper
public interface UserDao {
    User get(int id);
    User delete(int id);
    void insert(User user);
    void batchInsert(List<User> list);
    void update(User user);
    List<User> list();
}
