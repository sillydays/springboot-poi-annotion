package com.boot.controller;

import com.boot.entity.User;
import com.boot.service.ExcelService;
import com.boot.utils.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.CollectionUtils;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

/**
 * @Desc 测试注解版导入导出接口
 * @Author lixu
 * @Date 2021/10/14 11:20
 */

@RestController
@RequestMapping("/excel")
@Slf4j
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    /**
     * 导入 Excel
     */
    @PostMapping("/import")
    public void importExcel(@RequestParam("file") MultipartFile file) {
        List<User> users = ExcelUtil.readExcelToData(file, User.class);
        if (CollectionUtils.isEmpty(users)) {
            return;
        }
        for (User e : users) {
            log.info("id: {}, name: {}, sex: {}", e.getId(), e.getUserName(), e.getSex());
        }
    }

    /**
     * 导出 Excel
     */
    @GetMapping("/export/{fileName}")
    public void exportExcel(@PathVariable("fileName") String fileName, HttpServletResponse response) {
        List<User> userList = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            User user = User.builder()
                    .userName("李旭" + i)
                    .id((long) (i + 1))
                    .sex(i % 2)
                    .birthday(LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")))
                    .build();
            userList.add(user);
        }

        // 导出
        response.reset();
        String name;
        try {
             name = URLEncoder.encode(fileName + ".xlsx", "UTF-8");
        } catch (Exception e) {
            throw new RuntimeException("Encode Change Failure！");
        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        response.addHeader("Content-Disposition", "attachment;filename=" + name);

        try (OutputStream outputStream = new BufferedOutputStream(response.getOutputStream())) {
            byte[] bytes = ExcelUtil.readDataToExcel(userList, User.class);
            response.addHeader("Content-Length", "" + bytes.length);
            outputStream.write(bytes);
            outputStream.flush();
        } catch (IOException e) {
            throw new RuntimeException(e.getMessage());
        }
    }


}
