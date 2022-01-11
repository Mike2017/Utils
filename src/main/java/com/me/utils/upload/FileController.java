package com.me.utils.upload;

import com.me.utils.common.CommonResult;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import java.util.Objects;

// https://github.com/HelloSummer5/FileUploadDemo
@Controller
@RequestMapping("/file")
@Slf4j
public class FileController {

    @Value("${file.path}")
    private String filePath;

    @RequestMapping(value = "/upload", method = RequestMethod.POST)
    @ResponseBody
    public CommonResult<FileUploadResult> upload(@RequestParam("file") MultipartFile file) {
        // 要上传的目标文件存放路径
        String fileName = FileUtils.upload(file, filePath, file.getOriginalFilename());
        if(Objects.isNull(fileName)) {
            return CommonResult.failed();
        }
        String fileUrl = "http://localhost:8005/file/" + fileName;
        return CommonResult.success(FileUploadResult.buildResult(fileUrl));
    }
}
