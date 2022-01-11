package com.me.utils.upload;

import lombok.extern.slf4j.Slf4j;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;

/**
 * 文件上传工具包
 */
@Slf4j
public class FileUtils {

    /**
     *
     * @param file 文件
     * @param path 文件存放路径
     * @param fileName 源文件名
     * @return
     */
    public static String upload(MultipartFile file, String path, String fileName){
        try {
            // 生成新的文件名
            String newFileName = FileNameUtils.getFileName(fileName);
            File dest = new File(path + newFileName);
            //判断文件父目录是否存在
            if (!dest.getParentFile().exists()) {
                dest.getParentFile().mkdir();
            }
            //保存文件
            file.transferTo(dest);
            return newFileName;
        } catch (Exception e) {
            log.error("文件异常", e);
            return null;
        }
    }


}
