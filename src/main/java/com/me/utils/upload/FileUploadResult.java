package com.me.utils.upload;

import lombok.Data;

@Data
public class FileUploadResult {

    private String fileUrl;

    public static FileUploadResult buildResult(String url) {
        FileUploadResult result = new FileUploadResult();
        result.setFileUrl(url);
        return result;
    }
}
