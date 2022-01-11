package com.me.utils.excel;

import lombok.extern.slf4j.Slf4j;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Slf4j
public class ZipUtils {

    /**
     * 压缩单个excel文件的输出流 到zip输出流,注意zipOutputStream未关闭，需要交由调用者关闭之
     *
     * @param zipOutputStream   zip文件的输出流
     * @param excelOutputStream excel文件的输出流
     * @param filename     文件名可以带目录，例如 TestDir/test1.xlsx
     */
    public static void compressFileToZipStream(ZipOutputStream zipOutputStream,
                                                ByteArrayOutputStream excelOutputStream, String filename) {
        byte[] buf = new byte[1024];
        try {
            // Compress the files
            byte[] content = excelOutputStream.toByteArray();
            ByteArrayInputStream is = new ByteArrayInputStream(content);
            BufferedInputStream bis = new BufferedInputStream(is);
            // Add ZIP entry to output stream.
            zipOutputStream.putNextEntry(new ZipEntry(filename));
            // Transfer bytes from the file to the ZIP file
            int len;
            while ((len = bis.read(buf)) > 0) {
                zipOutputStream.write(buf, 0, len);
            }
            // Complete the entry
            //excelOutputStream.close();//关闭excel输出流
            zipOutputStream.closeEntry();
            bis.close();
            is.close();
            // Complete the ZIP file
        } catch (IOException e) {
            log.error("write zipOutputStream error", e);
        }
    }

}
