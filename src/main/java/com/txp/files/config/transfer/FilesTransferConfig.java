package com.txp.files.config.transfer;

import org.springframework.boot.web.servlet.MultipartConfigFactory;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.util.unit.DataSize;
import org.springframework.util.unit.DataUnit;

import javax.servlet.MultipartConfigElement;

/**
 *     BYTES("B", DataSize.ofBytes(1L)),
 *     KILOBYTES("KB", DataSize.ofKilobytes(1L)),
 *     MEGABYTES("MB", DataSize.ofMegabytes(1L)),
 *     GIGABYTES("GB", DataSize.ofGigabytes(1L)),
 *     TERABYTES("TB", DataSize.ofTerabytes(1L));
 */
@Configuration
public class FilesTransferConfig {
    @Bean
    public MultipartConfigElement multipartConfigElement() {
        MultipartConfigFactory factory = new MultipartConfigFactory();
        // 设置单个上传文件大小
        // 文件最大10M,DataUnit提供5中类型B,KB,MB,GB,TB
        factory.setMaxFileSize(DataSize.of(10, DataUnit.MEGABYTES));
        /// 设置总上传数据总大小10M
        factory.setMaxRequestSize(DataSize.of(50, DataUnit.MEGABYTES));
        return factory.createMultipartConfig();
    }
}
