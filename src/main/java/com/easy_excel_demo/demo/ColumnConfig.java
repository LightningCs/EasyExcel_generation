package com.easy_excel_demo.demo;

import lombok.Data;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.stereotype.Component;

import java.util.List;

@Data
@Component
@ConfigurationProperties(prefix = "export")
public class ColumnConfig {
    private Integer pace = 0;
    private List<String> columns;
}
