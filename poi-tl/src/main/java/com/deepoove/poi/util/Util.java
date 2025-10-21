package com.deepoove.poi.util;

import org.apache.commons.lang3.StringUtils;

import java.util.List;

public class Util {
    public String join(String delimiter, String defaultStr, List<Object> data) {
        if (delimiter == null) {
            delimiter = ",";
        }
        if (data == null) {
            return "";
        }
        if (defaultStr == null) {
            defaultStr = "——";
        }
        StringBuilder sb = new StringBuilder();
        for (Object s : data) {
            if (s == null) {
                continue;
            }
            if (StringUtils.isBlank(s.toString())) {
                continue;
            }
            // 默认的占位符好 不拼接
            if (defaultStr.equals(s.toString())) {
                continue;
            }
            sb.append(s).append(delimiter);
        }
        return sb.substring(0, sb.length() - delimiter.length());
    }

    public String join(String defaultResult, String delimiter, String defaultStr, List<Object> filter, List<Object> data) {
        if (delimiter == null) {
            delimiter = ",";
        }
        if (filter == null || data == null) {
            return "";
        }
        if (defaultStr == null) {
            defaultStr = "——";
        }
        if (defaultResult == null) {
            defaultResult = defaultStr;
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < filter.size(); i++) {
            Object s = filter.get(i);
            if (s == null) {
                continue;
            }
            if (StringUtils.isBlank(s.toString())) {
                continue;
            }
            if (defaultStr.equals(s.toString())) {
                continue;
            }
            sb.append(data.get(i) == null ? defaultResult : data.get(i)).append(delimiter);
        }
        return sb.substring(0, sb.length() - delimiter.length());
    }
}
