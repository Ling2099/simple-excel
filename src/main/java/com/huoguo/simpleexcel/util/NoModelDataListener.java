package com.huoguo.simpleexcel.util;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * 模型数据监听
 * @author Lizhenghuang
 */
public class NoModelDataListener extends AnalysisEventListener<Map<Integer, String>> {

    List<Map<Integer, String>> list = new ArrayList<>();

    @Override
    public void invoke(Map<Integer, String> integerStringMap, AnalysisContext analysisContext) {
        list.add(integerStringMap);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
