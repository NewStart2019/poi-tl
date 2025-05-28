package com.deepoove.poi.tl.util;

import com.deepoove.poi.util.ExcelTemplateRenderer;
import org.junit.jupiter.api.Test;
import org.springframework.expression.ExpressionParser;
import org.springframework.expression.spel.SpelCompilerMode;
import org.springframework.expression.spel.SpelParserConfiguration;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelTemplateRendererTest {

    @Test
    public void testExcelRender() throws Exception {
        ExcelTemplateRenderer renderer = new ExcelTemplateRenderer();

        Map<String, Object> model = new HashMap<>();
        List<Map<String, Object>> students = new ArrayList<>();

        Map<String, Object> s1 = new HashMap<>();
        s1.put("name", "张三");
        s1.put("age", 20);
        s1.put("score", 80);
        s1.put("sex", "男");
        students.add(s1);

        Map<String, Object> s2 = new HashMap<>();
        s2.put("name", "李四");
        s2.put("age", 22);
        s2.put("score", 90);
        s2.put("sex", "女");
        students.add(s2);

        model.put("students", students);
        model.put("student", s1);

        Map<String, Object> s3 = new HashMap<>();
        s3.put("name", "李四");
        s3.put("age", 22);
        s3.put("score", 15);
        s3.put("sex", "男");
        students.add(s3);

        // 渲染并输出文件
        renderer.render("src/test/resources/util/template.xlsx", model);
        renderer.save("target/output.xlsx");
    }

    @Test
    public void testSpel() {
        ExpressionParser parser;

        ClassLoader loader = getClass().getClassLoader();
        SpelParserConfiguration config = new SpelParserConfiguration(SpelCompilerMode.IMMEDIATE, loader,
            true, true, 1024);
        parser = new SpelExpressionParser(config);
        class MyObject {
            public String getName() {
                return "John";
            }
        }

        StandardEvaluationContext context = new StandardEvaluationContext(new MyObject());
        Map<String, Object> vars = new HashMap<>();
        vars.put("var1", 123);
        context.setVariables(vars);

        context.setVariable("inputValue", "Hello World");

        String result = parser.parseExpression("#inputValue").getValue(context, String.class);
        System.out.println(result);

        String name = parser.parseExpression("name").getValue(context, String.class);
        System.out.println(name);

        String input = "这是一个模板示例：[$name] 和 [ $student.age ] 以及 [ $user['info'] ]";

        // 正则表达式：匹配 [ $xxx ] 格式
        // String regex = "\\[\\s*\\$(.*?)\\s*\\]";
        String regex = "(\\[\\s*\\$.*?\\s*\\])";
        Pattern pattern = Pattern.compile(regex);
        Matcher matcher = pattern.matcher(input);

        while (matcher.find()) {
            String variableName = matcher.group(1); // 获取 () 中的内容
            System.out.println("找到变量表达式: " + variableName);
        }

        List<Object> data = new ArrayList<>();
        data.add("王五");
        data.add(2);
        Object[] data2 = data.toArray();
        System.out.printf(String.format("%s-%s", data2));

    }
}
