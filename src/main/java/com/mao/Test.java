package com.mao;

import freemarker.cache.ClassTemplateLoader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by mao on 2016/11/15.
 */

public class Test {
    public static void main(String[] args){
        Map<String, Object> dataMap = new HashMap<String, Object>();
        dataMap.put("name","王五");
        dataMap.put("sex","男");
        dataMap.put("phone","110");
        dataMap.put("address","安徽省合肥市");
        //这里map的key必须与模板中标记的名字一致才行，另外，这里生成的文件必须以“.doc结尾”，不能是“.docx”
        generateDocument(new File("D:/生成文件.doc"), "原始模板文件.ftl",dataMap);
    }

    /**
     * 根据模板和数据生成文档
     * @param outFile       目标文档文件
     * @param templateName  模板名称
     * @param dataMap       数据
     * @return
     */
    public static boolean generateDocument(File outFile,String templateName,Map<String, Object> dataMap){
        Configuration configuration=new Configuration(Configuration.VERSION_2_3_23);
        //配置模板文件加载的位置，因为放在resources/templates文件夹下
        configuration.setTemplateLoader(new ClassTemplateLoader(Test.class,"/templates") );
        configuration.setDefaultEncoding("UTF-8");
        Template template;
        try {
            //获取模板对象
            template = configuration.getTemplate(templateName,"UTF-8");
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        Writer out;
        FileOutputStream fos;
        try {
            fos = new FileOutputStream(outFile);
            OutputStreamWriter oWriter = new OutputStreamWriter(fos,"UTF-8");
            out = new BufferedWriter(oWriter);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
            return false;
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
            return false;
        }
        try {
            //填充数据
            template.process(dataMap, out);
            out.close();
        } catch (TemplateException e) {
            e.printStackTrace();
            return false;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
        return true;
    }
}
