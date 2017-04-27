package cn.itcast.common.excel.test.cn.itcast.common.excel.virtual.test;

import cn.itcast.common.excel.ExcelUtils;
import cn.itcast.common.excel.constants.ExcelType;
import com.alibaba.fastjson.JSONArray;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.*;

/**
 * Created by zhangtian on 2017/4/27.
 */
public class ExcelRowTest {

    private List<Map<Object,List<Object>>> datas = new ArrayList<Map<Object,List<Object>>>() ;

    @Before
    public void buildData() {
        Map<Object,List<Object>> op1 = new HashMap<Object,List<Object>>() ;

        List<Object> list1 = new ArrayList<Object>() ;
        QuestionOption option1 = new QuestionOption() ;
        option1.setId(UUID.randomUUID().toString());
        option1.setAnswerUsername("答题人");
        option1.setQuestionContent("十万个为什么？");
        option1.setQuestionType("题型类别");
        option1.setUserSchool("所在学校");

        QuestionAnswer answer1 = new QuestionAnswer() ;
        answer1.setId(UUID.randomUUID().toString());
        answer1.setAnswerUsername("张田");
        answer1.setQuestionContent("太阳为什么东升西落");
        answer1.setQuestionType("问答题");
        answer1.setUserSchool("景城学校");

        QuestionAnswer answer2 = new QuestionAnswer() ;
        answer2.setId(UUID.randomUUID().toString());
        answer2.setAnswerUsername("严加洋");
        answer2.setQuestionContent("月亮为什么晚上升起");
        answer2.setQuestionType("问答题");
        answer2.setUserSchool("星海中学");
        list1.add(answer1) ;
        list1.add(answer2) ;

        op1.put(option1, list1) ;
        datas.add(op1) ;

        Map<Object,List<Object>> op2 = new HashMap<Object, List<Object>>() ;
        List<Object> list2 = new ArrayList<Object>() ;

        QuestionOption option2 = new QuestionOption() ;
        option2.setId(UUID.randomUUID().toString());
        option2.setAnswerUsername("答题人");
        option2.setQuestionContent("唐诗宋词三百首？");
        option2.setQuestionType("题型类别");
        option2.setUserSchool("所在学校");

        QuestionAnswer answer3 = new QuestionAnswer() ;
        answer3.setId(UUID.randomUUID().toString());
        answer3.setAnswerUsername("于洋");
        answer3.setQuestionContent("床前明月光，疑似地上霜");
        answer3.setQuestionType("问答题");
        answer3.setUserSchool("金鸡湖学校");

        QuestionAnswer answer4 = new QuestionAnswer() ;
        answer4.setId(UUID.randomUUID().toString());
        answer4.setAnswerUsername("史可欣");
        answer4.setQuestionContent("大江东去浪淘尽");
        answer4.setQuestionType("问答题");
        answer4.setUserSchool("木渎实验中学");
        list2.add(answer3) ;
        list2.add(answer4) ;

        op2.put(option2, list2) ;
        datas.add(op2) ;

        System.out.println(JSONArray.toJSONString(datas));
    }

    @Test
    public void testVirtualRow() throws IOException {
        Workbook workbook = ExcelUtils.exportVirtualRowExcelData(datas, ExcelType.XLS, "demo") ;
        OutputStream out = new FileOutputStream(new File("C:\\Users\\zhangtian\\Desktop\\demo11.xlsx")) ;
        workbook.write(out);
        out.flush();
        out.close();
        workbook.close();
    }
}
