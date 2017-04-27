package cn.itcast.common.excel.test.cn.itcast.common.excel.virtual.test;

import cn.itcast.common.excel.annotation.ExcelColumn;

/**
 * Created by zhangtian on 2017/4/26.
 */
public class QuestionAnswer {
    @ExcelColumn
    private String id ;
    @ExcelColumn
    private String questionType ;
    @ExcelColumn
    private String questionContent ;
    @ExcelColumn
    private String userSchool ;
    @ExcelColumn
    private String questionProjectName ;
    @ExcelColumn
    private String answerUsername ;

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getQuestionType() {
        return questionType;
    }

    public void setQuestionType(String questionType) {
        this.questionType = questionType;
    }

    public String getQuestionContent() {
        return questionContent;
    }

    public void setQuestionContent(String questionContent) {
        this.questionContent = questionContent;
    }

    public String getUserSchool() {
        return userSchool;
    }

    public void setUserSchool(String userSchool) {
        this.userSchool = userSchool;
    }

    public String getQuestionProjectName() {
        return questionProjectName;
    }

    public void setQuestionProjectName(String questionProjectName) {
        this.questionProjectName = questionProjectName;
    }

    public String getAnswerUsername() {
        return answerUsername;
    }

    public void setAnswerUsername(String answerUsername) {
        this.answerUsername = answerUsername;
    }
}
