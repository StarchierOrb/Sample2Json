package me.starchier.json;

public class Question {
    private final String _id;
    private final String title;
    private final String comments;
    private final String typecode;
    private final String typename;
    private final Choose[] options;
    private final String examid;
    public Question(String _id, String title, String typecode, String typename, String comments, Choose[] options, String examid) {
        this._id = _id;
        this.title = title;
        this.comments = comments;
        this.typecode = typecode;
        this.typename = typename;
        this.options = options;
        this.examid = examid;
    }
}
