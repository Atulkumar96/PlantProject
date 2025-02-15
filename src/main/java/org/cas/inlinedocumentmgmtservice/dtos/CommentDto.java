package org.cas.inlinedocumentmgmtservice.dtos;

import lombok.NoArgsConstructor;

@NoArgsConstructor
public class CommentDto {
    private String author;
    private String initials;
    private String comment;
    private String parentComment;

    public CommentDto(String author, String initials, String comment, String parentComment) {
        this.author = author;
        this.initials = initials;
        this.comment = comment;
        this.parentComment = parentComment;
    }

    // Getters and Setters

    public String getComment() {
        return comment;
    }

    public void setText(String comment) {
        this.comment = comment;
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public String getInitials() {
        return initials;
    }

    public void setInitials(String initials) {
        this.initials = initials;
    }

    public String getParentComment() {
        return parentComment;
    }

    public void setParentComment(String parentComment) {
        this.parentComment = parentComment;
    }

    // End of Getters and Setters
}
