package org.cas.inlinedocumentmgmtservice.dtos;

import com.fasterxml.jackson.annotation.JsonFormat;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

@NoArgsConstructor
public class CommentDto {
    private String author;
    private String initials;
    private String comment;

    @JsonFormat(shape = JsonFormat.Shape.STRING, pattern = "MM-dd-yyyy' 'HH:mm:ss")
    private LocalDateTime commentDateTime;
    private String parentComment;

    public CommentDto(String author, String initials, String comment, LocalDateTime commentDateTime,String parentComment) {
        this.author = author;
        this.initials = initials;
        this.comment = comment;

        this.commentDateTime = commentDateTime;
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

    public LocalDateTime getCommentDateTime() {
        return commentDateTime;
    }

    public void setCommentDateTime(LocalDateTime commentDateTime) {
        this.commentDateTime = commentDateTime;
    }

    public String getParentComment() {
        return parentComment;
    }

    public void setParentComment(String parentComment) {
        this.parentComment = parentComment;
    }

    // End of Getters and Setters
}
