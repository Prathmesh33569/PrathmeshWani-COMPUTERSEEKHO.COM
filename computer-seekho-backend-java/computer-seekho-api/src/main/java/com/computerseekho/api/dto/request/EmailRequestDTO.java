package com.computerseekho.api.dto.request;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class EmailRequestDTO {

    private String name;
    private String email;
    private String message;
}
