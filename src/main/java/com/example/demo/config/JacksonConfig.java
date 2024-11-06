package com.example.demo.config;

import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.PropertyNamingStrategy;
import com.fasterxml.jackson.annotation.JsonInclude;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class JacksonConfig {

    @Bean
    public ObjectMapper objectMapper() {
        ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.setSerializationInclusion(JsonInclude.Include.NON_NULL);
        objectMapper.enable(SerializationFeature.FAIL_ON_EMPTY_BEANS);
        objectMapper.setPropertyNamingStrategy(PropertyNamingStrategy.SNAKE_CASE);
        objectMapper.enable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);
        return objectMapper;
    }
}