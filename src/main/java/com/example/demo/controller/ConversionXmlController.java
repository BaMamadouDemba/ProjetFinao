//AJOUT
package com.example.demo.controller;

import com.example.demo.service.ConvertionXmlService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ConversionXmlController {

    @Autowired
    private ConvertionXmlService convertionXmlService;

    @GetMapping("/export-xml")
    public ResponseEntity<Resource> exportXml() {
        String xmlData = convertionXmlService.convertDataToXml();

        ByteArrayResource resource = new ByteArrayResource(xmlData.getBytes());

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=data.xml")
                .contentType(MediaType.APPLICATION_XML)
                .body(resource);
    }
}
