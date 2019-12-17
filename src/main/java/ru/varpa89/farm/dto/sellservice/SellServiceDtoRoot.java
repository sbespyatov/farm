package ru.varpa89.farm.dto.sellservice;

import com.fasterxml.jackson.dataformat.xml.annotation.JacksonXmlProperty;
import com.fasterxml.jackson.dataformat.xml.annotation.JacksonXmlRootElement;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@JacksonXmlRootElement(localName = "Документы")
public class SellServiceDtoRoot {
    @JacksonXmlProperty(localName = "ОтдельныйДокумент")
    private List<SingleDocument> singleDocuments;
    @JacksonXmlProperty(isAttribute = true, localName = "l")
    private String documentStartDate;
    @JacksonXmlProperty(isAttribute = true, localName = "r")
    private String documentEndDate;
}
