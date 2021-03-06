package ru.varpa89.farm.service;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.springframework.stereotype.Service;
import ru.varpa89.farm.dto.DocumentHeader;
import ru.varpa89.farm.dto.DocumentTablePart;
import ru.varpa89.farm.dto.SingleDocument;

import java.math.BigDecimal;
import java.util.*;

@Slf4j
@Service
@RequiredArgsConstructor
public class SellService extends AbstractXlsService {
    private static final CellAddress INVOICE_NUMBER = new CellAddress("W27");
    private static final CellAddress INVOICE_DATE = new CellAddress("AC27");
    private static final CellAddress CLIENT_INFO = new CellAddress("H13");
    private static final CellAddress PRODUCTS = new CellAddress("A33");
    private static final CellAddress ORDER_NR = new CellAddress("BB24");
    private static final CellAddress ORDER_DATE = new CellAddress("BB25");
    private static final CellAddress GLN = new CellAddress("BB26");
    private static final CellAddress ADDRESS = new CellAddress("H22");

    private final ClientExtractor clientExtractor;

    public SingleDocument readFile(Workbook invoice) {
        final Sheet sheet = invoice.getSheetAt(0);
        log.info("Process sheet {}", sheet.getSheetName());

        final String invoiceNumber = getStringValue(sheet, INVOICE_NUMBER);
        final Date invoiceDate = getInvoiceDate(sheet, INVOICE_DATE);
        final String clientInfoValue = getStringValue(sheet, CLIENT_INFO);

        final ClientExtractor.ClientInfo clientInfo = clientExtractor.extractInfo(clientInfoValue);
        log.info("Address parsed: {}", clientInfo);

        final DocumentHeader documentHeader = DocumentHeader.builder()
                .clientName(clientInfo.getName())
                .kpp(clientInfo.getKpp())
                .inn(clientInfo.getInn())
                .addrName(getStringValue(sheet, ADDRESS))
                .numberTs(invoiceNumber)
                .date(getInvoiceFormattedDate(sheet, INVOICE_DATE))
                .typeRn("Продажа")
                .bonus(0)
                .promo(0)
                .firm("Мега Опт")
                .numberSf("")
                .orderNr(getStringValue(sheet, ORDER_NR))
                .orderDate(getStringValue(sheet, ORDER_DATE))
                .gln(getStringValue(sheet, GLN))
                .build();

        //firm - мега опт
        //NumberSF - пусто
        //OrderNR - форма
        //OrderDate - форма
        //GLN - форма
        //адрес - форма

        final List<DocumentTablePart> documentTableParts = extractDocumentTablePart(sheet);

        return new SingleDocument(invoiceNumber, documentHeader, documentTableParts, invoiceDate);
    }

    private List<DocumentTablePart> extractDocumentTablePart(Sheet sheet) {

        int rowIndex = PRODUCTS.getRow();
        List<DocumentTablePart> documentTableParts = new ArrayList<>();
        while (sheet.getRow(rowIndex).getCell(0).getCellType().equals(CellType.NUMERIC)) {

            Row row = sheet.getRow(rowIndex);

            final double lineNumber = row.getCell(0).getNumericCellValue();
            final double amount = row.getCell(42).getNumericCellValue();
            final double price = row.getCell(39).getNumericCellValue();
            final String unit = row.getCell(19).getStringCellValue();
            final String name = row.getCell(3).getStringCellValue();

            //nomenclature - код/артикул
            //plu - пустой
            //factor - сколько единиц в одной коробке - через артикул
            //quantity - столбец (10) количество разделить на factor
            //SummaNDS - 0
            // NDS - без ндс

            final String nomenclature = row.getCell(16).getStringCellValue();
            final Integer factor = nomenclatureToFactor.get(nomenclature);
            if (factor == null) {
                throw new RuntimeException("Неизвестная номенклатура " + nomenclature);
            }
            final int quantity = parseInteger(row.getCell(36).getNumericCellValue()) / factor;

            final DocumentTablePart documentTablePart = DocumentTablePart.builder()
                    .amount(parseBigDecimal(amount))
                    .price(parseBigDecimal(price))
                    .quantity(parseInteger(quantity))
                    .unit(unit)
                    .name(name)
                    .line(parseInteger(lineNumber))
                    .nomenclature(nomenclature)
                    .plu("")
                    .factor(factor)
                    .quantity(quantity)
                    .nds("Без НДС")
                    .ndsAmount(BigDecimal.ZERO)
                    .build();

            documentTableParts.add(documentTablePart);

            rowIndex++;
        }

        return documentTableParts;
    }
}
