package com.converter.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.concurrent.ConcurrentHashMap;

public class ExcelFileDownConverter {
    private String inputFileLocation;
    private String outputLocation;
    private HSSFWorkbook workBookOut;
    private Map<Short, CellStyle> formatCellStyleMap = new ConcurrentHashMap<>();

    public static void main(String[] args) {
        String filename = "/home/venkatesan/ROOT_DIR/sample-xlsx-file-for-testing.xlsx";
        ExcelFileDownConverter converter = new ExcelFileDownConverter(filename);
        converter.convert();
    }

    public ExcelFileDownConverter(String inputFileLocation) {
        Objects.requireNonNull(inputFileLocation, "Input file location should not be null or empty");
        Path path = Paths.get(inputFileLocation);
        String filename = path.getFileName().toString();
        if (!filename.endsWith(".xlsx")) {
            throw new RuntimeException("Please provide input file with .xlsx format");
        }
        this.inputFileLocation = path.toString();
        outputLocation = path.getParent()
                .toString()
                .concat(File.separator)
                .concat(filename.replace(".xlsx", ".xls"));
    }

    void convert() {
        Optional<XSSFWorkbook> wbIn = createInputWorkBook();
        if (!wbIn.isPresent()) {
            System.out.println("No valid xlsh workbook found");
            return;
        }
        workBookOut = new HSSFWorkbook();
        copySheets(wbIn.get());
        createOutputFile();
    }

    Optional<XSSFWorkbook> createInputWorkBook() {
        try (InputStream in = new FileInputStream(inputFileLocation)) {
            final XSSFWorkbook workbook = new XSSFWorkbook(in);
            return Optional.of(workbook);
        } catch (Exception ex) {
            System.out.println("Error while reading file :" + inputFileLocation);
            ex.printStackTrace();
        }
        return Optional.empty();
    }

    void copySheets(XSSFWorkbook wbIn) {
        wbIn.sheetIterator().forEachRemaining((sIn) -> {
            Sheet sOut = workBookOut.createSheet(sIn.getSheetName());
            sIn.rowIterator().forEachRemaining(row -> copyRow(sOut, row));
        });
    }

    void copyRow(Sheet sOut, Row rowIn) {
        Row rowOut = sOut.createRow(rowIn.getRowNum());
        rowIn.cellIterator().forEachRemaining(cell -> copyCell(rowOut, cell));
    }

    void copyCell(Row rowOut, Cell cellIn) {
        Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellType());

        if (cellIn.getCellType() == CellType.BOOLEAN) {
            cellOut.setCellValue(cellIn.getBooleanCellValue());
        } else if (cellIn.getCellType() == CellType.ERROR) {
            cellOut.setCellValue(cellIn.getErrorCellValue());
        } else if (cellIn.getCellType() == CellType.FORMULA) {
            cellOut.setCellFormula(cellIn.getCellFormula());
        } else if (cellIn.getCellType() == CellType.NUMERIC) {
            cellOut.setCellValue(cellIn.getNumericCellValue());
        } else if (cellIn.getCellType() == CellType.STRING) {
            cellOut.setCellValue(cellIn.getStringCellValue());
        }

        Short formatterKey = cellIn.getCellStyle().getDataFormat();
        updateCellStyleFromCache(cellIn, cellOut, formatterKey);
    }

    private void updateCellStyleFromCache(Cell cellIn, Cell cellOut, Short formatterKey) {
        CellStyle styleOut = Optional.ofNullable(formatCellStyleMap.get(formatterKey))
                .orElseGet(() -> {
                    CellStyle newCellStyle = workBookOut.createCellStyle();
                    formatCellStyleMap.put(formatterKey, newCellStyle);
                    return newCellStyle;
                });
        styleOut.setDataFormat(formatterKey);
        cellOut.setCellStyle(styleOut);
        cellOut.setCellComment(cellIn.getCellComment());
    }

    private void createOutputFile() {
        try (Workbook outBook = workBookOut;
             OutputStream out = new BufferedOutputStream(new FileOutputStream(outputLocation))) {
            outBook.write(out);
            System.out.println("File created successfully on: " + outputLocation);
        } catch (IOException e) {
            System.out.println("Error while writing into outputfile: " + outputLocation);
            e.printStackTrace();
        }
    }
}
