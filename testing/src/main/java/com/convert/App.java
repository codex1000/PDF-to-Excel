package com.convert;

import java.io.BufferedWriter;
import java.io.FileWriter;
//import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;

import com.aspose.pdf.*;
import com.aspose.cells.*;
import com.aspose.cells.Cells;

public final class App {
    private App() {

    }

     // The path to the documents directory.
     private static Path _dataDir = Paths.get("C:\\Users\\Okafor Handel\\OneDrive\\Desktop\\Test\\");


    public static void main(String[] args)  throws Exception {

        ConvertPDFtoExcelAdvanced_SaveXLSX();
        ConvertExceltoJson();

    }

    public static void ConvertPDFtoExcelAdvanced_SaveXLSX() {
        // Load PDF document
        Document pdfDocument = new Document(_dataDir + ".\\sample.pdf");

        // Instantiate ExcelSave Option object
        ExcelSaveOptions excelSave = new ExcelSaveOptions();
        excelSave.setFormat(ExcelSaveOptions.ExcelFormat.XLSX);

        // Save the output in XLS format
        pdfDocument.save("PDFToXLS_out.xlsx", excelSave);
    }

    public static void ConvertExceltoJson() throws Exception{

        // load XLSX file with an instance of Workbook
        Workbook workbook = new Workbook(".\\PDFToXLS_out.xlsx");
        // access CellsCollection of the worksheet containing data to be converted
        Cells cells = workbook.getWorksheets().get(0).getCells();
        // create & set ExportRangeToJsonOptions for advanced options
        ExportRangeToJsonOptions exportOptions = new ExportRangeToJsonOptions();
        // create a range of cells containing data to be exported
        Range range = cells.createRange(0, 0, cells.getLastCell().getRow() + 1, cells.getLastCell().getColumn() + 1);
        // export range as JSON data
        String jsonData = JsonUtility.exportRangeToJson(range, exportOptions);
        // write data to disc in JSON format
        BufferedWriter writer = new BufferedWriter(new FileWriter("output.json"));
        writer.write(jsonData);
        writer.close();
    }
}
